#!/usr/bin/env python3
# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
"""Experiment: a local model fixes layout symptoms through odf-tool.

Each round shows the model the ranked symptoms of the current .odt (rendered
with LibreOffice and compared with the reference PDF by layout_diff), the
involved paragraphs' current properties and effective indents, and the
property keys each symptom's hints allow. The model proposes one style
change, which is applied through `odf-tool protocol`, rendered and compared
again. The change is kept only if the layout improved and no new symptom
appeared; otherwise the next round starts again from the last kept document,
and the history tells the model what the reverted change caused.

    python3 experiments/fix_loop.py --reference ref.pdf --odt form.odt \
        --tool path/to/odf-tool --assets path/to/odf-rs \
        --model gemma-4-12b --work /tmp/loop

An .odt exported by Microsoft Word goes through prepare_word_odt.py first:
odf-tool refuses to edit it as exported.
"""
import argparse
import hashlib
import json
import pathlib
import shutil
import subprocess
import sys
import time
import urllib.request

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
import layout_diff  # noqa: E402

LETTER_SPACING = "style:text-properties/fo:letter-spacing"
TEXT_INDENT = "style:paragraph-properties/fo:text-indent"
MARGIN_LEFT = "style:paragraph-properties/fo:margin-left"
ALLOWED_KEYS = [LETTER_SPACING, TEXT_INDENT, MARGIN_LEFT]
# The keys each hint allows. A key the hint does not explain tends to remove
# one symptom by causing another (a negative margin for text that wraps
# pushes it out of its cell). fo:font-size is not offered: on the paragraph
# style it loses to the spans' own sizes, and runs of spaces take the Asian
# size, so it changed nothing on the forms tried.
HINT_KEYS = {
    "wide-gap": [LETTER_SPACING],
    "wider-text": [LETTER_SPACING],
    "narrower-text": [LETTER_SPACING],
    "wrapped-row": [LETTER_SPACING],
    "shifted-start": [TEXT_INDENT, MARGIN_LEFT],
}
CONTEXT_KEYS = ALLOWED_KEYS + [
    "style:text-properties/fo:font-size",
    "style:paragraph-properties/fo:margin-right",
    "style:paragraph-properties/fo:text-align",
    "style:paragraph-properties/fo:line-height",
]
# Symptoms shown to the model per round; the verdict uses all of them.
PROMPT_SYMPTOMS = 5
REPORT_LIMIT = 1000
# One row pushed onto another page weighs as much as this many points of shift.
ROW_ON_ANOTHER_PAGE_PT = 50.0
# A change is kept only if it lowers the score by at least this much.
MIN_IMPROVEMENT_PT = 0.5

SYSTEM = """You fix layout differences in an OpenDocument text document.
A tool rendered the document with LibreOffice (candidate) and compared it with the intended layout (reference).
You get the ranked symptoms. Each has an id, the paragraphs involved (address) with their current properties,
and allowedPropertyKeys: the only keys that may fix it.
Rules:
- Propose exactly one change: one symptom (prefer the first), one of its addresses, one of its allowedPropertyKeys.
  Each change is rendered and checked before the next round.
- Values carry units (cm, mm, in or pt). Prefer the smallest change that removes the symptom.
- Hints: "wide-gap" = a run of spaces no longer fits on the line; "wider-text" = the same text renders wider
  (on a right-aligned line its start moves left instead); "narrower-text" = narrower;
  "wrapped-row" = the end of the row above wrapped onto a line of its own ("extraRows").
  For these, tightening fo:letter-spacing a little (for example -0.01cm to -0.03cm) usually removes the extra line;
  for "narrower-text", loosen it instead.
- "shifted-start" = the text starts startShiftPt (in a break) or shiftPt (kind "shifted-line-start") points to the
  right (positive) or left (negative) of where it should. "indentsPt" is the marginLeft, marginRight and textIndent
  LibreOffice lays the paragraph out with, in points (a word instead of a number: it could not be determined).
  Set fo:text-indent (first line only) or fo:margin-left (every line) to the current value minus the shift.
- "history" lists earlier rounds. KEPT changes are in the document. REVERTED changes made the layout worse or caused
  "newSymptoms" and were undone: do not propose them again; try another value, key or symptom.
  "NO_EFFECT" means the key did not move the text at all: that key of that paragraph is not offered again.
- If no change is likely to help, return an empty "changes" list."""


def schema() -> dict:
    return {
        "type": "object",
        "properties": {
            "changes": {
                "type": "array",
                "maxItems": 1,
                "items": {
                    "type": "object",
                    "properties": {
                        "reason": {"type": "string"},
                        "symptomId": {"type": "integer"},
                        "address": {"type": "string"},
                        "propertyKey": {"enum": ALLOWED_KEYS},
                        "propertyValue": {"type": "string", "pattern": "^-?[0-9]+(\\.[0-9]+)?(cm|mm|in|pt)$"},
                    },
                    "required": ["reason", "symptomId", "address", "propertyKey", "propertyValue"],
                },
            }
        },
        "required": ["changes"],
    }


def style_operation(address: str, expected_style_name: str, assignments: dict[str, str]) -> dict:
    """One `set_style_properties` change carrying every property for one paragraph."""
    return {
        "name": "set_style_properties",
        "operationVersion": 8,
        "target": {"documentFamily": "text", "anchorType": "paragraph-path", "anchor": address},
        "arguments": {
            "styleProperties": [{"propertyKey": key, "propertyValue": value} for key, value in assignments.items()]
        },
        "preconditions": {"expectedStyleName": expected_style_name},
    }


def hints(symptom: dict) -> list[str]:
    found = [item["hint"] for item in symptom.get("breaks", []) + symptom.get("breaksInRowAbove", [])]
    if "hint" in symptom:
        found.append(symptom["hint"])
    return found


def allowed_keys(symptom: dict) -> list[str]:
    allowed = {key for hint in hints(symptom) for key in HINT_KEYS.get(hint, [])}
    return [key for key in ALLOWED_KEYS if key in allowed]


def fixable(report: dict, tried: set[tuple] | frozenset = frozenset()) -> list[dict]:
    """Symptoms naming a node and hinting at a cause, in report order, while
    some allowed key of some node has not turned out to have no effect."""
    return [
        symptom
        for symptom in report["symptoms"]
        if any((node["address"], key) not in tried for node in symptom.get("nodes", []) for key in allowed_keys(symptom))
    ]


def impact(symptom: dict) -> float:
    return abs(symptom["shiftPt"] if "shiftPt" in symptom else symptom["heightChangePt"])


def rows_on_another_page(report: dict) -> int:
    return sum(item["rowCount"] for item in report["rowsOnAnotherPage"])


def score(report: dict) -> float:
    return ROW_ON_ANOTHER_PAGE_PT * rows_on_another_page(report) + sum(impact(s) for s in report["symptoms"])


def layout_state(report: dict) -> tuple:
    """Everything the verdict looks at, to 0.1 pt."""
    return rows_on_another_page(report), [
        (s["kind"], s["page"], round(s["y"], 1), round(impact(s), 1)) for s in report["symptoms"]
    ]


def symptom_key(symptom: dict) -> tuple:
    return (symptom["kind"], symptom["page"], symptom.get("text") or symptom.get("row"))


def direction(symptom: dict) -> str:
    return "sideways" if symptom["kind"] == "shifted-line-start" else "vertical"


def is_new(symptom: dict, earlier: list[dict]) -> bool:
    """No earlier symptom moving text in the same direction names one of its
    nodes (or, without nodes, has its kind and text). A partial fix often
    turns a vertical shift into a taller block at the same paragraph; that
    is the same problem, while text of that paragraph moving sideways is not."""
    addresses = {node["address"] for node in symptom.get("nodes", [])}
    return not any(
        direction(other) == direction(symptom)
        and (addresses & {node["address"] for node in other.get("nodes", [])} or symptom_key(other) == symptom_key(symptom))
        for other in earlier
    )


def brief(symptom: dict) -> dict:
    return {
        "kind": symptom["kind"],
        "page": symptom["page"],
        "text": symptom.get("text") or symptom.get("row"),
        "pt": round(impact(symptom), 1),
    }


def judge(before: dict, after: dict) -> dict:
    """Whether the document behind `after` should replace the one behind `before`."""
    new = [brief(symptom) for symptom in after["symptoms"] if is_new(symptom, before["symptoms"])]
    if rows_on_another_page(after) > rows_on_another_page(before):
        reason = "MORE_ROWS_ON_ANOTHER_PAGE"
    elif new:
        reason = "NEW_SYMPTOM"
    elif layout_state(after) == layout_state(before):
        reason = "NO_EFFECT"
    elif score(after) > score(before) - MIN_IMPROVEMENT_PT:
        reason = "NO_IMPROVEMENT"
    else:
        reason = None
    return {"kept": reason is None, "reason": reason, "newSymptoms": new}


def check_change(change: dict, shown: list[dict], tried: set[tuple]) -> str | None:
    """Why the change cannot be applied, or None."""
    if not 1 <= change["symptomId"] <= len(shown):
        return "REJECTED_UNKNOWN_SYMPTOM"
    symptom = shown[change["symptomId"] - 1]
    if change["address"] not in {node["address"] for node in symptom["nodes"]}:
        return "REJECTED_ADDRESS_NOT_OFFERED"
    if change["propertyKey"] not in allowed_keys(symptom):
        return "REJECTED_KEY_NOT_ALLOWED"
    if (change["address"], change["propertyKey"]) in tried:
        return "REJECTED_NO_EFFECT"
    if (change["address"], change["propertyKey"], change["propertyValue"]) in tried:
        return "REJECTED_ALREADY_TRIED"
    return None


class Harness:
    def __init__(self, arguments: argparse.Namespace, work: pathlib.Path) -> None:
        self.arguments = arguments
        self.work = work

    def tool(self, request: dict) -> dict:
        completed = subprocess.run(
            [self.arguments.tool, "protocol", "--workspace", str(self.work), "--assets", self.arguments.assets],
            input=json.dumps(request),
            capture_output=True,
            text=True,
        )
        return json.loads(completed.stdout)

    def sha(self, name: str) -> str:
        return hashlib.sha256((self.work / name).read_bytes()).hexdigest()

    def project(self, name: str, **options: object) -> dict:
        """The `project` response, with the nodes of every window merged."""
        nodes: list[dict] = []
        offset = 0
        while True:
            response = self.tool(
                {
                    "protocolVersion": 1,
                    "requestId": "project",
                    "action": "project",
                    "source": name,
                    "expectedSourceSha256": self.sha(name),
                    "options": {**options, "offset": offset},
                }
            )
            if response["status"] != "success":
                raise RuntimeError(f"project {name}: {response['error']['code']}")
            nodes.extend(response["result"]["nodes"])
            offset = response["result"]["window"]["nextOffset"]
            if offset is None:
                response["result"]["nodes"] = nodes
                return response

    def apply(self, source: str, output: str, operations: list[dict]) -> str | None:
        """Plans and commits the operations as one request; returns the error code on failure."""
        base = {
            "protocolVersion": 1,
            "requestId": "apply",
            "action": "apply",
            "source": source,
            "expectedSourceSha256": self.sha(source),
            "output": output,
            "operations": operations,
        }
        plan = self.tool(dict(base, options={"mode": "plan", "validationProfile": "extended-odf"}))
        if plan["status"] != "success":
            errors = [item for item in plan.get("diagnostics", []) if item.get("severity") == "error"]
            if errors:
                first = errors[0]
                print(f"  {len(errors)} validation errors, first in {first.get('part')}: {first.get('subject')}", flush=True)
            return plan["error"]["code"]
        commit = self.tool(
            dict(
                base,
                options={"mode": "commit", "validationProfile": "extended-odf"},
                expectedPlanSha256=plan["result"]["planSha256"],
            )
        )
        return None if commit["status"] == "success" else commit["error"]["code"]

    def render(self, name: str) -> pathlib.Path:
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(self.work), str(self.work / name)],
            capture_output=True,
            timeout=180,
            check=True,
        )
        return self.work / (pathlib.Path(name).stem + ".pdf")

    def ask(self, content: dict) -> tuple[dict, float]:
        body = {
            "model": self.arguments.model,
            "temperature": 0,
            "max_tokens": self.arguments.max_tokens,
            # Reasoning models spend the whole budget thinking unless told not to.
            "chat_template_kwargs": {"enable_thinking": self.arguments.think},
            "messages": [
                {"role": "system", "content": SYSTEM},
                {"role": "user", "content": json.dumps(content, ensure_ascii=False)},
            ],
            "response_format": {"type": "json_schema", "json_schema": {"name": "changes", "schema": schema()}},
        }
        request = urllib.request.Request(
            self.arguments.endpoint, data=json.dumps(body).encode(), headers={"Content-Type": "application/json"}
        )
        started = time.monotonic()
        with urllib.request.urlopen(request, timeout=900) as response:
            data = json.load(response)
        choice = data["choices"][0]
        text = choice["message"].get("content") or ""
        reasoning = choice["message"].get("reasoning_content") or ""
        print(
            f"  finish={choice.get('finish_reason')} completionTokens={data['usage']['completion_tokens']} "
            f"reasoningChars={len(reasoning)}",
            flush=True,
        )
        try:
            answer = json.loads(text)
        except json.JSONDecodeError:
            print("  no parsable answer" + (" (the reasoning used up max_tokens)" if reasoning else ""), flush=True)
            answer = {"changes": []}
        return answer, time.monotonic() - started


def indents_pt(node: dict) -> dict:
    """`effectiveIndents` in points; an unresolved indent gives its reason instead."""
    return {
        name: indent["pt"] if indent["pt"] is not None else indent.get("unresolved")
        for name, indent in (node.get("effectiveIndents") or {}).items()
    }


def payload(shown: list[dict], report: dict, context: dict, history: list[dict]) -> dict:
    """`context` is a projection narrowed to the shown symptoms' addresses and CONTEXT_KEYS."""
    by_address = {node["address"]: node for node in context["result"]["nodes"]}
    symptoms = []
    for number, symptom in enumerate(shown, start=1):
        entry = {"id": number}
        for key in ("kind", "page", "shiftPt", "heightChangePt", "lineChange", "lineCount", "rowAbove", "text"):
            if symptom.get(key) is not None:
                entry[key] = symptom[key]
        breaks = symptom.get("breaks") or symptom.get("breaksInRowAbove") or []
        if breaks:
            entry["breaks"] = [
                {key: item[key] for key in ("textBeforeBreak", "textAfterBreak", "hint", "startShiftPt", "widthRatio", "gapPt")}
                for item in breaks
            ]
        for key in ("extraRows", "hint"):
            if key in symptom:
                entry[key] = symptom[key]
        entry["nodes"] = [
            {
                "address": node["address"],
                "excerpt": by_address[node["address"]]["excerpt"],
                "properties": {
                    item["propertyKey"]: item["value"] for item in by_address[node["address"]]["computedProperties"]
                },
                "indentsPt": indents_pt(by_address[node["address"]]),
            }
            for node in symptom["nodes"]
        ]
        entry["allowedPropertyKeys"] = allowed_keys(symptom)
        symptoms.append(entry)
    return {
        "goal": "remove the symptoms; no row may move to another page",
        "rowsOnAnotherPage": report["rowsOnAnotherPage"],
        "symptoms": symptoms,
        "history": history,
    }


def summary(report: dict) -> dict:
    return {
        "score": round(score(report), 1),
        "rowsOnAnotherPage": rows_on_another_page(report),
        "symptoms": [(s["kind"], s["page"], round(s["y"]), round(impact(s), 1)) for s in report["symptoms"][:8]],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--reference", required=True, help="PDF with the intended layout")
    parser.add_argument("--odt", required=True, help="document to fix (it is copied, not modified)")
    parser.add_argument("--tool", required=True, help="odf-tool binary")
    parser.add_argument("--assets", required=True, help="odf-rs checkout holding the protocol assets")
    parser.add_argument("--model", required=True)
    parser.add_argument("--endpoint", default="http://localhost:8418/v1/chat/completions")
    parser.add_argument("--think", action="store_true", help="let a reasoning model think before answering")
    parser.add_argument("--max-tokens", type=int, default=12000)
    parser.add_argument("--work", required=True, help="scratch directory, emptied first")
    parser.add_argument("--rounds", type=int, default=8, help="most changes proposed")
    arguments = parser.parse_args()
    work = pathlib.Path(arguments.work).resolve()
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    harness = Harness(arguments, work)
    reference = layout_diff.pdf_layout(arguments.reference)

    def evaluate(name: str) -> dict:
        projection_path = work / f"{pathlib.Path(name).stem}.project.json"
        projection_path.write_text(json.dumps(harness.project(name), ensure_ascii=False))
        nodes = layout_diff.load_projection_nodes([str(projection_path)])
        return layout_diff.compare(reference, layout_diff.pdf_layout(str(harness.render(name))), nodes, limit=REPORT_LIMIT)

    current = "round0.odt"
    shutil.copyfile(arguments.odt, work / current)
    report = evaluate(current)
    print(f"round 0: {json.dumps(summary(report), ensure_ascii=False)}", flush=True)
    history: list[dict] = []
    tried: set[tuple] = set()
    log: list[dict] = [{"round": 0, "state": summary(report)}]
    for number in range(1, arguments.rounds + 1):
        shown = fixable(report, tried)[:PROMPT_SYMPTOMS]
        if not shown:
            print("no symptom left that names a node and a cause", flush=True)
            break
        offered = sorted({node["address"] for symptom in shown for node in symptom["nodes"]})
        content = payload(shown, report, harness.project(current, addresses=offered, propertyKeys=CONTEXT_KEYS), history)
        answer, seconds = harness.ask(content)
        entry = {"round": number, "promptBytes": len(json.dumps(content, ensure_ascii=False)), "seconds": round(seconds, 1), "answer": answer}
        log.append(entry)
        changes = answer.get("changes", [])
        if not changes:
            print(f"round {number}: the model proposed no change", flush=True)
            break
        change = changes[0]
        record = {"round": number, **{key: change[key] for key in ("symptomId", "address", "propertyKey", "propertyValue")}}
        outcome = check_change(change, shown, tried)
        if outcome is None:
            tried.add((change["address"], change["propertyKey"], change["propertyValue"]))
            symptom = shown[change["symptomId"] - 1]
            record["symptom"] = brief(symptom)
            style_name = next(node["styleName"] for node in symptom["nodes"] if node["address"] == change["address"])
            attempt = f"round{number}.odt"
            operation = style_operation(change["address"], style_name, {change["propertyKey"]: change["propertyValue"]})
            outcome = harness.apply(current, attempt, [operation])
            if outcome is None:
                after = evaluate(attempt)
                verdict = judge(report, after)
                record.update(scoreBefore=round(score(report), 1), scoreAfter=round(score(after), 1))
                if verdict["kept"]:
                    outcome = "KEPT"
                    current, report = attempt, after
                else:
                    outcome = "REVERTED"
                    record["why"] = verdict["reason"]
                    if verdict["reason"] == "NO_EFFECT":
                        tried.add((change["address"], change["propertyKey"]))
                    if verdict["newSymptoms"]:
                        record["newSymptoms"] = verdict["newSymptoms"][:3]
        record["outcome"] = outcome
        history.append(record)
        entry["record"] = record
        entry["state"] = summary(report)
        print(
            f"round {number}: {change['address']} {change['propertyKey']}={change['propertyValue']} -> {outcome}"
            f"{' ' + record['why'] if 'why' in record else ''} "
            f"(score {record.get('scoreBefore')} -> {record.get('scoreAfter')}; model {seconds:.1f}s, "
            f"prompt {entry['promptBytes']} bytes; {change['reason'][:80]})",
            flush=True,
        )
    shutil.copyfile(work / current, work / "final.odt")
    print(f"final ({current}): {json.dumps(summary(report), ensure_ascii=False)}", flush=True)
    (work / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
