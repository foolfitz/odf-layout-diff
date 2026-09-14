#!/usr/bin/env python3
"""Experiment: a local model fixes layout symptoms through odf-tool.

Each iteration renders the current .odt with LibreOffice, compares it with
the reference PDF (layout_diff), shows the model the ranked symptoms plus the
involved paragraphs' current properties and effective indents, applies the
style changes it asks for through `odf-tool protocol`, and repeats. Changes
naming an address the model was not shown are rejected.

    python3 experiments/fix_loop.py --reference ref.pdf --odt form.odt \
        --tool path/to/odf-tool --assets path/to/odf-rs \
        --model gemma-4-12b --no-think --work /tmp/loop
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

ALLOWED_KEYS = [
    "style:text-properties/fo:letter-spacing",
    "style:text-properties/fo:font-size",
    "style:paragraph-properties/fo:text-indent",
    "style:paragraph-properties/fo:margin-left",
    "style:paragraph-properties/fo:margin-right",
]
CONTEXT_KEYS = ALLOWED_KEYS + [
    "style:paragraph-properties/fo:text-align",
    "style:paragraph-properties/fo:line-height",
]

SYSTEM = """You fix layout differences in an OpenDocument text document.
A tool rendered the document with LibreOffice (candidate) and compared it with the intended layout (reference).
You get the ranked symptoms. Each names the paragraphs involved (address, styleName) and their current properties.
Rules:
- Propose at most one style change per symptom, only for listed addresses, only with the allowed property keys.
- Values carry units (cm, mm, in, pt or %). Prefer the smallest change that removes the symptom.
- Hints: "wide-gap" = a run of spaces no longer fits on the line; "wider-text" = the same text renders wider.
  For both, tightening fo:letter-spacing a little (for example -0.01cm to -0.03cm) usually removes the extra line.
- "shifted-start" = the text starts startShiftPt points to the right (positive) or left (negative) of where it should.
  Move it back with fo:text-indent or fo:margin-left. "indentsPt" is the marginLeft, marginRight and textIndent
  LibreOffice lays the paragraph out with, in points (a word instead of a number: it could not be determined).
  Set an absolute value computed from it.
- A symptom without a hint or without an address: skip it.
- "history" lists your earlier changes and the symptoms that remained afterwards; use it to correct values."""


def schema() -> dict:
    return {
        "type": "object",
        "properties": {
            "changes": {
                "type": "array",
                "maxItems": 5,
                "items": {
                    "type": "object",
                    "properties": {
                        "reason": {"type": "string"},
                        "address": {"type": "string"},
                        "expectedStyleName": {"type": "string"},
                        "propertyKey": {"enum": ALLOWED_KEYS},
                        "propertyValue": {"type": "string", "pattern": "^-?[0-9]+(\\.[0-9]+)?(cm|mm|in|pt|%)$"},
                    },
                    "required": ["reason", "address", "expectedStyleName", "propertyKey", "propertyValue"],
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
            # Reasoning models spend most of the budget thinking before the answer.
            "max_tokens": 12000,
            "chat_template_kwargs": {"enable_thinking": not self.arguments.no_think},
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
            answer = {"changes": []}
        return answer, time.monotonic() - started


def indents_pt(node: dict) -> dict:
    """`effectiveIndents` in points; an unresolved indent gives its reason instead."""
    return {
        name: indent["pt"] if indent["pt"] is not None else indent.get("unresolved")
        for name, indent in (node.get("effectiveIndents") or {}).items()
    }


def payload(report: dict, context: dict, history: list[dict]) -> dict:
    """`context` is a projection narrowed to the symptoms' addresses and CONTEXT_KEYS."""
    by_address = {node["address"]: node for node in context["result"]["nodes"]}
    symptoms = []
    for symptom in report["symptoms"]:
        nodes = []
        for node in symptom.get("nodes", []):
            projected = by_address[node["address"]]
            nodes.append(
                {
                    **node,
                    "excerpt": projected["excerpt"],
                    "properties": {item["propertyKey"]: item["value"] for item in projected["computedProperties"]},
                    "indentsPt": indents_pt(projected),
                }
            )
        entry = {
            key: symptom[key]
            for key in ("kind", "page", "shiftPt", "heightChangePt", "lineChange", "rowAbove", "text")
            if key in symptom and symptom[key] is not None
        }
        breaks = symptom.get("breaks") or symptom.get("breaksInRowAbove") or []
        entry["breaks"] = [
            {key: item[key] for key in ("textBeforeBreak", "textAfterBreak", "hint", "startShiftPt", "widthRatio", "gapPt")}
            for item in breaks
        ]
        entry["nodes"] = nodes
        symptoms.append(entry)
    return {
        "goal": "remove the symptoms; no row may move to another page",
        "rowsOnAnotherPage": report["rowsOnAnotherPage"],
        "symptoms": symptoms,
        "allowedPropertyKeys": ALLOWED_KEYS,
        "history": history,
    }


def summary(report: dict) -> dict:
    return {
        "rowsOnAnotherPage": sum(item["rowCount"] for item in report["rowsOnAnotherPage"]),
        "symptoms": [
            (s["kind"], s["page"], round(s["y"]), s.get("shiftPt", s.get("heightChangePt"))) for s in report["symptoms"]
        ],
    }


def apply_changes(harness: Harness, current: str, target: str, changes: list[dict], offered: set[str]) -> list[str]:
    """Applies the changes to `current`, writing `target`; returns one outcome per change.

    Changes for the same paragraph and expected style become one operation, and
    all operations go in one request. That request fails as a whole, so on
    failure each operation is retried alone: one bad change does not block the
    others, and each outcome names its own error.
    """
    groups: dict[tuple[str, str], dict[str, str]] = {}
    keys: list[tuple[str, str] | None] = []
    for change in changes:
        if change["address"] not in offered:
            keys.append(None)
            continue
        key = (change["address"], change["expectedStyleName"])
        groups.setdefault(key, {})[change["propertyKey"]] = change["propertyValue"]
        keys.append(key)
    results: dict[tuple[str, str], str] = {}
    if groups and harness.apply(current, target, [style_operation(*key, a) for key, a in groups.items()]) is None:
        results = dict.fromkeys(groups, "APPLIED")
    else:
        source = current
        for index, (key, assignments) in enumerate(groups.items()):
            output = f"{pathlib.Path(target).stem}_{index}.odt"
            results[key] = harness.apply(source, output, [style_operation(*key, assignments)]) or "APPLIED"
            if results[key] == "APPLIED":
                source = output
        shutil.copyfile(harness.work / source, harness.work / target)
    return [results[key] if key else "REJECTED_ADDRESS_NOT_OFFERED" for key in keys]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--reference", required=True, help="PDF with the intended layout")
    parser.add_argument("--odt", required=True, help="document to fix (it is copied, not modified)")
    parser.add_argument("--tool", required=True, help="odf-tool binary")
    parser.add_argument("--assets", required=True, help="odf-rs checkout holding the protocol assets")
    parser.add_argument("--model", required=True)
    parser.add_argument("--endpoint", default="http://localhost:8418/v1/chat/completions")
    parser.add_argument("--no-think", action="store_true", help="ask the chat template to skip reasoning")
    parser.add_argument("--work", required=True, help="scratch directory, emptied first")
    parser.add_argument("--iterations", type=int, default=3)
    arguments = parser.parse_args()
    work = pathlib.Path(arguments.work).resolve()
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    shutil.copyfile(arguments.odt, work / "iter0.odt")
    harness = Harness(arguments, work)
    reference = layout_diff.pdf_layout(arguments.reference)
    history: list[dict] = []
    log = []
    for iteration in range(arguments.iterations + 1):
        current = f"iter{iteration}.odt"
        projection_path = work / f"iter{iteration}.project.json"
        projection_path.write_text(json.dumps(harness.project(current), ensure_ascii=False))
        nodes = layout_diff.load_projection_nodes([str(projection_path)])
        report = layout_diff.compare(reference, layout_diff.pdf_layout(str(harness.render(current))), nodes, limit=5)
        state = summary(report)
        print(f"iteration {iteration}: {json.dumps(state, ensure_ascii=False)}", flush=True)
        for item in history:
            if item["iteration"] == iteration - 1:
                item["afterwards"] = state
        entry = {"iteration": iteration, "state": state}
        log.append(entry)
        if (not report["rowsOnAnotherPage"] and not report["symptoms"]) or iteration == arguments.iterations:
            break
        offered = {node["address"] for symptom in report["symptoms"] for node in symptom.get("nodes", [])}
        context = (
            harness.project(current, addresses=sorted(offered), propertyKeys=CONTEXT_KEYS)
            if offered
            else {"result": {"nodes": []}}
        )
        content = payload(report, context, history)
        answer, seconds = harness.ask(content)
        entry.update({"promptBytes": len(json.dumps(content, ensure_ascii=False)), "seconds": round(seconds, 1), "answer": answer})
        changes = answer.get("changes", [])
        outcomes = apply_changes(harness, current, f"iter{iteration + 1}.odt", changes, offered)
        for change, outcome in zip(changes, outcomes):
            print(
                f"  {change['address']} {change['propertyKey']}={change['propertyValue']} -> {outcome} "
                f"({change['reason'][:90]})",
                flush=True,
            )
            history.append(
                {"iteration": iteration, **{k: change[k] for k in ("address", "propertyKey", "propertyValue")}, "outcome": outcome}
            )
        print(f"  model {seconds:.1f}s, prompt {entry['promptBytes']} bytes", flush=True)
    (work / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
