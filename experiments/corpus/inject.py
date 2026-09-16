# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
import re
BLOCK = re.compile(r'<config:config-item-set config:name="ooo:configuration-settings">.*</config:config-item-set>(?=\s*</office:settings>)', re.S)
def with_items(xml, items):
    """items: {name: value} (boolean) or {name: (type, value)}; handles self-closing roots."""
    if not re.search(r'<office:settings\b', xml):
        xml, n = re.subn(r'<office:document-settings\b([^>]*?)\s*/>', r'<office:document-settings\1><office:settings></office:settings></office:document-settings>', xml, count=1)
        if n != 1: raise RuntimeError('settings root not found')
    xml = xml.replace('<office:settings/>', '<office:settings></office:settings>')
    root = re.search(r'<office:document-settings\b[^>]*>', xml)
    if 'xmlns:ooo=' not in root.group(0):
        # config-item-set names are QNames: without this declaration LibreOffice ignores the whole file
        xml = xml[:root.end() - 1] + ' xmlns:ooo="http://openoffice.org/2004/office"' + xml[root.end() - 1:]
    if 'config:name="ooo:configuration-settings"' not in xml:
        xml = xml.replace('</office:settings>', '<config:config-item-set config:name="ooo:configuration-settings"></config:config-item-set></office:settings>', 1)
    for k, tv in items.items():
        t, v = tv if isinstance(tv, tuple) else ('boolean', tv)
        item = f'<config:config-item config:name="{k}" config:type="{t}">{v}</config:config-item>'
        pat = re.compile(rf'<config:config-item config:name="{k}" config:type="\w+">[^<]*</config:config-item>')
        if pat.search(xml): xml = pat.sub(item, xml, count=1); continue
        m = BLOCK.search(xml); i = m.start() + m.group(0).rfind('</config:config-item-set>')
        xml = xml[:i] + item + xml[i:]
    for k, tv in items.items():
        v = tv[1] if isinstance(tv, tuple) else tv
        if not re.search(rf'config:name="{k}" config:type="\w+">{re.escape(v)}</config:config-item>', xml):
            raise RuntimeError(f'injection check failed for {k}')
    if 'xmlns:ooo="http://openoffice.org/2004/office"' not in re.search(r'<office:document-settings\b[^>]*>', xml).group(0):
        raise RuntimeError('ooo namespace missing')
    return xml
