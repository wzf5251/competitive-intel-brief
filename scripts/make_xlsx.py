#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""零依赖 XLSX 生成器（仅用标准库），用于「竞品公开情报监控」的 Excel 台账。矩阵表九 sheet（字段矩阵/招采展开/附件清单/名单核对/新发现对手/待人工确认/冲突记录/字段覆盖/字段字典）、文档表十 sheet（本期变化/新发现对手/招采项目/省份执行/价格对照/附件清单/省份覆盖/竞品档案/政策日历/扫描覆盖）。

用法：
    python make_xlsx.py <spec.json> <output.xlsx>

spec.json 结构：
{
  "sheets": [
    {
      "name": "主表",
      "columns": [{"header": "竞品名称", "width": 16}, {"header": "挂网价", "width": 12}],
      "rows": [["乐提葆", 1280.0, "A 级"], ["竞品B", null, "待确认"]],
      "band": "deep",
      "highlight": {"companies": {"col": "竞品", "names": ["保妥适", "乔雅登"]},
                    "rows": [{"col": "字段", "in": ["集采中选价"], "style": "none"}]},
      "freeze": true,
      "autofilter": true
    }
  ]
}

说明：
- 首行固定为表头（深藏青底 + 白字 + 细边框），行高按内容估算，避免换行文字被裁切。
- 数字写为数值单元格，其余写为内联字符串；None / 空串写为空单元格。
- 单元格值整体是 http(s) 地址时，自动写成超链接（蓝色下划线），点开即达来源页。
- 冻结首行 + 自动筛选为默认开启，可在 spec 里关掉。

配色（数据一格不少，只把该先看的行/格顶出来）：
  公司色  重点竞品整行，每家一色：spec 的 highlight.companies.names 按顺序分配
          淡蓝 c1 / 淡紫 c2 / 淡青 c3 / 淡粉 c4，第五家起循环；
          旧写法 highlight.rows + style="deep" 仍可用（全表同一淡蓝）
  淡橙底  新发现对手整行（sheet 名「新发现对手」自动，或 band=new）
  淡黄整行  待人工确认（sheet 名自动）
  琥珀底  「未获取」（选了没查到，是要补的缺口）；「黄」级；置信度待确认；解析/核对需人工
  淡红底  冲突记录整行
  红/黄/绿底  「级别」列的 红 / 黄 / 绿，字体同步加粗
  灰字    「未展开 / 未查 / 不适用 / 待补」（本期没选或不成立，不是缺口）；「是否命中」为否；来源等级 C
优先级：单元格规则 > 行底色 > 无。行标统一用 sheet.band 或 highlight.rows[].style。
"""

import json
import math
import re
import sys
import zipfile

ACCENT = "1A3A5C"      # 深藏青，与 HTML 版式的强调色一致
HAIR = "BBBBBB"
FONT_NAME = "微软雅黑"

CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# ---------------------------------------------------------------- 样式字典

FONTS = {
    "n":     '<font><sz val="10"/><name val="%(f)s"/><color rgb="FF111111"/></font>',
    "h":     '<font><b/><sz val="10"/><name val="%(f)s"/><color rgb="FFFFFFFF"/></font>',
    "link":  '<font><sz val="10"/><name val="%(f)s"/><color rgb="FF1155CC"/><u/></font>',
    "red":   '<font><b/><sz val="10"/><name val="%(f)s"/><color rgb="FFB3261E"/></font>',
    "amber": '<font><b/><sz val="10"/><name val="%(f)s"/><color rgb="FF8A5A00"/></font>',
    "green": '<font><b/><sz val="10"/><name val="%(f)s"/><color rgb="FF2E6B3E"/></font>',
    "grey":  '<font><sz val="10"/><name val="%(f)s"/><color rgb="FF9A9A9A"/></font>',
    "focus": '<font><b/><sz val="10"/><name val="%(f)s"/><color rgb="FF1A3A5C"/></font>',
}

FILLS = {
    "none":    None,
    "gray125": "gray125",
    "head":    ACCENT,
    "deep":    "E8F0FA",   # 淡蓝：重点竞品（单色模式，兼容旧 spec）
    "c1":      "E8F0FA",   # 公司色 1：淡蓝
    "c2":      "EEEAFA",   # 公司色 2：淡紫
    "c3":      "E1F2F1",   # 公司色 3：淡青
    "c4":      "FBEFF3",   # 公司色 4：淡粉
    "warn":    "FFF6DE",   # 淡黄：待确认
    "new":     "FDEDE3",   # 淡橙：新发现对手
    "redbg":   "FADBD8",   # 淡红单元格：红级
    "amberbg": "FCECC5",   # 琥珀单元格：黄级
    "greenbg": "DFF0E4",   # 淡绿单元格：绿级
    "redrow":  "FBEAE7",   # 淡红整行：冲突记录
    "mute":    "F4F4F4",
}

# 重点公司按顺序分色：第一家常蓝、第二家紫、第三家青、第四家粉，第五家起循环
COMPANY_BANDS = ("c1", "c2", "c3", "c4")

# sheet 名 -> 默认整行底色，省得每份 spec 都声明
SHEET_BAND = {
    "待人工确认": "warn",
    "新发现对手": "new",
    "冲突记录": "redrow",
}

GREY_VALUES = ("未展开", "未查", "未覆盖", "待补")   # 「未获取」走琥珀底，不在此列

LEVEL_STYLE = {"红": ("red", "redbg"), "黄": ("amber", "amberbg"), "绿": ("green", "greenbg")}


class Styles:
    """按用到的组合动态生成 fonts / fills / cellXfs，避免枚举爆炸。"""

    def __init__(self):
        # 前四个索引固定，历史行为不变
        self._fonts = ["n", "h", "link"]
        self._fills = ["none", "gray125", "head"]
        self._xfs = [("n", "none", 0, "none"),
                     ("h", "head", 1, "center"),
                     ("n", "none", 1, "top"),
                     ("link", "none", 1, "top")]
        self._cache = {k: i for i, k in enumerate(self._xfs)}

    def xf(self, font="n", fill="none", border=1, align="top"):
        if font not in FONTS:
            font = "n"
        if fill not in FILLS:
            fill = "none"
        key = (font, fill, border, align)
        if key not in self._cache:
            self._xfs.append(key)
            self._cache[key] = len(self._xfs) - 1
        if font not in self._fonts:
            self._fonts.append(font)
        if fill not in self._fills:
            self._fills.append(fill)
        return self._cache[key]

    def xml(self):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
        out.append('<fonts count="%d">' % len(self._fonts))
        for k in self._fonts:
            out.append(FONTS[k] % {"f": FONT_NAME})
        out.append("</fonts>")

        out.append('<fills count="%d">' % len(self._fills))
        for k in self._fills:
            if k == "none":
                out.append('<fill><patternFill patternType="none"/></fill>')
            elif k == "gray125":
                out.append('<fill><patternFill patternType="gray125"/></fill>')
            else:
                out.append('<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/>'
                           '<bgColor indexed="64"/></patternFill></fill>' % FILLS[k])
        out.append("</fills>")

        out.append('<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
                   '<border><left style="thin"><color rgb="FF{HAIR}"/></left>'
                   '<right style="thin"><color rgb="FF{HAIR}"/></right>'
                   '<top style="thin"><color rgb="FF{HAIR}"/></top>'
                   '<bottom style="thin"><color rgb="FF{HAIR}"/></bottom><diagonal/></border></borders>'
                   .replace("{HAIR}", HAIR))

        out.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
        out.append('<cellXfs count="%d">' % len(self._xfs))
        for fk, ik, bd, al in self._xfs:
            f_idx = self._fonts.index(fk)
            i_idx = self._fills.index(ik)
            if bd == 0:
                out.append('<xf numFmtId="0" fontId="%d" fillId="%d" borderId="0" xfId="0"/>' % (f_idx, i_idx))
                continue
            attrs = '<xf numFmtId="0" fontId="%d" fillId="%d" borderId="1" xfId="0"' % (f_idx, i_idx)
            if fk != "n":
                attrs += ' applyFont="1"'
            if ik != "none":
                attrs += ' applyFill="1"'
            attrs += ' applyBorder="1" applyAlignment="1">'
            if al == "center":
                attrs += '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
            else:
                attrs += '<alignment vertical="top" wrapText="1"/></xf>'
            out.append(attrs)
        out.append("</cellXfs>")
        out.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
        out.append("</styleSheet>")
        return "".join(out)


def decide(header, value, band):
    """单元格该用什么字体与底色。返回 (font_key, fill_key)。"""
    h = (header or "").strip()
    v = "" if value is None else str(value).strip()
    font, fill = "n", band or "none"

    # 级别列：红 / 黄 / 绿
    if "级别" in h and v[:1] in LEVEL_STYLE:
        return LEVEL_STYLE[v[0]]
    # 置信度：待确认 / 待核实 / 低
    if "置信度" in h and any(k in v for k in ("待确认", "待核实", "低")):
        return ("amber", "warn")
    # 冲突列：只要不是「无」就有冲突
    if h == "冲突" and v not in ("", "—", "-", "－", "无", "无冲突", "不适用"):
        return ("red", fill)
    # 附件解析状态：非「已解析」「无附件（已核）」即需人工
    if "解析状态" in h and v and not v.startswith("已解析") and not v.startswith("无附件"):
        return ("amber", "warn")
    # 名单核对状态：需人工
    if "核对状态" in h and "人工" in v:
        return ("amber", "warn")
    # 是否命中 / 附件命中：否
    if "命中" in h and v[:1] in ("否", "无"):
        return ("grey", fill)
    # 停止供应 / 断供：红字提示，但保留行底色
    if "供应状态" in h and any(k in v for k in ("停止", "断供", "退出")):
        return ("red", fill)
    # 来源等级 C：只作线索
    if h == "等级" and v[:1].upper() == "C":
        return ("grey", fill)
    # 字段覆盖四态：未获取=选了查不到，要补（琥珀）；不适用=本产品不成立（灰字，不算缺口）
    if v == "未获取":
        return ("amber", "warn")
    if v == "不适用" and h not in ("价格口径",):
        return ("grey", fill)
    # 未查 / 未展开：灰字，提醒「这是没查，不是没有」
    if v in GREY_VALUES:
        return ("grey", fill)
    return (font, fill)


def row_band(sheet, headers, values):
    """整行底色：公司专属色优先，其次 highlight.rows，再次 sheet.band，最后按 sheet 名兜底。"""
    hl = sheet.get("highlight") or {}

    # 重点公司各给一色：names 的顺序决定颜色，同表里不同公司一眼分开
    comp = hl.get("companies") or {}
    names = [str(x) for x in (comp.get("names") or [])]
    col = comp.get("col") or "竞品"
    if names and col in headers:
        i = headers.index(col)
        cell = "" if i >= len(values) or values[i] is None else str(values[i]).strip()
        if cell in names:
            return COMPANY_BANDS[names.index(cell) % len(COMPANY_BANDS)]

    for rule in hl.get("rows") or []:
        col = rule.get("col")
        if col not in headers:
            continue
        i = headers.index(col)
        if i >= len(values):
            continue
        cell = "" if values[i] is None else str(values[i]).strip()
        cands = [str(c) for c in (rule.get("in") or [])]
        if "eq" in rule:
            cands.append(str(rule["eq"]))
        hit = any(cell == c for c in cands) or any(
            rule.get("fuzzy") and c and c in cell for c in cands)
        if hit:
            return rule.get("style") or "deep"
    band = sheet.get("band") or SHEET_BAND.get(sheet.get("name") or "", "")
    return band or "none"


# ---------------------------------------------------------------- XML 拼装

def esc(text):
    text = CTRL.sub("", str(text))
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


def col_letter(idx):
    """0 -> A, 25 -> Z, 26 -> AA"""
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def disp_width(text):
    """中文按 2 个字符宽估算。"""
    total = 0
    for ch in str(text):
        total += 2 if ord(ch) > 127 else 1
    return total


def est_height(values, widths):
    lines = 1
    for value, width in zip(values, widths):
        if value is None or value == "":
            continue
        usable = max(4, width - 2)
        one = 0
        need = 0
        for ch in str(value).split("\n"):
            one = math.ceil(max(1, disp_width(ch)) / usable)
            need += one
        lines = max(lines, need)
    return min(150, max(18, lines * 14 + 6))


def num_text(value):
    """数值转 XML 文本：整数值不带 .0，其余保留原值。"""
    if isinstance(value, int):
        return str(value)
    text = repr(float(value))
    if text.endswith(".0"):
        text = text[:-2]
    return text


def cell_xml(ref, value, style):
    if value is None or value == "":
        return '<c r="%s" s="%d"/>' % (ref, style)
    if isinstance(value, bool):
        value = "是" if value else "否"
    if isinstance(value, (int, float)):
        return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, num_text(value))
    return '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, style, esc(value))


URL_CELL = re.compile(r"^https?://[^\s<>\"']+$")


def sheet_xml(sheet, st):
    """返回 (worksheet XML, 超链接 rels XML 或 None)。

    单元格值整体是 http(s) 地址时，自动写成可点击超链接（蓝色下划线），
    用户点一下直接打开，不必复制粘贴。这是「矩阵能复核」的关键一环。
    底色与字体按 decide / row_band 自动选，重点行一眼能看出。
    """
    columns = sheet["columns"]
    headers = [c["header"] for c in columns]
    widths = [int(c.get("width", 16)) for c in columns]
    rows = sheet.get("rows", [])
    ncols = len(columns)
    nrows = len(rows) + 1
    last_col = col_letter(ncols - 1)
    ref_last = "%s%d" % (last_col, nrows)

    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
             '<dimension ref="A1:%s"/>' % ref_last]

    if sheet.get("freeze", True):
        parts.append('<sheetViews><sheetView workbookViewId="0">'
                     '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                     '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
                     '</sheetView></sheetViews>')
    else:
        parts.append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')

    parts.append('<sheetFormatPr defaultRowHeight="18"/>')

    # 列宽：title 列宽一点、其余按 spec
    parts.append("<cols>")
    for i, w in enumerate(widths):
        parts.append('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i + 1, i + 1, w))
    parts.append("</cols>")

    links = []          # (ref, rId, url)
    parts.append("<sheetData>")
    # 表头
    head_style = st.xf("h", "head", 1, "center")
    head_cells = "".join(cell_xml("%s1" % col_letter(i), c["header"], head_style) for i, c in enumerate(columns))
    parts.append('<row r="1" ht="30" customHeight="1">%s</row>' % head_cells)
    # 数据行
    for r, values in enumerate(rows, start=2):
        padded = list(values) + [None] * (ncols - len(values))
        band = row_band(sheet, headers, padded)
        out = []
        for i, v in enumerate(padded[:ncols]):
            ref = "%s%d" % (col_letter(i), r)
            font, fill = decide(headers[i], v, band)
            if isinstance(v, str) and URL_CELL.match(v.strip()):
                url = v.strip()
                rid = "rId%d" % (len(links) + 1)
                links.append((ref, rid, url))
                font = "link"
                v = url
            out.append(cell_xml(ref, v, st.xf(font, fill, 1, "top")))
        ht = est_height(padded[:ncols], widths)
        parts.append('<row r="%d" ht="%d" customHeight="1">%s</row>' % (r, ht, "".join(out)))
    parts.append("</sheetData>")

    if sheet.get("autofilter", True) and nrows > 1:
        parts.append('<autoFilter ref="A1:%s"/>' % ref_last)

    if links:
        parts.append("<hyperlinks>")
        for ref, rid, _ in links:
            parts.append('<hyperlink ref="%s" r:id="%s"/>' % (ref, rid))
        parts.append("</hyperlinks>")

    parts.append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
    parts.append("</worksheet>")

    rels = None
    if links:
        rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for _, rid, url in links:
            rels.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
                        ' Target="%s" TargetMode="External"/>' % (rid, esc(url)))
        rels.append("</Relationships>")
        rels = "".join(rels)

    return "".join(parts), rels


def build(spec_path, out_path):
    with open(spec_path, "r", encoding="utf-8") as fh:
        spec = json.load(fh)

    sheets = spec["sheets"]
    if not sheets:
        raise SystemExit("spec 里没有任何 sheet")

    names, used = [], set()
    for i, sh in enumerate(sheets, start=1):
        raw = re.sub(r"[\[\]:*?/\\]", "-", sh.get("name") or ("Sheet%d" % i))[:31]
        name = raw
        k = 2
        while name in used:
            name = raw[:28] + str(k)
            k += 1
        used.add(name)
        names.append(name)

    st = Styles()
    built = [sheet_xml(sh, st) for sh in sheets]

    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                     '<Default Extension="xml" ContentType="application/xml"/>',
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(1, len(sheets) + 1):
        content_types.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
    content_types.append("</Types>")

    workbook = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
                "<sheets>"]
    for i, name in enumerate(names, start=1):
        workbook.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (esc(name), i, i))
    workbook.append("</sheets></workbook>")

    wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(1, len(sheets) + 1):
        wb_rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i))
    wb_rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (len(sheets) + 1))
    wb_rels.append("</Relationships>")

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(content_types))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", "".join(workbook))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
        z.writestr("xl/styles.xml", st.xml())
        for i, (xml, rels) in enumerate(built, start=1):
            z.writestr("xl/worksheets/sheet%d.xml" % i, xml)
            if rels:
                z.writestr("xl/worksheets/_rels/sheet%d.xml.rels" % i, rels)

    return out_path


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    path = build(sys.argv[1], sys.argv[2])
    print("written:", path)
