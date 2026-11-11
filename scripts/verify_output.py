#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""产物机械复检（零依赖，仅标准库）。

用法：
    python verify_output.py <文档.md> [--matrix <矩阵.md>]

默认按同目录、同时间戳自动找 `竞品矩阵-<时间戳>.md`。

文档检查：占位符/样例残留、**八节齐全**、表格列数、变化行是否带级别+行号引用、
          引用行号是否存在（M/T/N 行）、待确认行是否被写成结论、未来发生日期、
          文件名时刻、基线是否存在、价差与价差率复算、HTML 是否引入文档之外的事实、Excel 表数与行数。
矩阵检查：**七节齐全**、七项元数据列是否齐全、每行七项是否非空、**产品名称与规格逐行非空**、
          来源网址是否逐行可打开（拒收「同上 / 见附件 / 只给附件文件名」）、
          置信度/价格口径/来源等级是否落在约定枚举内、行号是否连续唯一、
          字段组是否在四组内、是否混入叙述段落、**统计行三个数是否与明细对账**；
          招采展开区：T 行连续、产品规格可自证、**省份/规格核对行与表内行数一致、国采不少于 32 个省级单元**；
          附件清单区：附件网址指向可下载文件、解析状态五值、**每个项目都有记录、名单等式成立**；
          字段覆盖区：36 行、四态、合计行自洽。

退出码：0 通过（可含告警）｜1 存在错误｜2 用法或文件错误
结尾固定打印「复检计数：通过 N 项｜告警 M 条｜错误 K 条」——交付说明只能抄这行，不得自编项数。
"""
import os
import re
import sys
import zipfile
import datetime
import xml.etree.ElementTree as ET

ERRORS = []
WARNS = []
PASSED = []


def err(msg):
    ERRORS.append(msg)


def warn(msg):
    WARNS.append(msg)


def ok(msg):
    PASSED.append(msg)


CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
SEC_RE = re.compile(r"^##\s*([一二三四五六七八九十]+)、(.+?)\s*$", re.M)
PH_RE = re.compile(r"\{\{\s*[A-Za-z_][A-Za-z0-9_]*\s*\}\}")
LINK_RE = re.compile(r"https?://[^\s<>\"'\)\]\|，。；、）]+")
DATE_RE = re.compile(r"\b(20\d{2})([-/.])(\d{1,2})\2(\d{1,2})\b")
LEVELS = ("红", "黄", "绿")
CITE_RE = re.compile(r"\[([MTN])(\d+)\]")

META_COLS = ("来源名称", "来源网址", "页内位置", "等级", "获取日期", "置信度", "价格口径")
PRODUCT_COLS = ("产品名称", "规格")

# 「竞品」列要填竞品标识（品牌名/清单里的简称），不是企业全称。
# 企业全称属于 F01，只写在该行的「值」列；填错会让整张表按公司而不是按产品聚合。
COMPANY_HARD = ("有限公司", "有限责任公司", "股份有限公司")
COMPANY_SOFT = re.compile(r"(药业|制药|医药|集团|生物科技|医疗科技|医疗美容|生物制药)$")
GROUP_KEYS = ("基础标识", "准入状态", "价格", "招投标")
MATRIX_PREFIX = "竞品矩阵-"
DOC_PREFIX = "竞品监测-"

# 附件清单：解析状态五值（按关键词匹配，兼容全/半角括号）
ATTACH_STATUS_KW = ("已解析", "无附件", "图片型", "加密", "需登录")
ATTACH_EXT = re.compile(r"\.(xlsx?|csv|pdf|docx?|zip|rar|7z)(\?|#|$)", re.I)
NATIONAL_RE = re.compile(r"国采|第[一二三四五六七八九十\d]+批")

# 来源网址：允许的两种写法；写了别的就等于没给来源
NO_URL_PREFIX = "无公开网址"
NO_URL_MAX_SHARE = 0.3        # 「无公开网址」占比超过三成即告警，防滥用
URL_LIKE = re.compile(r"^https?://\S+$")
FILE_TAIL = re.compile(r"\.(xlsx|xls|pdf|docx?|csv|zip)$", re.I)
URL_DODGE = ("同上", "见上", "见附件", "见链接", "见上行", "附件")

# 产物文件名：<类型>-<日期>_<时刻>-<产品名称(规格)>[-冲突号].md
FILE_STAMP_RE = re.compile(r"(竞品矩阵|竞品监测)-(\d{4}-\d{2}-\d{2})_(\d{4})(?:-(.+))?\.md$")
WIN_BAD = re.compile(r'[\\/:*?"<>|]')

# 名单核对：核对状态三值
LIST_STATUS_KW = ("已核对", "名单需人工", "无名单")
# 新发现对手：六项非空硬约束
NEWCOMER_REQUIRED = ("企业名称", "产品名称", "规格", "来源网址", "页内位置", "发现路径")



def url_problem(value):
    """检查一格来源网址。返回 (级别, 说明)；合格返回 (None, None)。级别取 err / warn。"""
    v = (value or "").strip()
    if v in ("", "—", "-", "/"):
        return "err", "为空"
    if URL_LIKE.match(v):
        m = re.match(r"^https?://([^/?#]+)(/?)([?#].*)?$", v)
        if m and not m.group(3) and (not m.group(2) or m.group(2) == ""):
            return "warn", "只有域名，是首页不是具体页"
        if not re.match(r"^https?://[^/]+/.+", v):
            return "warn", "缺少页面路径，无法直达"
        return None, None
    if v.startswith(NO_URL_PREFIX):
        return None, None
    for tok in URL_DODGE:
        if v.startswith(tok):
            return "err", "写「%s」，用户无法据此打开来源" % tok
    if FILE_TAIL.search(v):
        return "err", "只给了文件名，缺附件下载地址"
    return "err", "不是 http(s) 地址（「%s」）" % v[:16]

# 枚举校验：写成约定外的值即为错（SKILL.md 七项元数据与 F20 的取值规则）
ENUM_GRADE = ("A", "B", "C")
ENUM_CONF = ("已确认", "较可能", "待确认")
# 价格口径：四类主口径 + 只出现在特定场景的三种（医保支付标准 / 供应价 / 挂网限价）
ENUM_PRICE_BASIS = ("挂网价", "集采中选价", "中标价", "零售价",
                    "医保支付标准", "供应价", "挂网限价", "不适用")
# 字段 → 允许的口径。填错口径等于把两个不同的价混比，是这份台账最致命的错
PRICE_BASIS_BY_FIELD = {
    "挂网价": ("挂网价", "挂网限价"),
    "挂网限价": ("挂网限价", "挂网价"),
    "集采中选价": ("集采中选价", "供应价"),
    "供应价": ("供应价", "集采中选价"),
    "中标价": ("中标价",),
    "医保支付标准": ("医保支付标准",),
    "零售价": ("零售价",),
}

# 值格式：价格类字段必须以数字开头（「约 1280」「面议」不算值）
PRICE_FIELDS = ("挂网价", "挂网限价", "集采中选价", "中标价", "医保支付标准", "零售价", "供应价")
DATE_FIELDS = ("有效期至", "价格日期", "执行周期", "公告发布日期", "投标截止日",
               "执行起始日", "采购周期止", "目录协议期")
NUM_START = re.compile(r"^[\d０-９]")
BANNED_WORDS = ("约", "大概", "大约", "左右", "预计", "估算", "推算", "拟按", "差不多", "应该有")
# 原文类字段：逐字引述，不套禁用词（官方原文写「约 3 个月」照抄，不算编造）
ORIGIN_FIELDS = ("适用范围", "公告标题")
# 这些值不是「数」，跳过格式校验（是合法占位，不是歪值）
EMPTY_VALUE = ("未展开", "未获取", "未查", "不适用", "未参与", "未披露", "未中选",
               "无", "不承担供应", "未公布", "—", "-", "/", "－")

# 36 个省级单元的标准名与常用简称（归一后才能判断「到底展开几个省」）
PROVINCES = (
    "北京市", "天津市", "河北省", "山西省", "内蒙古自治区", "辽宁省", "吉林省", "黑龙江省",
    "上海市", "江苏省", "浙江省", "安徽省", "福建省", "江西省", "山东省", "河南省",
    "湖北省", "湖南省", "广东省", "广西壮族自治区", "海南省", "重庆市", "四川省", "贵州省",
    "云南省", "西藏自治区", "陕西省", "甘肃省", "青海省", "宁夏回族自治区",
    "新疆维吾尔自治区", "新疆生产建设兵团",
)
PROV_ALIAS = {p: p for p in PROVINCES}
for _short, _full in (
    ("北京", "北京市"), ("天津", "天津市"), ("上海", "上海市"), ("重庆", "重庆市"),
    ("河北", "河北省"), ("山西", "山西省"), ("辽宁", "辽宁省"), ("吉林", "吉林省"),
    ("黑龙江", "黑龙江省"), ("江苏", "江苏省"), ("浙江", "浙江省"), ("安徽", "安徽省"),
    ("福建", "福建省"), ("江西", "江西省"), ("山东", "山东省"), ("河南", "河南省"),
    ("湖北", "湖北省"), ("湖南", "湖南省"), ("广东", "广东省"), ("海南", "海南省"),
    ("四川", "四川省"), ("贵州", "贵州省"), ("云南", "云南省"), ("陕西", "陕西省"),
    ("甘肃", "甘肃省"), ("青海", "青海省"), ("广西", "广西壮族自治区"),
    ("内蒙古", "内蒙古自治区"), ("西藏", "西藏自治区"), ("宁夏", "宁夏回族自治区"),
    ("新疆维吾尔自治区", "新疆维吾尔自治区"), ("新疆", "新疆维吾尔自治区"),
    ("新疆兵团", "新疆生产建设兵团"), ("兵团", "新疆生产建设兵团"),
    ("新疆生产建设兵团", "新疆生产建设兵团"),
):
    PROV_ALIAS[_short] = _full
PROV_BOGUS = ("全国", "多省", "若干", "其他", "其它", "各省", "部分省份", "华中", "华东", "西北", "西南", "华南", "华北")

# 字段覆盖表：四态
COVER_STATES = ("已取", "未展开", "未获取", "不适用")


def read(path):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def cells(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|") and not s.endswith("\\|"):
        s = s[:-1]
    return [c.strip() for c in re.split(r"(?<!\\)\|", s)]


def is_sep(row):
    return bool(row) and all(re.fullmatch(r":?-{2,}:?", c) for c in row if c != "")


def sections(text):
    ms = list(SEC_RE.finditer(text))
    out = []
    for i, m in enumerate(ms):
        end = ms[i + 1].start() if i + 1 < len(ms) else len(text)
        out.append((m.group(1), m.group(2), text.count("\n", 0, m.end()) + 1, text[m.end():end]))
    return out


def tables(block, base_line):
    res, cur = [], []
    for offset, line in enumerate(block.split("\n")):
        no = base_line + offset
        if line.strip().startswith("|"):
            cur.append((no, line))
        elif cur:
            res.append(cur)
            cur = []
    if cur:
        res.append(cur)
    out = []
    for blk in res:
        head = cells(blk[0][1])
        rows = []
        for no, line in blk[1:]:
            cs = cells(line)
            if is_sep(cs):
                continue
            rows.append((no, len(cs), line))
        out.append((blk[0][0], len(head), rows, head))
    return out


def check_common(text, path, expect_sections, label):
    ph = PH_RE.findall(text)
    if ph:
        err("%s 残留占位符 %d 处：%s" % (label, len(ph), "、".join(sorted(set(ph))[:6])))

    marks = [m for m in ("示意数据", "示例数据", "编的", "假数据", "FILL GUIDE", "待填", "TODO", "XXX") if m in text]
    if marks:
        err("%s 含未清理的占位/样例标记：%s" % (label, "、".join(marks)))

    secs = sections(text)
    got = sorted(CN.get(n, 99) for n, _, _, _ in secs)
    if got == list(range(1, expect_sections + 1)):
        ok("%s 章节齐全（%d 节）" % (label, expect_sections))
    else:
        err("%s 章节不完整：期望 1–%d 节，实得 %d 节（%s）"
            % (label, expect_sections, len(got), "、".join(str(g) for g in got)))

    bad = 0
    for name, _t, line_no, body in secs:
        for start, head_n, rows, _h in tables(body, line_no):
            for no, n, line in rows:
                if n != head_n:
                    bad += 1
                    if bad <= 6:
                        err("%s 第 %d 行表格列数 %d ≠ 表头 %d：%s"
                            % (label, no, n, head_n, line.strip()[:56]))
    if bad == 0:
        ok("%s 表格列数一致" % label)
    elif bad > 6:
        err("……%s 另有 %d 处列数不一致未列出" % (label, bad - 6))

    bad_fmt = set()
    # URL 里的 /2026/11/4/ 不是日期写法问题，先剔掉链接再比对分隔符
    for y, sep, m1, d1 in DATE_RE.findall(LINK_RE.sub(" ", text)):
        if sep != "-":
            bad_fmt.add("%s%s%s%s%s" % (y, sep, m1, sep, d1))
    if bad_fmt:
        warn("%s 日期分隔符不统一：%s" % (label, "、".join(sorted(bad_fmt)[:5])))
    else:
        ok("%s 日期格式统一" % label)

    base = os.path.basename(path)
    m = FILE_STAMP_RE.match(base)
    if m:
        expect = "%s %s:%s" % (m.group(2), m.group(3)[:2], m.group(3)[2:])
        if expect in text:
            ok("%s 文件名时刻与正文一致（%s）" % (label, expect))
        else:
            dm = re.search(r"(20\d{2}-\d{2}-\d{2})\s+(\d{2}:\d{2})", text)
            if dm and "%s %s" % (dm.group(1), dm.group(2)) != expect:
                err("%s 文件名时刻 %s 与正文 %s %s 不一致" % (label, expect, dm.group(1), dm.group(2)))
            else:
                warn("%s 正文未找到生成时刻，无法与文件名 %s 核对" % (label, expect))

        label_part = (m.group(4) or "").strip()
        if label_part:
            if WIN_BAD.search(label_part):
                err("%s 文件名产品段含 Windows 非法字符：%s（/ \\ : * ? \" < > | 需替换）"
                    % (label, label_part))
            elif "(" in label_part and ")" in label_part:
                ok("%s 文件名带产品名称(规格)：%s" % (label, label_part))
            elif "（" in label_part and "）" in label_part:
                warn("%s 文件名产品段用了全角括号「%s」：外层应为半角，写成 产品名称(规格)"
                     % (label, label_part))
            else:
                warn("%s 文件名产品段「%s」未按 产品名称(规格) 写，用户看不出盯的是哪个规格"
                     % (label, label_part))
        else:
            warn("%s 文件名未带产品名称(规格)：五份产物都要带同一段，便于一眼认领与归档"
                 % label)
    else:
        warn("%s 文件名不符合规范：%s" % (label, base))

    links = LINK_RE.findall(text)
    home = []
    for u in set(links):
        mm = re.match(r"https?://([^/]+)(/.*)?$", u)
        if mm and (not mm.group(2) or mm.group(2) in ("/", "")):
            home.append(u)
    if home:
        warn("%s 疑似首页顶替 %d 条：%s" % (label, len(home), "、".join(sorted(home)[:3])))
    elif links:
        ok("%s 链接均为具体页（%d 条）" % (label, len(set(links))))
    return secs


def check_aux_links(body, line_no, label, min_url_cols):
    """待确认区 / 冲突记录区的链接列：列要在，值要是能打开的地址。"""
    tbls = tables(body, line_no)
    if not tbls:
        return
    for start, head_n, rows, head in tbls:
        n_url_cols = sum(1 for h in head if "网址" in h or "链接" in h)
        if n_url_cols < min_url_cols:
            warn("%s缺链接列（至少需 %d 列含「网址」），用户无法据此复核"
                 % (label, min_url_cols))
            continue
        if not rows:
            continue
        idx = [i for i, h in enumerate(head) if "网址" in h or "链接" in h]
        bad = 0
        for no, n, line in rows:
            cs = cells(line)
            if len(cs) < head_n:
                continue
            if all(c in ("", "—", "-", "－") for c in cs):
                continue
            # 「本期无冲突」这类占位行没有网址是正常的，不该算不合规
            if "无冲突" in line or "本期无" in line:
                continue
            for j in idx:
                if j < len(cs):
                    sev, _ = url_problem(cs[j].strip())
                    if sev == "err":
                        bad += 1
        if bad:
            warn("%s有 %d 个链接格不合规：确无公开地址时写「无公开网址（原因）」"
                 % (label, bad))
        else:
            ok("%s链接列可用" % label)


def find_sec(secs, keyword):
    """按标题关键词找节，不依赖节序号——矩阵从 4 节扩到 7 节后，序号会漂。"""
    for name, title, line_no, body in secs:
        if keyword in title:
            return title, line_no, body
    return None


def data_rows(body, line_no):
    """数一个区里的数据行，忽略占位行（「本期无冲突」这类）。"""
    n = 0
    for _s, _h, rows, _hd in tables(body, line_no):
        for _no, _n, line in rows:
            cs = cells(line)
            if all(c in ("", "—", "-", "－") for c in cs):
                continue
            if any(k in line for k in ("无冲突", "本期无", "无待确认")):
                continue
            n += 1
    return n


def norm_txt(s):
    """去掉空白与全/半角括号差异，用于比对「产品名称列」与 F04 值是否同源。"""
    return re.sub(r"[\s（）()]", "", s or "")


def norm_prov(v):
    """省份归一：简称 → 标准名。归一后仍不在 32 个省级单元内，返回 None。

    这一步是防「山西」和「山西省」同时出现、去重后虚高成两个省，把漏省盖过去。
    """
    s = (v or "").strip().replace("　", "")
    s = re.sub(r"[（(].*?[)）]", "", s).strip()      # 去掉「（含兵团）」这类括注
    if s in PROV_ALIAS:
        return PROV_ALIAS[s]
    if s in PROV_BOGUS:
        return ""
    return None


def num_of(v):
    """从「268 元/支（100U，含税）」「-12」这类值里取第一个数（含负号），
    用于跨表对账与价差复算。价差必须保留符号，否则「贵了」会被算成「便宜了」。"""
    s = (v or "").replace(",", "").replace("−", "-").replace("－", "-").strip()
    m = re.search(r"(-?\d+(?:\.\d+)?)", s)
    return float(m.group(1)) if m else None


def check_coverage(body, line_no, full_text, m_rows):
    """字段覆盖表：36 行、四态、合计行自洽。这是「数据全不全」的唯一权威视图。"""
    tbls = tables(body, line_no)
    if not tbls:
        err("矩阵缺「字段覆盖」表：36 个字段一项不少地摊开，用户才能看出缺哪格")
        return

    for start, head_n, rows, head in tbls:
        i_fid = next((i for i, h in enumerate(head) if "字段号" in h), None)
        if i_fid is None:
            continue
        comp_cols = [i for i, h in enumerate(head)
                     if i >= 3 and h.strip() and "字段" not in h and "必填" not in h]
        if not comp_cols:
            warn("字段覆盖表没有竞品列：每家的覆盖情况要分列，不能合成一列")
            return

        seen = []
        counts = {i: {s: 0 for s in COVER_STATES} for i in comp_cols}
        bad_state = []
        total_row = None

        for no, n, line in rows:
            cs = cells(line)
            if len(cs) < head_n:
                continue
            fid = cs[i_fid].strip()
            if fid.startswith("合计") or "合计" in fid:
                total_row = (no, cs)
                continue
            m = re.fullmatch(r"F(\d{2})", fid)
            if not m:
                err("字段覆盖表第 %d 行字段号不合规（应 F01–F36）：%s" % (no, fid[:12]))
                continue
            seen.append(int(m.group(1)))
            for i in comp_cols:
                if i >= len(cs):
                    continue
                v = cs[i].strip()
                if v in counts[i]:
                    counts[i][v] += 1
                elif v:
                    bad_state.append((no, fid, head[i][:8], v[:10]))

        if bad_state:
            for no, fid, col, v in bad_state[:5]:
                err("字段覆盖第 %d 行「%s」（%s 列）取值「%s」不在四态内（已取/未展开/未获取/不适用）"
                    % (no, fid, col, v))
        elif seen:
            ok("字段覆盖表四态取值均合规")

        if sorted(seen) == list(range(1, 37)):
            ok("字段覆盖表 36 项齐全（F01–F36）")
        else:
            miss = [i for i in range(1, 37) if i not in seen]
            extra = [i for i in seen if i > 36 or seen.count(i) > 1]
            if miss:
                err("字段覆盖表缺 %d 个字段：%s；36 项一项不少，缺的写「未展开」而不是不列"
                    % (len(miss), "、".join("F%02d" % i for i in miss[:8])))
            if extra:
                err("字段覆盖表字段号重复或多出：%s" % "、".join("F%02d" % i for i in sorted(set(extra))[:6]))

        # 合计行必须等于逐列计数
        if total_row:
            tno, tcs = total_row
            for i in comp_cols:
                if i >= len(tcs):
                    continue
                cell = tcs[i].strip()
                if not cell or cell in ("—", "-"):
                    continue
                got = {}
                for st in COVER_STATES:
                    mm = re.search(r"(\d+)\s*" + st, cell)
                    if mm:
                        got[st] = int(mm.group(1))
                if not got:
                    warn("字段覆盖合计行第 %d 列「%s」没写四个计数，无法核对" % (tno, cell[:16]))
                    continue
                for st, v in got.items():
                    if counts[i][st] != v:
                        err("字段覆盖合计行与明细不符：%s 列「%s」写 %d，逐行数得 %d"
                            % (head[i][:8], st, v, counts[i][st]))
                if len(got) == 4 and sum(got.values()) != 36:
                    err("字段覆盖合计行 %s 列四项相加 %d ≠ 36" % (head[i][:8], sum(got.values())))
        elif seen:
            warn("字段覆盖表缺合计行（每家一行「x 已取／y 未展开／z 未获取／w 不适用」）")

        # 单家竞品时，未展开数必须等于 36 − 统计行的预设项数
        if len(comp_cols) == 1 and seen:
            mn = re.search(r"覆盖字段\s*(\d+)\s*/\s*36", full_text)
            open_n = counts[comp_cols[0]]["未展开"]
            if mn:
                expect = 36 - int(mn.group(1))
                if open_n != expect:
                    err("字段覆盖「未展开」%d 项，但统计行写覆盖 %s/36，应为 %d 项未展开——"
                        "预设名与实际取数不符" % (open_n, mn.group(1), expect))
                else:
                    ok("字段覆盖与统计行的预设项数一致（未展开 %d 项）" % open_n)
            if counts[comp_cols[0]]["已取"] == 0 and m_rows:
                err("字段覆盖表一个「已取」都没有，但主矩阵有 %d 行：覆盖表没跟着矩阵填" % len(m_rows))
        return

    err("字段覆盖表缺「字段号」列，无法核对 36 项是否齐全")


def check_tender(body, line_no, rows_map, full_text):
    """招采展开区：T 行编号连续、产品/规格逐行有值、省份数量与核对行自洽。返回 (项目集合, T 行数)。"""
    tbls = tables(body, line_no)
    if not tbls:
        warn("矩阵「招采展开」区没有表（本期无招采数据可写「本期无招采项目」）")
        return set(), 0, {}

    t_ids = []
    prov_by_proj = {}      # proj -> {标准省名}，用于国采「至少 32 个省级单元」判断
    prov_row_n = {}        # proj -> 省份执行行数（竞品 × 省），用于与核对行「已展开」对账
    prov_raw = {}          # proj -> {(竞品, 标准省名): [(原写法, T行号)]}，抓同一竞品同省重复占行
    specs_by_proj = {}
    projects = set()
    empty_prod = []
    tender_price = {}      # (竞品, 产品, 规格) -> [(数值, T行号)]，用于和主矩阵 F24 对账
    n_rows = 0

    for start, head_n, rows, head in tbls:
        i_prod = next((i for i, h in enumerate(head) if "产品名称" in h), None)
        i_spec = next((i for i, h in enumerate(head) if "规格" in h), None)
        i_prov = next((i for i, h in enumerate(head) if h.strip() == "省份"), None)
        i_proj = next((i for i, h in enumerate(head) if "项目" in h), None)
        i_comp = next((i for i, h in enumerate(head) if h.strip() == "竞品"), None)
        i_price = next((i for i, h in enumerate(head)
                        if "中选价" in h or "价格" in h and "单位" not in h), None)
        if i_prod is None or i_spec is None:
            err("招采展开表缺「产品名称」或「规格/型号」列，无法把价格对到具体产品")
        for no, n, line in rows:
            cs = cells(line)
            if len(cs) < head_n or all(c in ("", "—", "-") for c in cs):
                continue
            n_rows += 1
            rid = cs[0].strip()
            if not re.fullmatch(r"T\d{2,4}", rid):
                err("招采展开第 %d 行行号不合规（应 T01 起）：%s" % (no, rid[:16]))
                continue
            t_ids.append(int(rid[1:]))
            rows_map["T%02d" % int(rid[1:])] = {"line": no, "conf": "", "value": ""}
            if i_prod is not None and i_prod < len(cs) and cs[i_prod].strip() in ("", "—", "-"):
                empty_prod.append((no, rid, "产品名称"))
            if i_spec is not None and i_spec < len(cs) and cs[i_spec].strip() in ("", "—", "-"):
                empty_prod.append((no, rid, "规格/型号"))
            proj = cs[i_proj].strip() if (i_proj is not None and i_proj < len(cs)) else ""
            if proj:
                projects.add(proj)
            if i_prov is None:
                if proj:
                    specs_by_proj[proj] = specs_by_proj.get(proj, 0) + 1
                # 表一的中选价留一份，和主矩阵 F24 对账：两处写出不同的数，整份台账就不可信了
                if i_price is not None and i_price < len(cs):
                    key = (cs[i_comp].strip() if i_comp is not None and i_comp < len(cs) else "",
                           cs[i_prod].strip() if i_prod is not None and i_prod < len(cs) else "",
                           cs[i_spec].strip() if i_spec is not None and i_spec < len(cs) else "")
                    num = num_of(cs[i_price].strip())
                    if num is not None:
                        tender_price.setdefault(key, []).append((num, rid))
            elif i_prov < len(cs):
                p = cs[i_prov].strip()
                if p in PROV_BOGUS:
                    err("招采展开第 %d 行（%s）省份写「%s」：必须逐省一行，不能合并"
                        % (no, rid, p))
                elif p and p not in ("—", "-", "/") and proj:
                    std = norm_prov(p)
                    if std == "":
                        err("招采展开第 %d 行（%s）省份「%s」不是省级单元：只能写 31 省或新疆兵团，"
                            "「华中」「其他」这类写法和项目应展开数对不上" % (no, rid, p[:12]))
                    elif std is None:
                        err("招采展开第 %d 行（%s）省份「%s」不在 32 个省级单元内，写法需统一"
                            % (no, rid, p[:12]))
                    else:
                        # 去重必须带竞品：两家竞品在同一项目各列 32 省，是 64 行而不是 32 个重复
                        comp = cs[i_comp].strip() if (i_comp is not None and i_comp < len(cs)) else ""
                        prov_raw.setdefault(proj, {}).setdefault((comp, std), []).append((p, rid))
                        prov_by_proj.setdefault(proj, set()).add(std)
                        prov_row_n[proj] = prov_row_n.get(proj, 0) + 1

    if empty_prod:
        for no, rid, col in empty_prod[:5]:
            err("招采展开第 %d 行（%s）%s为空：每行必须能自证是哪个产品哪个规格" % (no, rid, col))
        if len(empty_prod) > 5:
            err("招采展开另有 %d 处产品名称/规格为空" % (len(empty_prod) - 5))
    else:
        ok("招采展开每行产品名称与规格齐全")

    if t_ids:
        nums = sorted(t_ids)
        dup = len(nums) - len(set(nums))
        gaps = [i for i in range(1, max(nums) + 1) if i not in nums]
        if dup:
            err("招采展开 T 行号重复 %d 个（同一行号用了两次）" % dup)
        elif gaps:
            err("招采展开 T 行号不连续，缺：%s" % "、".join("T%02d" % g for g in gaps[:6]))
        else:
            ok("招采展开 T 行连续唯一（T01–T%02d）" % len(nums))

    def parse_check(kind):
        pat = kind + r"[：:]\s*(.+?)\s*应展开\s*(\d+)\s*[｜|]\s*已展开\s*(\d+)\s*[｜|]\s*未展开\s*(\d+)"
        return {m.group(1): tuple(int(x) for x in m.groups()[1:])
                for m in re.finditer(pat, full_text)}

    def pick(checks, proj):
        if proj in checks:
            return checks[proj]
        for k, v in checks.items():
            if proj in k or k in proj:
                return v
        return None

    # 同一竞品同一省两行：写法不统一（山西/山西省）会让去重数虚高，把真正漏掉的省盖住
    for proj, bykey in prov_raw.items():
        dup = {k: v for k, v in bykey.items() if len(v) > 1}
        if dup:
            for (comp, std), hits in list(dup.items())[:3]:
                err("「%s」%s省份「%s」出现 %d 行（%s）：同一竞品一省一行，写法要统一，"
                    "否则去重后省份数虚高、漏省看不出来"
                    % (proj[:18], ("竞品「%s」" % comp) if comp else "", std, len(hits),
                       "、".join("%s@%s" % h for h in hits[:3])))

    # 省份：逐项目核对。「已展开」按行数（竞品 × 省）算，与「规格核对」的算法一致；
    # 「应展开」= 该项目下每家竞品各自应展开省数之和（两家国采 = 64）
    prov_checks = parse_check("省份核对")
    if prov_by_proj and not prov_checks:
        warn("缺省份核对行（每个集采项目一行：省份核对：项目名 应展开 N｜已展开 N｜未展开 N）")
    for proj, ps in prov_by_proj.items():
        if NATIONAL_RE.search(proj) and len(ps) < 32:
            err("「%s」只覆盖 %d 个省级单元；国采应为 31 省 + 新疆兵团 = 32 个（多家竞品按去重省数算）"
                % (proj[:24], len(ps)))
        v = pick(prov_checks, proj)
        if not v:
            if prov_checks:
                warn("「%s」没有省份核对行，无法判断是否查全" % proj[:24])
            continue
        should, done, missing = v
        if should != done + missing:
            err("「%s」省份核对不自洽：应展开 %d ≠ 已展开 %d + 未展开 %d"
                % (proj[:20], should, done, missing))
        rows_here = prov_row_n.get(proj, 0)
        if rows_here != done:
            err("「%s」省份执行表 %d 行，核对行写「已展开 %d」——行数按竞品×省算，必须与表一致"
                % (proj[:20], rows_here, done))
        elif not missing and done == should:
            ok("「%s」省份展开 %d/%d，无缺省" % (proj[:20], done, should))
        if missing:
            warn("「%s」省份未展开 %d 个，须在待确认区逐条写明缺哪一省" % (proj[:20], missing))

    # 规格：应展开 = 该项目下所有监控对象的规格记录总数（多竞品相加）
    spec_checks = parse_check("规格核对")
    if specs_by_proj and not spec_checks:
        warn("缺规格核对行（每个项目一行：规格核对：项目名 应展开 N｜已展开 N｜未展开 N）")
    for proj, cnt in specs_by_proj.items():
        v = pick(spec_checks, proj)
        if not v:
            continue
        should, done, missing = v
        if should != done + missing:
            err("「%s」规格核对不自洽：应展开 %d ≠ 已展开 %d + 未展开 %d"
                % (proj[:20], should, done, missing))
        if cnt != done:
            err("「%s」规格表 %d 条，核对行写「已展开 %d」" % (proj[:20], cnt, done))

    return projects, n_rows, tender_price


def check_attachments(body, line_no, projects, tender_rows, full_text, rows_map=None):
    """附件清单区三张表：附件登记（网址/解析状态）、名单核对（等式成立）、新发现对手（N 行六项非空）。

    这一节要回答「名单里的企业有没有全部进产物」——只登记附件不算查完。
    """
    tbls = tables(body, line_no)
    if not tbls:
        err("矩阵缺「附件清单」区：招采公告的名单、规格、价格都在附件里，不登记等于没查")
        return
    n_rows = 0
    bad_url = []
    bad_status = []
    hit_proj = set()
    parsed = 0
    att_status = {}          # 附件号 -> 解析状态

    list_rows = 0
    list_bad = []            # (line_no, msg) 名单核对等式
    list_status_bad = []
    checked_ids = set()
    candidate_total = 0

    newcomer_rows = 0
    newcomer_bad = []        # (line_no, msg) N 行缺项
    n_ids = []

    for start, head_n, rows, head in tbls:
        i_url = next((i for i, h in enumerate(head) if "附件网址" in h), None)
        if i_url is None:
            i_url = next((i for i, h in enumerate(head) if "网址" in h), None)
        i_st = next((i for i, h in enumerate(head) if "解析状态" in h), None)
        i_proj = next((i for i, h in enumerate(head) if h.strip() == "项目"), None)
        i_name = next((i for i, h in enumerate(head) if "附件名" in h), None)
        i_aid = next((i for i, h in enumerate(head) if "附件号" in h), None)

        # 3.2 名单核对
        i_should = next((i for i, h in enumerate(head) if "应识别" in h), None)
        i_watch = next((i for i, h in enumerate(head) if "已监控" in h), None)
        i_new = next((i for i, h in enumerate(head) if "新增候选" in h), None)
        i_other = next((i for i, h in enumerate(head) if "非竞品" in h), None)
        i_lst = next((i for i, h in enumerate(head) if "核对状态" in h), None)

        # 3.3 新发现对手
        i_ent = next((i for i, h in enumerate(head) if "企业名称" in h), None)

        for no, n, line in rows:
            cs = cells(line)
            if len(cs) < head_n or all(c in ("", "—", "-") for c in cs):
                continue

            if i_ent is not None:
                # ---- 3.3 新发现对手 ----
                newcomer_rows += 1
                rid = cs[0].strip()
                if not re.fullmatch(r"N\d{2,4}", rid):
                    newcomer_bad.append((no, "行号不合规（应 N01 起）：%s" % rid[:12]))
                else:
                    n_ids.append(int(rid[1:]))
                    if rows_map is not None:
                        ci = next((i for i, h in enumerate(head) if "置信度" in h), None)
                        rows_map["N%02d" % int(rid[1:])] = {
                            "line": no,
                            "conf": cs[ci].strip() if ci is not None and ci < len(cs) else "",
                            "value": "",
                        }
                for col in NEWCOMER_REQUIRED:
                    j = next((i for i, h in enumerate(head) if col in h), None)
                    if j is None:
                        newcomer_bad.append((no, "缺「%s」列，用户无法据此判断与复核" % col))
                    elif j >= len(cs) or cs[j].strip() in ("", "—", "-"):
                        newcomer_bad.append((no, "%s为空" % col))
                if i_url is not None and i_url < len(cs):
                    sev, msg = url_problem(cs[i_url].strip())
                    if sev == "err":
                        newcomer_bad.append((no, "来源网址%s" % msg))
                continue

            if i_should is not None:
                # ---- 3.2 名单核对 ----
                list_rows += 1
                if i_aid is not None and i_aid < len(cs):
                    checked_ids.add(cs[i_aid].strip())
                nums = {}
                for key, j in (("should", i_should), ("watch", i_watch),
                               ("new", i_new), ("other", i_other)):
                    if j is not None and j < len(cs):
                        mm = re.search(r"\d+", cs[j])
                        nums[key] = int(mm.group()) if mm else None
                if None in (nums.get("should"), nums.get("watch"),
                            nums.get("new"), nums.get("other")):
                    list_bad.append((no, "四列数字不全（应识别 / 已监控 / 新增候选 / 非竞品）"))
                else:
                    total = nums["watch"] + nums["new"] + nums["other"]
                    if nums["should"] != total:
                        list_bad.append(
                            (no, "等式不成立：应识别 %d ≠ 已监控 %d + 新增候选 %d + 非竞品 %d = %d"
                             % (nums["should"], nums["watch"], nums["new"], nums["other"], total)))
                    candidate_total += nums["new"]
                if i_lst is not None and i_lst < len(cs):
                    v = cs[i_lst].strip()
                    if v and not any(k in v for k in LIST_STATUS_KW):
                        list_status_bad.append((no, v[:14]))
                continue

            # ---- 3.1 附件登记 ----
            n_rows += 1
            if i_proj is not None and i_proj < len(cs):
                hit_proj.add(cs[i_proj].strip())
            if i_url is not None and i_url < len(cs):
                v = cs[i_url].strip()
                sev, msg = url_problem(v)
                if sev == "err":
                    bad_url.append((no, msg))
                elif v.startswith("http") and not ATTACH_EXT.search(v) and "down" not in v.lower():
                    # 公告页不是附件：可以登记，但要么是「无附件（已核）」，要么就该给附件直链
                    if i_name is not None and i_name < len(cs):
                        nm = cs[i_name].strip()
                        if nm and not nm.startswith("无附件") and ATTACH_EXT.search(nm):
                            bad_url.append((no, "附件名是文件但网址给的是公告页，应给附件直链"))
            if i_st is not None and i_st < len(cs):
                v = cs[i_st].strip()
                if v and not any(k in v for k in ATTACH_STATUS_KW):
                    bad_status.append((no, v[:12]))
                elif "已解析" in v:
                    parsed += 1
                if i_aid is not None and i_aid < len(cs):
                    att_status[cs[i_aid].strip()] = v

    if bad_url:
        for no, msg in bad_url[:5]:
            err("附件清单第 %d 行附件网址%s" % (no, msg))
    elif n_rows:
        ok("附件清单 %d 行网址均指向可下载附件或具体公告" % n_rows)

    if bad_status:
        for no, v in bad_status[:5]:
            err("附件清单第 %d 行解析状态「%s」不在五值内（已解析/无附件（已核）/图片型需人工/加密需密码/需登录）"
                % (no, v))
    elif n_rows:
        ok("附件清单解析状态均在约定五值内")

    if not n_rows and tender_rows:
        err("有 %d 条招采展开行，但附件清单一条都没有：每个公告都要登记附件，无附件也要写「无附件（已核）」"
            % tender_rows)
    elif not n_rows:
        pass
    else:
        missing = [p for p in projects
                   if p and not any(p in hp or hp in p for hp in hit_proj if hp)]
        if missing:
            warn("这些项目在附件清单里没有对应记录，无法判断是漏查还是无附件：%s"
                 % "、".join(missing[:3]))
        else:
            ok("附件清单覆盖全部招采项目（%d 个）" % len(projects))

        if parsed == 0:
            warn("附件清单没有任何「已解析」记录：确认是真的都是扫描件/需登录，而不是没打开")

    # ---- 名单核对：等式、状态、与「已解析」附件的对应关系 ----
    if list_bad:
        for no, msg in list_bad[:5]:
            err("名单核对第 %d 行%s" % (no, msg))
        if len(list_bad) > 5:
            err("名单核对另有 %d 行数字对不上" % (len(list_bad) - 5))
    elif list_rows:
        ok("名单核对 %d 行等式成立（应识别 = 已监控 + 新增候选 + 非竞品）" % list_rows)
    elif n_rows:
        err("有 %d 份附件却一条名单核对都没有：附件里的企业名单要逐份抄全，"
            "否则同组对手会被整批漏掉（本期一家同组企业都没有时，也要写核对行并填 0）"
            % n_rows)

    if list_status_bad:
        for no, v in list_status_bad[:5]:
            err("名单核对第 %d 行核对状态「%s」不在三值内（已核对/名单需人工（图片型）/无名单（已核））"
                % (no, v))

    lacking = [aid for aid, st in att_status.items()
               if "已解析" in (st or "") and aid and not any(aid in x for x in checked_ids if x)]
    if lacking:
        err("这些附件标了「已解析」却没有名单核对行：%s；把名单逐条抄出来才算查过"
            % "、".join(lacking[:4]))

    # ---- 新发现对手：N 行六项、行号、与候选数对齐 ----
    if newcomer_bad:
        for no, msg in newcomer_bad[:5]:
            err("新发现对手第 %d 行%s" % (no, msg))
        if len(newcomer_bad) > 5:
            err("新发现对手另有 %d 处缺项：六项非空是硬约束，缺一项就进待确认区"
                % (len(newcomer_bad) - 5))
    elif newcomer_rows:
        ok("新发现对手 %d 行，六项必备信息齐全" % newcomer_rows)

    if n_ids:
        nums = sorted(n_ids)
        dup = len(nums) - len(set(nums))
        gaps = [i for i in range(1, max(nums) + 1) if i not in nums]
        if dup:
            err("新发现对手 N 行号重复 %d 个" % dup)
        elif gaps:
            err("新发现对手 N 行号不连续，缺：%s" % "、".join("N%02d" % g for g in gaps[:6]))
        else:
            ok("新发现对手 N 行连续唯一（N01–N%02d）" % len(nums))

    if candidate_total and candidate_total != newcomer_rows:
        err("名单核对写「新增候选 %d 家」，新发现对手表只有 %d 行——名单里识别出的同组企业必须逐家落行"
            % (candidate_total, newcomer_rows))
    elif candidate_total and newcomer_rows:
        ok("名单核对新增候选 %d 家 = 新发现对手 %d 行（同源）" % (candidate_total, newcomer_rows))

    if not newcomer_rows and n_rows and "无新发现对手" not in full_text:
        warn("没有新发现对手行，也没写「本期无新发现对手（名单已核 N 份）」："
             "读者分不清是真没有，还是没查名单")

    m = re.search(r"附件核对[：:]\s*应查\s*(\d+)\s*[｜|]\s*已查\s*(\d+)\s*[｜|]\s*解析失败\s*(\d+)", full_text)
    if not m:
        if tender_rows:
            warn("缺附件核对行（附件核对：应查 N｜已查 N｜解析失败 N）")
    else:
        should, done, failed = (int(x) for x in m.groups())
        if should < done:
            err("附件核对：应查 %d 少于已查 %d，数字对不上" % (should, done))
        if done != n_rows:
            err("附件清单实际 %d 行，核对行写「已查 %d」" % (n_rows, done))
        if failed:
            warn("附件核对写解析失败 %d 个，须在待确认区逐条写明原因与页码" % failed)
        elif should == done and n_rows:
            ok("附件清单 %d/%d 全部已查" % (done, should))


def check_matrix(text, path):
    secs = check_common(text, path, 7, "矩阵")
    by_name = {}
    for name, title, line_no, body in secs:
        by_name[CN.get(name, 0)] = (title, line_no, body)

    rows_map = {}
    url_bad = []          # (severity, line_no, rid, msg)
    url_rows = 0
    no_url_rows = 0
    empty_prod = []
    bad_value = []        # 值格式/禁用词
    bad_basis = []        # 字段与价格口径不对应
    drift = []            # 「产品名称/规格」列与 F04/F06 值不同源
    if 1 in by_name:
        title, line_no, body = by_name[1]
        if "主矩阵" not in title:
            warn("矩阵第一节标题不是「主矩阵」：%s" % title)
        tbls = tables(body, line_no)
        if not tbls:
            err("矩阵第一节没有表格")
        for start, head_n, rows, head in tbls:
            idx = {h: i for i, h in enumerate(head)}
            missing = [c for c in META_COLS if not any(c in h for h in head)]
            if missing:
                err("矩阵缺元数据列：%s" % "、".join(missing))
                if "来源网址" in missing and any("来源位置" in h for h in head):
                    err("检测到旧版「来源位置」列：应拆为「来源网址」+「页内位置」，"
                        "网址必须独立成列，否则用户无法回站复核")
            else:
                ok("矩阵七项元数据列齐全")
            if any("产品名称" in h for h in head) and any("规格" in h for h in head):
                ok("矩阵含产品名称与规格/型号列")
            else:
                err("矩阵缺产品标识列（产品名称 / 规格-型号）：只写企业名，读者无法判断是哪个产品的价")
            p_prod = next((i for i, h in enumerate(head) if "产品名称" in h), None)
            p_spec = next((i for i, h in enumerate(head) if "规格" in h), None)
            gi = next((i for h, i in idx.items() if "字段组" in h), None)
            ii = next((i for h, i in idx.items() if h.strip() in ("行号", "#")), 0)
            vi = next((i for h, i in idx.items() if h.strip() == "值"), None)
            ci = next((i for h, i in idx.items() if "置信度" in h), None)
            li = next((i for h, i in idx.items() if h.strip() in ("等级", "来源等级")), None)
            ui = next((i for h, i in idx.items() if "来源网址" in h), None)
            bi = next((i for h, i in idx.items() if "价格口径" in h), None)
            ni = next((i for h, i in idx.items() if h.strip() == "备注"), None)
            di = next((i for h, i in idx.items() if "获取日期" in h), None)
            pi = next((i for h, i in idx.items() if h.strip() == "竞品"), None)
            fi = next((i for h, i in idx.items() if h.strip() == "字段"), None)
            comp_seen = {}
            ent_vals = set()
            for no, n, line in rows:
                cs = cells(line)
                if len(cs) < head_n:
                    continue
                rid = cs[ii].strip()
                if not re.fullmatch(r"M\d{2,4}", rid):
                    err("矩阵第 %d 行号不合规（应 M01 起）：%s" % (no, rid[:20]))
                    continue
                if rid in rows_map:
                    err("矩阵行号重复：%s（第 %d 行）" % (rid, no))
                rows_map[rid] = {
                    "line": no,
                    "conf": cs[ci].strip() if ci is not None and ci < len(cs) else "",
                    "value": cs[vi].strip() if vi is not None and vi < len(cs) else "",
                    "note": cs[ni].strip() if ni is not None and ni < len(cs) else "",
                    "date": cs[di].strip() if di is not None and di < len(cs) else "",
                    "field": cs[fi].strip() if fi is not None and fi < len(cs) else "",
                    "comp": cs[pi].strip() if pi is not None and pi < len(cs) else "",
                    "spec": cs[p_spec].strip() if p_spec is not None and p_spec < len(cs) else "",
                    "prod": cs[p_prod].strip() if p_prod is not None and p_prod < len(cs) else "",
                }
                if ui is not None and ui < len(cs):
                    raw = cs[ui].strip()
                    url_rows += 1
                    sev, msg = url_problem(raw)
                    if msg:
                        url_bad.append((sev, no, rid, msg))
                    elif raw.startswith(NO_URL_PREFIX):
                        no_url_rows += 1
                if pi is not None and pi < len(cs):
                    cv = cs[pi].strip()
                    if cv and cv not in comp_seen:
                        comp_seen[cv] = (no, rid)
                if fi is not None and fi < len(cs) and vi is not None and vi < len(cs):
                    fname = cs[fi].strip()
                    if "企业全称" in fname or "批文持有人" in fname or "上市许可持有人" in fname:
                        fval = cs[vi].strip()
                        if fval:
                            ent_vals.add(fval)
                if gi is not None and gi < len(cs):
                    g = cs[gi].strip()
                    if g and not any(k in g for k in GROUP_KEYS):
                        err("矩阵第 %d 行字段组不在四组内：%s" % (no, g))
                if li is not None and li < len(cs):
                    g = cs[li].strip().upper()
                    if g.startswith("C"):
                        err("矩阵第 %d 行（%s）为 C 级来源，C 级不得进主矩阵，应移入线索区"
                            % (no, rid))
                    elif g and not any(g.startswith(x) for x in ENUM_GRADE):
                        err("矩阵第 %d 行（%s）来源等级「%s」不在 A/B/C 内"
                            % (no, rid, cs[li].strip()[:10]))
                # 枚举校验：约定外的取值视为错（含「应该是」这类口语）
                bad_enum = []
                for col_name, allowed in (("置信度", ENUM_CONF), ("价格口径", ENUM_PRICE_BASIS)):
                    hit = [h for h in head if col_name in h]
                    if not hit:
                        continue
                    j = head.index(hit[0])
                    if j >= len(cs):
                        continue
                    v = cs[j].strip()
                    if v and not any(a in v for a in allowed):
                        bad_enum.append("%s「%s」应取 %s" % (col_name, v[:12], "/".join(allowed)))
                if bad_enum:
                    err("矩阵第 %d 行（%s）枚举取值不合规：%s"
                        % (no, rid, "；".join(bad_enum)))

                # 值本身的格式：元数据齐了但值写歪，用户照样不能直接用
                if fi is not None and vi is not None and fi < len(cs) and vi < len(cs):
                    fname = cs[fi].strip()
                    val = cs[vi].strip()
                    if val and val not in EMPTY_VALUE:
                        is_price = next((k for k in PRICE_FIELDS if k in fname), None)
                        is_date = next((k for k in DATE_FIELDS if k in fname), None)
                        # 原文类字段逐字引述，官方原文写「约 3 个月」也不能算编造
                        literal = any(k in fname for k in ORIGIN_FIELDS)
                        bad_word = [] if literal else [w for w in BANNED_WORDS if w in val]
                        if bad_word:
                            bad_value.append((no, rid, fname, "含禁用词「%s」" % bad_word[0]))
                        elif is_price and not NUM_START.match(val):
                            bad_value.append((no, rid, fname, "价格不是数字开头（「%s」）" % val[:14]))
                        elif is_date and not re.match(r"^20\d{2}-\d{2}-\d{2}", val):
                            bad_value.append((no, rid, fname, "日期不是 YYYY-MM-DD（「%s」）" % val[:14]))
                        # 字段与口径必须对得上：把挂网价标成中选价，整张表的价差全废
                        if is_price and bi is not None and bi < len(cs):
                            basis = cs[bi].strip()
                            allow = PRICE_BASIS_BY_FIELD.get(is_price)
                            if allow and basis and basis not in ("", "—", "-"):
                                if not any(a in basis for a in allow):
                                    bad_basis.append((no, rid, fname, basis, "/".join(allow)))
                        # 产品名称/规格列与 F04/F06 的值同源：两处各写一套，改一处就漂
                        if p_prod is not None and p_prod < len(cs) and "产品通用名" in fname:
                            pv = cs[p_prod].strip()
                            if pv and norm_txt(pv) != norm_txt(val):
                                drift.append((no, rid, "产品名称列「%s」≠ F04 值「%s」"
                                              % (pv[:14], val[:14])))
                        if p_spec is not None and p_spec < len(cs) and "规格型号" in fname:
                            sv = cs[p_spec].strip()
                            if sv and norm_txt(sv) != norm_txt(val):
                                drift.append((no, rid, "规格/型号列「%s」≠ F06 值「%s」"
                                              % (sv[:14], val[:14])))
                empty = []
                for c in META_COLS:
                    hit = [h for h in head if c in h]
                    if not hit:
                        continue
                    j = head.index(hit[0])
                    if j < len(cs) and cs[j].strip() in ("", "—", "-"):
                        empty.append(c)
                if empty:
                    err("矩阵第 %d 行（%s）元数据为空：%s；应移入待人工确认区"
                        % (no, rid, "、".join(empty)))
                for col, j in (("产品名称", p_prod), ("规格/型号", p_spec)):
                    if j is not None and j < len(cs) and cs[j].strip() in ("", "—", "-"):
                        empty_prod.append((no, rid, col))
        # 「竞品」列：填竞品标识，不是企业全称。填错会让全表按公司聚合，产品信息被压掉
        if pi is None:
            warn("矩阵没有「竞品」列：这列写竞品标识（品牌名/清单简称），与「产品名称」分工不同")
        elif comp_seen:
            strong, weak = [], []
            for v, (no, _rid) in comp_seen.items():
                if v in ent_vals or any(k in v for k in COMPANY_HARD):
                    strong.append((no, v))
                elif COMPANY_SOFT.search(v):
                    weak.append((no, v))
            if strong:
                for no, v in strong[:3]:
                    err("矩阵第 %d 行「竞品」列填的是企业名「%s」：这列要填竞品标识（watchlist 里的品牌名/简称，如 保妥适），"
                        "企业全称属于 F01，只写在该行「值」列" % (no, v[:20]))
                if len(strong) > 3:
                    err("矩阵另有 %d 个「竞品」取值也是企业名，一并改成竞品标识" % (len(strong) - 3))
            else:
                ok("「竞品」列未混入企业全称，与「产品名称」分工正确")
            if weak:
                warn("「竞品」列这些值像公司简称而非产品品牌，确认下：%s"
                     % "、".join(v[:12] for _, v in weak[:4]))

        if empty_prod:
            for no, rid, col in empty_prod[:5]:
                err("矩阵第 %d 行（%s）%s为空：每行必须写清是哪个产品、哪个规格" % (no, rid, col))
            if len(empty_prod) > 5:
                err("矩阵另有 %d 行产品名称/规格为空" % (len(empty_prod) - 5))
        elif rows_map:
            ok("主矩阵每行产品名称与规格非空")

        if bad_value:
            for no, rid, fname, msg in bad_value[:5]:
                err("矩阵第 %d 行（%s）「%s」的值%s：只写原文数值，含糊词一律进待确认区"
                    % (no, rid, fname[:12], msg))
            if len(bad_value) > 5:
                err("矩阵另有 %d 处值格式不合规（价格须数字开头、日期须 YYYY-MM-DD、禁「约/大概/预计」）"
                    % (len(bad_value) - 5))
        elif rows_map:
            ok("主矩阵值格式与禁用词检查通过")

        if bad_basis:
            for no, rid, fname, basis, allow in bad_basis[:4]:
                err("矩阵第 %d 行（%s）「%s」的价格口径写成「%s」，应为 %s"
                    % (no, rid, fname[:12], basis[:10], allow))
            if len(bad_basis) > 4:
                err("矩阵另有 %d 处字段与价格口径不对应（口径填错，价差就会算错）"
                    % (len(bad_basis) - 4))
        elif rows_map:
            ok("字段与价格口径一一对应")
        if drift:
            for no, rid, msg in drift[:4]:
                warn("矩阵第 %d 行（%s）%s：两处要写同一个值，改一处就漂" % (no, rid, msg))
            if len(drift) > 4:
                warn("另有 %d 行产品名称/规格与 F04/F06 值不同源" % (len(drift) - 4))
        m_nums = sorted(int(k[1:]) for k in rows_map if k.startswith("M"))
        if m_nums:
            if m_nums == list(range(1, len(m_nums) + 1)):
                ok("矩阵行号连续唯一（M01–M%02d）" % len(m_nums))
            else:
                gaps = [i for i in range(1, max(m_nums) + 1) if i not in m_nums]
                if gaps:
                    err("矩阵行号不连续，缺：%s" % "、".join("M%02d" % g for g in gaps[:6]))
        if url_rows and not url_bad:
            ok("矩阵 %d 行来源网址均为具体页（可逐行回站复核）" % url_rows)
        hard = [b for b in url_bad if b[0] == "err"]
        soft = [b for b in url_bad if b[0] == "warn"]
        for _, no, rid, msg in hard[:5]:
            err("矩阵第 %d 行（%s）来源网址%s" % (no, rid, msg))
        if len(hard) > 5:
            err("矩阵另有 %d 行来源网址不合规：逐行必须给可直接打开的完整地址" % (len(hard) - 5))
        for _, no, rid, msg in soft[:3]:
            warn("矩阵第 %d 行（%s）来源网址%s" % (no, rid, msg))
        if len(soft) > 3:
            warn("另有 %d 行来源网址可疑" % (len(soft) - 3))
        if url_rows and no_url_rows / float(url_rows) > NO_URL_MAX_SHARE:
            warn("「无公开网址」占 %d/%d 行（超过 %d%%）：请确认是内部资料，不是漏填网址"
                 % (no_url_rows, url_rows, int(NO_URL_MAX_SHARE * 100)))
    else:
        err("矩阵缺第一节「主矩阵」")

    # 招采展开与附件清单是硬节，逐省逐规格逐附件都要落到行
    tender = find_sec(secs, "招采展开")
    projects, tender_rows, tender_price = (set(), 0, {})
    if not tender:
        err("矩阵缺「二、招采展开」区：集采与招投标必须逐规格、逐省落行")
    else:
        projects, tender_rows, tender_price = check_tender(tender[2], tender[1], rows_map, text)

    # 跨表对账：主矩阵 F24 与招采表一的同规格中选价必须是同一个数
    if tender_price:
        clash = []
        checked = 0
        for rid, r in rows_map.items():
            if not rid.startswith("M") or "集采中选价" not in r.get("field", ""):
                continue
            key = (r.get("comp", ""), r.get("prod", ""), r.get("spec", ""))
            hits = tender_price.get(key)
            if not hits:
                continue
            mv = num_of(r.get("value", ""))
            if mv is None:
                continue
            checked += 1
            if not any(abs(mv - tv) < 0.01 for tv, _ in hits):
                clash.append((rid, r["line"], r.get("value", "")[:16],
                              "、".join("%s 写 %s" % (trid, tv) for tv, trid in hits[:2])))
        for rid, no, v, tinfo in clash[:4]:
            err("主矩阵第 %d 行（%s）集采中选价「%s」与招采表一对不上（%s）：同一事实两处必须同数"
                % (no, rid, v, tinfo))
        if not clash and checked:
            ok("主矩阵与招采表的中选价对账一致（%d 条）" % checked)

    attach = find_sec(secs, "附件清单")
    if not attach:
        err("矩阵缺「三、附件清单」区：公告附件是名单与价格的唯一出处")
    else:
        check_attachments(attach[2], attach[1], projects, tender_rows, text, rows_map)

    confirm = find_sec(secs, "待人工确认")
    if not confirm:
        err("矩阵缺「待人工确认区」")
    else:
        ok("矩阵含待人工确认区")
        check_aux_links(confirm[2], confirm[1], "待人工确认区", 1)

    conflict = find_sec(secs, "冲突记录")
    if not conflict:
        warn("矩阵缺「冲突记录」区（无冲突可写「本期无冲突」）")
    else:
        check_aux_links(conflict[2], conflict[1], "冲突记录区", 2)

    cover = find_sec(secs, "字段覆盖")
    if not cover:
        err("矩阵缺「六、字段覆盖」区：36 行摊开、每家一列，用户才能看出缺哪格；"
            "只给长表等于把缺口藏起来")
    else:
        check_coverage(cover[2], cover[1], text, rows_map)

    # 矩阵必须是全量快照：新增 + 变更 + 沿用 = M 行。只写变化行会让下期 diff 断档
    m_total = len([k for k in rows_map if k.startswith("M")])
    sm = re.search(r"矩阵\s*M?\s*(\d+)\s*行（新增\s*(\d+)\s*[／/]\s*变更\s*(\d+)\s*[／/]\s*沿用\s*(\d+)）", text)
    if sm:
        total, add, chg, keep = (int(x) for x in sm.groups())
        n_keep = len([1 for k, v in rows_map.items() if k.startswith("M")
                      and v.get("note", "").startswith("沿用")])
        if total != m_total:
            err("统计行写「矩阵 %d 行」，实际 M 行 %d 条" % (total, m_total))
        if add + chg + keep != total:
            err("统计行「新增 %d + 变更 %d + 沿用 %d = %d」≠ 矩阵 %d 行"
                % (add, chg, keep, add + chg + keep, total))
        if keep != n_keep:
            err("统计行写「沿用 %d」条，主矩阵备注以「沿用」开头的只有 %d 条——"
                "沿用行也要落表，不能只计数" % (keep, n_keep))
        if keep == 0 and total > 0 and "首期" not in text and "建档" not in text:
            warn("本期没有一条沿用行：若非首期建档，说明矩阵只写了变化行，下期无法当基线")
        if add + chg + keep == total and keep == n_keep:
            ok("矩阵为全量快照（新增 %d／变更 %d／沿用 %d）" % (add, chg, keep))
    elif m_total:
        warn("统计行未写「矩阵 M 行（新增 a／变更 b／沿用 p）」：无法判断矩阵是全量还是只写了变化行")

    # 统计行的「待确认 g 行」「冲突 h 处」必须与第四、五节的数据行数对得上
    for pat, label, sec in ((r"待确认\s*(\d+)\s*行", "待确认", confirm),
                            (r"冲突\s*(\d+)\s*处", "冲突", conflict)):
        hit = re.search(pat, text)
        if not hit or not sec:
            continue
        got = data_rows(sec[2], sec[1])
        if int(hit.group(1)) != got:
            err("统计行写「%s %s」，%s区实际 %d 行——统计数要与明细一致"
                % (label, hit.group(1), label, got))
        else:
            ok("统计行「%s %d」与明细一致" % (label, got))

    # 时效性：拿三个月前的挂网价去定价，比缺一格更危险
    today = datetime.date.today()
    stale = []
    for rid, r in rows_map.items():
        if not rid.startswith("M") or r.get("note", "").startswith("沿用"):
            continue
        f, d = r.get("field", ""), r.get("date", "")
        if not f or not d:
            continue
        if any(k in f for k in ("价", "挂网状态")):
            limit = 30
        elif any(k in f for k in ("注册证", "有效期", "医保", "目录", "适用范围")):
            limit = 90
        else:
            continue
        dm = re.match(r"(20\d{2})-(\d{2})-(\d{2})", d)
        if not dm:
            continue
        try:
            dt = datetime.date(int(dm.group(1)), int(dm.group(2)), int(dm.group(3)))
        except ValueError:
            continue
        age = (today - dt).days
        if age > limit:
            stale.append((rid, f, d, age, limit))
    if stale:
        for rid, f, d, age, limit in stale[:4]:
            warn("矩阵 %s「%s」获取于 %s，已 %d 天未复核（阈值 %d 天）：备注写「需复核（超 N 天）」或重查"
                 % (rid, f[:10], d, age, limit))
        if len(stale) > 4:
            warn("另有 %d 条价格/证照类数据超期未复核" % (len(stale) - 4))
    elif rows_map:
        ok("价格与证照类数据均在时效内")

    # 字段按预设取用后，必须自报本期覆盖范围，否则「按需选字段」不可追溯
    if re.search(r"覆盖字段\s*\d+\s*/\s*36", text):
        ok("矩阵统计行标注了本期字段覆盖范围")
    elif re.search(r"字段[：:]\s*预设", text):
        warn("矩阵只写了预设名，未在统计行写「覆盖字段 n/36」，无法判断实际填了多少")
    else:
        warn("矩阵未标注字段覆盖范围（建议统计行写「覆盖字段 n/36（预设 price）」）")

    prose = []
    fenced = False
    for no, line in enumerate(text.split("\n"), 1):
        s = line.strip()
        if s.startswith("```"):
            fenced = not fenced
            continue
        if fenced or not s:
            continue
        if s[0] in "#|-*>":
            continue
        # 统计行用全角竖线分隔字段，属于台账元信息，不算叙述段落
        if "｜" in s or "|" in s:
            continue
        if len(s) > 50 or (len(s) > 30 and s.endswith("。")):
            prose.append((no, s[:50]))
    if prose:
        err("矩阵出现 %d 处叙述段落（阶段一只许出矩阵，禁止段落与结论）：第 %d 行「%s…」"
            % (len(prose), prose[0][0], prose[0][1]))
    else:
        ok("矩阵无叙述段落")
    return rows_map


def check_price_compare(text):
    """文档价格对照：价差 = 竞品价 − 我方价；价差率的分母是我方价。

    这是用户拿去做定价决策的那张表，算错一次就误导一次，所以必须机械复算。
    """
    found = False
    checked = 0
    bad = []
    for start, head_n, rows, head in tables(text, 1):
        i_comp = next((i for i, h in enumerate(head) if "竞品价" in h), None)
        i_our = next((i for i, h in enumerate(head) if "我方" in h), None)
        i_gap = next((i for i, h in enumerate(head) if "价差" in h and "率" not in h), None)
        i_rate = next((i for i, h in enumerate(head) if "价差率" in h), None)
        if i_comp is None or i_our is None or i_gap is None:
            continue
        found = True
        for no, n, line in rows:
            cs = cells(line)
            if len(cs) < head_n or all(c in ("", "—", "-", "－") for c in cs):
                continue
            our_v = cs[i_our].strip() if i_our < len(cs) else ""
            if "未提供" in our_v:
                continue
            our = num_of(our_v)
            comp = num_of(cs[i_comp].strip()) if i_comp < len(cs) else None
            if our is None or comp is None or our == 0:
                continue
            gap = num_of(cs[i_gap].strip()) if i_gap < len(cs) else None
            if gap is None:
                bad.append((no, "价差没写数（算不出或留空），用户没法判断谁更贵"))
                continue
            checked += 1
            if abs(gap - (comp - our)) > 0.02:
                bad.append((no, "价差写 %g，竞品价 %g − 我方价 %g 应为 %g"
                            % (gap, comp, our, comp - our)))
                continue
            if i_rate is not None and i_rate < len(cs):
                raw = cs[i_rate].strip()
                rate = num_of(raw)
                if rate is None:
                    continue
                if "%" not in raw and abs(rate) < 1.5:
                    rate *= 100
                expect = (comp - our) / our * 100
                if abs(rate - expect) > 0.15:
                    bad.append((no, "价差率写 %g%%，按（竞品价−我方价）/我方价 应为 %.1f%%"
                                "——分母写成竞品价，结论会反" % (rate, expect)))
    if bad:
        for no, msg in bad[:5]:
            err("文档价差对照第 %d 行：%s" % (no, msg))
        if len(bad) > 5:
            err("价差对照另有 %d 行算错或缺失" % (len(bad) - 5))
    elif checked:
        ok("价格对照 %d 行价差与价差率复算一致" % checked)
    elif "价格对照" in text:
        warn("文档有「价格对照」节但没找到可复算的行（缺竞品价/我方价/价差列，或全是「未提供」）")


def check_doc(text, path, matrix_rows):
    secs = check_common(text, path, 8, "文档")
    by_name = {CN.get(n, 0): (t, l, b) for n, t, l, b in secs}

    cited = set((p, int(n)) for p, n in CITE_RE.findall(text))
    if cited:
        ok("文档引用矩阵行号 %d 个" % len(cited))
    else:
        warn("文档未出现任何 [M行号] / [T行号] 引用")

    if matrix_rows:
        missing = []
        for p, i in sorted(cited):
            key = "%s%02d" % (p, i)
            if key not in matrix_rows:
                missing.append(key)
        if missing:
            err("文档引用了矩阵中不存在的行号：%s" % "、".join(missing[:8]))
        else:
            ok("引用行号均存在于矩阵")

        bad_lines = []
        for no, line in enumerate(text.split("\n"), 1):
            if "待核实线索" in line:
                continue
            for m in CITE_RE.finditer(line):
                key = "%s%02d" % (m.group(1), int(m.group(2)))
                row = matrix_rows.get(key)
                if row and "待确认" in row["conf"]:
                    bad_lines.append((no, key))
        if bad_lines:
            for no, key in bad_lines[:5]:
                err("第 %d 行把待确认行 %s 写成了结论（须写「待核实线索」）" % (no, key))
        else:
            ok("待确认行未被写成结论")

    change_rows = 0
    if 2 in by_name:
        title, line_no, body = by_name[2]
        for start, head_n, rows, head in tables(body, line_no):
            for no, n, line in rows:
                if all(c in ("", "—", "-", "－") for c in cells(line)):
                    continue
                change_rows += 1
                if not any(lv in line for lv in LEVELS):
                    warn("文档第 %d 行变化未标级别：%s" % (no, line.strip()[:46]))
                if not CITE_RE.search(line):
                    err("文档第 %d 行变化缺 [M行号] 引用：%s" % (no, line.strip()[:46]))
                if not LINK_RE.search(line) and "依据行号" in "".join(head):
                    pass
        if change_rows:
            ok("变化清单 %d 行" % change_rows)
        else:
            ok("变化清单为空")
            if "无变化" not in body and "无新增" not in body:
                err("第二节既无变化行也未写「本期无变化」")

    # 发生日期列不得晚于当日
    if 2 in by_name:
        today = datetime.date.today()
        title, line_no, body = by_name[2]
        future = []
        checked = 0
        for start, head_n, rows, head in tables(body, line_no):
            idx = next((i for i, h in enumerate(head) if "发生" in h), None)
            if idx is None:
                continue
            for no, n, line in rows:
                cs = cells(line)
                if idx >= len(cs) or all(c in ("", "—", "-") for c in cs):
                    continue
                checked += 1
                for y, sep, m1, d1 in DATE_RE.findall(cs[idx]):
                    try:
                        d = datetime.date(int(y), int(m1), int(d1))
                    except ValueError:
                        err("文档第 %d 行发生日期非法：%s-%s-%s" % (no, y, m1, d1))
                        continue
                    if d > today:
                        future.append((no, d.isoformat()))
        if future:
            for no, d in future[:5]:
                err("文档第 %d 行变化发生日 %s 晚于今天" % (no, d))
        elif checked:
            ok("发生日期列 %d 条均不晚于当日" % checked)

    # 第三节「新发现对手」是硬节，名单里识别出的同组企业必须有落脚处
    if 3 in by_name:
        title, line_no, body = by_name[3]
        if "新发现对手" not in title:
            warn("文档第三节标题不是「新发现对手」：%s" % title)
        tbls = tables(body, line_no)
        newcomer_rows = 0
        if not tbls:
            if "无新发现对手" not in body:
                err("文档第三节没有表，也没写「本期无新发现对手（名单已核 N 份）」")
        for start, head_n, rows, head in tbls:
            i_ent = next((i for i, h in enumerate(head) if "企业名称" in h), None)
            i_prod = next((i for i, h in enumerate(head) if "产品名称" in h), None)
            i_spec = next((i for i, h in enumerate(head) if "规格" in h), None)
            i_cite = next((i for i, h in enumerate(head) if "依据" in h), None)
            if i_ent is None:
                err("文档新发现对手表缺「企业名称」列：用户要能认出这是哪家")
            if i_prod is None or i_spec is None:
                err("文档新发现对手表缺产品名称或规格列：只写公司名等于没写")
            for no, n, line in rows:
                cs = cells(line)
                if all(c in ("", "—", "-") for c in cs):
                    continue
                newcomer_rows += 1
                if i_cite is not None:
                    v = cs[i_cite].strip() if i_cite < len(cs) else ""
                    if not re.search(r"\[N\d+\]", v):
                        err("文档第 %d 行新对手缺 [N行号] 引用：%s" % (no, line.strip()[:40]))
                for j, col in ((i_ent, "企业名称"), (i_prod, "产品名称"), (i_spec, "规格/型号")):
                    if j is not None and (j >= len(cs) or cs[j].strip() in ("", "—", "-")):
                        err("文档第 %d 行新对手%s为空" % (no, col))
        if newcomer_rows:
            ok("文档新发现对手 %d 行" % newcomer_rows)
        n_in_matrix = len([k for k in matrix_rows if k.startswith("N")]) if matrix_rows else 0
        if matrix_rows and newcomer_rows != n_in_matrix:
            err("文档新发现对手 %d 行，矩阵 N 行 %d 条——两处必须同源" % (newcomer_rows, n_in_matrix))
    else:
        err("文档缺第三节「新发现对手」：名单里识别出的同组企业必须有落脚处，一家都不能只在附件里")

    # 文件名产品段可能含全角括号与全角斜杠；遇到全角竖线、感叹号、空白即止，
    # 否则「矩阵 A.md（M 行）｜基线 B.md」会被整段当成一个文件名
    refs = re.findall(r"竞品矩阵-\d{4}-\d{2}-\d{2}_\d{4}(?:-[^\\/:*?\"<>|｜（）\s]+)?\.md", text)
    folder = os.path.dirname(os.path.abspath(path))
    base = os.path.basename(path)
    missing_ref = [r for r in set(refs) if not os.path.exists(os.path.join(folder, r))]
    if missing_ref:
        warn("引用了不存在的矩阵基线：%s" % "、".join(missing_ref))
    elif refs:
        ok("矩阵基线引用可对上（%d 处）" % len(set(refs)))

    check_price_compare(text)
    return change_rows


def html_text(h):
    h = re.sub(r"<script.*?</script>|<style.*?</style>", " ", h, flags=re.S | re.I)
    h = re.sub(r"<[^>]+>", " ", h)
    for a, b in (("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"')):
        h = h.replace(a, b)
    return re.sub(r"\s+", " ", h)


def check_html(html_path, doc_text, matrix_text):
    h = read(html_path)
    ph = PH_RE.findall(h)
    if ph:
        err("HTML 残留占位符：%s" % "、".join(sorted(set(ph))[:6]))
    else:
        ok("HTML 无占位符残留")
    for mark in ("FILL GUIDE", "示意数据", "示例数据", "编的", "假数据"):
        if mark in h:
            err("HTML 含样例标记「%s」" % mark)
    for blk in ("需要动作", "持续关注", "政策日历", "执行中招采项目"):
        if blk not in h:
            warn("HTML 缺少「%s」区块" % blk)

    body_text = html_text(h)
    source = doc_text + "\n" + matrix_text
    hl = set(re.findall(r'href="(https?://[^"]+)"', h)) | set(LINK_RE.findall(body_text))
    src_links = set(u.rstrip("/") for u in LINK_RE.findall(source))
    extra = [u for u in hl if u.rstrip("/") not in src_links]
    if extra:
        warn("HTML 有 %d 条链接未在文档或矩阵出现：%s" % (len(extra), "、".join(extra[:3])))
    else:
        ok("HTML 链接均可回溯")

    nums = set(re.findall(r"\d{3,}(?:\.\d+)?%?", body_text)) | set(re.findall(r"\d+\.\d+%", body_text))
    src_nums = set(re.findall(r"\d{3,}(?:\.\d+)?%?", source)) | set(re.findall(r"\d+\.\d+%", source))
    miss = sorted(n for n in nums if n not in src_nums)
    if miss:
        warn("HTML 有 %d 个数字未在文档或矩阵出现：%s" % (len(miss), "、".join(miss[:8])))
    else:
        ok("HTML 关键数字均可回溯")


def check_xlsx(xlsx_path, change_rows):
    try:
        z = zipfile.ZipFile(xlsx_path)
    except Exception as e:
        err("Excel 无法打开：%s" % e)
        return
    try:
        wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
        names = re.findall(r'<sheet[^>]*name="([^"]+)"', wb)
        is_matrix = os.path.basename(xlsx_path).startswith(MATRIX_PREFIX)
        need = ("字段矩阵", "招采展开", "附件清单", "名单核对", "新发现对手", "待人工确认", "冲突记录", "字段覆盖", "字段字典") if is_matrix else \
               ("本期变化", "新发现对手", "招采项目", "省份执行", "价格对照", "附件清单", "省份覆盖", "竞品档案", "政策日历", "扫描覆盖")
        label = "矩阵 Excel" if is_matrix else "文档 Excel"
        ok("%s 可打开，%d sheet：%s" % (label, len(names), "、".join(names)))
        miss = [n for n in need if n not in names]
        if miss:
            warn("%s 缺少 sheet：%s" % (label, "、".join(miss)))
        for part in z.namelist():
            if part.endswith(".xml"):
                try:
                    ET.fromstring(z.read(part))
                except Exception:
                    err("%s 内 %s 不是合法 XML" % (label, part))
                    break
        else:
            ok("%s 内部 XML 合法" % label)

        styles = z.read("xl/styles.xml").decode("utf-8", "replace") \
            if "xl/styles.xml" in z.namelist() else ""
        if is_matrix:
            if any(c in styles for c in ("E8F0FA", "EEEAFA", "E1F2F1", "FBEFF3")):
                ok("矩阵 Excel 已给重点竞品整行上色（多家各一色），几十行里一眼看出先看哪几家、哪行是谁")
            else:
                warn("矩阵 Excel 没有重点竞品行底色：生成 spec 时用 highlight.companies 点名，否则全表一片白")
        else:
            tint = [c for c in ("FADBD8", "FCECC5", "DFF0E4") if c in styles]
            if tint:
                ok("文档 Excel 变化清单已按级别上色（红/黄/绿命中 %d 种）" % len(tint))
            elif "本期变化" in names:
                warn("文档 Excel 未按级别上色：「级别」列写 红/黄/绿 会自动带底色，检查列名是否被改过")
        body = "".join(z.read(p).decode("utf-8", "replace") for p in z.namelist() if p.endswith(".xml"))
        if "{{" in body or "TODO" in body:
            err("%s 单元格内有占位符或 TODO 残留" % label)
        if is_matrix:
            s1 = z.read("xl/worksheets/sheet1.xml").decode("utf-8", "replace")
            n_links = len(re.findall(r"<hyperlink ", s1))
            rels_part = "xl/worksheets/_rels/sheet1.xml.rels"
            if n_links:
                ok("矩阵 Excel 首表含 %d 个可点击来源链接" % n_links)
                if rels_part in z.namelist():
                    try:
                        ET.fromstring(z.read(rels_part))
                        ok("矩阵 Excel 链接关系文件合法")
                    except Exception:
                        err("矩阵 Excel 链接关系文件不是合法 XML，链接点不开")
                else:
                    err("矩阵 Excel 有 hyperlink 元素但缺 _rels 文件，链接点不开")
            else:
                warn("矩阵 Excel 首表未发现可点击链接：来源网址列要给完整 http(s) 地址")
        if not is_matrix:
            rows = len(re.findall(r"<row ", z.read("xl/worksheets/sheet1.xml").decode("utf-8", "replace")))
            data_rows = max(rows - 1, 0)
            if change_rows and data_rows != change_rows:
                warn("文档 Excel 首表 %d 条，MD 变化清单 %d 条，请确认同源" % (data_rows, change_rows))
            elif change_rows:
                ok("文档 Excel 首表条数与 MD 一致（%d）" % data_rows)
    finally:
        z.close()


def check_links_online(texts, limit=30, timeout=8, insecure=False):
    """真实打开产物里的链接，验证「可访问」不是自己声明的。

    只报实际返回的状态，不武断判定失效：
      2xx/3xx        可达
      401/403        被拦（可能需登录或有反爬，不等于失效）
      404/410        失效
      证书异常        需人工（企业网络 HTTPS 解密代理、站点证书过期或自签名）
      其他/超时       连不上（可能是本机网络或站点限制）

    证书默认走系统信任库校验。关掉校验会把「中间人返回的拦截页」当成 200 可达，
    那是假通过——比自己声明「可访问」更危险，所以 --insecure 必须显式指定。
    """
    import urllib.request
    import urllib.error
    import ssl

    urls = []
    for t in texts:
        for m in LINK_RE.finditer(t or ""):
            u = m.group(0).rstrip(".,;")
            if u not in urls:
                urls.append(u)
    if not urls:
        warn("产物中没有可校验的链接")
        return
    urls = urls[:limit]

    ctx = ssl.create_default_context()
    if insecure:
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36",
        "Accept-Language": "zh-CN,zh;q=0.9",
    }

    reach, blocked, dead, fail, certbad = [], [], [], [], []
    for u in urls:
        code = None
        ssl_err = None
        for method in ("HEAD", "GET"):
            try:
                req = urllib.request.Request(u, method=method, headers=headers)
                with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                    code = r.status
                break
            except urllib.error.HTTPError as e:
                code = e.code
                if code == 403 and method == "HEAD":
                    continue
                break
            except urllib.error.URLError as e:
                reason = getattr(e, "reason", None)
                if isinstance(reason, ssl.SSLError) or isinstance(reason, ssl.CertificateError):
                    ssl_err = "%s" % reason
                    break
                if method == "GET":
                    code = None
            except (ssl.SSLError, ssl.CertificateError) as e:
                ssl_err = "%s" % e
                break
            except Exception:
                if method == "GET":
                    code = None
        if ssl_err:
            certbad.append((ssl_err, u))
            continue
        if code is None:
            fail.append((None, u))
        elif 200 <= code < 400:
            reach.append((code, u))
        elif code in (401, 403):
            blocked.append((code, u))
        elif code in (404, 410):
            dead.append((code, u))
        else:
            fail.append((code, u))

    print("\n---- 链接实地校验（抽查 %d 条）----" % len(urls))
    if reach:
        for code, u in reach:
            print("  [%s] %s" % (code, u))
        ok("链接可达 %d 条（实开返回 2xx/3xx）" % len(reach))
    if blocked:
        for code, u in blocked:
            print("  [%s 被拦] %s" % (code, u))
        warn("链接被拦 %d 条（401/403，多为需登录或有反爬，**不等于失效**，需人工确认）：%s"
             % (len(blocked), "；".join(u for _, u in blocked[:3])))
    if certbad:
        for msg, u in certbad:
            print("  [证书异常] %s\n              %s" % (u, msg[:90]))
        warn("TLS/证书异常 %d 条（多为企业网络 HTTPS 解密代理、站点证书过期或自签名；"
             "**不判失效**，须人工点开，且不得据此认定页面内容可信）：%s"
             % (len(certbad), "；".join(u for _, u in certbad[:3])))
    if dead:
        for code, u in dead:
            print("  [%s] %s" % (code, u))
        err("链接 404/410 共 %d 条（应视为失效，改给栏目路径+检索词）：%s"
            % (len(dead), "；".join(u for _, u in dead[:3])))
    if fail:
        for code, u in fail:
            print("  [连不上 %s] %s" % (code if code else "无响应", u))
        warn("链接连不上 %d 条（可能是本机网络或站点限制，非必然失效，需人工确认）：%s"
             % (len(fail), "；".join(u for _, u in fail[:3])))
    if insecure:
        warn("本次以 --insecure 运行，**证书未校验**：拦截页可能被当成 200 可达，结果仅供排查，不得写进交付说明")
    print("  提示：403/超时/证书异常在本机与用户环境可能不同，最终以人工点开为准；"
          "但 404/410 基本可判失效。")


def report():
    """统一打印结论并给出退出码。"""
    print("\n==== 产物复检 ====")
    print("通过 %d 项" % len(PASSED))
    for p in PASSED:
        print("  [ok] %s" % p)
    if WARNS:
        print("\n告警 %d 条（需人工判断）" % len(WARNS))
        for w in WARNS:
            print("  [!] %s" % w)
    if ERRORS:
        print("\n错误 %d 条（必须修正并重新生成）" % len(ERRORS))
        for e in ERRORS:
            print("  [x] %s" % e)

    print("\n脚本判断不了、必须人工核对的 5 项：")
    print("  1. 每个值是否真取自所标来源，而非从同类产品反推")
    print("  2. 价格口径是否被写错：挂网价(上限) / 集采中选价(强制结算) / 中标价(成交)")
    print("  3. 绝对化措辞（首次、唯一、全部、均、领先）是否有两处独立来源")
    print("  4. 主体是否错位：批文持有人 / 生产企业 / 代理经销商 是否写成同一方")
    print("  5. 降幅与价差是否复算过：分子分母同规格同含税口径")

    if ERRORS:
        print("\n复检计数：通过 %d 项｜告警 %d 条｜错误 %d 条" % (len(PASSED), len(WARNS), len(ERRORS)))
        print("结论：FAIL —— 修正后重新生成并再次运行本脚本")
        return 1
    if WARNS:
        print("\n复检计数：通过 %d 项｜告警 %d 条｜错误 0 条" % (len(PASSED), len(WARNS)))
        print("结论：PASS（含告警）—— 告警项逐条确认后方可交付")
        return 0
    print("\n复检计数：通过 %d 项｜告警 0 条｜错误 0 条" % len(PASSED))
    print("结论：PASS")
    return 0


def main():
    args = [a for a in sys.argv[1:]]
    matrix_arg = None
    online = "--check-links" in args
    if online:
        args.remove("--check-links")
    insecure = "--insecure" in args
    if insecure:
        args.remove("--insecure")
    if "--matrix" in args:
        i = args.index("--matrix")
        if i + 1 < len(args):
            matrix_arg = args[i + 1]
            del args[i:i + 2]
    if not args:
        print("用法：python verify_output.py <文档.md 或 矩阵.md> [--matrix <矩阵.md>] "
              "[--check-links [--insecure]]")
        print("  --check-links  逐个实测产物里的链接（最多 30 条）")
        print("  --insecure     跳过证书校验（仅排查用；企业解密代理下会把拦截页误判成 200）")
        return 2
    doc = args[0]
    if not os.path.exists(doc):
        print("找不到文件：%s" % doc)
        return 2

    # 直接传矩阵文件时只查矩阵，不再要求配套文档（否则会误报「章节不完整」）
    if os.path.basename(doc).startswith(MATRIX_PREFIX):
        mtext = read(doc)
        check_matrix(mtext, doc)
        mx = os.path.splitext(doc)[0] + ".xlsx"
        if os.path.exists(mx):
            check_xlsx(mx, 0)
        else:
            warn("未找到矩阵 Excel：%s" % os.path.basename(mx))
        if online:
            check_links_online([mtext], insecure=insecure)
        return report()

    folder = os.path.dirname(os.path.abspath(doc))
    stem = os.path.splitext(os.path.basename(doc))[0]
    ts = stem.replace(DOC_PREFIX, "")

    if not matrix_arg:
        cand = os.path.join(folder, MATRIX_PREFIX + ts + ".md")
        if os.path.exists(cand):
            matrix_arg = cand
        else:
            others = sorted(f for f in os.listdir(folder) if f.startswith(MATRIX_PREFIX) and f.endswith(".md"))
            if others:
                matrix_arg = os.path.join(folder, others[-1])
                warn("未找到同时间戳矩阵，改用最近一期：%s" % others[-1])

    matrix_rows = {}
    if matrix_arg and os.path.exists(matrix_arg):
        mtext = read(matrix_arg)
        matrix_rows = check_matrix(mtext, matrix_arg)
    else:
        err("未找到矩阵文件——「先矩阵、后文档」是硬约束，缺矩阵的文档不予通过")
        mtext = ""

    doc_text = read(doc)
    change_rows = check_doc(doc_text, doc, matrix_rows)

    html_path = os.path.join(folder, DOC_PREFIX + ts + ".html")
    if os.path.exists(html_path):
        check_html(html_path, doc_text, mtext)
    else:
        warn("未找到同名 HTML：%s" % os.path.basename(html_path))

    if matrix_arg and os.path.exists(matrix_arg):
        mx = os.path.join(os.path.dirname(os.path.abspath(matrix_arg)),
                          os.path.splitext(os.path.basename(matrix_arg))[0] + ".xlsx")
        if os.path.exists(mx):
            check_xlsx(mx, 0)
        else:
            warn("未找到矩阵 Excel：%s" % os.path.basename(mx))

    dx = os.path.join(folder, stem + ".xlsx")
    if os.path.exists(dx):
        check_xlsx(dx, change_rows)
    else:
        warn("未找到文档 Excel：%s" % os.path.basename(dx))

    if online:
        html_txt = read(html_path) if os.path.exists(html_path) else ""
        check_links_online([mtext, doc_text, html_txt], insecure=insecure)

    return report()


if __name__ == "__main__":
    sys.exit(main())
