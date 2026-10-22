#!/usr/bin/env bash
# ============================================================================
# wps_pptx_repair.sh —— WPS 演示文稿（PPTX）跨系统文字错位检测与修复【交互式】
#
# 适用场景：Windows 上制作的 PPTX / POTX / PPSX 拷到统信 UOS、麒麟 Kylin
# 等国产 Linux 上，用 WPS 演示打开后出现：文字跑位、溢出色块、行距变大、
# 断行奇怪、表格被撑开、标题下移、符号变字母/方框……
#
# 本文件是「自包含 shell 脚本」：完整实现内嵌在本文件中，运行时释放到临时
# 文件后交给系统自带的 python3 执行。统信 UOS、麒麟、Ubuntu 等发行版都默认
# 预装 python3，且本工具不依赖任何 pip 第三方包，因此可以只拷这一个 .sh 文件
# 到目标机器直接使用（无需联网、无需安装依赖）：
#
#   chmod +x wps_pptx_repair.sh
#   ./wps_pptx_repair.sh                     # 交互式向导（推荐）
#   ./wps_pptx_repair.sh 课件.pptx           # 交互式，预载文件
#   ./wps_pptx_repair.sh 课件.pptx --scan    # 只体检，不改文件
#   ./wps_pptx_repair.sh 课件.pptx --fix -y  # 直接修复
#
# 也支持：sh wps_pptx_repair.sh --help
#
# 如果想指定解释器：WPS_REPAIR_PYTHON=/usr/bin/python3 ./wps_pptx_repair.sh
# ============================================================================
set -u

PY="${WPS_REPAIR_PYTHON:-}"
if [ -z "$PY" ]; then
  for _c in python3 python; do
    if command -v "$_c" >/dev/null 2>&1; then PY=$(command -v "$_c"); break; fi
  done
fi
if [ -z "$PY" ]; then
  echo "[x] 没有找到 python3。本工具需要系统自带的 python3（UOS/麒麟 默认已安装）。" >&2
  echo "    安装方式：sudo apt install python3   或   sudo dnf install python3" >&2
  exit 1
fi

TMP_PY=""
cleanup() {
  if [ -n "${TMP_PY:-}" ] && [ -f "$TMP_PY" ]; then rm -f "$TMP_PY"; fi
}
trap cleanup EXIT HUP INT TERM

TMP_PY=$(mktemp 2>/dev/null) || TMP_PY=$(mktemp -t wps_repair) || {
  echo "[x] 无法创建临时文件（检查 /tmp 是否可写）" >&2
  exit 1
}

cat > "$TMP_PY" <<'__WPS_PPTX_REPAIR_PY__'
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wps_pptx_repair.py —— WPS 演示文稿（PPTX）跨系统文字错位检测与修复工具

适用场景：Windows 上制作的 .pptx / .potx / .ppsx 迁移到统信 UOS、麒麟（Kylin）
等国产 Linux 发行版后，用 WPS 演示打开出现：
  * 文字溢出色块、行距变大、段落被挤到下一页或互相重叠
  * 中英文断行位置改变、标点跑到行首、文字整体上下/左右偏移
  * 符号（✓、→、项目符号）变成字母或方框
  * 表格文字换行、单元格被撑高

本工具做两件事：
  1. 体检：逐页逐形状定位可能导致错位的“隐患”，并给出量化的溢出估算；
  2. 修复：在不破坏原文件的前提下，生成修复版文件 + 备份 + 报告。

设计原则：
  * 纯标准库实现（麒麟/UOS 自带 python3 即可运行，无需 pip 安装任何依赖）；
  * 只改必要的 OOXML 属性，不重排、不“美化”版式，尽量保持原设计；
  * 默认只做“安全修复”，会改变视觉效果的修复单独成组、需要用户明确选择；
  * 任何写操作前自动备份，报告里记录改了什么，可随时回滚。

用法：
    python3 wps_pptx_repair.py                     # 交互式向导
    python3 wps_pptx_repair.py a.pptx              # 交互式处理指定文件
    python3 wps_pptx_repair.py a.pptx --scan       # 只体检，不改文件
    python3 wps_pptx_repair.py a.pptx --fix -y     # 非交互修复（推荐组）
    python3 wps_pptx_repair.py --help              # 查看全部参数
"""

import argparse
import bisect
import datetime
import json
import math
import os
import posixpath
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import textwrap
import zipfile
import xml.etree.ElementTree as ET

VERSION = "1.1.0"
TOOL_NAME = "WPS 演示文稿跨平台错位修复工具"

# ----------------------------------------------------------------------------
# 常量
# ----------------------------------------------------------------------------

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "a15": "http://schemas.microsoft.com/office/drawing/2012/main",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
for _p, _u in NS.items():
    try:
        ET.register_namespace(_p, _u)
    except ValueError:  # pragma: no cover
        pass


def qn(prefix, tag):
    """'a', 'rPr' -> '{a-namespace}rPr'"""
    return "{%s}%s" % (NS[prefix], tag)


def localname(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


EMU_PER_PT = 12700
EMU_PER_CM = 360000
DEFAULT_INSET_L = 91440
DEFAULT_INSET_R = 91440
DEFAULT_INSET_T = 45720
DEFAULT_INSET_B = 45720
DEFAULT_TABLE_CELL_MAR = (91440, 91440, 45720, 45720)

# OOXML 里允许出现的子元素顺序（用于安全插入，满足 schema 顺序要求）
RPR_ORDER = [
    "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill",
    "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
]
PPR_ORDER = [
    "lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct",
    "buSzPts", "buFontTx", "buFont", "buNone", "buAutoNum", "buChar",
    "buBlip", "tabLst", "defRPr", "extLst",
]
BODYPR_ORDER = [
    "prstTxWarp", "noAutofit", "normAutofit", "spAutoFit", "scene3d", "sp3d",
    "flatTx", "extLst",
]
TC_ORDER = ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill",
            "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "headers",
            "extLst"]

COLOR_ENABLED = (
    sys.stdout.isatty()
    and os.environ.get("NO_COLOR") is None
    and os.environ.get("TERM", "") not in ("dumb", "")
)


class C:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    RED = "\033[31m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    BLUE = "\033[34m"
    MAGENTA = "\033[35m"
    CYAN = "\033[36m"


def _c(text, color):
    if not COLOR_ENABLED or not color:
        return text
    return "%s%s%s" % (color, text, C.RESET)


def say(msg=""):
    print(msg)


def ok(msg):
    print("%s %s" % (_c("[OK]", C.GREEN), msg))


def warn(msg):
    print("%s %s" % (_c("[!]", C.YELLOW), msg))


def err(msg):
    print("%s %s" % (_c("[x]", C.RED), msg))


def info(msg):
    print("%s %s" % (_c("[i]", C.CYAN), msg))


def step(msg):
    print(_c("==> " + msg, C.BOLD))


def hr(char="-", width=78):
    print(_c(char * width, C.DIM))


def _display_width(text):
    """中文/全角字符按 2 列计算的显示宽度。"""
    width = 0
    for ch in text:
        width += 2 if ord(ch) > 0x2E7F and not ch.isascii() else 1
    return width


def pad_display(text, width):
    """按终端显示宽度右侧补空格。"""
    pad = max(0, width - _display_width(text))
    return text + " " * pad


def human_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return "%.1f %s" % (n, unit) if unit != "B" else "%d B" % n
        n /= 1024.0


def emu_to_pt(emu):
    return float(emu) / EMU_PER_PT


def emu_to_cm(emu):
    return float(emu) / EMU_PER_CM


def fmt_emu(emu):
    if emu is None:
        return "?"
    return "%.2fcm" % emu_to_cm(emu)


# ----------------------------------------------------------------------------
# UTF-8 / 终端兼容
# ----------------------------------------------------------------------------

def setup_console():
    """尽量保证中文在 UOS/Kylin 终端、重定向、无 locale 环境下都不乱码。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


# ----------------------------------------------------------------------------
# XML 读写（标准库 ElementTree，保留命名空间前缀与 mc:Ignorable 语义）
# ----------------------------------------------------------------------------

_XMLNS_RE = re.compile(r"""xmlns(?::([A-Za-z_][\w.\-]*))?\s*=\s*["']([^"']*)["']""")
_BOM = "\ufeff"


def _split_start_tag(body):
    """解析序列化结果中的第一个“元素起始标签”。

    返回 (标签起点, 标签名, 属性串, 自闭合标记) 或 None。
    属性值里的 > 已被 ElementTree 转义，这里仍按引号扫描，保证健壮。
    """
    start = body.find("<")
    if start < 0:
        return None
    i = start + 1
    n = len(body)
    while i < n and body[i] not in (" ", "\t", "\r", "\n", ">", "/"):
        i += 1
    tag = body[start + 1:i]
    if not tag:
        return None
    quote = None
    while i < n:
        ch = body[i]
        if quote:
            if ch == quote:
                quote = None
        elif ch in ("\"", "'"):
            quote = ch
        elif ch == ">":
            inner = body[start + 1 + len(tag):i]
            tail = "/" if inner.endswith("/") else ""
            if tail:
                inner = inner[:-1]
            return start, tag, inner, tail
        i += 1
    return None


class XmlDocument(object):
    """一个 OOXML part 的 DOM 包装，序列化时尽量保持原样。"""

    __slots__ = ("path", "root", "decls", "prolog", "dirty")

    def __init__(self, path, root, decls, prolog):
        self.path = path
        self.root = root
        self.decls = decls
        self.prolog = prolog
        self.dirty = False

    def serialize(self):
        body = ET.tostring(self.root, encoding="unicode")
        # 把原文件中声明过、但 ET 因为“未被使用”而丢掉的命名空间补回根元素，
        # 否则 mc:Ignorable="p14 v" 之类会引用到未声明的前缀。
        parsed = _split_start_tag(body)
        if parsed is None:  # pragma: no cover
            return self.prolog.encode("utf-8") + body.encode("utf-8")
        head, tag, attrs, tail = parsed
        existing = set()
        for pre, _uri in _XMLNS_RE.findall(attrs):
            existing.add(pre)
        additions = []
        for pre, uri in self.decls.items():
            key = pre or ""
            if key in existing:
                continue
            try:
                uri.encode("ascii")
            except UnicodeEncodeError:  # pragma: no cover
                continue
            if pre:
                additions.append(' xmlns:%s="%s"' % (pre, uri))
            else:
                additions.append(' xmlns="%s"' % uri)
        new_head = "<" + tag + attrs + "".join(additions) + tail + ">"
        body = body[:head] + new_head + body[head + len("<" + tag + attrs + tail + ">"):]
        return (self.prolog + body).encode("utf-8")


def parse_xml_bytes(data, path):
    text = data.decode("utf-8-sig", errors="strict")
    decls = {}
    for pre, uri in _XMLNS_RE.findall(text):
        key = pre or ""
        decls[key] = uri
        if pre:
            try:
                ET.register_namespace(pre, uri)
            except ValueError:  # pragma: no cover
                pass
    m = re.search(r"<(?![?!])", text)
    prolog = text[: m.start()] if m else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    root = ET.fromstring(text)
    return XmlDocument(path, root, decls, prolog)


# ----------------------------------------------------------------------------
# OOXML 工具函数
# ----------------------------------------------------------------------------

def children(el, tag=None):
    if el is None:
        return []
    if tag is None:
        return list(el)
    want = qn("a", tag) if ":" in tag else qn("p", tag)
    return [c for c in el if c.tag == want]


def first_child(el, tag):
    if el is None:
        return None
    want = tag if tag.startswith("{") else (
        qn("a", tag) if ":" in tag else qn("p", tag))
    for c in el:
        if c.tag == want:
            return c
    return None


def find_child(el, prefix, tag):
    if el is None:
        return None
    want = qn(prefix, tag)
    for c in el:
        if c.tag == want:
            return c
    return None


def get_attr(el, name, default=None):
    if el is None:
        return default
    return el.get(name, default)


def get_int(el, name, default=None):
    v = get_attr(el, name)
    if v is None:
        return default
    try:
        return int(v)
    except (TypeError, ValueError):
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return default


def insert_ordered(parent, child, order):
    """按 schema 顺序把 child 插入 parent；order 为本地名列表。"""
    name = localname(child.tag)
    try:
        idx = order.index(name)
    except ValueError:
        parent.append(child)
        return child
    for i, existing in enumerate(list(parent)):
        ename = localname(existing.tag)
        try:
            eidx = order.index(ename)
        except ValueError:
            continue
        if eidx > idx:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def ensure_child(parent, prefix, tag, order):
    el = find_child(parent, prefix, tag)
    if el is not None:
        return el
    el = ET.Element(qn(prefix, tag))
    if order:
        insert_ordered(parent, el, order)
    else:
        parent.append(el)
    return el


def remove_child(parent, prefix, tag):
    el = find_child(parent, prefix, tag)
    if el is not None:
        parent.remove(el)
    return el


def text_of(el):
    """收集 r / fld 里的文本。"""
    if el is None:
        return ""
    if localname(el.tag) == "t":
        return el.text or ""
    parts = []
    for c in el:
        if localname(c.tag) == "t":
            parts.append(c.text or "")
    return "".join(parts)


def shape_name(sp):
    for nv in ("nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"):
        holder = find_child(sp, "p", nv)
        if holder is not None:
            c = find_child(holder, "p", "cNvPr")
            if c is not None:
                return c.get("name") or "?"
    return "?"


def shape_id(sp):
    for nv in ("nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"):
        holder = find_child(sp, "p", nv)
        if holder is not None:
            c = find_child(holder, "p", "cNvPr")
            if c is not None:
                try:
                    return int(c.get("id"))
                except (TypeError, ValueError):
                    return None
    return None


def placeholder_info(sp):
    """返回 (type, idx)；非占位符返回 (None, None)。"""
    for nv in ("nvSpPr", "nvPicPr", "nvGraphicFramePr"):
        holder = find_child(sp, "p", nv)
        if holder is None:
            continue
        nvpr = find_child(holder, "p", "nvPr")
        if nvpr is None:
            continue
        ph = find_child(nvpr, "p", "ph")
        if ph is not None:
            return ph.get("type"), get_int(ph, "idx")
    return None, None


def xfrm_of(sp_pr):
    if sp_pr is None:
        return None
    return find_child(sp_pr, "a", "xfrm")


def read_xfrm(sp_pr):
    x = xfrm_of(sp_pr)
    if x is None:
        return None
    off = find_child(x, "a", "off")
    ext = find_child(x, "a", "ext")
    if off is None or ext is None:
        return None
    return {
        "x": get_int(off, "x", 0), "y": get_int(off, "y", 0),
        "cx": get_int(ext, "cx", 0), "cy": get_int(ext, "cy", 0),
    }


def write_xfrm(sp_pr, geom):
    x = ensure_child(sp_pr, "a", "xfrm", None)
    for tag in ("off", "ext"):
        old = find_child(x, "a", tag)
        if old is not None:
            x.remove(old)
    off = ET.SubElement(x, qn("a", "off"))
    off.set("x", str(int(geom["x"])))
    off.set("y", str(int(geom["y"])))
    ext = ET.SubElement(x, qn("a", "ext"))
    ext.set("cx", str(int(geom["cx"])))
    ext.set("cy", str(int(geom["cy"])))
    # xfrm 必须是 spPr 的第一个子元素
    if list(sp_pr)[0] is not x:
        sp_pr.remove(x)
        sp_pr.insert(0, x)


# ----------------------------------------------------------------------------
# PPTX 包
# ----------------------------------------------------------------------------

class PptxPackage(object):
    def __init__(self, path):
        self.path = path
        self._zip = zipfile.ZipFile(path, "r")
        self.names = [n for n in self._zip.namelist() if not n.endswith("/")]
        self._name_set = set(self.names)
        self._xml = {}
        self._rels = {}
        self.dirty = set()
        self.saved_path = None

    # -- 基础读取 -----------------------------------------------------------
    def has(self, name):
        return name in self._name_set

    def read(self, name):
        return self._zip.read(name)

    def xml(self, name):
        if name not in self._xml:
            self._xml[name] = parse_xml_bytes(self.read(name), name)
        return self._xml[name]

    def mark(self, name):
        self.dirty.add(name)

    def close(self):
        try:
            self._zip.close()
        except Exception:
            pass

    # -- 关系解析 -----------------------------------------------------------
    def relationships(self, part):
        if part in self._rels:
            return self._rels[part]
        rp = rels_part_name(part)
        result = {}
        if self.has(rp):
            root = self.xml(rp).root
            for rel in root:
                if localname(rel.tag) != "Relationship":
                    continue
                rid = rel.get("Id")
                target = rel.get("Target")
                mode = rel.get("TargetMode", "Internal")
                result[rid] = {
                    "type": rel.get("Type", ""),
                    "target": target,
                    "external": mode == "External",
                    "part": None if mode == "External" else resolve_part_target(part, target),
                }
        self._rels[part] = result
        return result

    def rel_by_type(self, part, suffix):
        for rid, rel in self.relationships(part).items():
            if rel["type"].endswith(suffix):
                return rel
        return None

    def related_part(self, part, suffix):
        rel = self.rel_by_type(part, suffix)
        if rel is None:
            return None
        return rel["part"] if rel["part"] in self._name_set else None

    def iter_parts(self, pattern):
        rx = re.compile(pattern)
        return [n for n in self.names if rx.match(n)]

    # -- 保存 ---------------------------------------------------------------
    def save(self, dest, keep_original_order=True):
        tmp = dest + ".tmp-%d" % os.getpid()
        names = self.names
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as z:
            for name in names:
                data = self.xml(name).serialize() if name in self.dirty else self.read(name)
                zi = zipfile.ZipInfo(name)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.date_time = (2024, 1, 1, 0, 0, 0)
                z.writestr(zi, data)
        os.replace(tmp, dest)
        self.saved_path = dest


def rels_part_name(part):
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def resolve_part_target(part, target):
    if target is None:
        return None
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(part)
    return posixpath.normpath(posixpath.join(base, target))


# 下面这些函数在后续代码中频繁使用，先把命名空间别名准备好
A = lambda tag: qn("a", tag)
P = lambda tag: qn("p", tag)
R = lambda tag: qn("r", tag)


# ============================================================================
# 字体：注册表、度量、替代映射
# ============================================================================

CJK_SERIF = [
    "Noto Serif CJK SC", "Source Han Serif SC", "Source Han Serif CN",
    "Noto Serif CJK TC", "AR PL UMing CN", "AR PL UMing TW",
    "WenQuanYi Zen Hei", "Droid Sans Fallback", "Noto Sans CJK SC",
]
CJK_SANS = [
    "Noto Sans CJK SC", "Source Han Sans SC", "Source Han Sans CN",
    "Noto Sans CJK TC", "WenQuanYi Micro Hei", "WenQuanYi Zen Hei",
    "Droid Sans Fallback", "Noto Serif CJK SC",
]
CJK_KAI = [
    "AR PL UKai CN", "AR PL UKai TW", "Noto Serif CJK SC", "Noto Sans CJK SC",
]
LATIN_SERIF = [
    "Times New Roman", "Liberation Serif", "Tinos", "Nimbus Roman",
    "Noto Serif", "DejaVu Serif", "FreeSerif",
]
LATIN_SANS = [
    "Arial", "Liberation Sans", "Arimo", "Nimbus Sans", "Noto Sans",
    "DejaVu Sans", "FreeSans",
]
LATIN_MONO = [
    "Courier New", "Liberation Mono", "Nimbus Mono PS", "DejaVu Sans Mono",
    "Noto Sans Mono", "FreeMono",
]
SYMBOL_FONT_NAMES = {
    "wingdings", "wingdings 2", "wingdings 3", "wingdings 1", "webdings",
    "symbol", "marlett", "mt extra", "dingbats", "zapf dingbats",
    "wingdings regular", "wingdings mt", "monotype sorts",
}

# Windows 常见字体 -> 首选替代（按优先级；运行时只挑选“本机已安装”的）
EXACT_SUBSTITUTES = {
    # 中文 Windows 字体
    "simsun": CJK_SERIF, "nsimsun": CJK_SERIF, "宋体": CJK_SERIF,
    "宋体-方正超大字符集": CJK_SERIF, "fangzheng": CJK_SERIF,
    "microsoft yahei": CJK_SANS, "微软雅黑": CJK_SANS,
    "microsoft yahei light": CJK_SANS, "微软雅黑 light": CJK_SANS,
    "microsoft jhenghei": ["Noto Sans CJK TC", "Noto Sans CJK SC"],
    "simhei": CJK_SANS, "黑体": CJK_SANS, "stheiti": CJK_SANS,
    "华文细黑": CJK_SANS, "stxihei": CJK_SANS,
    "dengxian": CJK_SANS, "等线": CJK_SANS, "dengxian light": CJK_SANS,
    "等线 light": CJK_SANS, "等线 light (正文)": CJK_SANS,
    "kaiti": CJK_KAI, "楷体": CJK_KAI, "kaiti_gb2312": CJK_KAI,
    "楷体_gb2312": CJK_KAI, "stkaiti": CJK_KAI, "华文楷体": CJK_KAI,
    "fangsong": CJK_SERIF, "仿宋": CJK_SERIF, "fangsong_gb2312": CJK_SERIF,
    "仿宋_gb2312": CJK_SERIF, "stfangsong": CJK_SERIF, "华文仿宋": CJK_SERIF,
    "youyuan": CJK_SANS, "幼圆": CJK_SANS,
    "lisu": CJK_SERIF, "隶书": CJK_SERIF,
    "stzhongsong": CJK_SERIF, "华文中宋": CJK_SERIF,
    "stcusong": CJK_SERIF, "华文粗宋": CJK_SERIF,
    "stxingkai": CJK_SERIF, "华文行楷": CJK_SERIF,
    "stxinwei": CJK_SERIF, "华文新魏": CJK_SERIF,
    "fzshusong_z04": CJK_SERIF, "方正书宋": CJK_SERIF,
    "fzxbsong_z46s": CJK_SERIF, "方正小标宋": CJK_SERIF,
    "fzhs": CJK_SANS, "方正黑体": CJK_SANS,
    "fzkt": CJK_KAI, "方正楷体": CJK_KAI,
    # 西文字体
    "times new roman": LATIN_SERIF, "times": LATIN_SERIF,
    "cambria": LATIN_SERIF, "georgia": LATIN_SERIF, "garamond": LATIN_SERIF,
    "book antiqua": LATIN_SERIF, "palatino linotype": LATIN_SERIF,
    "calibri": ["Carlito", "Liberation Sans", "Arimo", "Noto Sans", "DejaVu Sans"],
    "calibri light": ["Carlito", "Liberation Sans", "Noto Sans", "DejaVu Sans"],
    "candara": ["Carlito", "Liberation Sans", "DejaVu Sans"],
    "corbel": ["Carlito", "Liberation Sans", "DejaVu Sans"],
    "constantia": ["Caladea", "Liberation Serif", "DejaVu Serif"],
    "arial": LATIN_SANS, "helvetica": LATIN_SANS,
    "arial black": ["Liberation Sans", "DejaVu Sans"],
    "arial narrow": ["Liberation Sans Narrow", "Liberation Sans", "DejaVu Sans Condensed"],
    "tahoma": ["DejaVu Sans", "Liberation Sans", "Noto Sans"],
    "verdana": ["DejaVu Sans", "Liberation Sans", "Noto Sans"],
    "trebuchet ms": ["DejaVu Sans", "Liberation Sans", "Noto Sans"],
    "segoe ui": ["Noto Sans", "DejaVu Sans", "Liberation Sans"],
    "segoe ui light": ["Noto Sans Light", "Noto Sans", "DejaVu Sans"],
    "segoe ui semibold": ["Noto Sans Medium", "Noto Sans", "DejaVu Sans"],
    "franklin gothic book": ["Liberation Sans", "DejaVu Sans"],
    "century gothic": ["Liberation Sans", "DejaVu Sans"],
    "courier new": LATIN_MONO, "courier": LATIN_MONO,
    "consolas": ["Liberation Mono", "DejaVu Sans Mono", "Noto Sans Mono"],
    "lucida console": ["DejaVu Sans Mono", "Liberation Mono"],
    "lucida sans unicode": ["DejaVu Sans", "Noto Sans"],
    "ms mincho": ["Noto Serif CJK JP", "Noto Serif CJK SC"],
    "ms gothic": ["Noto Sans CJK JP", "Noto Sans CJK SC"],
    "meiryo": ["Noto Sans CJK JP", "Noto Sans CJK SC"],
    "malgun gothic": ["Noto Sans CJK KR", "Noto Sans CJK SC"],
    "batang": ["Noto Serif CJK KR", "Noto Serif CJK SC"],
    "gulim": ["Noto Sans CJK KR", "Noto Sans CJK SC"],
}

# 结合名称特征的模糊分类（用于表里没收录的 Windows 字体）
_HEURISTIC_RULES = [
    (("wingdings", "webdings", "marlett", "symbol", "dingbat", "mt extra"), "symbol"),
    (("楷", "kai", "ukai"), "cjk_kai"),
    (("仿宋", "fangsong", "song", "宋", "ming", "sun", "serif", "roman"), "cjk_serif"),
    (("雅黑", "yahei", "hei", "黑", "dengxian", "等线", "圆", "yuan", "sans", "gothic"), "cjk_sans"),
    (("times", "cambria", "georgia", "garamond", "palatino", "book"), "latin_serif"),
    (("courier", "consolas", "mono", "code"), "latin_mono"),
    (("arial", "calibri", "segoe", "tahoma", "verdana", "trebuchet", "helvetica"), "latin_sans"),
]


def norm_family(name):
    return re.sub(r"\s+", "", (name or "")).casefold()


def classify_family(name):
    """判断字体属于哪一类，返回 cjk_serif/cjk_sans/cjk_kai/latin_*/symbol/unknown。"""
    low = (name or "").strip().casefold()
    if not low:
        return "unknown"
    if low in SYMBOL_FONT_NAMES or low.split()[0] in SYMBOL_FONT_NAMES:
        return "symbol"
    key = norm_family(name)
    if key in EXACT_SUBSTITUTES:
        cands = EXACT_SUBSTITUTES[key]
        if cands is CJK_SERIF:
            return "cjk_serif"
        if cands is CJK_SANS:
            return "cjk_sans"
        if cands is CJK_KAI:
            return "cjk_kai"
        if cands is LATIN_SERIF:
            return "latin_serif"
        if cands is LATIN_SANS:
            return "latin_sans"
        if cands is LATIN_MONO:
            return "latin_mono"
        # 混合候选列表：用名字特征判断
    for keys, cls in _HEURISTIC_RULES:
        for k in keys:
            if k in low:
                return cls
    return "unknown"


def candidates_for(name, cls=None):
    key = norm_family(name)
    cands = EXACT_SUBSTITUTES.get(key)
    if cands:
        return list(cands)
    cls = cls or classify_family(name)
    table = {
        "cjk_serif": CJK_SERIF, "cjk_sans": CJK_SANS, "cjk_kai": CJK_KAI,
        "latin_serif": LATIN_SERIF, "latin_sans": LATIN_SANS,
        "latin_mono": LATIN_MONO, "symbol": [],
        "unknown": LATIN_SANS,
    }
    return list(table.get(cls, LATIN_SANS))


def is_cjk_char(ch):
    o = ord(ch)
    if 0x2E80 <= o <= 0x9FFF:
        return True
    if 0xF900 <= o <= 0xFAFF:
        return True
    if 0xFE30 <= o <= 0xFE4F:
        return True
    if 0xFF00 <= o <= 0xFFEF:
        return True
    if 0x20000 <= o <= 0x3FFFF:
        return True
    return False


def is_hangul_char(ch):
    o = ord(ch)
    return 0xAC00 <= o <= 0xD7AF or 0x1100 <= o <= 0x11FF


def is_kana_char(ch):
    o = ord(ch)
    return 0x3040 <= o <= 0x30FF


class FontParseError(Exception):
    pass


class SfntFont(object):
    """极简 TrueType/OpenType 度量读取器：只取 cmap / hmtx / head / hhea。"""

    def __init__(self, path, index=0):
        self.path = path
        self.index = index
        self.upem = 1000
        self.ascender = 800
        self.descender = -200
        self.line_gap = 0
        self._advances = []
        self._nmetrics = 0
        self._cmap = {}
        self._fmt12 = None
        self._fmt12_starts = []
        self._fmt4 = None
        self._glyph_cache = {}
        self._adv_cache = {}
        self._load()

    # -- 解析 ---------------------------------------------------------------
    def _load(self):
        try:
            with open(self.path, "rb") as fh:
                data = fh.read()
        except (IOError, OSError) as exc:
            raise FontParseError(str(exc))
        try:
            tables, base = self._table_directory(data)
            self._load_head(data, tables)
            self._load_hhea(data, tables)
            self._load_hmtx(data, tables)
            self._load_cmap(data, tables)
        except (struct.error, KeyError, IndexError, ValueError) as exc:
            raise FontParseError("字体表结构异常: %s" % exc)

    def _table_directory(self, data):
        if len(data) < 12:
            raise FontParseError("文件过短")
        tag = data[:4]
        if tag == b"ttcf":
            num = struct.unpack(">I", data[8:12])[0]
            if num == 0 or num > 4096:
                raise FontParseError("TTC 索引异常")
            offsets = struct.unpack(">%dI" % num, data[12:12 + 4 * num])
            idx = self.index if self.index < num else 0
            base = offsets[idx]
        else:
            base = 0
        if base + 12 > len(data):
            raise FontParseError("偏移越界")
        num_tables = struct.unpack(">H", data[base + 4:base + 6])[0]
        if num_tables <= 0 or num_tables > 512:
            raise FontParseError("表数量异常")
        tables = {}
        for i in range(num_tables):
            rec = base + 12 + 16 * i
            if rec + 16 > len(data):
                raise FontParseError("表目录越界")
            ttag = data[rec:rec + 4].decode("latin-1")
            off, length = struct.unpack(">II", data[rec + 8:rec + 16])
            tables[ttag] = (off, length)
        return tables, base

    def _load_head(self, data, tables):
        head = tables["head"][0]
        self.upem = struct.unpack(">H", data[head + 18:head + 20])[0]
        if not self.upem:
            raise FontParseError("unitsPerEm 为 0")

    def _load_hhea(self, data, tables):
        hhea = tables["hhea"][0]
        self.ascender, self.descender, self.line_gap = struct.unpack(
            ">hhh", data[hhea + 4:hhea + 10])
        self._nmetrics = struct.unpack(">H", data[hhea + 34:hhea + 36])[0]

    def _load_hmtx(self, data, tables):
        if self._nmetrics <= 0:
            raise FontParseError("hmtx 无度量")
        off, length = tables["hmtx"]
        count = min(self._nmetrics, length // 4)
        self._advances = list(struct.unpack(">%dH" % count, data[off:off + 2 * count]))[
            :count]

    def _load_cmap(self, data, tables):
        cmap_off = tables["cmap"][0]
        nsub = struct.unpack(">H", data[cmap_off + 2:cmap_off + 4])[0]
        if nsub <= 0 or nsub > 256:
            raise FontParseError("cmap 子表异常")
        best = None
        for i in range(nsub):
            rec = cmap_off + 4 + 8 * i
            pid, eid, off = struct.unpack(">HHI", data[rec:rec + 8])
            sub = cmap_off + off
            if sub + 6 > len(data):
                continue
            fmt = struct.unpack(">H", data[sub:sub + 2])[0]
            score = 0
            if fmt == 12:
                score = 100
            elif fmt == 4:
                score = 80
            elif fmt == 6:
                score = 50
            elif fmt == 0:
                score = 30
            elif fmt == 14:
                score = 0
            if pid in (3, 0) and score > (best[0] if best else 0):
                best = (score, fmt, sub, pid, eid)
        if best is None:
            raise FontParseError("没有可用的 Unicode cmap 子表")
        _score, fmt, sub, _pid, _eid = best
        if fmt == 12:
            self._parse_fmt12(data, sub)
        elif fmt == 4:
            self._parse_fmt4(data, sub)
        elif fmt == 6:
            self._parse_fmt6(data, sub)
        else:
            self._parse_fmt0(data, sub)

    def _parse_fmt12(self, data, sub):
        ngroups = struct.unpack(">I", data[sub + 12:sub + 16])[0]
        if ngroups > 200000:
            raise FontParseError("cmap 分组过多")
        groups = []
        for i in range(ngroups):
            rec = sub + 16 + 12 * i
            start, end, gid = struct.unpack(">III", data[rec:rec + 12])
            groups.append((start, end, gid))
        groups.sort()
        self._fmt12 = groups
        self._fmt12_starts = [g[0] for g in groups]

    def _parse_fmt4(self, data, sub):
        segx2 = struct.unpack(">H", data[sub + 6:sub + 8])[0]
        seg = segx2 // 2
        if seg <= 0:
            raise FontParseError("cmap 分段数为 0")
        end_o = sub + 14
        start_o = end_o + segx2 + 2
        delta_o = start_o + segx2
        range_o = delta_o + segx2
        table = {}
        for i in range(seg):
            end = struct.unpack(">H", data[end_o + 2 * i:end_o + 2 * i + 2])[0]
            start = struct.unpack(">H", data[start_o + 2 * i:start_o + 2 * i + 2])[0]
            delta = struct.unpack(">h", data[delta_o + 2 * i:delta_o + 2 * i + 2])[0]
            ro = struct.unpack(">H", data[range_o + 2 * i:range_o + 2 * i + 2])[0]
            if start == 0xFFFF:
                continue
            for cp in range(start, end + 1):
                if ro == 0:
                    gid = (cp + delta) & 0xFFFF
                else:
                    gi = range_o + 2 * i + ro + 2 * (cp - start)
                    if gi + 2 > len(data):
                        break
                    gid = struct.unpack(">H", data[gi:gi + 2])[0]
                    if gid:
                        gid = (gid + delta) & 0xFFFF
                if gid:
                    table[cp] = gid
        self._fmt4 = table

    def _parse_fmt6(self, data, sub):
        first = struct.unpack(">H", data[sub + 6:sub + 8])[0]
        count = struct.unpack(">H", data[sub + 8:sub + 10])[0]
        table = {}
        for i in range(count):
            rec = sub + 10 + 2 * i
            gid = struct.unpack(">H", data[rec:rec + 2])[0]
            if gid:
                table[first + i] = gid
        self._fmt4 = table

    def _parse_fmt0(self, data, sub):
        table = {}
        for cp in range(256):
            gid = data[sub + 6 + cp]
            if gid:
                table[cp] = gid
        self._fmt4 = table

    # -- 查询 ---------------------------------------------------------------
    def glyph_id(self, cp):
        try:
            return self._glyph_cache[cp]
        except KeyError:
            pass
        gid = None
        if self._fmt12 is not None:
            i = bisect.bisect_right(self._fmt12_starts, cp) - 1
            if i >= 0:
                start, end, gid0 = self._fmt12[i]
                if start <= cp <= end:
                    gid = gid0 + (cp - start)
        elif self._fmt4 is not None:
            gid = self._fmt4.get(cp)
        self._glyph_cache[cp] = gid
        return gid

    def has_glyph(self, ch):
        gid = self.glyph_id(ord(ch))
        return gid is not None and gid > 0

    def advance(self, cp):
        """返回以 em 为单位的字符步进宽度。"""
        try:
            return self._adv_cache[cp]
        except KeyError:
            pass
        if cp == 0x20:
            gid = self.glyph_id(cp)
            if gid is None:
                self._adv_cache[cp] = 0.25
                return 0.25
        gid = self.glyph_id(cp)
        if gid is None or not self._advances:
            value = 1.0 if is_cjk_char(chr(cp)) else 0.55
        elif gid < len(self._advances):
            value = self._advances[gid] / float(self.upem)
        else:
            value = self._advances[-1] / float(self.upem)
        self._adv_cache[cp] = value
        return value

    def line_em(self):
        height = (self.ascender - self.descender + self.line_gap) / float(self.upem)
        if height <= 0.3:
            height = 1.2
        return height


class FontRegistry(object):
    """通过 fontconfig 探测本机字体，并缓存字体度量。"""

    def __init__(self, extra_dirs=None):
        self.available = False
        self.families = {}       # norm -> display name
        self.files = {}          # norm -> (path, index)
        self._match_cache = {}
        self._font_cache = {}
        self._load()

    def _load(self):
        try:
            proc = subprocess.run(
                ["fc-list", "--format", "%{file}\t%{index}\t%{family}\n"],
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, timeout=60)
        except (OSError, subprocess.SubprocessError):
            return
        if proc.returncode != 0:
            return
        text = proc.stdout.decode("utf-8", "replace")
        for line in text.splitlines():
            fields = line.split("\t")
            if len(fields) < 3:
                continue
            path, idx, fams = fields[0], fields[1], "\t".join(fields[2:])
            try:
                index = int(idx or 0)
            except ValueError:
                index = 0
            for fam in fams.split(","):
                fam = fam.strip()
                if not fam:
                    continue
                key = norm_family(fam)
                self.families.setdefault(key, fam)
                self.files.setdefault(key, (path, index))
        self.available = bool(self.families)

    # -- 查询 ---------------------------------------------------------------
    def installed(self, name):
        return norm_family(name) in self.families

    def display_name(self, name):
        return self.families.get(norm_family(name), name)

    def match(self, name):
        """fc-match：返回 (family, path, index)；用于模拟“WPS 实际会用什么字体”。"""
        key = norm_family(name)
        if key in self._match_cache:
            return self._match_cache[key]
        result = None
        try:
            proc = subprocess.run(
                ["fc-match", "-f", "%{file}|%{index}|%{family}", name],
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, timeout=30)
            if proc.returncode == 0:
                out = proc.stdout.decode("utf-8", "replace").strip()
                parts = out.split("|")
                if len(parts) >= 3:
                    path = parts[0]
                    try:
                        idx = int(parts[1] or 0)
                    except ValueError:
                        idx = 0
                    fams = parts[2].split(",")
                    fam = fams[0].strip() if fams else name
                    result = (fam, path, idx)
        except (OSError, subprocess.SubprocessError):
            result = None
        self._match_cache[key] = result
        return result

    def font_for(self, name):
        """加载字体度量对象（按实际会被使用的字体）。"""
        key = norm_family(name)
        if key in self._font_cache:
            return self._font_cache[key]
        font = None
        entry = self.files.get(key)
        if entry is None:
            m = self.match(name)
            if m is not None:
                entry = (m[1], m[2])
        if entry is not None:
            try:
                font = SfntFont(entry[0], entry[1])
            except (FontParseError, OSError):
                font = None
        self._font_cache[key] = font
        return font

    def first_installed(self, candidates):
        for cand in candidates:
            if self.installed(cand):
                return self.display_name(cand)
        return None

    def best_fallback(self, name, cls=None):
        """为某个字体名挑一个本机可用的替代族。"""
        cls = cls or classify_family(name)
        for cand in candidates_for(name, cls):
            if self.installed(cand):
                return self.display_name(cand)
        m = self.match(name)
        if m is not None:
            return m[0]
        if cls.startswith("cjk"):
            return self.first_installed(CJK_SANS) or (m[0] if m else name)
        return self.first_installed(LATIN_SANS) or name

    def latin_companion(self, cjk_name):
        cls = classify_family(cjk_name)
        table = {"cjk_kai": LATIN_SERIF, "cjk_serif": LATIN_SERIF}.get(cls, LATIN_SANS)
        return (self.first_installed(table) or self.first_installed(LATIN_SANS)
                or self.best_fallback("Arial", "latin_sans"))

    def cjk_companion(self, latin_name):
        cls = classify_family(latin_name)
        table = {"latin_serif": CJK_SERIF, "latin_mono": CJK_SANS}.get(cls, CJK_SANS)
        return (self.first_installed(table) or self.first_installed(CJK_SANS)
                or self.best_fallback("宋体", "cjk_serif"))

    # -- 中文/西文覆盖 ------------------------------------------------------
    def has_cjk_coverage(self, name):
        font = self.font_for(name)
        if font is None:
            return None
        return font.has_glyph("中") and font.has_glyph("。")

    def coverage_missing(self, name, text):
        font = self.font_for(name)
        if font is None:
            return None
        missing = []
        for ch in text:
            if ch in "\r\n\t ":
                continue
            if not font.has_glyph(ch):
                missing.append(ch)
                if len(missing) >= 12:
                    break
        return missing


class FontDecision(object):
    __slots__ = ("name", "status", "target", "kind", "reason", "companion_latin",
                 "companion_cjk")

    def __init__(self, name, status, target=None, kind="unknown", reason="",
                 companion_latin=None, companion_cjk=None):
        self.name = name
        self.status = status          # ok / replace / symbol / unknown
        self.target = target
        self.kind = kind
        self.reason = reason
        self.companion_latin = companion_latin
        self.companion_cjk = companion_cjk

    def __repr__(self):  # pragma: no cover
        return "FontDecision(%r, %s -> %r)" % (self.name, self.status, self.target)


class FontResolver(object):
    """决定“某个字体名在本机应该变成什么”。"""

    def __init__(self, registry):
        self.registry = registry
        self._cache = {}

    def decide(self, name):
        key = norm_family(name)
        if key in self._cache:
            return self._cache[key]
        d = self._decide(name)
        self._cache[key] = d
        return d

    def _decide(self, name):
        registry = self.registry
        kind = classify_family(name)
        if kind == "symbol":
            return FontDecision(name, "symbol", None, kind,
                                "符号字体（Wingdings/Webdings/Symbol 等）无法用普通字体替代")
        if registry.available and registry.installed(name):
            return FontDecision(name, "ok", registry.display_name(name), kind, "本机已安装")
        target = registry.best_fallback(name, kind)
        if not registry.available:
            cands = candidates_for(name, kind)
            target = cands[0] if cands else name
            reason = "未检测到 fontconfig，按内置映射替换"
        else:
            reason = "本机缺失，替换为度量/风格最接近的字体"
        comp_latin = registry.latin_companion(name) if kind.startswith("cjk") else None
        comp_cjk = registry.cjk_companion(name) if kind in ("latin_serif", "latin_sans", "latin_mono") else None
        return FontDecision(name, "replace", target, kind, reason, comp_latin, comp_cjk)


# ============================================================================
# 文本度量与换行估算
# ============================================================================

class Measure(object):
    __slots__ = ("height_pt", "max_line_width_pt", "lines", "line_heights",
                 "text_len")

    def __init__(self, height_pt, max_line_width_pt, lines, line_heights, text_len):
        self.height_pt = height_pt
        self.max_line_width_pt = max_line_width_pt
        self.lines = lines
        self.line_heights = line_heights
        self.text_len = text_len


def _char_width(item, char, fallback_cjk, fallback_latin):
    font, size, spc = item
    if font is None:
        base = size * (1.0 if is_cjk_char(char) else 0.5)
        return base + spc
    if font.has_glyph(char):
        return font.advance(ord(char)) * size + spc
    # 字体缺字形时，渲染器会回退到系统字体
    fb = fallback_cjk if is_cjk_char(char) else fallback_latin
    if fb is not None and fb.has_glyph(char):
        return fb.advance(ord(char)) * size + spc
    return size * (1.0 if is_cjk_char(char) else 0.5) + spc


def _tokenize(items):
    """把 [(char, font, size, spc)] 拆成可断行的单元。"""
    units = []
    word = []

    def flush_word():
        if word:
            units.append(("word", word[:]))
            del word[:]

    for item in items:
        ch = item[0]
        if ch in " \t\u00a0":
            flush_word()
            units.append(("space", [item]))
        elif is_cjk_char(ch):
            flush_word()
            units.append(("char", [item]))
        else:
            word.append(item)
    flush_word()
    return units


def measure_paragraph(items, avail_w_pt, wrap=True, lnspc=None, spc_bef=0.0,
                      spc_aft=0.0, autofit_lnspc_reduction=0.0, fallback_cjk=None,
                      fallback_latin=None, empty_line_pt=0.0):
    """估算一段文本在给定宽度下的行数、行宽与总高度（单位 pt）。"""
    if not items:
        # 空段落（只有 a:endParaRPr 或 a:br）仍然占一行
        h = empty_line_pt
        if lnspc is not None:
            kind, val = lnspc
            if kind == "pts":
                h = val
            elif kind == "pct":
                h = h * val / 100000.0
        if autofit_lnspc_reduction:
            h = h * max(0.1, 1.0 - autofit_lnspc_reduction)
        return Measure(spc_bef + spc_aft + h, 0.0, 1, [h], 0)
    units = _tokenize(items)
    lines = []          # 每行: list[(char, font, size, spc)]
    cur = []
    cur_w = 0.0
    avail = max(avail_w_pt, 1.0)

    def newline():
        nonlocal cur, cur_w
        lines.append(cur)
        cur = []
        cur_w = 0.0

    for utype, chars in units:
        unit_w = sum(_char_width(it, it[0], fallback_cjk, fallback_latin) for it in chars)
        if not wrap:
            cur.extend(chars)
            cur_w += unit_w
            continue
        if utype == "space":
            if not cur:          # 行首空格不占位
                continue
            if cur_w + unit_w > avail + 0.01:
                newline()
                continue
            cur.extend(chars)
            cur_w += unit_w
            continue
        if cur and cur_w + unit_w > avail + 0.01:
            newline()
        if unit_w <= avail + 0.01:
            cur.extend(chars)
            cur_w += unit_w
        else:
            # 单个词比整行还宽：按字符硬断
            for it in chars:
                w = _char_width(it, it[0], fallback_cjk, fallback_latin)
                if cur and cur_w + w > avail + 0.01:
                    newline()
                cur.append(it)
                cur_w += w
    if cur or not lines:
        lines.append(cur)

    # 行尾空格不计入行宽，但排版上仍可能导致换行，这里取折中：不计宽度
    widths = []
    heights = []
    for line in lines:
        while line and line[-1][0] in " \t\u00a0":
            line = line[:-1]
        widths.append(sum(_char_width(it, it[0], fallback_cjk, fallback_latin)
                          for it in line))
        if not line:
            heights.append(empty_line_pt)
            continue
        base = 0.0
        max_size = 0.0
        for it in line:
            font, size, _spc = it
            max_size = max(max_size, size)
            if font is not None:
                base = max(base, font.line_em() * size)
        if base <= 0:
            base = 1.2 * max_size
        heights.append(base)

    # 行距
    if lnspc is not None:
        kind, val = lnspc
        if kind == "pts":
            heights = [val] * len(heights)
        elif kind == "pct":
            heights = [h * val / 100000.0 for h in heights]
    if autofit_lnspc_reduction:
        factor = max(0.1, 1.0 - autofit_lnspc_reduction)
        heights = [h * factor for h in heights]

    total = sum(heights) + spc_bef + spc_aft
    return Measure(total, max(widths) if widths else 0.0, len(lines), heights,
                   sum(len(line) for line in lines))


# ============================================================================
# 演示文稿模型：主题字体、占位符几何、文本样式继承
# ============================================================================

class ThemeFonts(object):
    __slots__ = ("part", "major_latin", "major_ea", "minor_latin", "minor_ea",
                 "script_major", "script_minor")

    def __init__(self, part):
        self.part = part
        self.major_latin = ""
        self.major_ea = ""
        self.minor_latin = ""
        self.minor_ea = ""
        self.script_major = {}
        self.script_minor = {}


def parse_theme_fonts(part, root):
    fonts = ThemeFonts(part)
    for scheme in root.iter():
        if localname(scheme.tag) != "fontScheme":
            continue
        for holder_name, target_major in (("majorFont", True), ("minorFont", False)):
            holder = find_child(scheme, "a", holder_name)
            if holder is None:
                continue
            latin = find_child(holder, "a", "latin")
            ea = find_child(holder, "a", "ea")
            if latin is not None:
                if target_major:
                    fonts.major_latin = latin.get("typeface") or ""
                else:
                    fonts.minor_latin = latin.get("typeface") or ""
            if ea is not None:
                if target_major:
                    fonts.major_ea = ea.get("typeface") or ""
                else:
                    fonts.minor_ea = ea.get("typeface") or ""
            for f in holder:
                if localname(f.tag) != "font":
                    continue
                script = f.get("script") or ""
                face = f.get("typeface") or ""
                if not script:
                    continue
                if target_major:
                    fonts.script_major[script] = face
                else:
                    fonts.script_minor[script] = face
        break
    return fonts


class Deck(object):
    """演示文稿结构信息：页序、版式/母版/主题关系、占位符几何、默认样式。"""

    def __init__(self, pkg, registry=None, resolver=None):
        self.pkg = pkg
        self.registry = registry
        self.resolver = resolver
        self.presentation_part = "ppt/presentation.xml"
        self.pres_root = pkg.xml(self.presentation_part).root
        self.sld_w = 12192000
        self.sld_h = 6858000
        sz = find_child(self.pres_root, "p", "sldSz")
        if sz is not None:
            self.sld_w = get_int(sz, "cx", self.sld_w)
            self.sld_h = get_int(sz, "cy", self.sld_h)
        self.default_text_style = find_child(self.pres_root, "p", "defaultTextStyle")
        self.embedded_fonts = []
        for efl in self.pres_root.iter():
            if localname(efl.tag) == "embeddedFontLst":
                for ef in efl:
                    if localname(ef.tag) != "embeddedFont":
                        continue
                    for child in ef:
                        if localname(child.tag) in ("font", "regular", "bold",
                                                    "italic", "boldItalic"):
                            face = child.get("typeface") or ""
                            if face:
                                self.embedded_fonts.append(face)
        self.slides = self._ordered_slides()
        self.slide_number = {part: i + 1 for i, part in enumerate(self.slides)}
        self._theme_cache = {}
        self._slide_theme = {}
        self._ph_cache = {}
        self._master_cache = {}
        self._layout_cache = {}
        self._txstyle_cache = {}
        # 兜底字体
        if registry is not None and registry.available:
            self.default_cjk = (registry.first_installed(CJK_SANS)
                                or registry.first_installed(CJK_SERIF)
                                or "Noto Sans CJK SC")
            self.default_latin = (registry.first_installed(LATIN_SANS)
                                  or "Liberation Sans")
        else:
            self.default_cjk = "Noto Sans CJK SC"
            self.default_latin = "Liberation Sans"
        self._build_geom_cache()

    # -- 页序 ---------------------------------------------------------------
    def _ordered_slides(self):
        ids = []
        lst = find_child(self.pres_root, "p", "sldIdLst")
        if lst is not None:
            rels = self.pkg.relationships(self.presentation_part)
            for sld in lst:
                rid = sld.get(qn("r", "id"))
                rel = rels.get(rid)
                if rel and rel["part"] and self.pkg.has(rel["part"]):
                    ids.append(rel["part"])
        if not ids:
            found = self.pkg.iter_parts(r"ppt/slides/slide\d+\.xml$")
            def num(p):
                m = re.search(r"(\d+)\.xml$", p)
                return int(m.group(1)) if m else 0
            ids = sorted(found, key=num)
        return ids

    # -- 关系链 -------------------------------------------------------------
    def layout_of(self, slide_part):
        if slide_part in self._layout_cache:
            return self._layout_cache[slide_part]
        part = self.pkg.related_part(slide_part, "/slideLayout")
        self._layout_cache[slide_part] = part
        return part

    def master_of(self, layout_part):
        if layout_part is None:
            return None
        if layout_part in self._master_cache:
            return self._master_cache[layout_part]
        part = self.pkg.related_part(layout_part, "/slideMaster")
        self._master_cache[layout_part] = part
        return part

    def theme_of_slide(self, slide_part):
        if slide_part in self._slide_theme:
            return self._slide_theme[slide_part]
        chain = [self.layout_of(slide_part)]
        master = None
        while chain and chain[-1] is not None and master is None:
            master = self.master_of(chain[-1])
            if master is None:
                break
        theme_part = None
        if master is not None:
            theme_part = self.pkg.related_part(master, "/theme")
        if theme_part is None:
            candidates = self.pkg.iter_parts(r"ppt/theme/theme\d+\.xml$")
            theme_part = candidates[0] if candidates else None
        self._slide_theme[slide_part] = theme_part
        return theme_part

    def theme_fonts(self, part):
        if part is None:
            return ThemeFonts("")
        if part not in self._theme_cache:
            self._theme_cache[part] = parse_theme_fonts(part, self.pkg.xml(part).root)
        return self._theme_cache[part]

    def theme_fonts_for_slide(self, slide_part):
        return self.theme_fonts(self.theme_of_slide(slide_part))

    # -- 占位符几何 ---------------------------------------------------------
    def _collect_placeholders(self, part):
        if part is None or not self.pkg.has(part):
            return {}
        if part in self._ph_cache:
            return self._ph_cache[part]
        result = {}
        root = self.pkg.xml(part).root
        for sp_tree in root.iter():
            if localname(sp_tree.tag) != "spTree":
                continue
            for sp in sp_tree:
                if localname(sp.tag) not in ("sp", "graphicFrame", "pic", "cxnSp"):
                    continue
                phtype, idx = placeholder_info(sp)
                if phtype is None and idx is None:
                    continue
                sp_pr = find_child(sp, "p", "spPr")
                xfrm = None
                if sp_pr is not None:
                    xfrm = read_xfrm(sp_pr)
                if xfrm is None:
                    xfrm = read_xfrm(find_child(sp, "p", "spPr"))
                key = (phtype, idx)
                result.setdefault(key, xfrm)
                if idx is not None:
                    result.setdefault((phtype, None), xfrm)
        self._ph_cache[part] = result
        return result

    def _build_geom_cache(self):
        # 预扫描版式与母版，后续 resolve_geom 只用缓存，避免反复解析
        for part in self.pkg.iter_parts(r"ppt/slideLayouts/slideLayout\d+\.xml$"):
            self._collect_placeholders(part)
        for part in self.pkg.iter_parts(r"ppt/slideMasters/slideMaster\d+\.xml$"):
            self._collect_placeholders(part)

    def resolve_geom(self, slide_part, sp):
        sp_pr = find_child(sp, "p", "spPr")
        geom = read_xfrm(sp_pr) if sp_pr is not None else None
        if geom is not None:
            return geom
        if localname(sp.tag) == "graphicFrame":
            xfrm = find_child(sp, "p", "xfrm")
            if xfrm is not None:
                off = find_child(xfrm, "a", "off")
                ext = find_child(xfrm, "a", "ext")
                if off is not None and ext is not None:
                    return {"x": get_int(off, "x", 0), "y": get_int(off, "y", 0),
                            "cx": get_int(ext, "cx", 0), "cy": get_int(ext, "cy", 0)}
        phtype, idx = placeholder_info(sp)
        if phtype is None and idx is None:
            return None
        chain = [slide_part]
        layout = self.layout_of(slide_part)
        if layout:
            chain.append(layout)
        master = self.master_of(layout) if layout else None
        if master:
            chain.append(master)
        for part in chain:
            table = self._collect_placeholders(part)
            for key in ((phtype, idx), (phtype, None), (None, idx), (None, None)):
                g = table.get(key)
                if g:
                    return dict(g)
        return None

    # -- 母版文本样式 -------------------------------------------------------
    def master_tx_styles(self, slide_part):
        layout = self.layout_of(slide_part)
        master = self.master_of(layout) if layout else None
        if master is None:
            return {}
        if master in self._txstyle_cache:
            return self._txstyle_cache[master]
        styles = {}
        root = self.pkg.xml(master).root
        tx = find_child(root, "p", "txStyles")
        if tx is not None:
            for name in ("titleStyle", "bodyStyle", "otherStyle"):
                holder = find_child(tx, "p", name)
                if holder is not None:
                    styles[name] = holder
        self._txstyle_cache[master] = styles
        return styles


# ---- 样式继承 ---------------------------------------------------------------

def style_from_defrpr(defrpr):
    st = {}
    if defrpr is None:
        return st
    for key in ("sz", "b", "i", "u", "spc", "baseline", "lang", "altLang"):
        v = defrpr.get(key)
        if v is not None:
            st[key] = v
    latin = find_child(defrpr, "a", "latin")
    ea = find_child(defrpr, "a", "ea")
    cs = find_child(defrpr, "a", "cs")
    sym = find_child(defrpr, "a", "sym")
    if latin is not None and latin.get("typeface"):
        st["latin"] = latin.get("typeface")
    if ea is not None and ea.get("typeface") is not None:
        st["ea"] = ea.get("typeface")
    if cs is not None and cs.get("typeface"):
        st["cs"] = cs.get("typeface")
    if sym is not None and sym.get("typeface"):
        st["sym"] = sym.get("typeface")
    return st


def style_from_rpr(rpr):
    return style_from_defrpr(rpr)


def merge_style(base, extra):
    out = dict(base)
    for k, v in extra.items():
        if v is None:
            continue
        if k in ("latin", "ea", "cs", "sym") and v == "":
            out[k] = ""
            continue
        out[k] = v
    return out


THEME_FONT_TOKENS = {
    "+mn-lt": ("minor_latin", "latin"),
    "+mj-lt": ("major_latin", "latin"),
    "+mn-ea": ("minor_ea", "ea"),
    "+mj-ea": ("major_ea", "ea"),
    "+mn-cs": ("minor_latin", "cs"),
    "+mj-cs": ("major_latin", "cs"),
}


class ShapeStyleContext(object):
    """解析某个形状内文本的“有效字体 / 字号 / 行距”。"""

    def __init__(self, deck, part, theme_fonts, sp, tx_body):
        self.deck = deck
        self.part = part
        self.theme = theme_fonts
        self.sp = sp
        self.tx_body = tx_body
        self.ph_type, self.ph_idx = placeholder_info(sp)
        self.theme_ea_ok = bool(
            (theme_fonts.minor_ea or "").strip() or (theme_fonts.major_ea or "").strip()
            or theme_fonts.script_minor.get("Hans")
            or theme_fonts.script_major.get("Hans"))
        self.base_defaults = self._build_defaults()

    def _master_style_holder(self):
        styles = self.deck.master_tx_styles(self.part)
        if self.ph_type in ("title", "ctrTitle"):
            return styles.get("titleStyle")
        if self.ph_type in (None,):
            return styles.get("otherStyle")
        if self.ph_type in ("body", "obj", "subTitle", "tbl", "chart", "dgm",
                            "media", "pic", "clipArt"):
            return styles.get("bodyStyle")
        return styles.get("otherStyle")

    def _lvl_defrpr(self, holder, lvl):
        if holder is None:
            return None
        want = "lvl%dpPr" % (lvl + 1)
        for el in holder:
            if localname(el.tag) == want:
                return find_child(el, "a", "defRPr")
        return None

    def _build_defaults(self):
        st = {
            "sz": "1800", "latin": "", "ea": "", "cs": "", "lang": None,
            "altLang": None, "spc": "0", "baseline": None,
        }
        master_holder = self._master_style_holder()
        st = merge_style(st, style_from_defrpr(self._lvl_defrpr(master_holder, 0)))
        # 形状自身 lstStyle 的第 1 级
        lst = find_child(self.tx_body, "a", "lstStyle")
        st = merge_style(st, style_from_defrpr(self._lvl_defrpr(lst, 0)))
        # 非占位符文本：库默认样式兜底
        if self.ph_type is None:
            pres_holder = self.deck.default_text_style
            st = merge_style(st, style_from_defrpr(self._lvl_defrpr(pres_holder, 0)))
        return st

    def level_style(self, lvl):
        st = dict(self.base_defaults)
        master_holder = self._master_style_holder()
        st = merge_style(st, style_from_defrpr(self._lvl_defrpr(master_holder, lvl)))
        lst = find_child(self.tx_body, "a", "lstStyle")
        st = merge_style(st, style_from_defrpr(self._lvl_defrpr(lst, lvl)))
        if self.ph_type is None:
            pres_holder = self.deck.default_text_style
            st = merge_style(st, style_from_defrpr(self._lvl_defrpr(pres_holder, lvl)))
        return st

    def paragraph_style(self, p):
        ppr = find_child(p, "a", "pPr")
        lvl = get_int(ppr, "lvl", 0) or 0
        st = self.level_style(lvl)
        if ppr is not None:
            st = merge_style(st, style_from_defrpr(find_child(ppr, "a", "defRPr")))
        return st

    def run_style(self, para_style, r_el):
        rpr = find_child(r_el, "a", "rPr") if r_el is not None else None
        st = merge_style(para_style, style_from_rpr(rpr))
        return self.resolve_fonts(st, rpr)

    def resolve_fonts(self, st, rpr=None):
        """把 +mn-lt 之类的主题字体引用解析成真实字体名。"""
        for token, (theme_attr, slot) in THEME_FONT_TOKENS.items():
            for k in ("latin", "ea", "cs"):
                if st.get(k) == token:
                    face = getattr(self.theme, theme_attr, "") or ""
                    if not face:
                        if slot == "ea":
                            face = (self.theme.script_minor.get("Hans")
                                    or self.theme.minor_ea or self.deck.default_cjk)
                        else:
                            face = self.deck.default_latin
                    st[k] = face
        # ea 缺失时：若 latin 本身是 CJK 字体则沿用，否则用主题/默认中文
        if not st.get("ea"):
            if classify_family(st.get("latin") or "").startswith("cjk"):
                st["ea"] = st["latin"]
            else:
                st["ea"] = (self.theme.minor_ea
                            or self.theme.script_minor.get("Hans")
                            or self.deck.default_cjk)
        if not st.get("latin"):
            st["latin"] = (self.theme.minor_latin or self.deck.default_latin)
        return st


# ---- 行距 / 段距 ------------------------------------------------------------

def parse_lnspc(ppr, style_sz):
    if ppr is None:
        return None
    holder = find_child(ppr, "a", "lnSpc")
    if holder is None:
        return None
    pct = find_child(holder, "a", "spcPct")
    if pct is not None:
        v = get_int(pct, "val", 100000)
        return ("pct", v)
    pts = find_child(holder, "a", "spcPts")
    if pts is not None:
        return ("pts", get_int(pts, "val", 0) / 100.0)
    return None


def parse_spc(ppr, tag, style_sz):
    if ppr is None:
        return 0.0
    holder = find_child(ppr, "a", tag)
    if holder is None:
        return 0.0
    pct = find_child(holder, "a", "spcPct")
    if pct is not None:
        v = get_int(pct, "val", 0)
        return (v / 100000.0) * (style_sz / 100.0) * 1.2
    pts = find_child(holder, "a", "spcPts")
    if pts is not None:
        return get_int(pts, "val", 0) / 100.0
    return 0.0


# ============================================================================
# 文本体：段落 / 运行 / 度量
# ============================================================================

def walk_tree_elements(sp_tree, out, depth=0):
    """收集 spTree 下的形状、表格、组合，用于逐项检查。"""
    if sp_tree is None or depth > 12:
        return
    for el in sp_tree:
        ln = localname(el.tag)
        if ln == "grpSp":
            out.append(("group", el))
            walk_tree_elements(el, out, depth + 1)
        elif ln in ("sp", "cxnSp"):
            out.append(("shape", el))
        elif ln == "graphicFrame":
            if find_child(el, "a", "graphic") is not None:
                out.append(("graphicFrame", el))


def table_of(graphic_frame):
    graphic = find_child(graphic_frame, "a", "graphic")
    if graphic is None:
        return None
    data = find_child(graphic, "a", "graphicData")
    if data is None:
        return None
    return find_child(data, "a", "tbl")


def read_bodypr(tx_body):
    if tx_body is None:
        return None
    return find_child(tx_body, "a", "bodyPr")


def body_insets(bodypr):
    return (
        get_int(bodypr, "lIns", DEFAULT_INSET_L),
        get_int(bodypr, "rIns", DEFAULT_INSET_R),
        get_int(bodypr, "tIns", DEFAULT_INSET_T),
        get_int(bodypr, "bIns", DEFAULT_INSET_B),
    )


def autofit_state(bodypr):
    """返回 (kind, font_scale, lnspc_reduction)。kind: none/norm/spauto"""
    if bodypr is None:
        return "none", 1.0, 0.0
    norm = find_child(bodypr, "a", "normAutofit")
    if norm is not None:
        scale = get_int(norm, "fontScale", 100000) / 100000.0
        red = get_int(norm, "lnSpcReduction", 0) / 100000.0
        return "norm", scale, red
    if find_child(bodypr, "a", "spAutoFit") is not None:
        return "spauto", 1.0, 0.0
    return "none", 1.0, 0.0


def build_paragraphs(style_ctx, tx_body, font_scale=1.0):
    """把 txBody 解析成可度量的段落结构。"""
    deck = style_ctx.deck
    registry = deck.registry
    paras = []
    if tx_body is None:
        return paras
    for p in tx_body:
        if localname(p.tag) != "p":
            continue
        ppr = find_child(p, "a", "pPr")
        para_style = style_ctx.paragraph_style(p)
        base_size = get_int(para_style, "sz", 1800) / 100.0
        lnspc = parse_lnspc(ppr, base_size)
        spc_bef = parse_spc(ppr, "spcBef", base_size)
        spc_aft = parse_spc(ppr, "spcAft", base_size)
        mar_l = (get_int(ppr, "marL", 0) or 0) / float(EMU_PER_PT)
        indent = (get_int(ppr, "indent", 0) or 0) / float(EMU_PER_PT)
        segments = [[]]
        max_size = base_size
        for child in p:
            ln = localname(child.tag)
            if ln in ("r", "fld"):
                try:
                    run_style = style_ctx.run_style(para_style, child)
                except Exception:  # pragma: no cover - 防御畸形文件
                    run_style = para_style
                text = text_of(child)
                size = get_int(run_style, "sz", 1800) / 100.0 * font_scale
                max_size = max(max_size, size)
                spc = get_int(run_style, "spc", 0) / 100.0
                latin = run_style.get("latin") or deck.default_latin
                ea = run_style.get("ea") or deck.default_cjk
                for ch in text:
                    if ch == "\n":
                        segments.append([])
                        continue
                    if ch == "\t":
                        ch = " "
                    name = ea if is_cjk_char(ch) else latin
                    font = registry.font_for(name) if registry is not None else None
                    segments[-1].append((ch, font, size, spc))
            elif ln == "br":
                segments.append([])
        paras.append({
            "segments": segments,
            "lnspc": lnspc,
            "spc_bef": spc_bef,
            "spc_aft": spc_aft,
            "marL": mar_l,
            "indent": indent,
            "algn": (ppr.get("algn") if ppr is not None else None),
            "empty_line": max_size * 1.2,
        })
    return paras


def measure_body(paras, inner_w_pt, wrap, lnspc_reduction, registry,
                 fallback_cjk=None, fallback_latin=None):
    """估算整个文本框所需高度与最大行宽（pt）。"""
    total_h = 0.0
    max_w = 0.0
    for para in paras:
        segs = para["segments"]
        n = len(segs)
        for i, seg in enumerate(segs):
            avail = inner_w_pt - para["marL"]
            if i == 0:
                avail -= para["indent"]
            avail = max(avail, 10.0)
            m = measure_paragraph(
                seg, avail, wrap, para["lnspc"],
                para["spc_bef"] if i == 0 else 0.0,
                para["spc_aft"] if i == n - 1 else 0.0,
                lnspc_reduction, fallback_cjk, fallback_latin,
                empty_line_pt=para["empty_line"])
            total_h += m.height_pt
            max_w = max(max_w, m.max_line_width_pt)
    return total_h, max_w


# ============================================================================
# 规则定义
# ============================================================================

RULES = [
    {
        "id": "R001", "group": "font", "severity": "high", "default_on": True,
        "name": "引用了本机不存在的字体",
        "desc": "Windows 字体（宋体/微软雅黑/Calibri 等）在 UOS/麒麟 上通常缺失，"
                "WPS 会用 fontconfig 回退字体，字宽、行高、字形全部改变，"
                "导致换行、溢出色块、行距异常。这是跨系统错位的头号原因。",
        "fix": "把字体名改写为本机已安装的度量/风格最接近字体（宋体→思源宋体，"
               "微软雅黑→思源黑体，Calibri→Carlito，Arial→Liberation Sans …）；"
               "或改用“字体方案”安装原版 Windows 字体 / 生成 fontconfig 别名。",
    },
    {
        "id": "R002", "group": "font", "severity": "high", "default_on": True,
        "name": "中文字符未显式指定 East Asian 字体",
        "desc": "运行属性里只有 <a:latin> 没有 <a:ea>，中文字符会落到主题/系统默认字体，"
                "而 PowerPoint 与 WPS 的默认回退链不同，位置自然对不上；"
                "另外中文 Windows 字体自带西文字形，WPS 回退后西文字宽变化也会改变断行。",
        "fix": "为含中文的运行补写 <a:ea>；若该字体同时承担西文，把 <a:latin> 拆成"
               "度量兼容的西文字体（如 Liberation Serif），中英文各自稳定。",
    },
    {
        "id": "R003", "group": "font", "severity": "mid", "default_on": True,
        "name": "主题字体缺少中文（Hans）映射",
        "desc": "theme 的 majorFont/minorFont 中 <a:ea> 为空或没有 script=\"Hans\" 条目时，"
                "中文会使用各实现自己的默认字体，Windows 与 WPS 取值不同。",
        "fix": "补写 <a:ea> 与 <a:font script=\"Hans\"> 映射到本机中文字体。",
    },
    {
        "id": "R004", "group": "font", "severity": "low", "default_on": True,
        "name": "演示文稿内嵌字体（WPS 会忽略）",
        "desc": "PPTX 可以嵌入字体，但 WPS for Linux 基本不读取 ppt/fonts/*.fntdata，"
                "嵌入式字体形同虚设，仍会走系统字体替换。",
        "fix": "无法把嵌入字体“解出来”直接生效，建议安装原版字体；"
               "或使用本工具替换字体名（R001）。",
    },
    {
        "id": "R005", "group": "font", "severity": "high", "default_on": True,
        "name": "符号字体缺失（Wingdings/Webdings/Symbol/Marlett）",
        "desc": "这些字体用“字母映射图形”的方式显示 ✓ → ★ 等符号。Linux 上没有这些字体时，"
                "同一个字符码会显示成字母，看起来就像“文字错位/乱码”。",
        "fix": "不要用普通字体替换（会变成字母）。请安装原版符号字体，"
               "或在 WPS 中把这些符号改成普通字符/图片。",
    },
    {
        "id": "R006", "group": "layout", "severity": "mid", "default_on": True,
        "name": "文本框内边距未显式设置（依赖应用默认值）",
        "desc": "lIns/rIns/tIns/bIns 缺省时，不同实现使用的默认内边距可能不一致，"
                "文字相对色块整体偏移、可用宽度变化导致换行位置不同。",
        "fix": "显式写入 PowerPoint 默认值（左右 0.1 英寸、上下 0.05 英寸），"
               "让不同系统上的排版基准一致。",
    },
    {
        "id": "R007", "group": "layout", "severity": "high", "default_on": True,
        "name": "PowerPoint 自动缩放残留（normAutofit）",
        "desc": "PowerPoint 会在保存时把“溢出时缩排文字”的结果写成 fontScale/lnSpcReduction；"
                "WPS 会按自己的字体度量重新计算或直接按比例套用，结果与 Windows 不一致。",
        "fix": "按当前字体的实际度量重新估算：不需要缩放就清除残留值（还原 100%），"
               "仍然溢出则写入实测需要的缩放比例。",
    },
    {
        "id": "R008", "group": "layout", "severity": "mid", "default_on": False,
        "name": "形状随文字自动调整高度（spAutoFit）",
        "desc": "spAutoFit 让形状高度由文字决定，WPS 与 PowerPoint 的字体度量不同，"
                "重算后的高度不同，形状会上下“长高/缩短”，压到相邻元素。",
        "fix": "改为 noAutofit，固定当前形状尺寸（文字可能略微溢出，"
               "可配合 R009 缩小字号）。会改变版式行为，默认不启用。",
    },
    {
        "id": "R009", "group": "layout", "severity": "high", "default_on": False,
        "name": "文本存在溢出风险（按 Linux 实际字体度量估算）",
        "desc": "用 WPS 实际会使用的替代字体度量估算：文字所需高度/宽度超过文本框。"
                "这类形状在 Windows 上刚好放得下，换字体后就溢出、压线、错位。",
        "fix": "写入实测需要的 normAutofit 缩放比例（不小于 25%），"
               "把文字缩到框内；会轻微改变字号，默认不启用。",
    },
    {
        "id": "R010", "group": "layout", "severity": "mid", "default_on": True,
        "name": "不换行文本框（wrap=\"none\"）宽度不足",
        "desc": "wrap=\"none\" 的文本框不自动换行，宽度完全依赖字体宽度。"
                "字体变宽后文字会越过文本框，看起来整体错位。",
        "fix": "按实测文本宽度以中心为基准扩宽形状（不超过页面边界）；"
               "若页面放不下则改为允许换行。",
    },
    {
        "id": "R011", "group": "layout", "severity": "mid", "default_on": True,
        "name": "表格单元格边距未显式设置",
        "desc": "表格单元格 marL/marR/marT/marB 缺省时，各实现的默认值不同，"
                "单元格文字会整体偏移，并改变换行位置与行高。",
        "fix": "显式写入 PowerPoint 默认单元格边距。",
    },
    {
        "id": "R012", "group": "layout", "severity": "mid", "default_on": False,
        "name": "表格行高不足以容纳文本",
        "desc": "固定行高 + 字体变宽变高后，WPS 会把单元格文字挤出边界或强行压缩行距。",
        "fix": "按实测内容增高行高（可能使表格整体变高）。默认不启用。",
    },
    {
        "id": "R013", "group": "layout", "severity": "low", "default_on": True,
        "name": "中文文本的语言标记不是中文",
        "desc": "run 的 lang 为 en-US 或缺失时，WPS 可能按西文规则处理中文断行/标点"
                "（行首禁则、中英文间距），断行位置与 Windows 不同。",
        "fix": "按内容把中文 run 的 lang 标记为 zh-CN（日文/韩文保持各自语言）。",
    },
    {
        "id": "R014", "group": "layout", "severity": "mid", "default_on": False,
        "name": "行距依赖字体度量（单倍/百分比行距）",
        "desc": "“单倍行距”或百分比行距按字体自身的 ascent/descent 计算。"
                "宋体行高约 1.0~1.2em，而思源黑体/宋体约 1.44em——换字体后"
                "整段文字会被“撑高”20%~40%，多行文本必然下移。",
        "fix": "把段落行距固定为磅值（默认 1.2 × 字号，可调），"
               "使行高不再随字体度量变化。会改变视觉行距，默认不启用。",
    },
    {
        "id": "R015", "group": "info", "severity": "low", "default_on": True,
        "name": "存在手动换行（a:br）",
        "desc": "手动换行是硬回车，字体宽度变化后无法自动重排，"
                "会表现为“断行位置奇怪/半行空白”。",
        "fix": "需人工在 WPS 中确认（本工具不自动删除硬回车）。",
    },
    {
        "id": "R016", "group": "info", "severity": "low", "default_on": True,
        "name": "组合形状的缩放比例不一致",
        "desc": "组合外框与子坐标空间比例不一致时，组合内文字被非等比缩放；"
                "WPS 与 PowerPoint 对组内文字缩放的处理不同，位置/字宽会偏移。",
        "fix": "建议在 WPS 里取消组合后重新组合，本工具只提示。",
    },
    {
        "id": "R017", "group": "info", "severity": "low", "default_on": True,
        "name": "竖排文字 / 文字变形（WordArt）",
        "desc": "竖排（vert=eaVert/vert270）与文字变形（prstTxWarp）依赖具体字体与渲染器，"
                "跨系统差异大，容易出现字距/位置偏移。",
        "fix": "如需完全一致，建议改为横排普通文本；本工具只提示。",
    },
    {
        "id": "R018", "group": "info", "severity": "high", "default_on": True,
        "name": "旧版二进制格式（.ppt/.dps/.wps）",
        "desc": "旧格式没有 OOXML 结构，无法做精细修复，且在 WPS/Linux 上兼容性最差。",
        "fix": "先用 WPS 或 LibreOffice 另存为 .pptx，再运行本工具。",
    },
    {
        "id": "R019", "group": "font", "severity": "mid", "default_on": False,
        "name": "替换字体缺少部分字形",
        "desc": "目标字体没有某些生僻字/繁体字/符号的字形时，渲染器会逐字回退，"
                "导致同一行内字宽突变、断行错位。",
        "fix": "改写为覆盖更全的中文字体（如思源黑体/思源宋体）。默认不启用。",
    },
    {
        "id": "R020", "group": "info", "severity": "info", "default_on": True,
        "name": "幻灯片尺寸非标准比例",
        "desc": "页面尺寸不是标准 16:9 / 4:3 时，WPS 可能启用“适应页面”缩放，"
                "整页元素等比缩放后看起来像“位置全偏了”。",
        "fix": "在 WPS 页面设置里确认尺寸与缩放比例；本工具只提示。",
    },
]
RULES_BY_ID = {r["id"]: r for r in RULES}
DEFAULT_RULES = {r["id"] for r in RULES if r["default_on"]}
SEVERITY_RANK = {"high": 0, "mid": 1, "low": 2, "info": 3}
SEVERITY_LABEL = {"high": "高", "mid": "中", "low": "低", "info": "提示"}


# ============================================================================
# 检测 / 修复引擎
# ============================================================================

class Finding(object):
    __slots__ = ("rule_id", "severity", "part", "slide_no", "shape_name",
                 "detail", "count", "fixed", "examples")

    def __init__(self, rule_id, part, detail="", count=1, fixed=0,
                 slide_no=None, shape_name=None, examples=None):
        self.rule_id = rule_id
        self.severity = RULES_BY_ID[rule_id]["severity"]
        self.part = part
        self.slide_no = slide_no
        self.shape_name = shape_name
        self.detail = detail
        self.count = count
        self.fixed = fixed
        self.examples = examples or []

    def location(self):
        bits = []
        if self.slide_no is not None:
            bits.append("第 %d 页" % self.slide_no)
        if self.shape_name:
            bits.append("「%s」" % self.shape_name)
        if not bits:
            bits.append(posixpath.basename(self.part))
        return " / ".join(bits)

    def to_dict(self):
        return {
            "rule": self.rule_id,
            "name": RULES_BY_ID[self.rule_id]["name"],
            "severity": self.severity,
            "part": self.part,
            "slide": self.slide_no,
            "shape": self.shape_name,
            "location": self.location(),
            "detail": self.detail,
            "count": self.count,
            "fixed": self.fixed,
            "examples": self.examples,
        }


class Change(object):
    __slots__ = ("rule_id", "part", "slide_no", "shape_name", "detail")

    def __init__(self, rule_id, part, detail, slide_no=None, shape_name=None):
        self.rule_id = rule_id
        self.part = part
        self.slide_no = slide_no
        self.shape_name = shape_name
        self.detail = detail

    def to_dict(self):
        return {"rule": self.rule_id, "part": self.part, "slide": self.slide_no,
                "shape": self.shape_name, "detail": self.detail}


def _set_run_font(rpr, slot, value):
    """在 rPr 中安全地写入/更新 <a:latin>/<a:ea>/<a:cs>。"""
    el = find_child(rpr, "a", slot)
    if el is None:
        el = ET.Element(qn("a", slot))
        insert_ordered(rpr, el, RPR_ORDER)
    if (el.get("typeface") or "") == value:
        return False
    el.set("typeface", value)
    return True


def _ensure_rpr(run):
    rpr = find_child(run, "a", "rPr")
    if rpr is None:
        rpr = ET.Element(qn("a", "rPr"))
        run.insert(0, rpr)
    return rpr


def _replace_body_autofit(bodypr, kind, font_scale=None, lnspc_red=None):
    """把 bodyPr 的自动调整子元素替换成指定类型。"""
    for tag in ("noAutofit", "normAutofit", "spAutoFit"):
        remove_child(bodypr, "a", tag)
    if kind == "norm":
        el = ET.Element(qn("a", "normAutofit"))
        if font_scale is not None:
            el.set("fontScale", str(int(font_scale)))
        if lnspc_red is not None:
            el.set("lnSpcReduction", str(int(lnspc_red)))
        insert_ordered(bodypr, el, BODYPR_ORDER)
        return el
    if kind == "spauto":
        el = ET.Element(qn("a", "spAutoFit"))
    else:
        el = ET.Element(qn("a", "noAutofit"))
    insert_ordered(bodypr, el, BODYPR_ORDER)
    return el


class Engine(object):
    def __init__(self, deck, enabled=None, apply=False, options=None):
        self.deck = deck
        self.pkg = deck.pkg
        self.registry = deck.registry
        self.resolver = deck.resolver
        self.enabled = set(DEFAULT_RULES if enabled is None else enabled)
        self.apply = apply
        self.options = options or {}
        self.line_ratio = float(self.options.get("line_ratio", 1.2))
        self.shrink_min = float(self.options.get("shrink_min", 0.25))
        self.findings = []
        self.changes = []
        self.rewritten_fonts = {}
        self.symbol_fonts = {}
        self.metrics = {
            "slides": len(deck.slides), "shapes": 0, "measured": 0,
            "overflow": 0, "worst_ratio": 1.0, "runs": 0,
            "hard_breaks": 0, "tables": 0,
        }
        self._fb_cjk = None
        self._fb_latin = None
        if self.registry is not None and self.registry.available:
            self._fb_cjk = self.registry.font_for(self.deck.default_cjk)
            self._fb_latin = self.registry.font_for(self.deck.default_latin)

    # -- 小工具 -------------------------------------------------------------
    def on(self, rid):
        return rid in self.enabled

    def _touch(self, part):
        self.pkg.mark(part)

    def _finding(self, rid, part, detail="", count=1, fixed=0, slide_no=None,
                 shape_name=None, examples=None):
        self.findings.append(Finding(rid, part, detail, count, fixed,
                                     slide_no, shape_name, examples))

    def _change(self, rid, part, detail, slide_no=None, shape_name=None):
        self.changes.append(Change(rid, part, detail, slide_no, shape_name))

    def _font_parts(self):
        for name in self.pkg.names:
            if not name.startswith("ppt/"):
                continue
            if not name.endswith(".xml"):
                continue
            if "/_rels/" in name:
                continue
            yield name

    # -- 主流程 -------------------------------------------------------------
    def run(self):
        self.check_format()
        self.check_embedded_fonts()
        self.check_fonts()
        self.check_theme_fonts()
        self.normalize_runs()
        self.check_slides()
        self.check_global()
        return self

    # -- R018 文件格式 ------------------------------------------------------
    def check_format(self):
        ext = os.path.splitext(self.pkg.path)[1].lower()
        if ext in (".ppt", ".dps", ".wps", ".dpt"):
            self._finding("R018", self.pkg.path,
                          "文件为旧版二进制格式：%s" % ext, count=1)
            return
        if ext not in (".pptx", ".potx", ".ppsx", ".pptm", ".potm", ".ppsm"):
            self._finding("R018", self.pkg.path,
                          "未知扩展名：%s，按 OOXML 尝试处理" % ext, count=1)

    # -- R004 内嵌字体 ------------------------------------------------------
    def check_embedded_fonts(self):
        if not self.deck.embedded_fonts:
            return
        faces = sorted(set(self.deck.embedded_fonts))
        self._finding("R004", "ppt/presentation.xml",
                      "内嵌字体 %d 个，WPS for Linux 一般不会加载" % len(faces),
                      count=len(faces), examples=faces[:8])

    # -- R001 / R005 字体缺失 ----------------------------------------------
    def check_fonts(self):
        missing = {}
        for part in self._font_parts():
            doc = self.pkg.xml(part)
            changed = False
            # 内嵌字体（p:embeddedFontLst）只提示（R004），不要改写它的字体名，
            # 否则会与其中的字体数据对不上。
            embedded = set()
            for holder in doc.root.iter():
                if localname(holder.tag) == "embeddedFontLst":
                    for sub in holder.iter():
                        embedded.add(id(sub))
            for el in doc.root.iter():
                if id(el) in embedded:
                    continue
                ln = localname(el.tag)
                if ln not in ("latin", "ea", "cs", "sym", "font", "buFont"):
                    continue
                face = el.get("typeface")
                if not face or face.startswith("+"):
                    continue
                decision = self.resolver.decide(face)
                if decision.status == "symbol":
                    rec = self.symbol_fonts.setdefault(face, {"count": 0, "parts": set()})
                    rec["count"] += 1
                    rec["parts"].add(part)
                    continue
                if decision.status != "replace":
                    continue
                rec = missing.setdefault(face, {
                    "count": 0, "part_set": set(), "target": decision.target,
                    "fixed": 0,
                })
                rec["count"] += 1
                rec["part_set"].add(part)
                if self.apply and self.on("R001"):
                    el.set("typeface", decision.target)
                    if el.get("panose") is not None:
                        del el.attrib["panose"]
                    changed = True
                    self.rewritten_fonts[face] = decision.target
                    rec["fixed"] += 1
            if changed:
                self._touch(part)
        for face, rec in sorted(missing.items(), key=lambda kv: -kv[1]["count"]):
            self._finding(
                "R001", sorted(rec["part_set"])[0],
                "字体「%s」本机不存在（出现 %d 处）→ 替代为「%s」"
                % (face, rec["count"], rec["target"]),
                count=rec["count"], fixed=rec["fixed"], examples=[face])
        for face, rec in sorted(self.symbol_fonts.items()):
            self._finding(
                "R005", sorted(rec["parts"])[0],
                "符号字体「%s」缺失（出现 %d 处）：字符会显示成字母或方框，不能自动替换"
                % (face, rec["count"]),
                count=rec["count"], examples=[face])

    # -- R003 主题字体 ------------------------------------------------------
    def check_theme_fonts(self):
        tc_cjk = None
        if self.registry is not None and self.registry.available:
            tc_cjk = self.registry.first_installed(
                ["Noto Serif CJK TC", "Noto Sans CJK TC", "Noto Serif CJK SC",
                 "Noto Sans CJK SC"])
        for part in self.pkg.iter_parts(r"ppt/theme/theme\d+\.xml$"):
            doc = self.pkg.xml(part)
            issues = []
            changed = False
            for scheme in doc.root.iter():
                if localname(scheme.tag) != "fontScheme":
                    continue
                for holder_tag in ("majorFont", "minorFont"):
                    holder = find_child(scheme, "a", holder_tag)
                    if holder is None:
                        continue
                    ea = find_child(holder, "a", "ea")
                    if ea is None:
                        issues.append("%s 缺少 <a:ea>" % holder_tag)
                        if self.apply and self.on("R003"):
                            ea = ET.Element(qn("a", "ea"))
                            insert_ordered(holder, ea, ["latin", "ea", "cs"])
                            changed = True
                    if ea is not None and not (ea.get("typeface") or "").strip():
                        issues.append("%s 的 <a:ea> 为空" % holder_tag)
                        if self.apply and self.on("R003"):
                            ea.set("typeface", self.deck.default_cjk)
                            changed = True
                    hans = None
                    hant = None
                    for f in holder:
                        if localname(f.tag) != "font":
                            continue
                        if f.get("script") == "Hans":
                            hans = f
                        elif f.get("script") == "Hant":
                            hant = f
                    if hans is None:
                        issues.append("%s 缺少 script=\"Hans\" 映射" % holder_tag)
                        if self.apply and self.on("R003"):
                            f = ET.Element(qn("a", "font"))
                            f.set("script", "Hans")
                            f.set("typeface", self.deck.default_cjk)
                            _append_before_extlst(holder, f)
                            changed = True
                    elif not (hans.get("typeface") or "").strip():
                        issues.append("%s 的 Hans 映射为空" % holder_tag)
                        if self.apply and self.on("R003"):
                            hans.set("typeface", self.deck.default_cjk)
                            changed = True
                    if hant is None and tc_cjk:
                        issues.append("%s 缺少 script=\"Hant\" 映射" % holder_tag)
                        if self.apply and self.on("R003"):
                            f = ET.Element(qn("a", "font"))
                            f.set("script", "Hant")
                            f.set("typeface", tc_cjk)
                            _append_before_extlst(holder, f)
                            changed = True
                break
            if issues:
                fixed = len(issues) if (self.apply and self.on("R003")) else 0
                self._finding("R003", part, "；".join(issues), count=len(issues),
                              fixed=fixed)
                if fixed and changed:
                    self._touch(part)
                    self._change("R003", part,
                                 "补写主题中文字体映射（%d 处）" % fixed)

    # -- R002 / R013 / R019 运行级字体与语言 --------------------------------
    def normalize_runs(self):
        stats = {"R002": 0, "R013": 0, "R019": 0}
        examples = {"R002": [], "R013": [], "R019": []}
        parts_touched = {}
        for slide_part in self.deck.slides:
            slide_no = self.deck.slide_number.get(slide_part)
            theme = self.deck.theme_fonts_for_slide(slide_part)
            doc = self.pkg.xml(slide_part)
            tree = _sp_tree(doc.root)
            items = []
            walk_tree_elements(tree, items)
            for kind, el in items:
                if kind != "shape":
                    continue
                tx_body = find_child(el, "p", "txBody")
                if tx_body is None:
                    continue
                name = shape_name(el)
                ctx = ShapeStyleContext(self.deck, slide_part, theme, el, tx_body)
                changed = False
                for p in tx_body:
                    if localname(p.tag) != "p":
                        continue
                    para_style = ctx.paragraph_style(p)
                    for run in p:
                        if localname(run.tag) not in ("r", "fld"):
                            continue
                        text = text_of(run)
                        if not text:
                            continue
                        self.metrics["runs"] += 1
                        try:
                            style = ctx.run_style(para_style, run)
                        except Exception:  # pragma: no cover
                            continue
                        res = self._normalize_one_run(run, text, style, ctx,
                                                      slide_part, slide_no, name)
                        if res:
                            changed = True
                            for rid in res:
                                stats[rid] += 1
                                if len(examples[rid]) < 5:
                                    snippet = text.strip()[:24]
                                    examples[rid].append("第%s页/%s：%s"
                                                         % (slide_no, name, snippet))
                if changed:
                    self._touch(slide_part)
                    parts_touched[slide_part] = parts_touched.get(slide_part, 0) + 1
        for rid, count in stats.items():
            if count <= 0:
                continue
            if rid == "R002":
                detail = "为中文运行补写/拆分中英文字体（%d 处）" % count
            elif rid == "R013":
                detail = "修正中文文本的语言标记（%d 处）" % count
            else:
                detail = "替换为字形覆盖更全的字体（%d 处）" % count
            fixed = count if (self.apply and self.on(rid)) else 0
            self._finding(rid, sorted(parts_touched)[0] if parts_touched else "",
                          detail, count=count, fixed=fixed,
                          examples=examples[rid])
            if fixed:
                self._change(rid, "", detail)

    def _normalize_one_run(self, run, text, style, ctx, slide_part, slide_no, name):
        changed_rules = []
        latin_name = style.get("latin") or self.deck.default_latin
        ea_name = style.get("ea") or self.deck.default_cjk
        latin_cls = classify_family(latin_name)
        has_cjk = any(is_cjk_char(ch) for ch in text)
        has_latin = any(ord(ch) < 128 and ch.isalnum() for ch in text)
        rpr = find_child(run, "a", "rPr")
        explicit_ea = find_child(rpr, "a", "ea") if rpr is not None else None
        explicit_latin = find_child(rpr, "a", "latin") if rpr is not None else None

        # R002：补 ea / 拆分中英文
        if has_cjk and self.on("R002"):
            ea_decision = self.resolver.decide(ea_name)
            target_ea = ea_decision.target or ea_name
            ea_missing = (explicit_ea is None
                          or not (explicit_ea.get("typeface") or "").strip())
            split_latin = has_latin and latin_cls.startswith("cjk")
            need_ea = ea_missing and (split_latin or not ctx.theme_ea_ok)
            if need_ea:
                if self.apply:
                    rpr = _ensure_rpr(run)
                    if _set_run_font(rpr, "ea", target_ea):
                        changed_rules.append("R002")
                else:
                    changed_rules.append("R002")
            if split_latin:
                comp = self.resolver.decide(latin_name).companion_latin
                if comp and comp != latin_name:
                    if self.apply:
                        rpr = _ensure_rpr(run)
                        if _set_run_font(rpr, "latin", comp):
                            changed_rules.append("R002")
                        _set_run_font(rpr, "ea", target_ea)
                    elif explicit_latin is None or explicit_latin.get("typeface") != comp:
                        changed_rules.append("R002")
        # R013：语言标记
        if has_cjk and self.on("R013"):
            lang = (style.get("lang") or "").strip()
            script = self._script_of(text)
            desired = {"ja": "ja-JP", "ko": "ko-KR", "zh": "zh-CN"}[script]
            if not lang.lower().startswith(script):
                if self.apply:
                    rpr = _ensure_rpr(run)
                    rpr.set("lang", desired)
                    if not rpr.get("altLang"):
                        rpr.set("altLang", "en-US")
                changed_rules.append("R013")
        # R019：字形覆盖（中文看 ea 字体，西文看 latin 字体）
        if self.on("R019") and self.registry is not None:
            checks = []
            if has_cjk:
                checks.append((ea_name,
                               "".join(ch for ch in text if is_cjk_char(ch))))
            if has_latin:
                checks.append((latin_name,
                               "".join(ch for ch in text
                                       if ord(ch) < 128 and not ch.isspace())))
            for font_name, chars in checks:
                if not chars:
                    continue
                missing = self.registry.coverage_missing(font_name, chars)
                if missing:
                    if self.apply:
                        fixed = self._fix_coverage(run, chars, font_name)
                        if fixed:
                            changed_rules.append("R019")
                    else:
                        changed_rules.append("R019")
        return sorted(set(changed_rules))

    @staticmethod
    def _script_of(text):
        if any(is_kana_char(ch) for ch in text):
            return "ja"
        if any(is_hangul_char(ch) for ch in text):
            return "ko"
        return "zh"

    def _fix_coverage(self, run, text, font_name):
        candidates = CJK_SANS + CJK_SERIF
        chosen = None
        for cand in candidates:
            if not self.registry.installed(cand):
                continue
            font = self.registry.font_for(cand)
            if font is None:
                continue
            if all(font.has_glyph(ch) for ch in text if not ch.isspace()):
                chosen = cand
                break
        if chosen is None:
            return False
        rpr = _ensure_rpr(run)
        # 缺字形的是中文字符时，换 ea 字体即可；西文字形缺失不能靠中文字体解决，
        # 这里只处理中文覆盖，避免把西文强行套上中文字体。
        return _set_run_font(rpr, "ea", chosen)


def elems(parent, prefix, tag):
    if parent is None:
        return []
    want = qn(prefix, tag)
    return [c for c in parent if c.tag == want]


def _sp_tree(root):
    for el in root.iter():
        if localname(el.tag) == "spTree":
            return el
    return None


def _ppr_of(p):
    return find_child(p, "a", "pPr")


def _ensure_ppr(p):
    ppr = find_child(p, "a", "pPr")
    if ppr is None:
        ppr = ET.Element(qn("a", "pPr"))
        p.insert(0, ppr)
    return ppr


class EngineGeometryMixin(object):
    """逐页、逐形状的版式检查与修复。"""

    # -- 需要外部提供的状态（见 Engine）------------------------------------

    def _measure(self, ctx, tx_body, scale, red, inner_w, wrap):
        paras = build_paragraphs(ctx, tx_body, font_scale=scale)
        return measure_body(paras, inner_w, wrap, red, self.registry,
                            self._fb_cjk, self._fb_latin)

    def _check_group(self, grp, slide_part, slide_no):
        if not self.on("R016"):
            return
        grp_pr = find_child(grp, "p", "grpSpPr")
        xfrm = find_child(grp_pr, "a", "xfrm") if grp_pr is not None else None
        if xfrm is None:
            return
        ext = find_child(xfrm, "a", "ext")
        ch_ext = find_child(xfrm, "a", "chExt")
        if ext is None or ch_ext is None:
            return
        cx, cy = get_int(ext, "cx", 0), get_int(ext, "cy", 0)
        ccx, ccy = get_int(ch_ext, "cx", 0), get_int(ch_ext, "cy", 0)
        if cx <= 0 or cy <= 0 or ccx <= 0 or ccy <= 0:
            return
        rx, ry = cx / float(ccx), cy / float(ccy)
        if abs(rx - 1.0) > 0.02 or abs(ry - 1.0) > 0.02:
            self._finding("R016", slide_part,
                          "组合「%s」缩放比例 x=%.2f / y=%.2f" % (shape_name(grp), rx, ry),
                          slide_no=slide_no, shape_name=shape_name(grp))

    def _check_table(self, frame, slide_part, slide_no, theme):
        tbl = table_of(frame)
        if tbl is None:
            return
        self.metrics["tables"] += 1
        grid = find_child(tbl, "a", "tblGrid")
        cols = [get_int(g, "w", 0) for g in elems(grid, "a", "gridCol")]
        default_col = cols[-1] if cols else 0
        name = shape_name(frame)
        missing_margins = 0
        cells = 0
        row_findings = []
        for tr in elems(tbl, "a", "tr"):
            row_h = get_int(tr, "h", 0)
            row_need = 0.0
            cells_in_row = 0
            for ci, tc in enumerate(elems(tr, "a", "tc")):
                tx_body = find_child(tc, "a", "txBody")
                if tx_body is None or not _has_visible_text(tx_body):
                    continue
                cells += 1
                cells_in_row += 1
                tc_pr = find_child(tc, "a", "tcPr")
                has_missing = any(
                    tc_pr is None or tc_pr.get(attr) is None
                    for attr in ("marL", "marR", "marT", "marB"))
                if has_missing:
                    missing_margins += 1
                    if self.apply and self.on("R011"):
                        if tc_pr is None:
                            tc_pr = ET.Element(qn("a", "tcPr"))
                            tc.insert(0, tc_pr)
                        for attr, val in zip(("marL", "marR", "marT", "marB"),
                                             DEFAULT_TABLE_CELL_MAR):
                            if tc_pr.get(attr) is None:
                                tc_pr.set(attr, str(val))
                ml = get_int(tc_pr, "marL", DEFAULT_TABLE_CELL_MAR[0]) if tc_pr is not None else DEFAULT_TABLE_CELL_MAR[0]
                mr = get_int(tc_pr, "marR", DEFAULT_TABLE_CELL_MAR[1]) if tc_pr is not None else DEFAULT_TABLE_CELL_MAR[1]
                mt = get_int(tc_pr, "marT", DEFAULT_TABLE_CELL_MAR[2]) if tc_pr is not None else DEFAULT_TABLE_CELL_MAR[2]
                mb = get_int(tc_pr, "marB", DEFAULT_TABLE_CELL_MAR[3]) if tc_pr is not None else DEFAULT_TABLE_CELL_MAR[3]
                col_w = cols[ci] if ci < len(cols) else default_col
                if col_w <= 0:
                    continue
                inner_w = (col_w - ml - mr) / float(EMU_PER_PT)
                if inner_w <= 4:
                    continue
                ctx = ShapeStyleContext(self.deck, slide_part, theme, frame, tx_body)
                bodypr = read_bodypr(tx_body)
                _kind, scale, red = autofit_state(bodypr)
                try:
                    need_h, _max_w = self._measure(ctx, tx_body, scale, red,
                                                   inner_w, True)
                except Exception:  # pragma: no cover
                    continue
                cell_need = need_h + (mt + mb) / float(EMU_PER_PT)
                row_need = max(row_need, cell_need)
            if row_h > 0 and row_need > row_h / float(EMU_PER_PT) * 1.03:
                row_findings.append((row_h, row_need))
                if self.apply and self.on("R012"):
                    tr.set("h", str(int(row_need * EMU_PER_PT)))
        if missing_margins:
            fixed = missing_margins if (self.apply and self.on("R011")) else 0
            self._finding("R011", slide_part,
                          "表格「%s」有 %d 个单元格缺少显式边距" % (name, missing_margins),
                          count=missing_margins, fixed=fixed,
                          slide_no=slide_no, shape_name=name)
            if fixed:
                self._touch(slide_part)
                self._change("R011", slide_part, "补写表格单元格边距（%d 格）" % fixed,
                             slide_no=slide_no, shape_name=name)
        if row_findings:
            fixed = len(row_findings) if (self.apply and self.on("R012")) else 0
            self._finding("R012", slide_part,
                          "表格「%s」有 %d 行高度不足（最大 %.1fpt < 需要 %.1fpt）"
                          % (name, len(row_findings),
                             row_findings[0][0] / float(EMU_PER_PT),
                             row_findings[0][1]),
                          count=len(row_findings), fixed=fixed,
                          slide_no=slide_no, shape_name=name)
            if fixed:
                self._touch(slide_part)
                self._change("R012", slide_part, "增高表格行（%d 行）" % fixed,
                             slide_no=slide_no, shape_name=name)

    def _check_shape_text(self, sp, tx_body, slide_part, slide_no, theme):
        name = shape_name(sp)
        if not _has_visible_text(tx_body):
            return
        bodypr = read_bodypr(tx_body)
        if bodypr is None:
            if self.apply:
                bodypr = ET.Element(qn("a", "bodyPr"))
                tx_body.insert(0, bodypr)
                self._touch(slide_part)
            else:
                bodypr = ET.Element(qn("a", "bodyPr"))

        # ---- R006 内边距 ----
        if self.on("R006"):
            missing = [a for a in ("lIns", "rIns", "tIns", "bIns")
                       if bodypr.get(a) is None]
            if missing:
                fixed = 0
                if self.apply:
                    defaults = {"lIns": DEFAULT_INSET_L, "rIns": DEFAULT_INSET_R,
                                "tIns": DEFAULT_INSET_T, "bIns": DEFAULT_INSET_B}
                    for attr in missing:
                        bodypr.set(attr, str(defaults[attr]))
                    self._touch(slide_part)
                    fixed = len(missing)
                self._finding("R006", slide_part,
                              "「%s」缺少内边距 %s" % (name, "/".join(missing)),
                              count=len(missing), fixed=fixed,
                              slide_no=slide_no, shape_name=name)
                if fixed:
                    self._change("R006", slide_part,
                                 "补写文本框内边距（%d 项）" % fixed,
                                 slide_no=slide_no, shape_name=name)

        # ---- R017 竖排 / 文字变形 ----
        vert = bodypr.get("vert")
        warp = find_child(bodypr, "a", "prstTxWarp")
        if (vert not in (None, "horz")) or warp is not None:
            kind = "竖排(%s)" % vert if vert not in (None, "horz") else "文字变形"
            self._finding("R017", slide_part, "「%s」使用了%s" % (name, kind),
                          slide_no=slide_no, shape_name=name)

        # ---- R015 手动换行 ----
        breaks = sum(len(elems(p, "a", "br"))
                     for p in elems(tx_body, "a", "p"))
        if breaks:
            self.metrics["hard_breaks"] += breaks
            self._finding("R015", slide_part,
                          "「%s」含 %d 个手动换行" % (name, breaks),
                          count=breaks, slide_no=slide_no, shape_name=name)

        geom = self.deck.resolve_geom(slide_part, sp)
        autofit_kind, scale, red = autofit_state(bodypr)

        # ---- R007 自动缩放残留 ----
        if self.on("R007") and autofit_kind == "norm" and (abs(scale - 1.0) > 1e-6 or red > 1e-6):
            norm = find_child(bodypr, "a", "normAutofit")
            applied = False
            detail = ""
            if geom is not None:
                l, r, t, b = body_insets(bodypr)
                inner_w = (geom["cx"] - l - r) / float(EMU_PER_PT)
                inner_h = (geom["cy"] - t - b) / float(EMU_PER_PT)
                ctx = ShapeStyleContext(self.deck, slide_part, theme, sp, tx_body)
                wrap = bodypr.get("wrap", "square") != "none"
                if inner_w > 4 and inner_h > 4:
                    need_cur, _ = self._measure(ctx, tx_body, scale, red, inner_w, wrap)
                    need_full, _ = self._measure(ctx, tx_body, 1.0, 0.0, inner_w, wrap)
                    if need_full <= inner_h * 1.02 and need_cur <= inner_h * 1.02:
                        if self.apply:
                            if norm.get("fontScale") is not None:
                                del norm.attrib["fontScale"]
                            if norm.get("lnSpcReduction") is not None:
                                del norm.attrib["lnSpcReduction"]
                            self._touch(slide_part)
                            applied = True
                            detail = "清除自动缩放残留（fontScale=%.1f%%）" % (scale * 100)
                        scale, red = 1.0, 0.0
                    elif need_cur > inner_h * 1.02:
                        new_scale = inner_h / need_cur * scale * 0.98
                        if self.apply:
                            if norm is None:
                                norm = _replace_body_autofit(bodypr, "norm")
                            norm.set("fontScale", str(max(1000, int(new_scale * 100000 / 100) * 100)))
                            norm.set("lnSpcReduction", "0")
                            self._touch(slide_part)
                            applied = True
                            detail = "按实测重算自动缩放（%.1f%%→%.1f%%）" % (
                                scale * 100, max(1.0, new_scale) * 100)
                        scale = max(1.0, new_scale)
                        red = 0.0
            fixed = 1 if applied else 0
            self._finding("R007", slide_part,
                          "「%s」残留自动缩放 fontScale=%.1f%% lnSpcReduction=%.1f%%"
                          % (name, scale * 100 if not applied else 0,
                             red * 100 if not applied else 0),
                          fixed=fixed, slide_no=slide_no, shape_name=name)
            if applied:
                self._change("R007", slide_part, detail, slide_no=slide_no,
                             shape_name=name)

        # ---- R008 spAutoFit ----
        if self.on("R008") and autofit_kind == "spauto":
            fixed = 0
            if self.apply:
                _replace_body_autofit(bodypr, "none")
                self._touch(slide_part)
                fixed = 1
                self._change("R008", slide_part, "spAutoFit → noAutofit（固定尺寸）",
                             slide_no=slide_no, shape_name=name)
            self._finding("R008", slide_part, "「%s」形状高度随文字自动变化" % name,
                          fixed=fixed, slide_no=slide_no, shape_name=name)

        # ---- R014 行距固定 ----
        if self.on("R014"):
            ctx = ShapeStyleContext(self.deck, slide_part, theme, sp, tx_body)
            cnt = 0
            for p in elems(tx_body, "a", "p"):
                ppr = find_child(p, "a", "pPr")
                style = ctx.paragraph_style(p)
                size = get_int(style, "sz", 1800) / 100.0
                ln = parse_lnspc(ppr, size)
                if ln is not None and ln[0] not in ("pct",):
                    continue
                cnt += 1
                if self.apply:
                    ratio = 1.0
                    if ln is not None and ln[0] == "pct":
                        ratio = ln[1] / 100000.0
                    target = max(4.0, size * self.line_ratio * ratio)
                    ppr = _ensure_ppr(p)
                    holder = find_child(ppr, "a", "lnSpc")
                    if holder is None:
                        holder = ET.Element(qn("a", "lnSpc"))
                        insert_ordered(ppr, holder, PPR_ORDER)
                    for child in list(holder):
                        holder.remove(child)
                    pts = ET.SubElement(holder, qn("a", "spcPts"))
                    pts.set("val", str(int(target * 100)))
                    self._touch(slide_part)
            if cnt:
                fixed = cnt if (self.apply and self.on("R014")) else 0
                self._finding("R014", slide_part,
                              "「%s」有 %d 段行距随字体度量变化" % (name, cnt),
                              count=cnt, fixed=fixed,
                              slide_no=slide_no, shape_name=name)
                if fixed:
                    self._change("R014", slide_part,
                                 "行距固定为 %d%% 字号（%d 段）" % (self.line_ratio * 100, fixed),
                                 slide_no=slide_no, shape_name=name)

        # ---- 度量 / R009 / R010 ----
        if geom is None:
            return
        l, r, t, b = body_insets(bodypr)
        inner_w = (geom["cx"] - l - r) / float(EMU_PER_PT)
        inner_h = (geom["cy"] - t - b) / float(EMU_PER_PT)
        if inner_w <= 4 or inner_h <= 4:
            return
        if vert not in (None, "horz"):
            return
        autofit_kind, scale, red = autofit_state(bodypr)
        ctx = ShapeStyleContext(self.deck, slide_part, theme, sp, tx_body)
        wrap = bodypr.get("wrap", "square") != "none"
        need_h, max_w = self._measure(ctx, tx_body, scale, red, inner_w, wrap)
        self.metrics["measured"] += 1
        ratio_h = need_h / inner_h if inner_h > 0 else 1.0
        ratio_w = (max_w / inner_w) if (not wrap and inner_w > 0) else 1.0
        ratio = max(ratio_h, ratio_w)
        if ratio > self.metrics["worst_ratio"]:
            self.metrics["worst_ratio"] = ratio

        # R010：不换行文本框宽度不足
        if not wrap and ratio_w > 1.03:
            fixed = 0
            if self.apply:
                sp_pr = find_child(sp, "p", "spPr")
                needed_cx = int(max_w * EMU_PER_PT) + l + r + int(0.02 * EMU_PER_PT)
                delta = needed_cx - geom["cx"]
                placed = None
                if delta > 0 and sp_pr is not None:
                    # 先试“居中扩宽”，再试“向右扩宽”“向左扩宽”
                    for nx, ncx in ((geom["x"] - delta // 2, geom["cx"] + delta),
                                    (geom["x"], geom["cx"] + delta),
                                    (geom["x"] - delta, geom["cx"] + delta)):
                        if nx >= 0 and nx + ncx <= self.deck.sld_w:
                            placed = (nx, ncx)
                            break
                if placed is not None:
                    write_xfrm(sp_pr, {"x": placed[0], "y": geom["y"],
                                       "cx": placed[1], "cy": geom["cy"]})
                    geom = {"x": placed[0], "y": geom["y"],
                            "cx": placed[1], "cy": geom["cy"]}
                    inner_w = (geom["cx"] - l - r) / float(EMU_PER_PT)
                    detail = "扩宽 %.2fcm → %.2fcm" % (
                        emu_to_cm(needed_cx - delta), emu_to_cm(geom["cx"]))
                else:
                    bodypr.set("wrap", "square")
                    wrap = True
                    detail = "页面放不下，改为允许自动换行"
                self._touch(slide_part)
                fixed = 1
                self._change("R010", slide_part, detail, slide_no=slide_no,
                             shape_name=name)
                # 修复后重新度量，避免 R009 依据旧数据重复缩字
                try:
                    need_h, max_w = self._measure(ctx, tx_body, scale, red,
                                                  inner_w, wrap)
                    ratio_h = need_h / inner_h if inner_h > 0 else 1.0
                    ratio_w = (max_w / inner_w) if (not wrap and inner_w > 0) else 1.0
                except Exception:  # pragma: no cover
                    ratio_h, ratio_w = ratio_h, ratio_w
            self._finding("R010", slide_part,
                          "「%s」不换行文本框宽度不足（需要 %.1fpt > 可用 %.1fpt）"
                          % (name, max_w, inner_w),
                          fixed=fixed, slide_no=slide_no, shape_name=name)

        # R010 可能已经改动几何/换行方式，这里用最新数据决定 R009
        ratio = max(ratio_h, ratio_w)

        # R009：溢出风险
        if ratio > 1.03:
            self.metrics["overflow"] += 1
            fixed = 0
            detail = ""
            if self.apply and self.on("R009"):
                needed = min(1.0 / ratio_h if ratio_h > 0 else 1.0,
                             1.0 / ratio_w if ratio_w > 0 else 1.0) * 0.98
                new_scale = max(self.shrink_min, scale * needed)
                if new_scale < scale - 1e-6:
                    norm = find_child(bodypr, "a", "normAutofit")
                    if norm is None:
                        norm = _replace_body_autofit(bodypr, "norm")
                    norm.set("fontScale", str(max(1000, int(new_scale * 100000 / 100) * 100)))
                    norm.set("lnSpcReduction", "0")
                    self._touch(slide_part)
                    fixed = 1
                    detail = "缩小字号到 %.1f%% 以消除溢出" % (new_scale * 100)
                    self._change("R009", slide_part, detail, slide_no=slide_no,
                                 shape_name=name)
            self._finding("R009", slide_part,
                          "「%s」预计溢出：高 %.0f%%/宽 %.0f%%（需要 %.1fpt，可用 %.1fpt）"
                          % (name, ratio_h * 100 - 100, ratio_w * 100 - 100,
                             need_h, inner_h),
                          fixed=fixed, slide_no=slide_no, shape_name=name)


# 把几何检查方法并入 Engine（避免主类过长拆文件，保持单文件可分发）
def _attach_mixin(cls, mixin):
    for name in dir(mixin):
        if name.startswith("__"):
            continue
        attr = getattr(mixin, name)
        if callable(attr) and not hasattr(cls, name):
            setattr(cls, name, attr)


_attach_mixin(Engine, EngineGeometryMixin)


def _engine_check_slides(self):
    for slide_part in self.deck.slides:
        slide_no = self.deck.slide_number.get(slide_part)
        theme = self.deck.theme_fonts_for_slide(slide_part)
        doc = self.pkg.xml(slide_part)
        tree = _sp_tree(doc.root)
        if tree is None:
            continue
        items = []
        walk_tree_elements(tree, items)
        for kind, el in items:
            if kind == "group":
                self._check_group(el, slide_part, slide_no)
            elif kind == "graphicFrame":
                if table_of(el) is not None:
                    try:
                        self._check_table(el, slide_part, slide_no, theme)
                    except Exception:  # pragma: no cover - 防御畸形表格
                        pass
            else:
                tx_body = find_child(el, "p", "txBody")
                if tx_body is None:
                    continue
                self.metrics["shapes"] += 1
                try:
                    self._check_shape_text(el, tx_body, slide_part, slide_no, theme)
                except Exception:  # pragma: no cover - 单个形状失败不影响整体
                    continue


def _engine_check_global(self):
    if not self.on("R020"):
        return
    standard = {
        (12192000, 6858000),   # 16:9
        (9144000, 6858000),    # 4:3
        (12192000, 7620000),   # 16:10
        (10058400, 6858000),   # 16:9 (10in)
        (10692000, 7560000),   # 接近 A4
        (7560000, 10692000),   # A4 竖版
    }
    if (self.deck.sld_w, self.deck.sld_h) not in standard:
        self._finding(
            "R020", "ppt/presentation.xml",
            "页面尺寸 %.2f × %.2f cm（非标准比例）"
            % (emu_to_cm(self.deck.sld_w), emu_to_cm(self.deck.sld_h)))


def _engine_summarize(self):
    summary = {}
    for f in self.findings:
        rec = summary.setdefault(f.rule_id, {
            "rule": f.rule_id, "name": RULES_BY_ID[f.rule_id]["name"],
            "severity": f.severity, "count": 0, "fixed": 0, "locations": 0,
        })
        rec["count"] += f.count
        rec["fixed"] += f.fixed
        rec["locations"] += 1
    for rid in summary:
        summary[rid]["severity"] = RULES_BY_ID[rid]["severity"]
    return summary


def _engine_high_count(self):
    total = 0
    for f in self.findings:
        if f.severity == "high":
            total += f.count
    return total


def _engine_to_dict(self):
    return {
        "metrics": dict(self.metrics),
        "findings": [f.to_dict() for f in self.findings],
        "changes": [c.to_dict() for c in self.changes],
        "summary": _engine_summarize(self),
        "rewritten_fonts": dict(self.rewritten_fonts),
    }


Engine.check_slides = _engine_check_slides
Engine.check_global = _engine_check_global
Engine.summarize = _engine_summarize
Engine.high_count = _engine_high_count
Engine.to_dict = _engine_to_dict


# ============================================================================
# 单个文件的处理流程
# ============================================================================

class FileResult(object):
    def __init__(self, path):
        self.path = path
        self.ok = False
        self.error = ""
        self.output_path = None
        self.backup_path = None
        self.prefix = None          # Engine（修复前扫描）
        self.fixed = None           # Engine（修复过程）
        self.postfix = None         # Engine（对输出文件复检）
        self.font_mode = "rewrite"
        self.duration = 0.0
        self.render = None

    def scan_summary(self):
        return self.prefix.summarize() if self.prefix else {}


def _new_engine(deck, enabled, apply, options):
    return Engine(deck, enabled=enabled, apply=apply, options=options)


def _open_deck(path, registry, resolver):
    pkg = PptxPackage(path)
    deck = Deck(pkg, registry=registry, resolver=resolver)
    return pkg, deck


def scan_file(path, enabled, registry, resolver, options=None):
    """只读体检。返回 (engine, deck, pkg)；调用方负责 pkg.close()。"""
    pkg, deck = _open_deck(path, registry, resolver)
    engine = _new_engine(deck, enabled, apply=False, options=options)
    engine.run()
    return engine, deck, pkg


def _default_output_path(src, out_dir=None, suffix=".fixed", overwrite=False):
    src_dir = os.path.dirname(os.path.abspath(src))
    stem, ext = os.path.splitext(os.path.basename(src))
    target_dir = os.path.abspath(out_dir) if out_dir else src_dir
    if overwrite:
        return os.path.abspath(src)
    return os.path.join(target_dir, stem + suffix + ext)


def _make_backup(src, out_dir=None):
    src_dir = os.path.dirname(os.path.abspath(src))
    stem, ext = os.path.splitext(os.path.basename(src))
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    target_dir = os.path.abspath(out_dir) if out_dir else src_dir
    backup = os.path.join(target_dir, "%s.orig-%s%s" % (stem, stamp, ext))
    shutil.copy2(src, backup)
    return backup


def process_file(path, enabled, options=None, registry=None, resolver=None,
                 out_dir=None, suffix=".fixed", overwrite=False, backup=True,
                 font_mode="rewrite", do_render_check=False):
    """扫描 + 修复 + 复检。返回 FileResult。"""
    started = datetime.datetime.now()
    result = FileResult(path)
    result.font_mode = font_mode
    options = dict(options or {})
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        result.error = "文件不存在：%s" % path
        return result
    ext = os.path.splitext(path)[1].lower()
    if ext in (".ppt", ".dps", ".wps", ".dpt"):
        result.error = "旧版二进制格式无法直接修复，请先在 WPS/LibreOffice 中另存为 .pptx"
        return result
    if ext not in (".pptx", ".potx", ".ppsx", ".pptm", ".potm", ".ppsm"):
        result.error = "不支持的文件类型：%s" % ext
        return result

    # 体检（只读）始终用全部规则：检测得越全，报告越有用；
    # 真正改写文件时才只应用用户勾选的规则。
    scan_rules = set(RULES_BY_ID)
    # 1) 修复前体检
    pkg = deck = None
    try:
        pkg, deck = _open_deck(path, registry, resolver)
        engine = _new_engine(deck, scan_rules, apply=False, options=options)
        engine.run()
        result.prefix = engine
        pkg.close()
        pkg = None
    except (zipfile.BadZipFile, KeyError) as exc:
        result.error = "不是有效的 PPTX 文件或文件损坏：%s" % exc
        return result
    except Exception as exc:  # pragma: no cover
        result.error = "体检失败：%s: %s" % (type(exc).__name__, exc)
        return result

    # 2) 修复
    fix_rules = set(enabled)
    if font_mode in ("alias", "install", "none"):
        fix_rules.discard("R001")
    out_path = _default_output_path(path, out_dir, suffix, overwrite)
    try:
        pkg, deck = _open_deck(path, registry, resolver)
        engine = _new_engine(deck, fix_rules, apply=True, options=options)
        engine.run()
        result.fixed = engine
        if backup:
            result.backup_path = _make_backup(path, out_dir)
        if not os.path.isdir(os.path.dirname(out_path)):
            os.makedirs(os.path.dirname(out_path))
        pkg.save(out_path)
        result.output_path = out_path
        pkg.close()
        pkg = None
    except Exception as exc:
        result.error = "修复失败：%s: %s" % (type(exc).__name__, exc)
        if pkg is not None:
            pkg.close()
        return result

    # 3) 复检
    try:
        engine2, _deck2, pkg2 = scan_file(out_path, scan_rules, registry, resolver, options)
        result.postfix = engine2
        pkg2.close()
    except Exception as exc:  # pragma: no cover
        result.error = "修复完成但复检失败：%s" % exc

    if do_render_check:
        try:
            result.render = compare_render(path, out_path)
        except Exception as exc:  # pragma: no cover
            result.render = {"ok": False, "error": str(exc)}

    result.ok = True
    result.duration = (datetime.datetime.now() - started).total_seconds()
    return result


# ============================================================================
# 字体方案：Windows 字体导入、fontconfig 别名、撤销
# ============================================================================

def fontconf_dir():
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    return os.path.join(base, "fontconfig", "conf.d")


def alias_conf_path():
    return os.path.join(fontconf_dir(), "99-wps-windows-fonts.conf")


def user_font_dir():
    base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    return os.path.join(base, "fonts", "windows-fonts")


def _run_fc_cache():
    try:
        proc = subprocess.run(["fc-cache", "-f"],
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=180)
        return proc.returncode == 0, proc.stderr.decode("utf-8", "replace")[:400]
    except (OSError, subprocess.SubprocessError) as exc:
        return False, str(exc)


def _xml_escape(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def install_font_aliases(mappings, do_fc_cache=True):
    """生成 fontconfig 别名，让 WPS 把缺失字体统一替换成指定字体。

    mappings: [(原字体名, 目标字体名), ...]
    返回 (ok, message)
    """
    mappings = [(a, b) for a, b in mappings if a and b and a != b]
    if not mappings:
        return False, "没有需要建立别名的字体映射"
    path = alias_conf_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    backup = None
    if os.path.exists(path):
        backup = path + ".bak"
        shutil.copy2(path, backup)
    lines = [
        '<?xml version="1.0"?>',
        '<!DOCTYPE fontconfig SYSTEM "fonts.dtd">',
        '<!-- 由 %s 自动生成，可用同工具的“撤销字体别名”删除 -->' % TOOL_NAME,
        '<fontconfig>',
    ]
    for src, dst in mappings:
        s, d = _xml_escape(src), _xml_escape(dst)
        lines.append('  <alias binding="same"><family>%s</family>'
                     '<prefer><family>%s</family></prefer></alias>' % (s, d))
    for src, dst in mappings:
        s, d = _xml_escape(src), _xml_escape(dst)
        lines.append('  <match target="pattern">')
        lines.append('    <test name="family" qual="any" compare="eq"><string>%s</string></test>' % s)
        lines.append('    <edit name="family" mode="assign_replace" binding="same">'
                     '<string>%s</string></edit>' % d)
        lines.append('  </match>')
    lines.append('</fontconfig>')
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    ok_cache = True
    msg = ""
    if do_fc_cache:
        ok_cache, err = _run_fc_cache()
        if not ok_cache:
            msg = "（fc-cache 执行失败：%s，可能需要重新登录或手动运行 fc-cache -f）" % err
    return True, "已写入 %s，共 %d 条映射%s" % (path, len(mappings), msg)


def remove_font_aliases(do_fc_cache=True):
    path = alias_conf_path()
    removed = False
    if os.path.exists(path):
        os.remove(path)
        removed = True
    backup = path + ".bak"
    if os.path.exists(backup) and not removed:
        os.replace(backup, path)
        removed = True
    if not removed:
        return False, "没有找到本工具生成的字体别名文件"
    if do_fc_cache:
        _run_fc_cache()
    return True, "已移除字体别名：%s" % path


WIN_FONT_EXTS = (".ttf", ".ttc", ".otf", ".otc")


def import_windows_fonts(src_dir, registry=None):
    """把 Windows 字体目录（C:\\Windows\\Fonts）复制到用户字体目录并刷新缓存。"""
    src_dir = os.path.abspath(os.path.expanduser(src_dir))
    if not os.path.isdir(src_dir):
        return False, "目录不存在：%s" % src_dir
    dest = user_font_dir()
    os.makedirs(dest, exist_ok=True)
    copied = 0
    names = []
    for name in sorted(os.listdir(src_dir)):
        if not name.lower().endswith(WIN_FONT_EXTS):
            continue
        src = os.path.join(src_dir, name)
        if not os.path.isfile(src):
            continue
        try:
            shutil.copy2(src, os.path.join(dest, name))
            copied += 1
            names.append(name)
        except (IOError, OSError):
            continue
    if not copied:
        return False, "该目录下没有找到字体文件（.ttf/.ttc/.otf）"
    ok_cache, err = _run_fc_cache()
    msg = "已复制 %d 个字体到 %s" % (copied, dest)
    if not ok_cache:
        msg += "；fc-cache 失败：%s" % err
    if registry is not None:
        registry._load()
    return True, msg


# ============================================================================
# 渲染验证（LibreOffice / PDF）
# ============================================================================

def _soffice_path():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    return None


def _pdfinfo_pages(pdf_path):
    exe = shutil.which("pdfinfo")
    if not exe:
        return None
    try:
        out = subprocess.run([exe, pdf_path], stdout=subprocess.PIPE,
                             stderr=subprocess.DEVNULL, timeout=60).stdout
        m = re.search(rb"Pages:\s+(\d+)", out)
        if m:
            return int(m.group(1))
    except (OSError, subprocess.SubprocessError):
        pass
    return None


def _pdf_text(pdf_path):
    exe = shutil.which("pdftotext")
    if not exe:
        return None
    try:
        out = subprocess.run([exe, "-layout", pdf_path, "-"],
                             stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                             timeout=120).stdout
        return out.decode("utf-8", "replace")
    except (OSError, subprocess.SubprocessError):
        return None


def render_pptx(path, workdir):
    """用 LibreOffice 把 PPTX 渲染成 PDF，返回信息字典。"""
    soffice = _soffice_path()
    if not soffice:
        return {"ok": False, "error": "未安装 LibreOffice（soffice），无法渲染验证"}
    os.makedirs(workdir, exist_ok=True)
    profile = os.path.join(workdir, "lo-profile")
    home = os.path.join(workdir, "lo-home")
    runtime = os.path.join(workdir, "lo-runtime")
    for d in (profile, home, runtime):
        os.makedirs(d, exist_ok=True)
    env = dict(os.environ)
    env.update({
        "HOME": home,
        "XDG_RUNTIME_DIR": runtime,
        "SAL_USE_VCLPLUGIN": "svp",
        "LANG": "zh_CN.UTF-8",
    })
    import pathlib
    profile_uri = pathlib.Path(profile).as_uri()
    cmd = [soffice, "-env:UserInstallation=" + profile_uri,
           "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", workdir, path]
    try:
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=300, env=env)
    except (OSError, subprocess.SubprocessError) as exc:
        return {"ok": False, "error": "LibreOffice 调用失败：%s" % exc}
    stem = os.path.splitext(os.path.basename(path))[0]
    pdf = os.path.join(workdir, stem + ".pdf")
    if not os.path.exists(pdf):
        return {"ok": False,
                "error": "LibreOffice 没有生成 PDF",
                "stderr": proc.stderr.decode("utf-8", "replace")[-400:]}
    pages = _pdfinfo_pages(pdf)
    text = _pdf_text(pdf)
    return {"ok": True, "pdf": pdf, "pages": pages, "text": text,
            "chars": len(text or "")}


def _normalize_text(text):
    return re.sub(r"\s+", "", text or "")


def compare_render(before_path, after_path):
    """渲染修复前后的文件，比较页数与文字内容，确认修复没有丢内容。"""
    workdir = tempfile.mkdtemp(prefix="wps-repair-render-")
    try:
        a = render_pptx(before_path, os.path.join(workdir, "before"))
        b = render_pptx(after_path, os.path.join(workdir, "after"))
        if not a.get("ok") or not b.get("ok"):
            return {"ok": False,
                    "before": {k: v for k, v in a.items() if k != "text"},
                    "after": {k: v for k, v in b.items() if k != "text"},
                    "error": a.get("error") or b.get("error")}
        ta, tb = _normalize_text(a["text"]), _normalize_text(b["text"])
        same_pages = (a["pages"] is not None and a["pages"] == b["pages"])
        common = sum(1 for x, y in zip(ta, tb) if x == y)
        similarity = common / float(max(len(ta), len(tb), 1))
        return {
            "ok": True,
            "pages_before": a["pages"], "pages_after": b["pages"],
            "chars_before": len(ta), "chars_after": len(tb),
            "same_pages": same_pages, "text_similarity": round(similarity, 4),
            "content_preserved": same_pages and similarity >= 0.97,
            "before_pdf": a["pdf"], "after_pdf": b["pdf"],
        }
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


# ============================================================================
# 环境探测
# ============================================================================

def read_os_release():
    data = {}
    for path in ("/etc/os-release", "/usr/lib/os-release"):
        if not os.path.exists(path):
            continue
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as fh:
                for line in fh:
                    line = line.strip()
                    if not line or "=" not in line or line.startswith("#"):
                        continue
                    k, v = line.split("=", 1)
                    data[k.strip()] = v.strip().strip('"').strip("'")
            break
        except (IOError, OSError):
            continue
    return data


def detect_wps():
    candidates = [
        "/usr/bin/wps", "/opt/kingsoft/wps-office/office6/wps",
        "/opt/kingsoft/wps-office/wps", "/usr/bin/wpp",
    ]
    found = None
    for path in candidates:
        if os.path.exists(path):
            found = path
            break
    if found is None:
        found = shutil.which("wps") or shutil.which("wpp")
    version = ""
    # 不要执行 `wps --version`：部分 WPS for Linux 版本会直接拉起图形界面。
    # 优先从安装包元数据 / version 文件读取。
    for cmd in (["dpkg-query", "-W", "-f", "${Version}", "wps-office"],
                ["rpm", "-q", "--qf", "%{VERSION}", "wps-office"]):
        try:
            proc = subprocess.run(cmd, stdout=subprocess.PIPE,
                                  stderr=subprocess.DEVNULL, timeout=5)
            text = proc.stdout.decode("utf-8", "replace").strip()
            if proc.returncode == 0 and text:
                version = text
                break
        except (OSError, subprocess.SubprocessError):
            continue
    if not version:
        for vf in ("/opt/kingsoft/wps-office/version",
                   "/opt/kingsoft/wps-office/office6/version"):
            try:
                with open(vf, "r", encoding="utf-8", errors="replace") as fh:
                    version = fh.read().strip()[:40]
                    break
            except (IOError, OSError):
                continue
    return {"path": found, "version": version}


def detect_libreoffice():
    exe = _soffice_path()
    version = ""
    if exe:
        try:
            proc = subprocess.run([exe, "--version"], stdout=subprocess.PIPE,
                                  stderr=subprocess.DEVNULL, timeout=20,
                                  env=dict(os.environ, HOME=os.path.expanduser("~")))
            m = re.search(r"(\d+\.\d+\.\d+(?:\.\d+)?)",
                          proc.stdout.decode("utf-8", "replace"))
            if m:
                version = m.group(1)
        except (OSError, subprocess.SubprocessError):
            pass
    return {"path": exe, "version": version}


def environment_info(registry):
    osr = read_os_release()
    return {
        "os_name": osr.get("PRETTY_NAME") or osr.get("NAME") or (os.uname().sysname if hasattr(os, "uname") else ""),
        "os_id": osr.get("ID", ""),
        "os_version": osr.get("VERSION_ID", ""),
        "kernel": os.uname().release if hasattr(os, "uname") else "",
        "python": sys.version.split()[0],
        "wps": detect_wps(),
        "libreoffice": detect_libreoffice(),
        "fontconfig": bool(registry.available),
        "font_count": len(registry.families) if registry.available else 0,
        "cjk_fonts": [n for n in (CJK_SANS + CJK_SERIF + CJK_KAI)
                      if registry.available and registry.installed(n)][:8],
        "pdftotext": bool(shutil.which("pdftotext")),
    }


def print_environment(env):
    hr("=")
    say(_c("环境自检", C.BOLD))
    say("  操作系统 : %s %s" % (env["os_name"], env["kernel"]))
    say("  Python   : %s" % env["python"])
    wps = env["wps"]
    if wps["path"]:
        say("  WPS      : %s %s" % (wps["path"], wps["version"] or ""))
    else:
        say("  WPS      : 未检测到（不影响修复，只是无法自动验证 WPS 渲染）")
    lo = env["libreoffice"]
    if lo["path"]:
        say("  LibreOffice: %s %s（可用于渲染验证）" % (lo["path"], lo["version"]))
    else:
        say("  LibreOffice: 未安装（可选，仅影响渲染验证）")
    if env["fontconfig"]:
        say("  fontconfig : 可用，共 %d 个字体族" % env["font_count"])
        say("  中文字体   : %s" % ("、".join(env["cjk_fonts"]) or "未发现"))
    else:
        say("  fontconfig : 不可用（将只使用内置字体映射，无法判断字体是否已安装）")
    say("  PDF 工具   : pdftotext %s" % ("可用" if env["pdftotext"] else "不可用"))
    hr("=")


# ============================================================================
# 控制台输出与报告
# ============================================================================

def print_rule_summary(engine, max_detail=6):
    summary = engine.summarize()
    if not summary:
        ok("未发现已知的跨平台错位隐患。")
        return
    total = sum(rec["count"] for rec in summary.values())
    high = sum(rec["count"] for rec in summary.values() if rec["severity"] == "high")
    say("")
    say(_c("问题汇总", C.BOLD) +
        "：共 %d 项（高危 %d 项），涉及 %d 类原因" % (total, high, len(summary)))
    say("%-6s %-4s %-28s %6s %6s" % ("规则", "级别", "原因", "数量", "已修复"))
    hr()
    for rid in sorted(summary, key=lambda r: (SEVERITY_RANK[summary[r]["severity"]], r)):
        rec = summary[rid]
        color = C.RED if rec["severity"] == "high" else (
            C.YELLOW if rec["severity"] in ("mid", "low") else C.DIM)
        say("%-6s %-4s %-28s %6d %6d" % (
            rid, _c(SEVERITY_LABEL[rec["severity"]], color),
            rec["name"][:26], rec["count"], rec["fixed"]))
    hr()
    # 明细
    shown = {}
    for f in engine.findings:
        detail_level = SEVERITY_RANK[f.severity]
        if detail_level > 1 or shown.get(f.rule_id, 0) >= max_detail:
            continue
        shown[f.rule_id] = shown.get(f.rule_id, 0) + 1
        say("  %s %s  %s" % (_c("[%s]" % f.rule_id, C.MAGENTA), f.location(), f.detail))
        for ex in f.examples[:2]:
            say("      · %s" % ex)
    truncated = [rid for rid in summary if shown.get(rid, 0) < len(
        [x for x in engine.findings if x.rule_id == rid])]
    if truncated:
        info("部分低风险明细已省略（可在报告文件中查看全部）。")
    if engine.metrics.get("measured"):
        say("  度量了 %d 个文本框，其中 %d 个存在溢出风险，最严重 %.0f%%" % (
            engine.metrics["measured"], engine.metrics["overflow"],
            (engine.metrics["worst_ratio"] - 1.0) * 100 if engine.metrics["worst_ratio"] > 1 else 0))


def print_changes(engine, limit=40):
    if not engine.changes:
        info("没有对文件做任何修改。")
        return
    say(_c("已修改内容（%d 处）" % len(engine.changes), C.BOLD))
    for ch in engine.changes[:limit]:
        loc = ("第 %d 页 / " % ch.slide_no) if ch.slide_no else ""
        if ch.shape_name:
            loc += "「%s」" % ch.shape_name
        say("  %s %s%s" % (_c(ch.rule_id, C.GREEN), loc, ch.detail))
    if len(engine.changes) > limit:
        info("……其余 %d 条见报告文件。" % (len(engine.changes) - limit))


def _md_rule_section(engine):
    summary = engine.summarize()
    lines = ["## 诊断结果", ""]
    if not summary:
        lines.append("未发现已知的跨平台错位隐患。")
        return lines
    lines.append("| 规则 | 级别 | 原因 | 数量 | 已修复 |")
    lines.append("| --- | --- | --- | ---: | ---: |")
    for rid in sorted(summary, key=lambda r: (SEVERITY_RANK[summary[r]["severity"]], r)):
        rec = summary[rid]
        lines.append("| %s | %s | %s | %d | %d |" % (
            rid, SEVERITY_LABEL[rec["severity"]], rec["name"], rec["count"], rec["fixed"]))
    lines.append("")
    lines.append("### 逐项明细")
    lines.append("")
    by_rule = {}
    for f in engine.findings:
        by_rule.setdefault(f.rule_id, []).append(f)
    for rid in sorted(by_rule, key=lambda r: (SEVERITY_RANK[RULES_BY_ID[r]["severity"]], r)):
        rule = RULES_BY_ID[rid]
        lines.append("#### %s %s（%s）" % (rid, rule["name"],
                                     SEVERITY_LABEL[rule["severity"]]))
        lines.append("")
        lines.append("> %s" % rule["desc"])
        lines.append("")
        lines.append("修复方式：%s" % rule["fix"])
        lines.append("")
        for f in by_rule[rid][:60]:
            ex = ("；例：" + "、".join(f.examples[:4])) if f.examples else ""
            lines.append("- %s：%s（%d 项）%s" % (f.location(), f.detail, f.count, ex))
        if len(by_rule[rid]) > 60:
            lines.append("- …… 其余 %d 条省略" % (len(by_rule[rid]) - 60))
        lines.append("")
    return lines


def build_report_markdown(result, env=None):
    rule = RULES_BY_ID
    prefix = result.prefix
    postfix = result.postfix
    lines = []
    lines.append("# WPS 演示文稿错位检测报告")
    lines.append("")
    lines.append("- 文件：`%s`" % result.path)
    lines.append("- 生成时间：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("- 工具版本：%s v%s" % (TOOL_NAME, VERSION))
    if result.backup_path:
        lines.append("- 原文件备份：`%s`" % result.backup_path)
    if result.output_path:
        lines.append("- 修复后文件：`%s`" % result.output_path)
    if result.error:
        lines.append("- **错误**：%s" % result.error)
    lines.append("")
    if env:
        lines.append("## 运行环境")
        lines.append("")
        lines.append("- 操作系统：%s" % env.get("os_name", ""))
        lines.append("- Python：%s" % env.get("python", ""))
        wps = env.get("wps") or {}
        lines.append("- WPS：%s" % (wps.get("path") or "未检测到"))
        lines.append("- fontconfig：%s" % ("可用" if env.get("fontconfig") else "不可用"))
        lines.append("- 中文字体：%s" % ("、".join(env.get("cjk_fonts") or []) or "无"))
        lines.append("")
    if prefix:
        lines.append("## 文件概况")
        lines.append("")
        m = prefix.metrics
        lines.append("- 幻灯片：%d 页" % m.get("slides", 0))
        lines.append("- 文本形状：%d 个，实测 %d 个" % (m.get("shapes", 0), m.get("measured", 0)))
        lines.append("- 表格：%d 个" % m.get("tables", 0))
        lines.append("- 文本 run：%d 个" % m.get("runs", 0))
        lines.append("- 手动换行：%d 处" % m.get("hard_breaks", 0))
        if m.get("measured"):
            lines.append("- 溢出风险形状：%d 个（最严重超出 %.0f%%）" % (
                m.get("overflow", 0),
                max(0.0, (m.get("worst_ratio", 1.0) - 1.0) * 100)))
        lines.append("")
        lines.extend(_md_rule_section(prefix))
        if result.fixed is not None and result.fixed.changes:
            lines.append("## 修复记录")
            lines.append("")
            for ch in result.fixed.changes:
                loc = ("第 %d 页 / " % ch.slide_no) if ch.slide_no else ""
                if ch.shape_name:
                    loc += "「%s」" % ch.shape_name
                lines.append("- %s：%s%s" % (ch.rule_id, loc, ch.detail))
            lines.append("")
        if result.fixed is not None and result.fixed.rewritten_fonts:
            lines.append("### 字体替换表")
            lines.append("")
            lines.append("| 原字体（Windows） | 替换为 |")
            lines.append("| --- | --- |")
            for src, dst in sorted(result.fixed.rewritten_fonts.items()):
                lines.append("| %s | %s |" % (src, dst))
            lines.append("")
        if postfix is not None:
            lines.append("## 修复后复检")
            lines.append("")
            s = postfix.summarize()
            pp = postfix.metrics
            if s:
                lines.append("仍有以下问题（可能需要 WPS 中人工调整）：")
                lines.append("")
                lines.append("| 规则 | 级别 | 原因 | 数量 |")
                lines.append("| --- | --- | --- | ---: |")
                for rid in sorted(s, key=lambda r: (SEVERITY_RANK[s[r]["severity"]], r)):
                    rec = s[rid]
                    lines.append("| %s | %s | %s | %d |" % (
                        rid, SEVERITY_LABEL[rec["severity"]], rec["name"], rec["count"]))
                lines.append("")
            else:
                lines.append("未再发现已知错位隐患。")
                lines.append("")
            if pp.get("measured"):
                lines.append("- 溢出风险形状：%d 个（修复前 %d 个）" % (
                    pp.get("overflow", 0), prefix.metrics.get("overflow", 0)))
                lines.append("- 最严重溢出：%.0f%%（修复前 %.0f%%）" % (
                    max(0.0, (pp.get("worst_ratio", 1.0) - 1.0) * 100),
                    max(0.0, (prefix.metrics.get("worst_ratio", 1.0) - 1.0) * 100)))
                lines.append("")
        if result.render is not None:
            r = result.render
            lines.append("## 渲染验证（LibreOffice）")
            lines.append("")
            if r.get("ok"):
                lines.append("- 页数：修复前 %s → 修复后 %s %s" % (
                    r.get("pages_before"), r.get("pages_after"),
                    "（一致）" if r.get("same_pages") else "（不一致！）"))
                lines.append("- 文字内容相似度：%.2f%%" % (r.get("text_similarity", 0) * 100))
                lines.append("- 内容保持：%s" % ("通过" if r.get("content_preserved") else "需人工确认"))
            else:
                lines.append("- 渲染验证未通过/不可用：%s" % r.get("error", ""))
            lines.append("")
    lines.append("## 建议的人工检查清单（WPS 演示）")
    lines.append("")
    lines.append("1. 打开修复后的文件，逐页检查**标题和正文是否仍在色块内**；")
    lines.append("2. 检查表格：文字是否被挤到单元格边界、行高是否异常；")
    lines.append("3. 检查项目符号、箭头、对勾等符号是否显示为字母/方框；")
    lines.append("4. 检查“自动调整”相关形状是否遮住相邻元素；")
    lines.append("5. 若仍有细微偏差，在 WPS 中选中该文本框，“开始/格式 → 字体”确认实际使用的字体；")
    lines.append("6. 如需与 Windows 完全一致，最彻底的办法是安装原版 Windows 字体（见工具“字体兼容方案”）。")
    lines.append("")
    lines.append("## 回滚方法")
    lines.append("")
    if result.backup_path:
        lines.append("原始文件已备份为 `%s`，直接改名替换修复版即可回滚；"
                     "若使用了 fontconfig 字体别名，可在工具中选择“撤销字体别名”。"
                     % result.backup_path)
    else:
        lines.append("本工具默认不覆盖原文件；删除修复版文件即可。"
                     "若使用了 fontconfig 字体别名，可在工具中选择“撤销字体别名”。")
    lines.append("")
    return "\n".join(lines)


def build_report_dict(result, env=None):
    return {
        "tool": TOOL_NAME,
        "version": VERSION,
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "environment": env or {},
        "file": {
            "path": result.path,
            "output": result.output_path,
            "backup": result.backup_path,
            "font_mode": result.font_mode,
            "error": result.error,
            "duration_sec": result.duration,
        },
        "before": result.prefix.to_dict() if result.prefix else None,
        "fix": {
            "changes": [c.to_dict() for c in result.fixed.changes] if result.fixed else [],
            "rewritten_fonts": dict(result.fixed.rewritten_fonts) if result.fixed else {},
        },
        "after": result.postfix.to_dict() if result.postfix else None,
        "render_check": _strip_temp_paths(result.render),
    }


def write_reports(result, env=None, out_dir=None):
    base_dir = os.path.abspath(out_dir or os.path.dirname(result.path))
    stem = os.path.splitext(os.path.basename(result.path))[0]
    json_path = os.path.join(base_dir, stem + ".report.json")
    md_path = os.path.join(base_dir, stem + ".report.md")
    try:
        with open(json_path, "w", encoding="utf-8") as fh:
            json.dump(build_report_dict(result, env), fh, ensure_ascii=False, indent=2)
        with open(md_path, "w", encoding="utf-8") as fh:
            fh.write(build_report_markdown(result, env))
        return json_path, md_path
    except (IOError, OSError):
        return None, None


# ============================================================================
# 补充工具
# ============================================================================

def _append_before_extlst(parent, child):
    """按 DrawingML 习惯把新元素插到 extLst 之前。"""
    ext = find_child(parent, "a", "extLst")
    if ext is not None:
        parent.insert(list(parent).index(ext), child)
    else:
        parent.append(child)


def _strip_temp_paths(render):
    if not render:
        return render
    out = dict(render)
    for key in ("before_pdf", "after_pdf"):
        out.pop(key, None)
    return out


def _has_visible_text(tx_body):
    if tx_body is None:
        return False
    for el in tx_body.iter():
        if localname(el.tag) == "t" and (el.text or "").strip():
            return True
    return False


# ============================================================================
# 交互输入
# ============================================================================

STDIN_EOF = False


def prompt(label, default="", allow_empty=True):
    """读一行输入；EOF（非交互环境）时返回默认值并记录状态。"""
    global STDIN_EOF
    suffix = ""
    if default not in (None, ""):
        suffix = " [%s]" % default
    try:
        raw = input("%s%s: " % (label, suffix))
    except EOFError:
        STDIN_EOF = True
        return default or ""
    except KeyboardInterrupt:
        say("")
        raise
    raw = raw.strip()
    if not raw:
        return default or ""
    return raw


def yes_no(label, default=True):
    hint = "(Y/n)" if default else "(y/N)"
    ans = prompt("%s %s" % (label, hint), "y" if default else "n")
    low = ans.strip().lower()
    if low in ("y", "yes", "是", "1", "ok"):
        return True
    if low in ("n", "no", "否", "0"):
        return False
    return default


def ask_rules_manual():
    raw = prompt("输入规则号（逗号分隔，例如 R001,R009,R014）", "")
    return rules_from_arg(raw)


# ============================================================================
# 路径 / 规则 / 字体映射辅助
# ============================================================================

PRESENTATION_EXTS = (".pptx", ".potx", ".ppsx", ".pptm", ".potm", ".ppsm")
LEGACY_EXTS = (".ppt", ".dps", ".wps", ".dpt")


def list_in_dir(path):
    found = []
    try:
        names = os.listdir(path)
    except (IOError, OSError):
        return found
    for name in sorted(names):
        full = os.path.join(path, name)
        if os.path.isfile(full) and os.path.splitext(name)[1].lower() in (
                PRESENTATION_EXTS + LEGACY_EXTS):
            found.append(full)
    return found


def expand_paths(items):
    """展开文件/目录/通配符，去重。返回 (files, errors)。"""
    import glob as _glob
    files = []
    errors = []
    seen = set()
    for item in items:
        if not item:
            continue
        item = item.strip().strip('"').strip("'")
        item = os.path.expanduser(item)
        matches = [item]
        if any(ch in item for ch in "*?["):
            matches = _glob.glob(item)
            if not matches:
                errors.append("没有匹配到文件：%s" % item)
                continue
        for m in matches:
            if os.path.isdir(m):
                found = list_in_dir(m)
                if not found:
                    errors.append("目录中没有演示文稿：%s" % m)
                for f in found:
                    ap = os.path.abspath(f)
                    if ap not in seen:
                        seen.add(ap)
                        files.append(ap)
            elif os.path.isfile(m):
                ap = os.path.abspath(m)
                if ap not in seen:
                    seen.add(ap)
                    files.append(ap)
            else:
                errors.append("路径不存在：%s" % m)
    return files, errors


FONT_GROUP_RULES = ["R001", "R002", "R003", "R004", "R005", "R013", "R019"]
LAYOUT_GROUP_RULES = ["R006", "R007", "R008", "R009", "R010", "R011", "R012",
                      "R014", "R015", "R016", "R017", "R018", "R020"]


def rules_from_arg(value):
    if value is None or value == "":
        return set(DEFAULT_RULES)
    value = value.strip()
    low = value.lower()
    if low in ("default", "safe", "推荐"):
        return set(DEFAULT_RULES)
    if low in ("all", "全部"):
        return set(RULES_BY_ID)
    if low in ("font", "字体"):
        return set(FONT_GROUP_RULES)
    if low in ("layout", "版式"):
        return set(LAYOUT_GROUP_RULES)
    ids = set()
    for token in re.split(r"[,\s]+", value.upper()):
        if not token:
            continue
        if token in RULES_BY_ID:
            ids.add(token)
        elif re.match(r"^R\d{1,3}$", token):
            token2 = "R%03d" % int(token[1:])
            if token2 in RULES_BY_ID:
                ids.add(token2)
    return ids or set(DEFAULT_RULES)


def collect_missing_fonts(engines, resolver):
    names = set()
    for engine in engines:
        if engine is None:
            continue
        for f in engine.findings:
            if f.rule_id == "R001":
                for name in f.examples:
                    if name:
                        names.add(name)
    mappings = []
    for name in sorted(names):
        decision = resolver.decide(name)
        if decision.target and decision.target != name:
            mappings.append((name, decision.target))
    return mappings


def print_rules_table():
    say(_c("规则清单（%d 条）" % len(RULES), C.BOLD))
    hr()
    for rule in RULES:
        flag = "默认" if rule["default_on"] else "可选"
        say("%s %s [%s/%s] %s" % (
            _c(rule["id"], C.MAGENTA), rule["name"],
            SEVERITY_LABEL[rule["severity"]], flag,
            _c("可自动修复" if rule["default_on"] else "需确认", C.DIM)))
    hr()


def print_rule_detail(rid):
    rule = RULES_BY_ID.get(rid.upper())
    if rule is None:
        err("未知规则：%s" % rid)
        return False
    say(_c("%s %s" % (rule["id"], rule["name"]), C.BOLD))
    say("级别：%s   默认：%s   分组：%s" % (
        SEVERITY_LABEL[rule["severity"]],
        "启用" if rule["default_on"] else "需手动启用", rule["group"]))
    say("")
    say("为什么会导致错位：")
    say(textwrap.fill(rule["desc"], 76, initial_indent="  ", subsequent_indent="  "))
    say("")
    say("修复方式：")
    say(textwrap.fill(rule["fix"], 76, initial_indent="  ", subsequent_indent="  "))
    return True


# ============================================================================
# 交互式界面
# ============================================================================

class InteractiveApp(object):
    def __init__(self, initial_paths=None, options=None):
        self.registry = FontRegistry()
        self.resolver = FontResolver(self.registry)
        self.env = environment_info(self.registry)
        self.options = options or {"line_ratio": 1.2, "shrink_min": 0.25}
        self.files = []
        self.scans = {}
        self.results = []
        self.out_dir = None
        self.add_paths(initial_paths or [])

    # -- 基础 ---------------------------------------------------------------
    def add_paths(self, items):
        files, errors = expand_paths(items)
        for e in errors:
            err(e)
        for f in files:
            if f not in self.files:
                self.files.append(f)
        if files:
            ok("已载入 %d 个文件（当前共 %d 个）" % (len(files), len(self.files)))

    def reload_fonts(self):
        self.registry = FontRegistry()
        self.resolver = FontResolver(self.registry)
        self.env = environment_info(self.registry)

    def _all_rules(self):
        return set(RULES_BY_ID)

    # -- 主循环 -------------------------------------------------------------
    def run(self):
        self.banner()
        while True:
            if STDIN_EOF:
                say("")
                info("输入已结束，退出。")
                return 0
            try:
                self.main_menu()
                choice = prompt("请选择操作", "3" if not self.files else "4")
            except (KeyboardInterrupt, EOFError):
                say("")
                say("已退出。")
                return 0
            if choice in ("0", "q", "Q", "退出"):
                say("再见。")
                return 0
            try:
                if choice == "1":
                    self.select_files()
                elif choice == "2":
                    print_environment(self.env)
                elif choice == "3":
                    self.scan_many()
                elif choice == "4":
                    self.repair_wizard()
                elif choice == "5":
                    self.render_last()
                elif choice == "6":
                    self.font_menu()
                elif choice == "7":
                    self.show_help()
                else:
                    warn("无效选择，请输入 0-7。")
            except (KeyboardInterrupt, EOFError):
                say("")
                info("已取消当前操作。")
            except Exception as exc:  # pragma: no cover
                err("操作失败：%s: %s" % (type(exc).__name__, exc))

    def banner(self):
        say("")
        say(_c("╔" + "═" * 72 + "╗", C.CYAN))
        say(_c("║", C.CYAN) + pad_display("  %s  v%s" % (TOOL_NAME, VERSION), 72)
            + _c("║", C.CYAN))
        say(_c("║", C.CYAN) + pad_display("  专治：Windows → UOS/麒麟 迁移后 WPS 演示文字错位", 72)
            + _c("║", C.CYAN))
        say(_c("╚" + "═" * 72 + "╝", C.CYAN))
        say("  系统：%s" % self.env.get("os_name", ""))
        wps = self.env.get("wps") or {}
        say("  WPS ：%s" % (wps.get("path") or "未检测到"))
        if not self.env.get("fontconfig"):
            warn("fontconfig 不可用：将无法检测字体是否已安装，只做保守替换。")
        say("")

    def main_menu(self):
        hr("=")
        say(_c("主菜单", C.BOLD) + "   当前文件：%d 个" % len(self.files))
        if self.files:
            for i, f in enumerate(self.files[:6], 1):
                say("  %d) %s" % (i, os.path.basename(f)))
            if len(self.files) > 6:
                say("  … 其余 %d 个" % (len(self.files) - 6))
        say("")
        say("  [1] 选择演示文稿（支持拖入路径 / 目录 / 通配符）")
        say("  [2] 环境与字体体检")
        say("  [3] 扫描诊断（只读，不改任何文件）")
        say("  [4] 修复向导（扫描 → 选择修复项 → 备份并修复）")
        say("  [5] 渲染验证（LibreOffice 对比修复前后）")
        say("  [6] 字体兼容方案（导入 Windows 字体 / 生成或撤销字体别名）")
        say("  [7] 帮助：原因说明、修复原理与推荐流程")
        say("  [0] 退出")
        hr("=")

    def select_files(self):
        say("提示：可直接把文件从文件管理器拖到终端，支持多个路径；也可输入目录。")
        raw = prompt("路径（留空返回）", "")
        if not raw:
            return
        try:
            import shlex
            items = shlex.split(raw)
        except ValueError:
            items = [raw]
        self.add_paths(items)

    # -- 扫描 ---------------------------------------------------------------
    def ensure_files(self):
        while not self.files:
            self.select_files()
            if not self.files:
                if not yes_no("还没有选择文件，是否返回主菜单？", True):
                    continue
                return False
        return True

    def scan_many(self):
        if not self.ensure_files():
            return
        self.scans = {}
        for idx, path in enumerate(self.files, 1):
            say("")
            step("[%d/%d] 体检：%s" % (idx, len(self.files), os.path.basename(path)))
            if os.path.splitext(path)[1].lower() in LEGACY_EXTS:
                err("旧版二进制格式（.ppt/.dps/.wps）无法直接体检，"
                    "请先在 WPS 中另存为 .pptx。")
                continue
            try:
                engine, deck, pkg = scan_file(path, self._all_rules(),
                                              self.registry, self.resolver,
                                              self.options)
                pkg.close()
                self.scans[path] = engine
                print_rule_summary(engine)
            except Exception as exc:
                err("体检失败：%s: %s" % (type(exc).__name__, exc))
        say("")
        info("扫描完成。可以用 [4] 修复，或 [6] 处理字体缺字问题。")

    # -- 修复向导 -----------------------------------------------------------
    def choose_rule_set(self):
        rule_groups = [
            ("1", "推荐修复（安全项，默认）", DEFAULT_RULES),
            ("2", "仅字体问题（安全项）", set(FONT_GROUP_RULES) & DEFAULT_RULES),
            ("3", "仅版式问题（安全项）", set(LAYOUT_GROUP_RULES) & DEFAULT_RULES),
            ("4", "推荐 + 溢出缩字（R009）", DEFAULT_RULES | {"R009"}),
            ("5", "全部规则（含改变视觉效果的项）", set(RULES_BY_ID)),
        ]
        say("")
        say(_c("选择修复范围", C.BOLD))
        for key, label, ids in rule_groups:
            say("  [%s] %s（%d 条规则）" % (key, label, len(ids)))
        say("  [6] 自定义规则号，例如 R001,R009")
        say("")
        for key, label, ids in rule_groups:
            if key == "1":
                say(_c("推荐项说明：", C.DIM) + "字体替换/中文字体拆分/主题字体/内边距/"
                    "自动缩放残留/表格边距/语言标记/不换行宽度")
        choice = prompt("请选择", "1")
        if choice == "6":
            raw = prompt("输入规则号（逗号分隔）", "")
            ids = rules_from_arg(raw)
            if not ids:
                warn("没有识别到规则号，使用推荐项。")
                return set(DEFAULT_RULES)
            return ids
        for key, label, ids in rule_groups:
            if choice == key:
                return set(ids)
        warn("无效选择，使用推荐项。")
        return set(DEFAULT_RULES)

    def _missing_font_mappings(self):
        engines = [self.scans.get(f) for f in self.files]
        return collect_missing_fonts(engines, self.resolver)

    def apply_font_strategy(self):
        mappings = self._missing_font_mappings()
        say("")
        say(_c("字体处理方式", C.BOLD))
        if mappings:
            say("  检测到 %d 个本机缺失的字体，例如：" % len(mappings))
            for src, dst in mappings[:8]:
                say("    %s  →  %s" % (src, dst))
        else:
            say("  未在本机发现缺失字体（或尚未扫描）。")
        say("")
        say("  [1] 改写文件内字体名（推荐：输出文件可独立分发，换电脑也不缺字）")
        say("  [2] 生成系统 fontconfig 别名（不改文件；WPS 统一替换，适合长期/批量）")
        say("  [3] 从 Windows 字体目录导入（最保真；需要能访问 Windows 的 Fonts 目录）")
        say("  [4] 跳过字体处理")
        choice = prompt("请选择", "1")
        if choice == "2":
            if not mappings:
                warn("没有缺失字体，无需生成别名。")
                return "rewrite"
            if yes_no("将写入 ~/.config/fontconfig/conf.d（可用工具撤销），继续？", True):
                okflag, msg = install_font_aliases(mappings)
                if okflag:
                    ok(msg)
                    ok("提示：已运行的 WPS 需要重启才能读到新别名。")
                    self.reload_fonts()
                    return "alias"
                err(msg)
            return "rewrite"
        if choice == "3":
            src = prompt("Windows 字体目录（如 /mnt/c/Windows/Fonts）", "")
            if not src:
                warn("未输入目录，改为改写文件内字体名。")
                return "rewrite"
            okflag, msg = import_windows_fonts(src, self.registry)
            if okflag:
                ok(msg)
                self.reload_fonts()
                return "install"
            err(msg)
            return "rewrite"
        if choice == "4":
            return "none"
        return "rewrite"

    def repair_wizard(self):
        if not self.ensure_files():
            return
        say("")
        step("第 1 步：先做一次体检")
        existing = [f for f in self.files if f in self.scans]
        if len(existing) != len(self.files):
            self.scan_many()
        else:
            info("使用上次扫描结果；如需重新扫描请选择 [3]。")
        total_high = sum((self.scans.get(f).high_count() if self.scans.get(f) else 0)
                         for f in self.files)
        if not total_high:
            info("未发现高危问题，仍可继续做一致性修复。")

        step("第 2 步：选择修复范围")
        enabled = self.choose_rule_set()

        step("第 3 步：字体策略")
        font_mode = self.apply_font_strategy()
        if font_mode in ("alias", "install", "none"):
            enabled.discard("R001")

        step("第 4 步：输出设置")
        default_out = "与原文件同目录"
        out = prompt("输出目录（留空 = %s）" % default_out, "")
        self.out_dir = os.path.abspath(os.path.expanduser(out)) if out else None
        overwrite = yes_no("是否直接覆盖原文件？（强烈建议否，先对比再替换）", False)

        say("")
        hr("=")
        say(_c("即将执行", C.BOLD))
        say("  文件：%d 个" % len(self.files))
        say("  规则：%s" % ", ".join(sorted(enabled)))
        say("  字体策略：%s" % self.font_mode_label(font_mode))
        say("  输出：%s" % (
            "覆盖原文件" if overwrite else
            "同目录 xxx.fixed.pptx" if not self.out_dir else self.out_dir))
        say("  原文件会先自动备份。")
        hr("=")
        if not yes_no("确认开始修复？", True):
            info("已取消。")
            return

        step("第 5 步：修复")
        results = []
        for idx, path in enumerate(self.files, 1):
            say("")
            step("[%d/%d] 修复：%s" % (idx, len(self.files), os.path.basename(path)))
            res = process_file(
                path, enabled, options=self.options,
                registry=self.registry, resolver=self.resolver,
                out_dir=self.out_dir, overwrite=overwrite,
                font_mode=font_mode)
            results.append(res)
            if not res.ok:
                err(res.error)
                continue
            ok("已生成：%s" % res.output_path)
            if res.backup_path:
                info("原文件备份：%s" % res.backup_path)
            print_changes(res.fixed)
            if res.postfix is not None:
                print_rule_summary(res.postfix)
            json_path, md_path = write_reports(res, self.env, self.out_dir)
            if md_path:
                info("报告：%s" % md_path)
        self.results = results

        step("第 6 步：完成后验证")
        good = [r for r in results if r.ok]
        say("  成功 %d 个，失败 %d 个。" % (len(good), len(results) - len(good)))
        for r in results:
            if not r.ok:
                err("%s：%s" % (os.path.basename(r.path), r.error))
        if good and yes_no("是否用 LibreOffice 渲染对比修复前后（确认没丢内容）？", False):
            self.render_results(good)
        say("")
        info("请在 WPS 中打开修复后的文件逐页确认；如需回滚，用备份文件替换即可。")

    @staticmethod
    def font_mode_label(mode):
        return {
            "rewrite": "改写文件内字体名",
            "alias": "fontconfig 别名（不改文件字体）",
            "install": "已导入 Windows 字体",
            "none": "跳过字体处理",
        }.get(mode, mode)

    # -- 渲染验证 -----------------------------------------------------------
    def render_results(self, results):
        for res in results:
            if not res.output_path:
                continue
            say("")
            step("渲染对比：%s" % os.path.basename(res.path))
            try:
                r = compare_render(res.path, res.output_path)
            except Exception as exc:
                err("渲染失败：%s" % exc)
                continue
            res.render = r
            if not r.get("ok"):
                err(r.get("error", "渲染失败"))
                continue
            say("  页数：%d → %d %s" % (
                r.get("pages_before"), r.get("pages_after"),
                "（一致）" if r.get("same_pages") else "（不一致）"))
            say("  文字内容相似度：%.2f%%" % (r.get("text_similarity", 0) * 100))
            if r.get("content_preserved"):
                ok("内容保持通过（页数与文字基本一致）")
            else:
                warn("内容可能有变化，请人工对比 PDF/WPS。")
            json_path, md_path = write_reports(res, self.env, self.out_dir)
            if md_path:
                info("报告已更新：%s" % md_path)

    def render_last(self):
        if not self.results:
            warn("还没有修复结果。请先执行 [4] 修复向导。")
            return
        self.render_results([r for r in self.results if r.ok])

    # -- 字体菜单 -----------------------------------------------------------
    def font_menu(self):
        while True:
            say("")
            say(_c("字体兼容方案", C.BOLD))
            say("  [1] 导入 Windows 字体（复制到用户字体目录并刷新缓存）")
            say("  [2] 生成/更新 fontconfig 字体别名（基于当前文件的缺失字体）")
            say("  [3] 撤销本工具生成的字体别名")
            say("  [4] 查询某个字体在本机的实际匹配结果")
            say("  [5] 查看 Windows 字体 → Linux 字体映射表")
            say("  [0] 返回主菜单")
            choice = prompt("请选择", "0")
            if choice == "0":
                return
            if choice == "1":
                self._menu_import_fonts()
            elif choice == "2":
                self._menu_make_aliases()
            elif choice == "3":
                okflag, msg = remove_font_aliases()
                (ok if okflag else warn)(msg)
                if okflag:
                    self.reload_fonts()
            elif choice == "4":
                name = prompt("字体名（例如 宋体 / SimSun / Calibri）", "")
                if name:
                    self._show_font_match(name)
            elif choice == "5":
                self._show_font_table()

    def _menu_import_fonts(self):
        say("常见目录：/mnt/c/Windows/Fonts（WSL/双系统）、U 盘里的 Windows/Fonts")
        src = prompt("Windows 字体目录", "")
        if not src:
            return
        okflag, msg = import_windows_fonts(src, self.registry)
        (ok if okflag else err)(msg)
        if okflag:
            self.reload_fonts()
            info("字体已导入；重新执行 [3]/[4] 时就会按“已安装字体”处理。")

    def _menu_make_aliases(self):
        if not self.scans:
            if self.files:
                self.scan_many()
            else:
                warn("请先选择文件并扫描，或手动输入映射。")
        mappings = self._missing_font_mappings()
        if not mappings:
            warn("没有检测到缺失字体。")
            say("也可以手动建立映射，格式：原字体=目标字体，每行一个。")
            raw = prompt("映射（留空返回）", "")
            if raw:
                manual = []
                for line in raw.replace("；", ";").replace("\n", ";").split(";"):
                    if "=" in line:
                        a, b = line.split("=", 1)
                        manual.append((a.strip(), b.strip()))
                mappings = manual
        if not mappings:
            return
        say("将生成以下映射：")
        for src, dst in mappings:
            say("  %s  →  %s" % (src, dst))
        if yes_no("写入 ~/.config/fontconfig/conf.d/99-wps-windows-fonts.conf？", True):
            okflag, msg = install_font_aliases(mappings)
            (ok if okflag else err)(msg)
            if okflag:
                info("请重启 WPS 使别名生效；撤销请选择 [3]。")
                self.reload_fonts()

    def _show_font_match(self, name):
        if not self.registry.available:
            warn("fontconfig 不可用，无法查询。")
            return
        if self.registry.installed(name):
            ok("「%s」本机已安装：%s" % (name, self.registry.display_name(name)))
        else:
            warn("「%s」本机未安装。" % name)
        m = self.registry.match(name)
        if m:
            say("  实际匹配：%s（%s）" % (m[0], m[1]))
        decision = self.resolver.decide(name)
        say("  本工具建议替换为：%s" % (decision.target or "（无需替换）"))
        if decision.companion_latin:
            say("  西文搭配字体：%s" % decision.companion_latin)
        font = self.registry.font_for(name)
        if font is not None:
            say("  行高度量：%.2f em（单倍行距 100%% 时的行高/字号）" % font.line_em())

    def _show_font_table(self):
        rows = [
            ("宋体 / SimSun / NSimSun", "Noto Serif CJK SC / AR PL UMing CN"),
            ("微软雅黑 / Microsoft YaHei", "Noto Sans CJK SC"),
            ("黑体 / SimHei", "Noto Sans CJK SC"),
            ("楷体 / KaiTi", "AR PL UKai CN"),
            ("仿宋 / FangSong", "Noto Serif CJK SC / AR PL UMing CN"),
            ("等线 / DengXian", "Noto Sans CJK SC"),
            ("Calibri", "Carlito / Liberation Sans"),
            ("Cambria", "Caladea / Liberation Serif"),
            ("Times New Roman", "Liberation Serif / Tinos"),
            ("Arial", "Liberation Sans / Arimo"),
            ("Courier New", "Liberation Mono"),
            ("Wingdings / Webdings / Symbol", "无替代：必须安装原字体"),
        ]
        hr()
        say("%-34s %s" % ("Windows 字体", "Linux 替代（优先已安装者）"))
        hr()
        for src, dst in rows:
            say("%-34s %s" % (src, dst))
        hr()
        say("实际选择顺序由本机已安装字体决定；用 [4] 可查看具体字体名的匹配结果。")

    # -- 帮助 ---------------------------------------------------------------
    def show_help(self):
        say("")
        say(_c("为什么会错位（按影响从大到小）", C.BOLD))
        causes = [
            "① 字体缺失/替换：宋体、微软雅黑、Calibri 等在 UOS/麒麟 上通常没有，"
            "WPS 用 fontconfig 回退字体，字宽与行高全变。",
            "② 行高差异：宋体单倍行距约 1.0~1.2em，思源黑体/宋体约 1.44em，"
            "换字体后多行文本被“撑高”，整段下移。",
            "③ 中英文混排：中文 Windows 字体自带西文字形，"
            "Linux 回退后西文字宽不同，断行位置改变。",
            "④ 自动缩放残留：PowerPoint 写进文件里的 fontScale/lnSpcReduction，"
            "WPS 按另一套字体度量重算。",
            "⑤ 内边距/单元格边距默认值不一致，文字整体偏移几个点。",
            "⑥ 不换行文本框、随文字自动调整高度的形状，"
            "在不同度量下宽度/高度变化，遮挡相邻元素。",
            "⑦ 符号字体（Wingdings/Webdings）：字符显示成字母，看起来像乱码。",
            "⑧ 语言标记与断行规则：lang=en-US 的中文按西文规则断行。",
            "⑨ 组合缩放、竖排/文字变形：WPS 与 PowerPoint 处理方式不同。",
        ]
        for c in causes:
            say(textwrap.fill(c, 76, initial_indent="  ", subsequent_indent="    "))
        say("")
        say(_c("推荐流程", C.BOLD))
        steps = [
            "1) 主菜单 [2] 看环境与中文字体；",
            "2) 主菜单 [1] 选择文件 → [3] 扫描，确认问题分布；",
            "3) 若缺字体很多：先到 [6] 导入 Windows 字体（最保真）或生成字体别名；",
            "4) [4] 修复向导：选“推荐修复”，字体策略建议“改写文件内字体名”；",
            "5) 在 WPS 中打开 xxx.fixed.pptx 逐页确认；",
            "6) 如仍有个别页面溢出，重新运行 [4] 并勾选 R009（溢出缩字）或 R014（固定行距）。",
        ]
        for s in steps:
            say("  " + s)
        say("")
        say(_c("安全与回滚", C.BOLD))
        say("  · 默认不覆盖原文件，修复版为 xxx.fixed.pptx；原文件同时备份为 xxx.orig-时间.pptx；")
        say("  · 所有修改都有报告（.report.md / .report.json）；")
        say("  · 字体别名写在用户目录，可用 [6]/[3] 随时撤销；")
        say("  · 导入的字体在 ~/.local/share/fonts/windows-fonts，删除目录即可。")
        say("")
        say(_c("说明", C.BOLD))
        say("  · 溢出是“按替代字体度量做的估算”，与 WPS 实际渲染可能有 1~2% 误差；")
        say("  · 渲染验证使用 LibreOffice，仅用于确认“没丢内容/页数没变”，不能完全代表 WPS。")
        say("")


# ============================================================================
# 非交互式批量处理与命令行入口
# ============================================================================

def run_batch(args, registry, resolver, env):
    files, errors = expand_paths(args.paths)
    for e in errors:
        err(e)
    if not files:
        err("没有可处理的文件。")
        return 1
    if args.rules is None and args.scan and not args.fix:
        rules = set(RULES_BY_ID)
    else:
        rules = rules_from_arg(args.rules)
    if args.rules_all:
        rules = set(RULES_BY_ID)
    if args.font_mode == "none":
        rules.discard("R001")

    font_mode = args.font_mode
    if font_mode == "ask":
        font_mode = "rewrite"

    # 安装/别名模式的前置动作
    if args.fix and font_mode == "install":
        src = args.fonts_dir
        if not src:
            err("--font-mode install 需要配合 --fonts-dir 指定 Windows 字体目录。")
            return 1
        okflag, msg = import_windows_fonts(src, registry)
        (ok if okflag else err)(msg)
        if not okflag:
            return 1
        registry = FontRegistry()
        resolver = FontResolver(registry)
        env = environment_info(registry)
    elif args.fix and font_mode == "alias":
        engines = []
        for path in files:
            try:
                engine, _deck, pkg = scan_file(path, rules, registry, resolver, args_options(args))
                pkg.close()
                engines.append(engine)
            except Exception as exc:
                err("体检失败 %s：%s" % (path, exc))
        mappings = collect_missing_fonts(engines, resolver)
        if mappings:
            okflag, msg = install_font_aliases(mappings)
            (ok if okflag else err)(msg)
            registry = FontRegistry()
            resolver = FontResolver(registry)
            env = environment_info(registry)
        else:
            info("未检测到缺失字体，跳过别名生成。")

    options = args_options(args)
    results = []
    for idx, path in enumerate(files, 1):
        say("")
        step("[%d/%d] %s" % (idx, len(files), path))
        if args.scan and not args.fix:
            if os.path.splitext(path)[1].lower() in LEGACY_EXTS:
                err("旧版二进制格式无法直接体检，请先另存为 .pptx：%s" % path)
                res = FileResult(path)
                res.error = "旧版二进制格式"
                results.append(res)
                continue
            try:
                engine, deck, pkg = scan_file(path, rules, registry, resolver, options)
                pkg.close()
            except Exception as exc:
                err("体检失败：%s: %s" % (type(exc).__name__, exc))
                results.append(FileResult(path))
                results[-1].error = str(exc)
                continue
            print_rule_summary(engine)
            res = FileResult(path)
            res.ok = True
            res.prefix = engine
            res.duration = 0.0
            json_path, md_path = write_reports(res, env, args.out_dir)
            if md_path:
                info("报告：%s" % md_path)
            results.append(res)
            continue

        out_dir = args.out_dir
        if args.output:
            if len(files) > 1:
                err("--output 只能用于单个文件。")
                return 1
            res = process_file(path, rules, options, registry, resolver,
                               out_dir=out_dir, suffix=args.suffix,
                               overwrite=False, backup=not args.no_backup,
                               font_mode=font_mode,
                               do_render_check=args.render_check)
            if res.ok:
                # 自定义输出路径
                try:
                    shutil.move(res.output_path, os.path.abspath(args.output))
                    res.output_path = os.path.abspath(args.output)
                except (IOError, OSError) as exc:
                    err("无法写入 %s：%s" % (args.output, exc))
        else:
            res = process_file(path, rules, options, registry, resolver,
                               out_dir=out_dir, suffix=args.suffix,
                               overwrite=args.overwrite,
                               backup=not args.no_backup,
                               font_mode=font_mode,
                               do_render_check=args.render_check)
        results.append(res)
        if not res.ok:
            err(res.error)
            continue
        ok("输出：%s" % res.output_path)
        if res.backup_path:
            info("备份：%s" % res.backup_path)
        print_changes(res.fixed)
        if res.postfix is not None:
            print_rule_summary(res.postfix)
        if res.render is not None:
            print_render_result(res.render)
        json_path, md_path = write_reports(res, env, args.out_dir)
        if md_path:
            info("报告：%s" % md_path)

    say("")
    good = [r for r in results if r.ok]
    say("完成：成功 %d / %d" % (len(good), len(results)))
    if args.json:
        payload = [build_report_dict(r, env) for r in results]
        say(json.dumps(payload, ensure_ascii=False, indent=2))
    if args.scan and not args.fix:
        high = sum(1 for r in results if r.prefix and r.prefix.high_count() > 0)
        return 2 if high else 0
    return 0 if len(good) == len(results) else 1


def args_options(args):
    opts = {"line_ratio": getattr(args, "line_ratio", 1.2),
            "shrink_min": getattr(args, "shrink_min", 0.25)}
    return opts


def print_render_result(r):
    if not r.get("ok"):
        warn("渲染验证不可用：%s" % r.get("error", ""))
        return
    say("  渲染验证：页数 %s → %s，文字相似度 %.2f%%，%s" % (
        r.get("pages_before"), r.get("pages_after"),
        r.get("text_similarity", 0) * 100,
        "内容保持通过" if r.get("content_preserved") else "需人工确认"))


def build_argparser():
    parser = argparse.ArgumentParser(
        prog="wps_pptx_repair",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="%s v%s —— 检测并修复 WPS 演示文稿在 UOS/麒麟 上的文字错位问题"
                    % (TOOL_NAME, VERSION),
        epilog=textwrap.dedent("""\
            示例：
              ./wps_pptx_repair.sh                          # 交互式向导（推荐）
              ./wps_pptx_repair.sh 课件.pptx                # 交互式，预载文件
              ./wps_pptx_repair.sh 课件.pptx --scan         # 只体检
              ./wps_pptx_repair.sh 课件.pptx --fix -y       # 直接修复（推荐项）
              ./wps_pptx_repair.sh ./slides --fix -y --font-mode rewrite
              ./wps_pptx_repair.sh 课件.pptx --fix -y --rules all --render-check
              ./wps_pptx_repair.sh --import-fonts /mnt/c/Windows/Fonts
              ./wps_pptx_repair.sh 课件.pptx --install-aliases
              ./wps_pptx_repair.sh --remove-aliases
              ./wps_pptx_repair.sh --list-rules
              ./wps_pptx_repair.sh --explain R014
            """))
    parser.add_argument("paths", nargs="*", help="PPTX 文件 / 目录 / 通配符（可多个）")
    parser.add_argument("--scan", action="store_true", help="只体检，不改文件")
    parser.add_argument("--fix", action="store_true", help="执行修复（非交互）")
    parser.add_argument("-y", "--yes", action="store_true", help="非交互确认")
    parser.add_argument("--rules", default=None,
                        help="修复规则：default/all/font/layout 或 R001,R009 …"
                             "（--scan 时默认检测全部规则）")
    parser.add_argument("--rules-all", action="store_true", help="等价于 --rules all")
    parser.add_argument("--font-mode", default="rewrite",
                        choices=["rewrite", "alias", "install", "none"],
                        help="字体策略：rewrite 改写文件（默认）/ alias 系统别名 / "
                             "install 先导入字体 / none 跳过")
    parser.add_argument("--fonts-dir", default=None,
                        help="--font-mode install 时使用的 Windows 字体目录")
    parser.add_argument("--out-dir", default=None, help="输出目录")
    parser.add_argument("--output", default=None, help="指定单个文件输出路径")
    parser.add_argument("--suffix", default=".fixed", help="输出文件名后缀（默认 .fixed）")
    parser.add_argument("--overwrite", action="store_true", help="覆盖原文件（仍会备份）")
    parser.add_argument("--no-backup", action="store_true", help="不备份原文件")
    parser.add_argument("--render-check", action="store_true",
                        help="用 LibreOffice 渲染对比修复前后（需已安装）")
    parser.add_argument("--line-ratio", type=float, default=1.2,
                        help="R014 固定行距的倍率（默认 1.2）")
    parser.add_argument("--shrink-min", type=float, default=0.25,
                        help="R009 缩小字号的下限（默认 0.25）")
    parser.add_argument("--json", action="store_true", help="在标准输出附加 JSON 报告")
    parser.add_argument("--install-aliases", action="store_true",
                        help="扫描文件中的缺失字体并生成 fontconfig 别名")
    parser.add_argument("--remove-aliases", action="store_true", help="撤销字体别名")
    parser.add_argument("--import-fonts", default=None, metavar="DIR",
                        help="导入 Windows 字体目录到用户字体目录")
    parser.add_argument("--list-rules", action="store_true", help="列出全部规则")
    parser.add_argument("--explain", default=None, metavar="RID",
                        help="解释某条规则（如 R001）")
    parser.add_argument("--version", action="store_true", help="显示版本")
    return parser


def main(argv=None):
    setup_console()
    parser = build_argparser()
    args = parser.parse_args(argv)

    if args.version:
        say("%s v%s" % (TOOL_NAME, VERSION))
        return 0
    if args.list_rules:
        print_rules_table()
        return 0
    if args.explain:
        return 0 if print_rule_detail(args.explain) else 1

    if args.remove_aliases:
        okflag, msg = remove_font_aliases()
        (ok if okflag else warn)(msg)
        return 0 if okflag else 1

    registry = FontRegistry()
    resolver = FontResolver(registry)

    if args.import_fonts:
        okflag, msg = import_windows_fonts(args.import_fonts, registry)
        (ok if okflag else err)(msg)
        return 0 if okflag else 1

    if args.install_aliases:
        files, errors = expand_paths(args.paths)
        for e in errors:
            err(e)
        if not files:
            err("请指定要扫描的演示文稿。")
            return 1
        rules = set(RULES_BY_ID)
        engines = []
        for path in files:
            try:
                engine, _deck, pkg = scan_file(path, rules, registry, resolver)
                pkg.close()
                engines.append(engine)
            except Exception as exc:
                err("体检失败 %s：%s" % (path, exc))
        mappings = collect_missing_fonts(engines, resolver)
        if not mappings:
            info("未检测到缺失字体，无需生成别名。")
            return 0
        for src, dst in mappings:
            say("  %s → %s" % (src, dst))
        okflag, msg = install_font_aliases(mappings)
        (ok if okflag else err)(msg)
        return 0 if okflag else 1

    if args.paths and (args.scan or args.fix):
        env = environment_info(registry)
        return run_batch(args, registry, resolver, env)

    if not sys.stdin.isatty():
        err("当前环境不是交互式终端。请使用 --scan 或 --fix 参数。")
        parser.print_help()
        return 2

    app = InteractiveApp(args.paths, args_options(args))
    return app.run()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        say("")
        sys.exit(130)
__WPS_PPTX_REPAIR_PY__

"$PY" "$TMP_PY" "$@"
STATUS=$?
exit "$STATUS"
