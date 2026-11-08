# -*- coding: utf-8 -*-
"""
燃气工程设计变更单 —— 变更原因读写核心库

知识来源：南昌燃气工程设计变更单历史归档
  E:\\2024\\2024结算相关\\1.材料平差\\设计变更\\hk  (136 份 .xls 变更单)

本模块提供：
  1. ChangeOrder  —— 变更单读写（xlrd 读 / Excel COM 写）
  2. Corpus       —— 历史语料库加载与相似案例检索
  3. ReasonBuilder—— 依据历史规律合成变更原因

环境要求：
  - Python 3.12/3.13
  - xlrd>=2.0 (读 .xls)
  - pywin32 (写 .xls，需本机安装 Excel)
"""

import os
import re
import sys
import glob
import json
import shutil
import tempfile

try:
    import xlrd
except ImportError:  # pragma: no cover
    raise SystemExit("缺少依赖 xlrd，请执行: pip install xlrd")

if hasattr(sys.stdout, "buffer"):
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


# ---------------------------------------------------------------- 常量

#: 历史变更单归档根目录（可多个）
#:
#: 取值优先级：
#:   1) 环境变量 GAS_CHANGE_ARCHIVE（多个路径用 ; 分隔）
#:   2) 本文件同级的 config.json 中的 "archive_roots"（数组）或 "archive_root"（字符串）
#:   3) 下面的默认值（本机路径）
#:
#: 用 `python fill_change_reason.py reindex` 可重建语料缓存。
_DEFAULT_ARCHIVE = [
    r"E:\2024\2024结算相关\1.材料平差\设计变更\hk",
    r"D:\Documents\桌面\历史变更单",
]


def _resolve_archive_roots():
    """解析归档根目录列表：环境变量 > config.json > 默认值。

    显式设置即使路径不存在也予以采纳，这样拼错路径时 doctor 能明确报错，
    而不是静默回退到默认路径。
    """
    roots = []
    env = os.environ.get("GAS_CHANGE_ARCHIVE")
    if env and env.strip():
        roots = [p for p in re.split(r"[;|]", env.replace("\n", ";")) if p.strip()]
    else:
        cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        if os.path.exists(cfg):
            try:
                import json
                with open(cfg, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                v = data.get("archive_roots")
                if isinstance(v, list):
                    roots = [str(x) for x in v if str(x).strip()]
                elif data.get("archive_root"):
                    roots = [str(data["archive_root"])]
            except Exception:
                pass
    if not roots:
        roots = list(_DEFAULT_ARCHIVE)
    return [os.path.abspath(r.strip()) for r in roots]


ARCHIVE_ROOTS = _resolve_archive_roots()

#: 兼容旧接口：单一根目录（第一个存在的）
ARCHIVE_ROOT = next((r for r in ARCHIVE_ROOTS if os.path.isdir(r)),
                    ARCHIVE_ROOTS[0] if ARCHIVE_ROOTS else "")

#: 变更单打印表名关键字（排除作废/旧版）
PRINT_SHEET_KEY = "变更单打印"

#: 变更原因占位符
PLACEHOLDER = "[变更原因]"

#: 申请单结尾语（314 份语料中 213 份使用）
TAIL_APP = "兹提出设计变更申请单，请予以审批。"

#: 工程量变更单结尾语
TAIL_QTY = "兹提出工程量变更单。"

#: 罕见结尾语（历史上是笔误，勿模仿）
TAIL_APP_RARE = "拟调整如下材料，兹提出变更申请单。"

#: 申请单结尾语的已知变体（按 314 份语料频次排序）
#: 老归档 99% 用标准句；新归档 40% 用「兹提出变更申请单，请予以审批。」
TAIL_APP_VARIANTS = [
    "兹提出设计变更申请单，请予以审批。",   # 213
    "兹提出变更申请单，请予以审批。",       # 72
    "兹提出变更单申请，请予以审批。",       # 15
    "兹提出工程量变更申请单，请予以审批。",  # 4
    "兹提出设计变更清单，请予以审批。",     # 1
]

#: 识别正文是否已自带结尾语（避免重复拼接）
_TAIL_RE = re.compile(r"(兹提出[^。；]*?(?:申请单|变更单|变更清单|工程量变更申请单)[^。；]*?。)\s*$")


def detect_tail(text):
    """识别文本结尾是否已是结尾语；是则返回该结尾语，否则空串。"""
    m = _TAIL_RE.search(str(text or "").strip())
    return m.group(1) if m else ""


# ---------------------------------------------------------------- 工具

def _norm(name):
    """材料名归一化：去括号注释、去规格后缀、小写。

    '20#钢无缝管_D76_4mm_双层熔结环氧粉末防腐' -> '20#钢无缝管_d76'
    '钢制弯头_20#碳钢_DN80_4mm_1.5D_90°'      -> '钢制弯头_dn80_90°'
    """
    s = str(name).strip()
    s = re.sub(r"（[^）]*）", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    # 去尾部规格说明：_4mm_... / _1.6MPa_...
    s = re.sub(r"_\d+(?:\.\d+)?mm.*$", "", s)
    s = re.sub(r"_\d+(?:\.\d+)?MPa.*$", "", s)
    return s.strip().lower()


def _short_spec(name):
    """截断材料名到规格段，用于正文引用（对齐历史措辞）。

    '钢制单头牙_DN15_外牙_50cm' -> '钢制单头牙_DN15'
    '钢制弯头_DN80_90°'         -> '钢制弯头_DN80_90°'
    '20#钢无缝管_D76'            -> '20#钢无缝管_D76'
    """
    s = str(name).strip()
    parts = s.split("_")
    for i, p in enumerate(parts[1:], start=1):
        if re.match(r"^(DN|dn|D|d)\d+$", p):
            keep = parts[: i + 1]
            if i + 1 < len(parts) and "°" in parts[i + 1]:
                keep.append(parts[i + 1])
            return "_".join(keep)
    return s


def _clean(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


# ---------------------------------------------------------------- 变更单读取

class ChangeOrder(object):
    """一个设计变更单工作簿的只读视图。

    典型工作表：'1-初步比对' / '2-汇总' / '587-变更单打印' /
                 '4-汇总打印版 (新版)' / '5-变更估价' / '@变更单EXCEL模板'
    """

    def __init__(self, path):
        self.path = os.path.abspath(path)
        self.name = os.path.basename(self.path)
        self._wb = None
        self._print_sheet = None
        self._merged = None

    # -- 基础访问 ------------------------------------------------------

    @property
    def workbook(self):
        if self._wb is None:
            self._wb = xlrd.open_workbook(self.path, formatting_info=True)
        return self._wb

    @property
    def sheet_names(self):
        return [s.name for s in self.workbook.sheets()]

    def sheet(self, name):
        try:
            return self.workbook.sheet_by_name(name)
        except Exception:
            return None

    @property
    def print_sheet(self):
        """定位 '*-变更单打印' 工作表（排除作废/旧版）。"""
        if self._print_sheet is not None:
            return self._print_sheet
        cands = [s for s in self.workbook.sheets()
                 if PRINT_SHEET_KEY in s.name and "作废" not in s.name and "旧版" not in s.name]
        self._print_sheet = cands[0] if cands else None
        return self._print_sheet

    @property
    def merged(self):
        if self._merged is None:
            self._merged = self.workbook.sheet_by_name(self.print_sheet.name).merged_cells \
                if self.print_sheet else []
        return self._merged

    # -- 字段提取 ------------------------------------------------------

    def header(self):
        """返回 (工程编号, 工程名称)。"""
        sh = self.print_sheet
        if sh is None:
            return "", ""
        for r in range(min(10, sh.nrows)):
            if _clean(sh.cell_value(r, 0)) == "工程编号":
                return (_clean(sh.cell_value(r, 2)), _clean(sh.cell_value(r, 7)))
        return "", ""

    def _section_materials(self, sh, label_col, name_col, qty_col):
        """读取『增加材料:』/『减少材料:』区块。

        label_col/name_col/qty_col 分别为标签列、材料名列、数量列。
        以 A 列出现『变更预算』作为区块结束标志，避免把签字栏
        （预算人员/施工单位/监理单位…）误读成材料。
        """
        out = []
        label = None
        for r in range(8, min(40, sh.nrows)):
            cell = _clean(sh.cell_value(r, label_col))
            if cell in ("增加材料:", "增加材料："):
                label = "add"
                continue
            if cell in ("减少材料:", "减少材料："):
                label = "red"
                continue
            if label is None:
                continue
            # 表格下方的签字栏不是材料
            col_a = _clean(sh.cell_value(r, 0))
            if col_a.startswith(("变更预算", "施工单位", "监理单位", "变更内容")):
                break
            nm = _clean(sh.cell_value(r, name_col))
            if not nm:
                continue
            if nm.startswith(("变更预算", "施工单位", "监理单位", "变更内容",
                              "预算人员", "日 期", "签 章")):
                break
            out.append({"name": nm, "qty": _clean(sh.cell_value(r, qty_col)),
                        "norm": _norm(nm)})
        return out

    def materials(self):
        """返回 (增加材料 list, 减少材料 list)。"""
        sh = self.print_sheet
        if sh is None:
            return [], []
        add = self._section_materials(sh, 0, 0, 6)
        red = self._section_materials(sh, 7, 7, 13)
        return add, red

    def reason_slots(self):
        """定位变更原因的写入位置。

        返回 dict:
          app : (sheet_name, row, col)  申请单「设计变更原因」右侧合并区左上角
          qty : (sheet_name, row, col)  工程量变更单「[变更原因]」合并区左上角
        """
        sh = self.print_sheet
        if sh is None:
            return {}
        out = {}
        # 申请单：A 列 = '设计变更原因'，取值右侧合并区（B5 起）
        for r in range(min(12, sh.nrows)):
            if _clean(sh.cell_value(r, 0)) == "设计变更原因":
                for (rlo, rhi, clo, chi) in self.merged:
                    if rlo == r and clo == 1:
                        out["app"] = (sh.name, rlo, clo)
                        break
                else:
                    out["app"] = (sh.name, r, 1)
                break
        # 工程量变更单：定位「工程量变更单」标题行之下、行号最小的整行合并区
        # （xlrd 的 merged_cells 未按行排序，不能依赖遍历顺序，必须取 min(rlo)）
        title_r = None
        for r in range(sh.nrows):
            if "工程量变更单" in _clean(sh.cell_value(r, 0)):
                title_r = r
                break
        cands = []
        for (rlo, rhi, clo, chi) in self.merged:
            if clo == 0 and chi >= 5 and (title_r is None or rlo > title_r):
                cands.append((rlo, clo))
        if cands:
            rlo, clo = min(cands)          # 标题之下第一个整行合并区 = [变更原因]
            out["qty"] = (sh.name, rlo, clo)
        return out

    def current_reason(self):
        """读取已有变更原因（申请单侧），返回文本；无则空串。"""
        sh = self.print_sheet
        if sh is None:
            return ""
        for r in range(min(12, sh.nrows)):
            if _clean(sh.cell_value(r, 0)) == "设计变更原因":
                for c in range(1, sh.ncols):
                    v = _clean(sh.cell_value(r, c))
                    if v:
                        return v
        return ""

    def needs_fill(self):
        """True 表示变更原因尚未填写。"""
        cur = self.current_reason()
        return (not cur) or cur == PLACEHOLDER

    def read_draft(self):
        """读取两处变更原因单元格的现有内容（用户草稿）。

        返回 dict:
          app : 申请单现状文本（无则空串）
          qty : 工程量变更单现状文本（无则空串）
          app_addr / qty_addr : 对应 A1 地址
          has_draft : 是否存在可作为提示词的草稿
        """
        out = {"app": "", "qty": "", "app_addr": "", "qty_addr": "", "has_draft": False}
        slots = self.reason_slots()
        sh = self.print_sheet
        if sh is None:
            return out
        for key in ("app", "qty"):
            if key not in slots:
                continue
            _, r, c = slots[key]
            val = _clean(sh.cell_value(r, c))
            addr = _addr(r, c)
            if key == "app":
                out["app"], out["app_addr"] = val, addr
            else:
                out["qty"], out["qty_addr"] = val, addr
            if val and val != PLACEHOLDER:
                out["has_draft"] = True
        return out


# ---------------------------------------------------------------- 历史语料

class Corpus(object):
    """历史变更原因语料库。

    支持多个归档目录（root 可为字符串或列表）。
    """

    def __init__(self, root=None, cache=None):
        if root is None:
            roots = list(ARCHIVE_ROOTS)
        elif isinstance(root, (list, tuple)):
            roots = [os.path.abspath(r) for r in root if str(r).strip()]
        else:
            roots = [os.path.abspath(str(root))]
        self.roots = roots
        self.root = roots[0] if roots else ""
        self.cache = cache
        self.records = []
        if cache and os.path.exists(cache):
            with open(cache, "r", encoding="utf-8") as fh:
                self.records = json.load(fh)
        else:
            self.records = self._scan()
            if cache:
                try:
                    with open(cache, "w", encoding="utf-8") as fh:
                        json.dump(self.records, fh, ensure_ascii=False, indent=1)
                except OSError:
                    pass

    # -- 扫描 ----------------------------------------------------------

    def _scan(self):
        recs = []
        seen = set()
        files = []
        for root in self.roots:
            if not root or not os.path.isdir(root):
                continue
            for f in sorted(glob.glob(os.path.join(root, "**", "*.xls"), recursive=True)):
                if ".workbuddy" in f:
                    continue
                key = os.path.normcase(os.path.abspath(f))
                if key in seen:
                    continue
                seen.add(key)
                files.append(f)
        for f in files:
            try:
                co = ChangeOrder(f)
                sh = co.print_sheet
            except Exception:
                continue
            if sh is None:
                continue
            try:
                code, name = co.header()
                add, red = co.materials()
                reason = co.current_reason()
            except Exception:
                continue
            recs.append({
                "file": os.path.basename(f),
                "path": f,
                "code": code,
                "name": name,
                "reason": reason,
                "add": add,
                "red": red,
                "add_norm": sorted(set(x["norm"] for x in add)),
                "red_norm": sorted(set(x["norm"] for x in red)),
            })
        return recs

    # -- 检索 ----------------------------------------------------------

    @staticmethod
    def jaccard(a, b):
        A, B = set(a), set(b)
        if not (A | B):
            return 0.0
        return len(A & B) / len(A | B)

    def similar(self, add_norm, red_norm, top=8, min_score=0.0):
        """按材料增删集合相似度返回最相近的历史案例。

        分数 = 0.5 * Jaccard(增加) + 0.5 * Jaccard(减少)
        """
        scored = []
        for r in self.records:
            if not r["reason"] or r["reason"] == PLACEHOLDER:
                continue
            s = (0.5 * self.jaccard(add_norm, r["add_norm"])
                 + 0.5 * self.jaccard(red_norm, r["red_norm"]))
            if s >= min_score:
                scored.append((s, r))
        scored.sort(key=lambda x: -x[0])
        return scored[:top]

    def by_code(self, code):
        """按工程编号找历史记录（同名工程的其他标段可作参照）。"""
        return [r for r in self.records if r["code"] == code]

    # -- 润色场景的增强检索 --------------------------------------------

    #: 中文停用词（用于从草稿抽关键词）
    _STOP = set("因故该工程原定我部经与及其有关进行需要根据实际情况现场核实商议后决定"
                "造成相关材料有所增减减少增加如下问题出现情况时间工作要求以上"
                "本次变更理由关于对于同时并且然后因此所以但是"
                "一二三四五六七八九十的了在是有和就不人都一个"
                "予以提出审批兹申请单工程量变更设计变更原因"
                "相关对应配套管材管件材料变动有所" )

    @classmethod
    def _blocks(cls, text):
        """切分出语义块（不做块内滑窗）——这些词适合展示。"""
        t = str(text or "")
        t = re.sub(r"[（(][^）)]*[）)]", " ", t)          # 去括号注释
        t = re.sub(r"\d+(?:\.\d+)?\s*(?:米|m|个|片|套|处|根|副|件|支)", " ", t)
        t = re.sub(r"[^一-鿿A-Za-z_]+", " ", t)
        parts = re.split(r"[，,。;；、\s]+|"
                         r"(?:因|故|经|的|及|并|后|时|而|就|也|还|又|再)", t)
        out = set()
        for seg in parts:
            seg = seg.strip()
            if len(seg) < 3 or any(ch.isdigit() for ch in seg):
                continue
            if all(ch in cls._STOP for ch in seg):
                continue
            out.add(seg)
        return out

    @classmethod
    def _keywords(cls, text):
        """匹配用关键词 = 语义块 + 块内子串。

        块内滑窗能产生「接入口」「灶具接入口」这类更短的匹配词
        （历史原因的切块上下文与草稿不同，长块常匹配不上），
        但滑窗碎片不适合展示，故展示时另用 _blocks。
        """
        out = set(cls._blocks(text))
        for seg in cls._blocks(text):
            for n in (5, 4, 3):
                for i in range(len(seg) - n + 1):
                    w = seg[i:i + n]
                    if any(ch.isdigit() for ch in w):
                        continue
                    if all(ch in cls._STOP for ch in w):
                        continue
                    out.add(w)
        return out

    @classmethod
    def _display_words(cls, text, add=None, red=None):
        """可展示的命中词。

        = 语义块 + 块内 4–6 字子串 + 材料基类名，
          全部剔除程序性套话（报告公司…现场核实…等）。
        块内子串保证「矮立管移位」这类具体表述能被历史原文匹配到。
        """
        out = set()
        for seg in cls._blocks(text):
            if len(seg) >= 4 and not cls._is_proc(seg):
                out.add(seg)
            for n in (6, 5, 4):
                for i in range(len(seg) - n + 1):
                    w = seg[i:i + n]
                    if any(ch.isdigit() for ch in w):
                        continue
                    if cls._is_proc(w):
                        continue
                    if all(ch in cls._STOP for ch in w):
                        continue
                    out.add(w)
        for w in cls._material_words(add) | cls._material_words(red):
            if len(w) >= 3 and not cls._is_proc(w):
                out.add(w)
        return out

    @classmethod
    def _material_words(cls, items):
        """从材料清单提取基类名（如 '镀锌弯头_DN40_90°_内丝' → '镀锌弯头'）。

        草稿通常会提到具体材料，用材料名做关键词比中文分词可靠得多。
        """
        out = set()
        for x in items or []:
            name = x["name"] if isinstance(x, dict) else str(x)
            base = re.split(r"[_（(\s]", name)[0]
            if len(base) >= 2 and not any(ch.isdigit() for ch in base):
                out.add(base)
        return out

    #: 程序性套话：可作为「文风一致」的弱信号，但不具区分度，展示时排除
    _PROC = ("报告公司", "相关代表", "人员代表", "现场核实", "商议后决定", "核实后决定",
             "兹提出", "申请单", "请予以审批", "工程量变更单", "变更申请单",
             "公司相关", "代表到", "关代表", "表到现", "到现场", "场核实")

    @classmethod
    def _is_proc(cls, word):
        return any(p in word for p in cls._PROC)

    def related(self, add_norm, red_norm, code="", draft="", top=8,
                add=None, red=None):
        """为「润色」场景检索相关历史案例。

        与 similar() 的区别：多叠加两个维度，
        避免材料差异大时相似度全为 0 而返回无关案例——
          ① 材料 Jaccard（权重 3）
          ② 同类别工程（细分类型码相同，权重 1）
          ③ 草稿关键词与历史原因的命中（权重 1.5）

        关键词权重按 IDF 折减：像「报告公司相关代表到现场核实」这种
        历史语料里人人都有���通用套话，DF 极高、区分度为零，应当降权；
        而「灶具接入口位置」这类具体表述 DF 低，权重高。

        关键词来源：草稿语义块 + 增/减材料基类名。
        返回 (综合分, 材料分, 命中理由, 记录) 列表。
        """
        kw = self._keywords(draft)
        kw |= self._material_words(add) | self._material_words(red)
        cat = (code or "")[4:5] if code and len(code) >= 5 else ""
        n_all = max(1, len([r for r in self.records if r["reason"]]))
        scored = []
        for r in self.records:
            if not r["reason"] or r["reason"] == PLACEHOLDER:
                continue
            mat = (0.5 * self.jaccard(add_norm, r["add_norm"])
                   + 0.5 * self.jaccard(red_norm, r["red_norm"]))
            same_cat = bool(cat) and (r["code"] or "")[4:5] == cat
            hits = []
            hit_idf = 0.0
            for k in kw:
                if k in r["reason"]:
                    hits.append(k)
                    hit_idf += self._idf(k, n_all)
            total_idf = sum(self._idf(k, n_all) for k in kw) or 1.0
            kw_ratio = hit_idf / total_idf
            score = mat * 3.0 + (1.0 if same_cat else 0.0) + kw_ratio * 3.0
            if score <= 0:
                continue
            why = []
            if mat > 0:
                why.append("材料相似 %.2f" % mat)
            if same_cat:
                why.append("同类工程")
            # 展示用词：语义块 + 材料基类名（不含滑窗碎片与程序性套话）
            disp = self._display_words(draft, add, red)
            good = [k for k in disp if k in r["reason"] and self._idf(k, n_all) >= 0.3]
            good.sort(key=lambda k: (-len(k), -self._idf(k, n_all)))
            picked = []
            for k in good:
                if not any(k in p for p in picked):
                    picked.append(k)
            if picked:
                why.append("关键词命中：" + "、".join(picked[:3]))
            elif hits:
                why.append("文风与同类工程一致")
            scored.append((score, mat, "；".join(why), r))
        scored.sort(key=lambda x: -x[0])
        return scored[:top]

    #: 关键词文档频率缓存
    _df_cache = None

    def _doc_freq(self, word):
        """关键词在历史原因原文中的文档频率（缓存）。

        直接对**原始原因文本**做子串判断，而不是对转换后的词集——
        否则「报告公司相关代表到现场核实」这类套话因切分方式不同
        而统计失真，无法被 IDF 识别为低区分度词。
        """
        if Corpus._df_cache is None:
            df = {}
            for r in self.records:
                if not r["reason"] or r["reason"] == PLACEHOLDER:
                    continue
                txt = r["reason"]
                for w in set(self._keywords(txt)) | self._material_words(r["add"]) \
                        | self._material_words(r["red"]):
                    df[w] = df.get(w, 0) + 1
            Corpus._df_cache = df
        return Corpus._df_cache.get(word, 0)

    def _raw_doc_freq(self, word):
        """在历史原因**原文**中直接统计包含该词的文档数。

        必须基于原文子串匹配，不能只查 _keywords 词集：
        不同文本的切块结果不同（"报告公司相关代表…" vs "报告公司相关人员代表…"），
        若按词集统计，套话会被误判为"罕见词"而获得最高权重。
        """
        cache = getattr(Corpus, "_raw_df_cache", None)
        if cache is None:
            cache = {}
            for r in self.records:
                if not r["reason"] or r["reason"] == PLACEHOLDER:
                    continue
                txt = r["reason"]
                for w in self._keywords(txt):
                    if w in txt:
                        cache[w] = cache.get(w, 0) + 1
            Corpus._raw_df_cache = cache
        if word in cache:
            return cache[word]
        # 缓存未覆盖（来自草稿的切块）→ 直接遍历原文统计
        n = 0
        for r in self.records:
            if r["reason"] and word in r["reason"]:
                n += 1
        cache[word] = n
        return n

    def _idf(self, word, n_all):
        """IDF：越少见的词权重越高。

        用原始文本的文档频率计算，能正确识别「人人都有」的通用套话。
        """
        df = self._raw_doc_freq(word) / float(n_all)
        return 1.0 - min(df, 0.95)

    def stats(self):
        filled = [r for r in self.records if r["reason"] and r["reason"] != PLACEHOLDER]
        return {
            "total": len(self.records),
            "filled": len(filled),
            "pending": len(self.records) - len(filled),
        }


# ---------------------------------------------------------------- 原因合成

#: 触发因子 -> (关键词, 原因模板)
#: 模板中 {n} 为数量占位，{q} 会被替换为实际数量
FACTORS = [
    {
        "key": "pe_to_steel",
        "name": "PE管改钢管（材质转换）",
        "material_kw": ["pe管", "pe警示带", "pe管_dn", "鞍型", "热熔弯头_dn", "电熔"],
        "steel_kw": ["钢无缝管", "钢制弯头", "钢制三通", "钢制大小头", "焊接法兰", "热镀锌钢管", "镀锌"],
        "template": ("因原设计{old}在该处无法满足敷设及安装要求，经报告公司相关代表到现场核实并商议后决定，"
                     "改为采用{new}敷设，配套管件及连接方式相应调整，故造成相关管材管件有所增减"),
    },
    {
        "key": "instrument",
        "name": "仪表/阀门连接件",
        "material_kw": ["单头牙", "压力表", "铜球阀", "四氟垫", "金属缠绕垫", "过滤网", "跨接线"],
        "template": "因{what}需要{conn}，故增加{item}",
    },
    {
        "key": "flange_blind",
        "name": "法兰盲板封堵预留",
        "material_kw": ["盲板", "焊接法兰", "丝扣法兰"],
        "template": "因{what}采用焊接法兰加盲板进行封堵预留，故增加{item}",
    },
    {
        "key": "underground",
        "name": "地下管线避让",
        "material_kw": ["弯头", "三通", "套筒", "变径"],
        "template": ("该工程位于{where}，地下管线及其他设施较多，需根据现场开挖实际情况调整管线敷设位置及长度，"
                     "并增加管件绕行避让障碍，故造成相关管材管件有所增减"),
    },
    {
        "key": "protect_groove",
        "name": "埋深不够加保护沟",
        "material_kw": ["保护沟", "钢板"],
        "template": ("因{what}段管道埋深不够，根据公司相关要求需敷设{how}保护沟，"
                     "故造成相关材料有所增减"),
    },
    {
        "key": "collision",
        "name": "碰接方式变更",
        "material_kw": ["鞍型电熔", "不停输", "停气碰接", "碰接"],
        "template": ("原定碰接方式因{why}无法实施，经报告公司相关代表到现场核实商议后决定，"
                     "改为{how}碰接，故造成相关材料有所增减"),
    },
    {
        "key": "reserved",
        "name": "增加预留",
        "material_kw": ["预留", "端帽", "球阀阀门井"],
        "template": "因{why}需增加预留，故增加{item}，造成相关材料有所增加",
    },
]

#: 结尾语
TAILS = {
    "app": TAIL_APP,
    "qty": TAIL_QTY,
}


class ReasonBuilder(object):
    """依据材料清单生成变更原因。

    主路径走「材料级推理」（material_reason 模块）：
    从增/减材料反推**工程动作**，再按历史文风组装。

    保留本类作为兼容入口；material_reason 不可用时退化为模板兜底。
    """

    def __init__(self, corpus):
        self.corpus = corpus

    # -- 特征识别 ------------------------------------------------------

    @staticmethod
    def detect(add, red):
        """识别触发因子，返回命中列表。"""
        try:
            from material_reason import derive_actions
            return sorted(set(a.kind for a in derive_actions(add, red, "", "")))
        except ImportError:
            pass
        add_n = [x["norm"] for x in add]
        red_n = [x["norm"] for x in red]
        all_n = add_n + red_n
        hits = []
        if (any(("pe管" in n or "pe警示" in n) for n in red_n)
                and any(("钢无缝管" in n or "钢制" in n or "镀锌" in n) for n in add_n)):
            hits.append("pe_to_steel")
        if any(k in n for n in all_n
               for k in ("单头牙", "压力表", "铜球阀", "四氟垫", "金属缠绕垫", "过滤网", "跨接线")):
            hits.append("instrument")
        if any("盲板" in n for n in add_n):
            hits.append("flange_blind")
        if any("保护沟" in n for n in all_n):
            hits.append("protect_groove")
        return hits

    # -- 正文合成 ------------------------------------------------------

    def build(self, add, red, name="", code="", location="", extra=None):
        """生成变更原因正文（不含结尾语）。

        返回 (正文, 命中因子列表, 参考案例 list)
        """
        extra = (extra or "").strip()
        add_norm = sorted(set(x["norm"] for x in add))
        red_norm = sorted(set(x["norm"] for x in red))
        refs = self.corpus.similar(add_norm, red_norm, top=5)

        try:
            from material_reason import compose
            body, acts = compose(add, red, name=name, code=code,
                                  location=location, extra=extra)
            return body, sorted(set(a.kind for a in acts)), refs
        except ImportError:
            pass

        return self._build_fallback(add, red, name, location, extra), self.detect(add, red), refs

    # -- 兜底（material_reason 不可用时）-------------------------------

    def _build_fallback(self, add, red, name, location, extra):
        segs = []
        if extra:
            segs.append(extra.rstrip("。；; ") + "，")
        loc = location or self.infer_location(name)
        segs.append(self._site_clause(loc))
        body = "".join(segs)
        return body if body.endswith("；") else body + "；"

    @staticmethod
    def infer_location(name):
        """从工程名称推断现场场景。"""
        try:
            from material_reason import infer_scene
            s = infer_scene(name)
            if s:
                return s + "内"
        except ImportError:
            pass
        return ""

    @staticmethod
    def _site_clause(loc):
        if loc and any(k in loc for k in ("内", "厂房", "车间", "锅炉房", "厨房")):
            return ("因{loc}设备及设施较多，需根据现场实际情况调整管位，并增加管件绕行避让障碍，"
                    "故造成相关管材管件有所增减；".format(loc=loc))
        if loc:
            return ("该工程位于{loc}，地下管线及其他设施较多，需根据现场开挖实际情况"
                    "调整管线敷设位置及长度，并增加管件绕行避让障碍，"
                    "故造成相关管材管件有所增减；".format(loc=loc))
        return ("因该工程地下管线及其他设施较多，需根据现场开挖实际情况调整管线敷设位置"
                "及长度，并增加管件绕行避让障碍，故造成相关管材管件有所增减；")


# ---------------------------------------------------------------- 写入 (.xls)

class ExcelWriter(object):
    """通过 Excel COM 写入 .xls（xlrd 只能读）。

    .xls 老格式必须走 Excel COM；写入采用「先复制到临时文件 → 改 → 覆盖原文件」
    的方式，避免 Excel 独占锁导致原文件损坏。
    """

    def __init__(self):
        self._app = None

    def _excel(self):
        if self._app is None:
            import win32com.client
            self._app = win32com.client.DispatchEx("Excel.Application")
            self._app.Visible = False
            self._app.DisplayAlerts = False
        return self._app

    def close(self):
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None

    def write(self, path, sheet_name, cells):
        """写入多个单元格。

        cells: {"B5": "文本", "A49": "文本"}
        返回 True 表示成功。
        """
        src = os.path.abspath(path)
        tmp_dir = tempfile.mkdtemp(prefix="wbx_")
        tmp = os.path.join(tmp_dir, os.path.basename(src))
        shutil.copy2(src, tmp)
        ok = False
        wb = None
        try:
            app = self._excel()
            wb = app.Workbooks.Open(tmp)
            ws = wb.Worksheets(sheet_name)
            for addr, val in cells.items():
                ws.Range(addr).Value = val
            wb.Save()
            wb.Close()
            wb = None
            shutil.copy2(tmp, src)
            ok = True
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return ok


def fill_reason(path, app_text, qty_text=None):
    """把变更原因写入变更单的申请单与工程量变更单两个位置。

    app_text : 申请单「设计变更原因」正文（不含结尾语亦可，脚本会补全）
    qty_text : 工程量变更单正文；None 时自动由 app_text 派生

    不生成备份副本——写入前请自行确认源文件另有留存。
    """
    co = ChangeOrder(path)
    slots = co.reason_slots()
    if "app" not in slots:
        raise RuntimeError("未找到申请单变更原因单元格位置：%s" % path)

    # 安全校验：
    #   自动生成（未指定正文）→ 目标格已有内容则中止，避免盲目覆盖；
    #   显式指定正文（--body，AI 润色结果）→ 允许覆盖。
    sh = co.print_sheet
    existing = {}
    for key in ("app", "qty"):
        if key not in slots:
            continue
        _, r, c = slots[key]
        cur = _clean(sh.cell_value(r, c))
        if cur and cur != PLACEHOLDER:
            existing[key] = (r, c, cur)

    if existing and not app_text:
        detail = "；".join("%s→%s：%s" % (k, _addr(r, c), v[:40])
                           for k, (r, c, v) in existing.items())
        raise RuntimeError(
            "该变更单的变更原因已填写，为避免覆盖已有内容已中止。\n"
            "  已填写位置：%s\n"
            "  若是用户草稿需润色：请用 polish 命令取素材，再以 --body 写入润色结果；\n"
            "  若确需重写：请用 --body 指定新正文。" % detail
        )

    if existing:
        # 覆盖已有内容且不生成备份——原内容将被永久替换，需明确告知
        for key, (r, c, v) in existing.items():
            sys.stderr.write(
                "[警告] %s 单元格 %s 原有内容将被覆盖（不保留备份）：%s\n"
                % (key.upper(), _addr(r, c), v[:50])
            )

    app_text = (app_text or "").strip()

    # 结尾语：沿用该文件既有的风格（同一单位/同一批变更单的写法可能不同）
    tail_app = TAIL_APP
    existing_tail = detect_tail(existing.get("app", ("", "", ""))[2] if existing else "")
    if existing_tail:
        tail_app = existing_tail
    elif qty_text is None and detect_tail(app_text):
        tail_app = detect_tail(app_text)

    if app_text.endswith(tail_app):
        # 已含结尾语：仅规范化分隔符
        app_full = app_text
    else:
        # 正文 → 规范为「正文；结尾语」
        core = app_text.rstrip("；;。 ")
        app_full = (core + "；" + tail_app) if core else tail_app

    # 正文（不含结尾语），末尾统一以「；」衔接
    body = app_full[: -len(tail_app)].rstrip("；; ") if app_full.endswith(tail_app) else app_full
    qty_full = qty_text or (body + "；" + TAIL_QTY)
    cells = {}
    sh, r, c = slots["app"]
    cells[_addr(r, c)] = app_full
    if "qty" in slots:
        sh2, r2, c2 = slots["qty"]
        cells[_addr(r2, c2)] = qty_full

    w = ExcelWriter()
    try:
        ok = w.write(path, sh, cells)
    finally:
        w.close()
    return ok, cells, slots


def _addr(r, c):
    """0-based (row, col) -> A1 地址。"""
    s = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return "%s%d" % (s, r + 1)
