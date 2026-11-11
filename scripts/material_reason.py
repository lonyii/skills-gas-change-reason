# -*- coding: utf-8 -*-
"""
材料级推理引擎 —— 从变更材料清单反推**工程动作**，编制专属变更原因

方法论依据（历史语料实证）
--------------------------------
材料变动最复杂的历史案例（NC25E073 变动 40 项，正文 191 字）**只点名 0–1 种材料**，
靠「工程动作」承载信息：

    「将调压柜前的dn90PE球阀取消，改为定向钻出土点处安装dn110PE球阀」
    「改为架空敷设，且为保证用气量，矮立管改为DN150」

而少量、点状的专项材料才会被点名并带数量：

    「故增加鞍型电熔管件dn110-2个」
    「故增加钢制单头牙_DN15-4个」
    「故增加焊接法兰及相关配件2套」

因此本引擎的映射是  **材料 → 工程动作 → 原因表述**，
不是把材料清单复述一遍。动作类型：
    管材替换 / 接口规格调整 / 支管增设 / 预留封堵 /
    仪表连接 / 保护沟 / 防撞护栏
"""

__author__ = "lonyii"
__email__ = "lonyii@qq.com"

import re

# ---------------------------------------------------------------- 规格解析

_SPEC_RE = re.compile(r"(?:^|[_\-\s])(DN|dn|D|d)(\d+(?:\.\d+)?)")

_PE_KW = ("pe管", "pe100", "pe警示", "鞍型", "热熔", "电熔")
_STEEL_KW = ("钢无缝管", "无缝钢管", "热镀锌钢管", "镀锌钢管", "钢管")
_GALV_KW = ("镀锌",)
_PROTECT_KW = ("保护沟", "钢板", "警示板", "警示保护板")

#: 结尾语
TAIL_APP = "兹提出设计变更申请单，请予以审批。"
TAIL_QTY = "兹提出工程量变更单。"

#: 正文长度上限（对齐历史最长 234 字）
MAX_LEN = 234

#: 超长时的降级顺序：先降次要动作，保住主线（管材替换 / 接口调整）
_DOWNGRADE_ORDER = [
    ("anti_impact", "protect", "cap_reserve", "instrument"),
    ("branch_add", "resize"),
]


def parse_spec(name):
    """'钢制三通_DN65/15' → ('DN', 65.0) ; '20#钢无缝管_D76' → ('D', 76.0)"""
    m = _SPEC_RE.search(str(name))
    if not m:
        return None
    g1 = m.group(1)
    kind = "DN" if g1.upper() == "DN" else ("dn" if g1 == "dn" else g1)
    try:
        return (kind, float(m.group(2)))
    except ValueError:
        return None


def _qty_num(qty):
    if qty is None:
        return 0.0
    m = re.search(r"(\d+(?:\.\d+)?)", str(qty))
    return float(m.group(1)) if m else 0.0


def _is_pipe(norm):
    """是否为管材（可参与管段替换判定）。

    排除管件与附件：支架、法兰、转换、套管、护栏、井盖等都不是管材。
    """
    if "管" not in norm:
        return False
    if any(k in norm for k in ("弯头", "三通", "套筒", "变径", "支架", "法兰", "转换",
                               "套管", "护栏", "井盖", "警示", "保护", "球阀", "闸阀",
                               "切断阀", "过滤器", "补偿", "波纹", "放散", "表")):
        return False
    # 必须以管材类词结尾（管 / 钢管 / 软管）
    return norm.endswith(("管", "钢管", "软管")) or "管_" in norm or "_管" in norm


def _base(norm):
    return norm.split("_")[0]


def _main_spec(item):
    """三通/变径的主规格：三通_DN65/15 → DN65（斜杠前）；变径_dn200/110 → dn200。"""
    name = item["name"] if isinstance(item, dict) else str(item)
    m = re.search(r"(DN|dn|D|d)(\d+)(?:[/_](?:DN|dn|D|d)?\d+)?", name)
    if m:
        return "%s%s" % (m.group(1), m.group(2))
    return None


#: 正文引用：基类名用历史惯用简称
_ALIAS = {
    "钢制三通": "三通", "钢制弯头": "弯头", "钢制大小头": "大小头",
    "钢制单头牙": "单头牙", "20#钢无缝管": "无缝钢管", "20钢无缝管": "无缝钢管",
    "电熔90°弯头": "电熔弯头", "90°热熔弯头": "热熔弯头",
}


def _cite(name):
    """正文引用名：基类简称 + 规格段（与历史点名风格一致）。"""
    s = str(name)
    parts = s.split("_")
    for i, p in enumerate(parts[1:], start=1):
        if re.match(r"^(DN|dn|D|d)\d+$", p):
            keep = parts[: i + 1]
            if i + 1 < len(parts) and "°" in parts[i + 1]:
                keep.append(parts[i + 1])
            parts = keep
            break
    if parts[0] in _ALIAS:
        parts[0] = _ALIAS[parts[0]]
    return "_".join(parts)


# ---------------------------------------------------------------- 变更模型

class Action(object):
    """一个工程动作。"""

    def __init__(self, kind, text, materials=None):
        self.kind = kind
        self.text = text
        self.materials = materials or []

    def __repr__(self):
        return "<Action %s: %s>" % (self.kind, self.text)


def infer_scene(name, code=""):
    """推断场景短语。

    优先用工程编号的类别（依据《工程建设项目分类、编码及命名指引》附表1），
    再用工程名称细化：
      编号 → 项目性质（工业/商业/公福/市政/小区）
      名称 → 具体场景（生产线/车间/锅炉房/食堂/厨房）

    名称中的具体场景优先于类别（锅炉类工程横跨工业与商业）。
    """
    n = str(name or "")
    for kw, label in (("生产线", "该生产线内"), ("车间", "该车间内"), ("锅炉", "该锅炉房内"),
                      ("食堂", "该食堂内"), ("厨房", "该厨房内"),
                      ("餐饮", "该厨房内"), ("食品", "该车间内")):
        if kw in n:
            return label
    if code:
        try:
            from project_code import profile
            p = profile(code)
            if p.get("scene") and p["scene"] != "该工程":
                return p["scene"]
        except ImportError:
            pass
    return ""


def infer_group(code):
    """取工程编号归一类别（工业/商业/公福/市政/小区…），失败返回空串。"""
    if not code:
        return ""
    try:
        from project_code import group_of
        return group_of(code)
    except ImportError:
        return ""


def derive_actions(add, red, name="", code=""):
    """从增/减材料清单推导工程动作列表。"""
    acts = []
    A = [dict(x) for x in add]
    R = [dict(x) for x in red]

    # --- 1. 管材段替换 → 敷设方式/材质调整 ---
    # 约束：替换对必须「量级相当 + 公称直径接近」。否则容易把
    # 调压柜移位等连带增减误判成管材替换（历史 NC25D004 即此类）。
    red_pipes = [x for x in R if _is_pipe(x["norm"]) and not any(k in x["norm"] for k in _PROTECT_KW)]
    add_pipes = [x for x in A if _is_pipe(x["norm"]) and not any(k in x["norm"] for k in _PROTECT_KW)]
    best = None
    for rp in red_pipes:
        for ap in add_pipes:
            # 同名管材不算「替换」（那是同规格的数量增减）
            if _base(rp["norm"]) == _base(ap["norm"]) and \
                    parse_spec(rp["name"]) == parse_spec(ap["name"]):
                continue
            rq, aq = _qty_num(rp["qty"]), _qty_num(ap["qty"])
            if rq <= 0 or aq <= 0 or not (0.6 <= aq / rq <= 1.7):
                continue
            rs, as_ = parse_spec(rp["name"]), parse_spec(ap["name"])
            if not rs or not as_:
                continue
            # 公称直径须接近。dn110 ↔ D76 在燃气工程中是常见配套
            # （PE管 dn110 对应无缝钢管 D76），故放宽下限。
            if rs[1] and as_[1]:
                ratio = as_[1] / rs[1]
                if not (0.65 <= ratio <= 1.35):
                    continue
            score = abs(aq - rq) / max(aq, rq)
            if best is None or score < best[0]:
                best = (score, rp, ap)
    if best:
        _, rp, ap = best
        acts.append(Action(
            "pipe_swap",
            "原设计{old}改用{new}敷设".format(old=_cite(rp["name"]), new=_cite(ap["name"])),
            [rp["name"], ap["name"]]))

    # --- 2. 三通 规格迁移 → 接口规格调整 ---
    # 只认「同系列、主规格不同、且减少件确实被替换掉」：
    #   三通_DN25/15 → 三通_DN65/15   ✓（支管接口改径）
    #   变径_dn200/110 → 变径_dn160/110 ✗（同系列不同组合，非改径）
    red_j = [x for x in R if "三通" in x["norm"]]
    add_j = [x for x in A if "三通" in x["norm"]]
    resize_parts = []
    for aj in add_j:
        aspec = parse_spec(aj["name"])
        if not aspec:
            continue
        for rj in red_j:
            if _base(aj["norm"]) != _base(rj["norm"]):
                continue
            rspec = parse_spec(rj["name"])
            if not rspec or rspec == aspec:
                continue
            # 主规格必须不同（三通_DN65/15 → DN65）
            if _main_spec(aj) == _main_spec(rj):
                continue
            # 数量须匹配（1 换 1），否则不是改径而是增配
            aq, rq = _qty_num(aj["qty"]), _qty_num(rj["qty"])
            if rq <= 0 or aq <= 0 or aq / rq > 2:
                continue
            resize_parts.append("{rj}改为{aj}".format(
                rj=_cite(rj["name"]), aj=_cite(aj["name"])))
            break

    # --- 3. 弯头 规格迁移 → 主管/管件规格调整 ---
    # 法兰不做「改径」判定：法兰的增减多来自封堵预留而非更换。
    for kw in ("弯头",):
        reds = [x for x in R if kw in x["norm"]]
        adds = [x for x in A if kw in x["norm"]]
        for ax in adds:
            aspec = parse_spec(ax["name"])
            if not aspec:
                continue
            for rx in reds:
                rspec = parse_spec(rx["name"])
                if not rspec or rspec == aspec or _base(ax["norm"]) != _base(rx["norm"]):
                    continue
                # 数量须匹配（1 换 1），否则属增配而非改径
                aq, rq = _qty_num(ax["qty"]), _qty_num(rx["qty"])
                if rq <= 0 or aq <= 0 or aq / rq > 2:
                    continue
                resize_parts.append("{rx}改为{ax}".format(
                    rx=_cite(rx["name"]), ax=_cite(ax["name"])))
                break

    if resize_parts:
        # 同一「原」件对应多个「改为」目标时归并，避免重复点名
        merged = []
        seen = {}
        for p in resize_parts:
            src, dst = p.split("改为")
            if src in seen:
                seen[src] = seen[src] + "、" + dst
            else:
                seen[src] = dst
        for src, dst in seen.items():
            merged.append("{src}改为{dst}".format(src=src, dst=dst))
        acts.append(Action(
            "resize",
            "{acts}，故造成相关管件有所增减".format(acts="、".join(merged)),
            resize_parts))

    # --- 4. 支管数量大增 → 增设支管 ---
    cited = set(resize_parts)
    branch_items = []
    for ax in A:
        q = _qty_num(ax["qty"])
        spec = parse_spec(ax["name"])
        if not spec or q < 3:
            continue
        if not any(k in ax["norm"] for k in ("弯头", "三通", "球阀", "法兰")):
            continue
        if any(_cite(ax["name"]) in a for a in cited):
            continue
        branch_items.append(ax)
    # 同类动作合并：多件同属「增设支管」时合成一句，避免正文重复
    if branch_items:
        detail = "、".join("{n}-{q}".format(n=_cite(b["name"]), q=b["qty"]) for b in branch_items)
        total = sum(int(_qty_num(b["qty"])) for b in branch_items)
        acts.append(Action(
            "branch_add",
            "需按现场实际布置增设{n}处支管，故增加{detail}".format(n=total, detail=detail),
            [b["name"] for b in branch_items]))

    # --- 5. 成套预留封堵 ---
    flange = next((x for x in A if "法兰" in x["norm"]), None)
    blind = next((x for x in A if "盲板" in x["norm"]), None)
    gasket = next((x for x in A if ("四氟垫" in x["norm"] or "金属缠绕垫" in x["norm"])), None)
    if flange and blind:
        names = [flange["name"], blind["name"]]
        if gasket:
            names.append(gasket["name"])
        acts.append(Action(
            "cap_reserve",
            "预留管口采用焊接法兰加盲板封堵，故增加{names}及相关配件".format(
                names="、".join(_cite(x) for x in names)),
            names))

    # --- 5b. 支架替代（物资供应部不再提供固定/落地支架，改用导向支架）---
    guide = next((x for x in A if "导向支架" in x["norm"]), None)
    fixed = next((x for x in R if ("固定支架" in x["norm"] or "落地支架" in x["norm"]
                                  or "吊架" in x["norm"])), None)
    if guide and fixed:
        acts.append(Action(
            "support_swap",
            "因物资供应部不再提供{old}，需改用{new}代替，故增加{gn}-{gq}、减少{fn}-{fq}".format(
                old=_cite(fixed["name"]), new=_cite(guide["name"]),
                gn=_cite(guide["name"]), gq=guide["qty"],
                fn=_cite(fixed["name"]), fq=fixed["qty"]),
            [guide["name"], fixed["name"]]))

    # --- 5c. 碰接方式 / 接入口变更 ---
    # 特征：新增鞍型电熔变径三通等碰接件（改为不停输/三通碰接），
    #       或 PE管 与 钢塑转换/镀锌管/法兰 等互换（换接入口）。
    _SADDLE_KW = ("鞍型电熔变径", "鞍型电熔管件", "电熔鞍型", "鞍型旁通", "鞍型电熔")
    _ELEC_KW = ("电熔三通", "电熔弯头", "热熔三通", "热熔弯头", "热熔变径", "电熔变径")
    tap_conn = next((x for x in A if any(k in x["norm"] for k in _SADDLE_KW)), None)
    tap_removed = next((x for x in R if any(k in x["norm"] for k in _SADDLE_KW)), None)
    if tap_conn and tap_removed:
        acts.append(Action(
            "tap_change",
            "原{t0}改为{t1}碰接，故{t1n}增加{t1q}、{t0n}减少{t0q}".format(
                t0=_cite(tap_removed["name"]), t1=_cite(tap_conn["name"]),
                t1n=_cite(tap_conn["name"]), t1q=tap_conn["qty"],
                t0n=_cite(tap_removed["name"]), t0q=tap_removed["qty"]),
            [tap_conn["name"], tap_removed["name"]]))
    elif tap_removed:
        # 鞍型电熔管件（不停输碰接件）取消，改用常规熔接件（三通/弯头）碰接。
        # 这是历史语料中最高频的碰接变更形态（27 例）：
        #   减 鞍型电熔管件_BFPE110  →  增 电熔弯头_dn63_90° / 电熔三通_dn110×90
        # 常见缘由：不停输设备无法进场、作业面不够、埋深不足影响安全、
        #           探明实际管位与图纸不符、要求停气碰接。
        elec_added = [x for x in A if any(k in x["norm"] for k in _ELEC_KW)]
        if elec_added:
            add_txt = "、".join("%s-%s" % (_cite(x["name"]), x["qty"]) for x in elec_added[:3])
            acts.append(Action(
                "tap_change",
                "取消{t0n}-{t0q}不停输碰接，改为{add}常规熔接碰接".format(
                    t0n=_cite(tap_removed["name"]), t0q=tap_removed["qty"], add=add_txt),
                [tap_removed["name"]] + [x["name"] for x in elec_added[:3]]))
        else:
            # 无替代熔接件可点名时，用概括表述（见 _BRIEF）
            acts.append(Action(
                "tap_change",
                "取消{t0n}-{t0q}不停输碰接，改为停气常规熔接碰接".format(
                    t0n=_cite(tap_removed["name"]), t0q=tap_removed["qty"]),
                [tap_removed["name"]]))
    else:
        # PE 管与「钢塑转换 / 镀锌管 / 焊接法兰」互换 → 换接入口
        # 注意：必须用原始 name 过滤——norm 已剥离括号内容，
        # 「热镀锌钢管_DN25（自制防撞护栏）」在 norm 里看不出是防撞件。
        def _is_tap_material(x):
            raw = str(x["name"])
            if any(k in raw for k in ("防撞", "护栏", "防撞墩", "支架", "标示", "警示", "标志")):
                return False
            return any(k in x["norm"] for k in ("钢塑转换", "镀锌钢管", "焊接法兰", "钢制无缝管"))

        pe_reduced = [x for x in R if "pe管" in x["norm"] and _qty_num(x["qty"]) >= 1]
        steel_added = [x for x in A if _is_tap_material(x)]
        if pe_reduced and steel_added:
            acts.append(Action(
                "tap_change",
                "原气源接入口调整，改由{new}接出，故{n}增加{nq}、{old}减少{oq}".format(
                    new=_cite(steel_added[0]["name"]), n=_cite(steel_added[0]["name"]),
                    nq=steel_added[0]["qty"], old=_cite(pe_reduced[0]["name"]),
                    oq=pe_reduced[0]["qty"]),
                [steel_added[0]["name"], pe_reduced[0]["name"]]))

    # --- 5d. 管材路径改道（管长大幅重排，几十米级增减）---
    def _pipe_len(items):
        return sum(_qty_num(x["qty"]) for x in items if _is_pipe(x["norm"]))

    red_len, add_len = _pipe_len(R), _pipe_len(A)
    drill_add = next((x for x in A if ("定向钻" in x["norm"] or "包覆" in x["norm"])), None)
    drill_red = next((x for x in R if ("定向钻" in x["norm"] or "包覆" in x["norm"])), None)
    if drill_add and drill_red:
        acts.append(Action(
            "route_change",
            "原定向钻{d0}改为{d1}敷设，故{d1n}增加{d1q}、{d0n}减少{d0q}".format(
                d0=_cite(drill_red["name"]), d1=_cite(drill_add["name"]),
                d1n=_cite(drill_add["name"]), d1q=drill_add["qty"],
                d0n=_cite(drill_red["name"]), d0q=drill_red["qty"]),
            [drill_add["name"], drill_red["name"]]))
    elif (red_len >= 20 and add_len >= 5 and abs(red_len - add_len) >= 15):
        # 埋地管长大幅重排（老段取消、新段另敷）→ 敷设路径改道
        def _real_pipe(x):
            raw = str(x["name"])
            if any(k in raw for k in ("防撞", "护栏", "防撞墩", "支架", "标示", "警示")):
                return False
            return _is_pipe(x["norm"])

        big_red = max((x for x in R if _real_pipe(x)),
                      key=lambda x: _qty_num(x["qty"]), default=None)
        big_add = max((x for x in A if _real_pipe(x)),
                      key=lambda x: _qty_num(x["qty"]), default=None)
        if big_red and big_add:
            acts.append(Action(
                "route_change",
                "原敷设路径无法实施，经报告公司相关代表到现场核实并商议后决定改道，"
                "故{an}增加{aq}、{rn}减少{rq}".format(
                    an=_cite(big_add["name"]), aq=big_add["qty"],
                    rn=_cite(big_red["name"]), rq=big_red["qty"]),
                [big_add["name"], big_red["name"]]))

    # --- 6. 仪表连接件 ---
    tap = next((x for x in A if "单头牙" in x["norm"]), None)
    if tap:
        acts.append(Action(
            "instrument",
            "压力表及直嘴铜球阀DN15需要钢制单头牙_DN15连接，故增加{iname}-{iq}".format(
                iname=_cite(tap["name"]), iq=tap["qty"]),
            [tap["name"]]))
    for x in A:
        if "过滤网" in x["norm"]:
            acts.append(Action(
                "instrument",
                "该工程所用的流量计为老表，故安装时增加{iname}-{iq}".format(
                    iname=_cite(x["name"]), iq=x["qty"]),
                [x["name"]]))
            break

    # --- 7. 保护沟 ---
    for x in A:
        if "保护沟" in x["norm"]:
            acts.append(Action(
                "protect",
                "部分管段埋深不够，根据公司相关要求需敷设{pn}-{pq}".format(
                    pn=_cite(x["name"]), pq=x["qty"]),
                [x["name"]]))
            break

    # --- 8. 防撞护栏 ---
    for x in A:
        if any(k in x["norm"] for k in ("防撞", "护栏", "防撞墩")):
            acts.append(Action(
                "anti_impact",
                "该处车辆进出频繁，为保证燃气管道安全，自制防撞护栏，故增加{an}-{aq}".format(
                    an=_cite(x["name"]), aq=x["qty"]),
                [x["name"]]))
            break

    return acts


# ---------------------------------------------------------------- 原因组装

#: 各动作类型的原因前缀（只提供「因由」，动作文本自带结论）
#: {sc} 会被替换为场景短语（不带「该」前缀时自动补「该」）
_LEAD = {
    "pipe_swap": "因{sc}设备布置及用气接口与原设计不符，经报告公司相关代表到现场核实并商议后决定，",
    "resize": "因{sc}用气量及接口规格需相应调整，",
    "branch_add": "因{sc}设备及设施较多，",
    "cap_reserve": "为便于后续施工，",
    "instrument": "因该处需安装计量及控制仪表，",
    "protect": "因该段管线的覆土深度不足，",
    "anti_impact": "因该处车辆进出频繁，为保证燃气管道安全，",
    "support_swap": "因物资供应部不再提供该类配件，",
    "tap_change": "因{sc}气源接入口位置及方式需调整，经报告公司相关代表到现场核实并商议后决定，",
    "route_change": "因{sc}施工条件及敷设路径需调整，经报告公司相关代表到现场核实并商议后决定，",
}

def fallback_clause(scene="", location=""):
    """无任何工程动作可推时的兜底句（按室内 / 室外分措辞）。

    室内场景（厨房、食堂、车间、锅炉房、厂区、商业场所）不能用
    「地下管线较多、按现场开挖调整敷设位置」那套市政口径——那会把
    室内厨房、厂内设备、调压器迁改等工程错误归因。
    """
    loc = str(location or "").strip()
    sc = str(scene or "").strip()
    if not sc and loc and any(k in loc for k in
                              ("内", "厨房", "食堂", "车间", "锅炉房", "厂区", "商业区")):
        sc = loc if loc.startswith("该") else "该" + loc
    if sc:
        if not sc.startswith("该"):
            sc = "该" + sc
        return ("因{sc}设备及设施较多，需根据现场实际情况调整管位，"
                "并增加管件绕行避让障碍，故造成相关管材管件有所增减；".format(sc=sc))
    if loc:
        return ("该工程位于{loc}，地下管线及其他设施较多，需根据现场开挖实际情况"
                "调整管线敷设位置及长度，并增加管件绕行避让障碍，"
                "故造成相关管材管件有所增减；".format(loc=loc))
    return ("因该工程地下管线及其他设施较多，需根据现场开挖实际情况调整管线敷设位置"
            "及长度，并增加管件绕行避让障碍，故造成相关管材管件有所增减；")


def compose(add, red, name="", code="", location="", extra="", actions=None):
    """基于工程动作组装变更原因正文。

    长度自适应（对齐历史规律）：
      材料变动越多 → 正文越概括。历史上变动 40 项的案例正文仅点名 0–1 种材料；
      变动少时才逐项点名。完整明细由变更单的增/减材料栏承载。

    返回 (正文, 动作列表)。正文以「；」分隔各分句，末尾保留「；」。
    """
    acts = actions if actions is not None else derive_actions(add, red, name, code)
    scene = infer_scene(name, code)
    segs = []

    if extra:
        segs.append(extra.rstrip("。；; ") + "，")

    # 变动规模 → 点名密度（对齐历史：材料越多，正文越概括）
    n_mat = len(add) + len(red)
    if n_mat <= 8:
        cited_kinds = None            # 全部点名
    elif n_mat <= 16:
        cited_kinds = {"pipe_swap", "resize", "branch_add", "cap_reserve", "instrument"}
    else:
        cited_kinds = {"pipe_swap", "resize"}

    # 逐句组装；「故造成…增减」的结论只在末句出现一次（历史写法）
    body_segs = []
    for a in acts:
        if cited_kinds is not None and a.kind not in cited_kinds:
            brief = _BRIEF.get(a.kind)
            if not brief:
                continue
            body_segs.append(_BRIEF_LEAD.get(a.kind, "因现场实际情况，").format(
                sc=(scene or "该工程")) + brief)
            continue
        lead = _LEAD.get(a.kind, "因现场实际情况，")
        if "{sc}" in lead:
            lead = lead.format(sc=(scene or "该工程"))
        body_segs.append(lead + a.text)

    if body_segs:
        # 末句补结论
        last = body_segs[-1]
        if "故造成" not in last and "故增加" not in last and "故取消" not in last:
            body_segs[-1] = last + "，故造成相关管材管件有所增减"
        body = "；".join(body_segs) + "；"

        # 超长时逐步降级：把点名型分句改概括，优先压缩最次要的动作
        for order in (_DOWNGRADE_ORDER):
            if len(body) + len(TAIL_APP) <= MAX_LEN:
                break
            for kind in order:
                for i, a in enumerate(acts):
                    if a.kind != kind or i >= len(body_segs):
                        continue
                    brief = _BRIEF.get(kind)
                    if not brief:
                        continue
                    body_segs[i] = _BRIEF_LEAD.get(kind, "因现场实际情况，").format(
                        sc=(scene or "该工程")) + brief
                    body = "；".join(body_segs) + "；"
                    if len(body) + len(TAIL_APP) <= MAX_LEN:
                        break
                if len(body) + len(TAIL_APP) <= MAX_LEN:
                    break
    else:
        # 兜底：按室内/室外分措辞（与 core._site_clause 共用同一实现）
        body = fallback_clause(scene=scene, location=location)

    return body, acts


#: 不点名材料时的概括表述（整句使用，不与前缀重复）
_BRIEF = {
    "resize": "故相关接口及管件规格已作相应调整",
    "branch_add": "需按现场实际布置增设多处支管",
    "cap_reserve": "预留管口已用焊接法兰加盲板封堵",
    "instrument": "计量及控制仪表已配套安装相应连接件",
    "protect": "部分管段已按要求敷设保护沟",
    "anti_impact": "已自制防撞护栏",
    "support_swap": "原支架已改用导向支架代替",
    "tap_change": "碰接方式已改用常规熔接碰接",
    "route_change": "敷设路径已作相应调整",
}

#: 概括表述的前缀
_BRIEF_LEAD = {
    "resize": "因{sc}用气量及接口规格需相应调整，",
    "branch_add": "因{sc}设备及设施较多，",
    "cap_reserve": "为便于后续施工，",
    "instrument": "因该处需安装计量及控制仪表，",
    "protect": "因该段管线的覆土深度不足，",
    "anti_impact": "因该处车辆进出频繁，为保证燃气管道安全，",
}


# ---------------------------------------------------------------- 同项目定稿复用

_TAIL_PAT = re.compile(r"兹提出[^。；]*。")

#: 定稿复用时允许补充的分句类型
#: branch_add 的「N 处支管」由弯头数量推得、易失真，故不补
_REUSE_KINDS = ("pipe_swap", "resize", "instrument", "cap_reserve", "protect",
                "anti_impact", "support_swap", "tap_change", "route_change")

_CONCL_KW = ("故造成", "故增加", "故减少", "故取消", "造成相关", "造成部分材料",
             "造成对应的", "材料有所", "材料变更如下")


def reuse_project_reason(project_text, add, red, name="", code="", extra=""):
    """以同项目历史定稿为骨架，补入本次材料推出的工程动作。

    材料只能证明「改了什么」，证明不了「为什么」。同项目定稿里已有
    「设计管位顶部有蒸汽管道」「厨房设施与设计管位安全间距不足」这类
    现场事实，比纯材料推理可靠，因此优先沿用其口径，仅补充本次实际增减
    涉及、但定稿未提及的工程动作。
    """
    try:
        from gas_change_core import _norm as _cn
    except Exception:
        def _cn(s):
            return str(s).lower()

    core = _TAIL_PAT.sub("", str(project_text or "")).strip()
    # 定稿常见收束语「综上所述」「综上原因」「综上」不带句号，需单独剥掉
    for w in ("综上所述", "综上原因", "综上"):
        core = core.replace(w, "")
    core = core.rstrip("；;。 ").strip("；;，, ")
    segs = [s.strip() for s in re.split(r"[；;]", core) if s.strip()]
    if extra:
        segs.insert(0, extra.rstrip("。；; "))

    acts = derive_actions(add, red, name, code)
    clauses = []
    for a in acts:
        if a.kind not in _REUSE_KINDS:
            continue
        lead = _LEAD.get(a.kind, "")
        if not lead or "{sc}" in lead:
            continue
        if any((_cite(m) in core) or (_base(_cn(m)) in core)
               for m in (a.materials or [])):
            continue
        clause = lead + a.text
        if not any(k in clause for k in ("故造成", "故增加", "故减少", "故取消")):
            clause += "，故造成相关材料有所增减"
        clauses.append(clause)
        if len(clauses) >= 2:
            break

    has_concl = any(k in core for k in _CONCL_KW)
    if not clauses and not has_concl:
        clauses = ["故造成相关管材管件有所增减"]

    body = ""
    while True:
        parts = segs + clauses
        body = "；".join(p for p in parts if p) + "；"
        if len(body) + len(TAIL_APP) <= MAX_LEN or not clauses:
            break
        clauses.pop()

    return body
