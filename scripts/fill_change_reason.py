# -*- coding: utf-8 -*-
"""
燃气工程变更单 —— 变更原因生成 CLI

用法:
  # 1) 只分析，不写文件（推荐先看结果）
  python fill_change_reason.py analyze "<变更单.xls>"

  # 2) 分析并写入
  python fill_change_reason.py fill "<变更单.xls>" [--location "南昌县老城区"] [--extra "现场补充事实"]

  # 3) 单元格已有用户草稿 → 输出润色素材包，AI 润色后回写
  python fill_change_reason.py polish "<变更单.xls>"
  python fill_change_reason.py fill "<变更单.xls>" --body "<AI润色后的正文>"

  # 4) 重建历史语料库缓存
  python fill_change_reason.py reindex

  # 5) 检索历史相似案例
  python fill_change_reason.py search "<变更单.xls>" [--top 8]
"""

import os
import sys
import json
import glob
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gas_change_core import (  # noqa: E402
    ChangeOrder, Corpus, ReasonBuilder, fill_reason,
    ARCHIVE_ROOT, ARCHIVE_ROOTS, PLACEHOLDER, TAIL_APP, TAIL_QTY,
)

CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "corpus_cache.json")


def cmd_polish(args):
    """输出「润色素材包」：用户草稿 + 材料推导 + 历史案例 + 文风规范。

    供 AI 参照历史语料把草稿润色成规范文案，再以 fill --body 写回。
    """
    co = ChangeOrder(args.file)
    if co.print_sheet is None:
        print("[错误] 未找到 '*变更单打印' 工作表。")
        return 2
    code, name = co.header()
    add, red = co.materials()
    draft = co.read_draft()

    print("=" * 70)
    print("润色素材包 —— %s  %s" % (code, name))
    print("=" * 70)

    # 1) 用户草稿（提示词）
    print("\n【1】用户草稿（作为提示词，润色时须保留其中的现场事实）")
    if draft["has_draft"]:
        if draft["app"]:
            print("  申请单 %s：" % draft["app_addr"])
            print("    %s" % draft["app"])
        if draft["qty"] and draft["qty"] != draft["app"]:
            print("  工程量单 %s：" % draft["qty_addr"])
            print("    %s" % draft["qty"])
    else:
        print("  （无草稿 —— 单元格为空或为占位符，可直接用 fill 自动生成）")

    # 2) 工程编号分类
    print("\n【2】工程分类")
    try:
        from project_code import describe, parse_code
        p = parse_code(code)
        if p:
            print("  %s（%s） %s" % (p["cat_name"], p["group"], describe(code)))
    except ImportError:
        pass

    # 3) 材料增删
    print("\n【3】材料变动（润色时须与之一致）")
    print("  增加 %d 项：" % len(add))
    for x in add:
        print("    + %-42s %s" % (x["name"], x["qty"]))
    print("  减少 %d 项：" % len(red))
    for x in red:
        print("    - %-42s %s" % (x["name"], x["qty"]))

    # 4) 推导出的工程动作
    print("\n【4】材料 → 工程动作（自动推理，供交叉验证草稿）")
    try:
        from material_reason import derive_actions
        acts = derive_actions(add, red, name, code)
        if acts:
            for a in acts:
                print("    [%s] %s" % (a.kind, a.text))
        else:
            print("    （未命中既有动作类型 —— 草稿信息为主要依据）")
    except ImportError:
        pass

    # 5) 相似历史案例（润色的措辞参照）
    corpus = Corpus(ARCHIVE_ROOTS, CACHE)
    add_n = sorted(set(x["norm"] for x in add))
    red_n = sorted(set(x["norm"] for x in red))
    draft_text = draft["app"] or draft["qty"] or ""
    print("\n【5】相关历史案例（润色的措辞参照）")
    hits = corpus.related(add_n, red_n, code=code, draft=draft_text,
                          top=args.top, add=add, red=red)
    if not hits:
        print("  （未找到相关案例，请以草稿事实 + 文风规范为准）")
    for score, mat, why, r in hits:
        print("  [%.2f] %s" % (score, r["file"]))
        print("      关联：%s" % (why or "同类"))
        print("      原文：%s" % r["reason"])

    # 6) 文风规范
    print("\n【6】文风规范（历史 132 份语料归纳）")
    for line in _STYLE_RULES:
        print("  " + line)

    print("\n" + "=" * 70)
    print("下一步：AI 参照以上素材润色出正文，然后写回——")
    print('  python fill_change_reason.py fill "%s" --body "<润色后的正文>"' % args.file)
    print("=" * 70)
    return 0


_STYLE_RULES = [
    "① 结构：分句以「；」分隔，末句以「故造成…有所增减/减少/增加」收尾，结尾语前保留「；」",
    "② 结尾语（314 份语料）：申请单 69% 用「兹提出设计变更申请单，请予以审批。」；",
    "   23% 用「兹提出变更申请单，请予以审批。」——若同一单位/同批单据用的是后者，就沿用后者；",
    "   工程量变更单用「兹提出工程量变更单。」",
    "③ 每句以「因 / 该工程 / 原定 / 我部」开头，陈述事实后再给结论",
    "④ 少量点状专项材料要点名并带数量（如「故增加钢制单头牙_DN15-4个」）；",
    "   管长类总量变化用概括表述（「故造成相关管材管件有所增减」），不要罗列全部材料",
    "⑤ 决策过程写「经报告公司相关代表到现场核实并商议后决定」",
    "⑥ 长度 38–234 字（历史区间），多数 70–150 字",
    "⑦ 保留草稿中的现场事实（甲方要求 / 阻工 / 库房缺货 / 设备变动等），只改措辞不改事实",
    "",
    "新语料补充的常见句式：",
    "  · 因图纸/设计与现场不符：「因设计图碰接点位置与现场不符…」",
    "  · 物资供应问题：「由于物资供应部不再提供支架，需施工单位自行购买…」",
    "  · 配合施工方要求：「该工程为道路配合工程，我部需配合…」",
    "  · 结论收尾变体：「综上所述，兹提出变更申请单，请予以审批。」",
    "  · 变更清单式：「…造成相关材料增减如下；拟调整如下材料，兹提出变更申请单。」",
]


def cmd_analyze(args):
    co = ChangeOrder(args.file)
    if co.print_sheet is None:
        print("[错误] 未找到 '*变更单打印' 工作表，文件可能不是变更单。")
        print("       现有工作表: %s" % ", ".join(co.sheet_names))
        return 2
    code, name = co.header()
    add, red = co.materials()
    print("=" * 70)
    print("工程编号: %s    工程名称: %s" % (code, name))
    # 工程编号分类解读（依据《工程建设项目分类、编码及命名指引》附表1）
    try:
        from project_code import describe, parse_code
        desc = describe(code)
        if desc:
            p = parse_code(code)
            group = p["group"]
            label = p["cat_name"] if group == p["cat_name"] else "%s（%s）" % (p["cat_name"], group)
            print("项目分类: %s" % label)
            print("编码含义: %s" % desc)
    except ImportError:
        pass
    print("工作表  : %s" % co.print_sheet.name)
    print("=" * 70)
    print("\n增加材料 (%d):" % len(add))
    for x in add:
        print("   + %-45s %s" % (x["name"], x["qty"]))
    print("\n减少材料 (%d):" % len(red))
    for x in red:
        print("   - %-45s %s" % (x["name"], x["qty"]))

    corpus = Corpus(ARCHIVE_ROOTS, CACHE)
    rb = ReasonBuilder(corpus)
    body, hits, refs = rb.build(add, red, name=name, code=code,
                                location=args.location or "", extra=args.extra or "")

    # 展示推导出的工程动作（材料→动作的推理过程）
    try:
        from material_reason import derive_actions
        acts = derive_actions(add, red, name, code)
        if acts:
            print("\n" + "-" * 70)
            print("材料 → 工程动作 推理:")
            for a in acts:
                print("   [%s] %s" % (a.kind, a.text))
    except ImportError:
        pass

    print("\n" + "-" * 70)
    print("命中触发因子: %s" % (", ".join(hits) if hits else "无"))
    print("-" * 70)
    print("\n【生成的变更原因正文】\n")
    print(body.rstrip("；") + "；" + TAIL_APP)
    print("\n【工程量变更单版本结尾】\n")
    print(body.rstrip("；") + "；" + TAIL_QTY)

    print("\n" + "-" * 70)
    print("参考历史案例 (材料相似度):")
    for s, r in refs:
        print("  [%.3f] %s" % (s, r["file"]))
        print("        %s" % (r["reason"][:110] + ("..." if len(r["reason"]) > 110 else "")))
    return 0


def cmd_fill(args):
    corpus = Corpus(ARCHIVE_ROOTS, CACHE)
    rb = ReasonBuilder(corpus)
    co = ChangeOrder(args.file)
    if co.print_sheet is None:
        print("[错误] 未找到 '*变更单打印' 工作表。")
        return 2
    code, name = co.header()
    add, red = co.materials()
    body, hits, refs = rb.build(add, red, name=name, code=code,
                                location=args.location or "", extra=args.extra or "")
    if args.body:
        body = args.body.rstrip("。；;") + "。"

    ok, cells, slots = None, None, None
    try:
        ok, cells, slots = fill_reason(args.file, body)
    except RuntimeError as exc:
        print("[已中止] %s" % exc)
        return 1
    print("写入完成: %s" % ("成功" if ok else "失败"))
    print("  工作表: %s" % slots["app"][0])
    for addr, val in cells.items():
        print("  %-6s <- %s" % (addr, val[:80] + ("..." if len(val) > 80 else "")))
    print("  提示: 不生成备份副本，原变更原因已被覆盖")
    return 0 if ok else 1


def cmd_reindex(args):
    cache = args.out or CACHE
    corpus = Corpus(ARCHIVE_ROOTS, None)   # 强制重扫
    with open(cache, "w", encoding="utf-8") as fh:
        json.dump(corpus.records, fh, ensure_ascii=False, indent=1)
    st = corpus.stats()
    print("重建语料库: %s" % cache)
    print("  文件总数 %d | 已填原因 %d | 待填 %d" % (st["total"], st["filled"], st["pending"]))
    return 0


def cmd_doctor(args):
    """环境自检：依赖、归档路径、Excel 可用性。"""
    ok = True
    print("=" * 62)
    print("燃气工程变更原因生成 —— 环境自检")
    print("=" * 62)

    print("\n[1] Python 版本")
    v = sys.version_info
    print("    %d.%d.%d  %s" % (v.major, v.minor, v.micro, sys.executable))
    if v < (3, 8):
        print("    ✗ 需要 Python 3.8 及以上")
        ok = False
    else:
        print("    ✓ 满足要求")

    print("\n[2] 依赖库")
    try:
        import xlrd
        print("    ✓ xlrd %s" % xlrd.__VERSION__)
    except ImportError:
        print("    ✗ 缺少 xlrd   → pip install xlrd")
        ok = False
    try:
        import win32com.client  # noqa: F401
        print("    ✓ pywin32（写入 .xls 需要）")
    except ImportError:
        print("    ✗ 缺少 pywin32 → pip install pywin32")
        ok = False
    except Exception as exc:
        print("    ⚠ pywin32 已安装但导入异常：%s" % str(exc)[:60])
        print("      写入功能可能受限，请确认已安装 Microsoft Excel")

    print("\n[3] Excel（写入 .xls 需要本机安装 Excel）")
    try:
        import win32com.client
        app = win32com.client.DispatchEx("Excel.Application")
        print("    ✓ Excel %s" % app.Version)
        app.Quit()
    except Exception as exc:
        print("    ✗ 不可用：%s" % str(exc)[:70])
        print("      仅查看/生成原因不受影响；写入需安装 Microsoft Excel")
        ok = False

    print("\n[4] 历史变更单归档目录")
    total_xls = 0
    for root in ARCHIVE_ROOTS:
        print("    · %s" % root)
        if not os.path.isdir(root):
            print("      ✗ 目录不存在")
            ok = False
            continue
        n = len(glob.glob(os.path.join(root, "**", "*.xls"), recursive=True))
        total_xls += n
        if n:
            print("      ✓ 找到 %d 份 .xls" % n)
        else:
            print("      ⚠ 目录存在但没有 .xls 文件")
    if not ARCHIVE_ROOTS:
        print("    ✗ 未配置任何归档目录")
        ok = False
    elif total_xls == 0 and ok:
        print("    ⚠ 所有归档目录中都没有 .xls —— 检索相似案例将不可用")
        print("      请设置环境变量 GAS_CHANGE_ARCHIVE（多个路径用 ; 分隔），")
        print('      或在 scripts/config.json 写入 {"archive_roots": ["路径1", "路径2"]}')

    print("\n[5] 语料库缓存")
    if os.path.exists(CACHE):
        try:
            import json
            d = json.load(open(CACHE, encoding="utf-8"))
            print("    ✓ %s（%d 条）" % (os.path.basename(CACHE), len(d)))
            print("      若归档目录有变动，请运行：python fill_change_reason.py reindex")
        except Exception as exc:
            print("    ⚠ 缓存无法读取：%s" % str(exc)[:50])
    else:
        print("    ⚠ 无缓存，请运行：python fill_change_reason.py reindex")

    print("\n" + "=" * 62)
    print("结论: %s" % ("环境就绪 ✓" if ok else "存在问题，请按上方 ✗ 项处理"))
    print("=" * 62)
    return 0 if ok else 1


def cmd_search(args):
    corpus = Corpus(ARCHIVE_ROOTS, CACHE)
    co = ChangeOrder(args.file)
    add, red = co.materials()
    add_n = sorted(set(x["norm"] for x in add))
    red_n = sorted(set(x["norm"] for x in red))
    for s, r in corpus.similar(add_n, red_n, top=args.top):
        print("[%.3f] %s" % (s, r["file"]))
        print("   原因: %s" % r["reason"])
        print()
    return 0


def main():
    ap = argparse.ArgumentParser(description="燃气工程变更单变更原因生成")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("analyze", help="分析变更单并预览原因")
    p.add_argument("file")
    p.add_argument("--location", default="", help="现场位置描述，如「南昌县老城区」")
    p.add_argument("--extra", default="", help="补充的现场事实描述")
    p.set_defaults(func=cmd_analyze)

    p = sub.add_parser("fill", help="生成并写入变更原因")
    p.add_argument("file")
    p.add_argument("--location", default="")
    p.add_argument("--extra", default="")
    p.add_argument("--body", default="", help="直接指定正文，覆盖自动生成")
    p.set_defaults(func=cmd_fill)

    p = sub.add_parser("reindex", help="重建历史语料库")
    p.add_argument("--out", default="")
    p.set_defaults(func=cmd_reindex)

    p = sub.add_parser("polish", help="输出润色素材包（草稿+材料+案例+文风），供 AI 润色")
    p.add_argument("file")
    p.add_argument("--top", type=int, default=5)
    p.set_defaults(func=cmd_polish)

    p = sub.add_parser("doctor", help="环境自检（依赖/归档目录/Excel）")
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("search", help="检索相似历史案例")
    p.add_argument("file")
    p.add_argument("--top", type=int, default=8)
    p.set_defaults(func=cmd_search)

    args = ap.parse_args()
    sys.exit(args.func(args))


if __name__ == "__main__":
    main()
