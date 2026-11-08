# -*- coding: utf-8 -*-
"""
燃气工程变更原因生成 —— 一键安装

用法（AI 或人工均可执行）：
    python install.py                     # 自动安装到用户级 skills 目录
    python install.py --target <目录>      # 安装到指定目录
    python install.py --archive <目录>     # 指定历史变更单归档目录
    python install.py --no-deps           # 跳过依赖安装
    python install.py --check-only        # 只做环境体检，不安装

脚本会依次完成：
    1. 检查 Python 版本
    2. 安装缺失依赖（xlrd / pywin32）
    3. 复制 skill 到 skills 目录
    4. 配置历史变更单归档目录
    5. 运行 doctor 自检并给出结论

全程只在缺失时才动手，不会覆盖已有可用依赖。
"""

import os
import re
import sys
import glob
import json
import shutil
import argparse
import subprocess

SKILL_NAME = "gas-change-reason"

#: (pip 包名, 用于检测的导入模块名)
#: pywin32 没有顶层 pywin32 模块，只有 win32com 子包，检测时须用后者
NEED = [("xlrd", "xlrd"), ("pywin32", "win32com.client")]


def log(tag, msg):
    print("[%s] %s" % (tag, msg))


def ok(msg):
    log("OK", msg)


def warn(msg):
    log("WARN", msg)


def err(msg):
    log("ERROR", msg)


def info(msg):
    log("INFO", msg)


# ---------------------------------------------------------------- 1. Python

def check_python():
    v = sys.version_info
    log("INFO", "Python %d.%d.%d  (%s)" % (v.major, v.minor, v.micro, sys.executable))
    if v < (3, 8):
        err("需要 Python 3.8 及以上")
        return False
    ok("Python 版本满足要求")
    return True


# ---------------------------------------------------------------- 2. 依赖

def has_module(modname):
    """检测模块是否可导入（用于依赖检测）。"""
    try:
        __import__(modname)
        return True
    except ImportError:
        return False
    except Exception:
        # 模块存在但导入时出错（如 COM 未注册），也算已安装
        return True


def install_deps(force=False):
    missing = [pkg for pkg, mod in NEED if not has_module(mod)]
    if not missing:
        ok("依赖已齐备：%s" % "、".join(pkg for pkg, _ in NEED))
        return True
    info("缺少依赖：%s" % "、".join(missing))
    cmd = [sys.executable, "-m", "pip", "install"] + missing
    info("执行：%s" % " ".join(cmd))
    try:
        r = subprocess.run(cmd, capture_output=True, text=True)
    except Exception as exc:
        err("pip 调用失败：%s" % exc)
        return False
    if r.returncode != 0:
        err("pip 安装失败，请手动执行：%s" % " ".join(cmd))
        err((r.stderr or r.stdout or "")[-600:])
        return False
    still = [pkg for pkg, mod in missing if not has_module(mod)]
    if still:
        err("安装后仍无法导入：%s" % "、".join(still))
        return False
    ok("依赖安装完成：%s" % "、".join(missing))
    return True


# ---------------------------------------------------------------- 3. 安装位置

def default_skill_dir():
    home = os.path.expanduser("~")
    return os.path.join(home, ".workbuddy", "skills", SKILL_NAME)


def copy_skill(src_root, target):
    """把 skill 目录复制到 target。已存在则备份后覆盖。"""
    target = os.path.abspath(target)
    parent = os.path.dirname(target)
    if not os.path.isdir(parent):
        os.makedirs(parent)
    if os.path.isdir(target):
        bak = target + ".bak"
        if os.path.isdir(bak):
            shutil.rmtree(bak)
        shutil.move(target, bak)
        info("已备份原目录 → %s" % bak)
    shutil.copytree(src_root, target,
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo", ".git"))
    ok("已安装到 %s" % target)
    return target


# ---------------------------------------------------------------- 4. 归档目录

def looks_like_archive(path):
    """判断目录内是否有变更单（含 *变更单打印 工作表的 .xls）。"""
    try:
        import xlrd
    except ImportError:
        return len(glob.glob(os.path.join(path, "**", "*.xls"), recursive=True)) > 0
    files = glob.glob(os.path.join(path, "**", "*.xls"), recursive=True)
    for f in files[:40]:
        try:
            wb = xlrd.open_workbook(f)
            if any("变更单打印" in s.name for s in wb.sheets()):
                return True
        except Exception:
            continue
    return False


def probe_archive(hint=None):
    """探测历史变更单归档目录（可多个）。

    顺序：显式参数 > 环境变量 > 常见位置扫描
    返回目录列表。
    """
    if hint:
        return [os.path.abspath(p) for p in re.split(r"[;|]", hint.replace("\n", ";")) if p.strip()]
    env = os.environ.get("GAS_CHANGE_ARCHIVE")
    if env and env.strip():
        return [os.path.abspath(p) for p in re.split(r"[;|]", env.replace("\n", ";")) if p.strip()]

    # 常见位置扫描
    roots = []
    for base in (os.path.expanduser("~"), "E:\\", "D:\\", "C:\\"):
        if os.path.isdir(base):
            roots.append(base)
    keys = ("结算", "变更", "燃气", "设计", "材料", "平差")
    for base in roots:
        if base in ("E:\\", "D:\\", "C:\\"):
            try:
                subs = os.listdir(base)
            except Exception:
                continue
            for s in subs:
                if any(k in s for k in keys):
                    roots.append(os.path.join(base, s))
    found = []
    for r in roots:
        if r not in found and looks_like_archive(r):
            found.append(os.path.abspath(r))
    return found or None


def write_config(skill_dir, archives):
    """写入归档目录配置（支持多个目录）。"""
    cfg = os.path.join(skill_dir, "scripts", "config.json")
    paths = archives if isinstance(archives, (list, tuple)) else [archives]
    with open(cfg, "w", encoding="utf-8") as fh:
        json.dump({"archive_roots": list(paths)}, fh, ensure_ascii=False, indent=2)
    ok("已写入配置：%s" % cfg)


# ---------------------------------------------------------------- 5. 自检

def run_doctor(skill_dir):
    script = os.path.join(skill_dir, "scripts", "fill_change_reason.py")
    if not os.path.exists(script):
        err("找不到 %s" % script)
        return False
    r = subprocess.run([sys.executable, script, "doctor"],
                       capture_output=True, text=True, cwd=os.path.dirname(script))
    out = (r.stdout or "") + (r.stderr or "")
    for line in out.splitlines():
        print("    " + line)
    return r.returncode == 0


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description="燃气工程变更原因生成 —— 一键安装")
    ap.add_argument("--target", default="", help="skill 安装目录（默认用户级 skills 目录）")
    ap.add_argument("--archive", default="", help="历史变更单归档目录（多个用 ; 分隔）")
    ap.add_argument("--no-deps", action="store_true", help="跳过依赖安装")
    ap.add_argument("--check-only", action="store_true", help="只做环境体检，不安装")
    args = ap.parse_args()

    print("=" * 66)
    print("燃气工程变更原因生成 —— 一键安装")
    print("=" * 66)

    if not check_python():
        return 2

    if args.check_only:
        return 0 if install_deps() else 1

    if not args.no_deps:
        if not install_deps():
            warn("依赖安装未完全成功，继续安装；若后续报错请手动处理")
        print("")

    # skill 源目录 = 本脚本上级
    src_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if not os.path.exists(os.path.join(src_root, "SKILL.md")):
        err("当前目录不是 skill 根目录（缺少 SKILL.md）")
        return 2

    target = args.target or default_skill_dir()
    skill_dir = copy_skill(src_root, target)
    print("")

    # 归档目录
    archives = probe_archive(args.archive)
    if archives:
        write_config(skill_dir, archives)
        total = 0
        for a in archives:
            n = len(glob.glob(os.path.join(a, "**", "*.xls"), recursive=True))
            total += n
            ok("归档目录可用：%s（%d 份 .xls）" % (a, n))
        ok("语料来源合计 %d 份变更单" % total)
    else:
        warn("未能自动定位历史变更单归档目录")
        info("仍可正常使用（检索相似案例将使用随包语料）")
        info("如需指定，运行：python install.py --archive <你的归档目录>")
    print("")

    # 自检
    log("INFO", "运行 doctor 自检")
    good = run_doctor(skill_dir)

    print("")
    print("=" * 66)
    if good:
        print("安装完成 ✓")
        print("")
        print("下一步：")
        print("  1) 把变更单交给 Agent，说「帮我填写这份变更单的变更原因：<路径>」")
        print("  2) 或直接命令行使用：")
        print("       cd \"%s\"" % os.path.join(skill_dir, "scripts"))
        print("       python fill_change_reason.py analyze <变更单.xls>")
        print("       python fill_change_reason.py fill    <变更单.xls>")
    else:
        print("安装基本完成，但自检有未通过项 ⚠")
        print("请查看上方 ✗ 项，按提示处理后重新运行：")
        print("  python \"%s\"" % os.path.join(skill_dir, "scripts", "fill_change_reason.py"))
    print("=" * 66)
    return 0 if good else 1


if __name__ == "__main__":
    sys.exit(main())
