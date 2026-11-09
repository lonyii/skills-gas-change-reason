#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""gas-change-reason —— 自动更新模块。

从 GitHub 仓库检查新版本并就地升级。

设计要点
--------
1. **只走匿名 HTTPS**，不需要 GitHub Token（仓库公开）。
2. **保留本地配置**：``scripts/config.json`` 记录了同事自己的归档目录路径，
   升级时必须保留，否则升级后又要重新配置。
3. **原子替换 + 自动回滚**：新版本先解压到临时目录，校验通过后才替换旧目录；
   若替换过程出错，自动用备份目录回滚，避免把技能装坏。
4. **幂等**：已是最新版本时不做任何写入。

用法
----
    python update.py --check      # 只检查，不更新
    python update.py              # 检查并更新到最新版
    python update.py --rollback   # 回滚到上一个版本
"""
from __future__ import print_function

import io
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

try:
    from urllib.request import Request, urlopen
    from urllib.error import HTTPError, URLError
except ImportError:                                   # Python 2 兜底
    from urllib2 import Request, urlopen, HTTPError, URLError

#: 仓库地址
REPO = "lonyii/skills-gas-change-reason"
GITHUB_REPO = "https://github.com/%s" % REPO
#: 版本清单（放在仓库根的 VERSION 文件，内容形如 "1.1.0"）
RAW_VERSION = "https://raw.githubusercontent.com/%s/main/VERSION" % REPO
#: 源码 zip（仓库快照，含全部文件）
CODE_ZIP = "https://codeload.github.com/%s/zip/refs/heads/main" % REPO
#: Release 资产（若有）
RELEASE_API = "https://api.github.com/repos/%s/releases/latest" % REPO

#: 技能根目录（本文件位于 <skill>/scripts/ 下）
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VERSION_FILE = os.path.join(SKILL_DIR, "VERSION")
CONFIG_FILE = os.path.join(SKILL_DIR, "scripts", "config.json")

#: 备份目录——必须放在技能目录**之外**。
#: 若放在内部（如 <skill>/.update_backup），「把技能目录移进备份目录」会变成
#: "把目录移进它自己的子目录"，Windows 直接报
#: "Cannot move a directory into itself"。
BACKUP_DIR = os.path.join(
    os.path.dirname(SKILL_DIR), ".%s.update_backup" % os.path.basename(SKILL_DIR))

#: 不纳入版本管理的文件
IGNORE_DIRS = {"__pycache__", ".git", ".idea", ".vscode", ".update_backup"}
IGNORE_FILES = {".pyc", ".pyo", "config.json"}

UA = "gas-change-reason-skill-updater"


# ------------------------------------------------------------------ 输出

def _out(msg):
    sys.stdout.write(msg + "\n")
    sys.stdout.flush()


def _err(msg):
    sys.stderr.write(msg + "\n")
    sys.stderr.flush()


def _ok(msg):
    _out("[OK] %s" % msg)


def _info(msg):
    _out("[INFO] %s" % msg)


def _warn(msg):
    _out("[WARN] %s" % msg)


def _fail(msg):
    _err("[ERROR] %s" % msg)


# ------------------------------------------------------------------ 版本

def local_version():
    """读取本地版本号；读不到返回 "0.0.0"。"""
    if os.path.exists(VERSION_FILE):
        try:
            with open(VERSION_FILE, "r", encoding="utf-8") as fh:
                v = fh.read().strip()
            if v:
                return v
        except Exception:
            pass
    # 退而从 SKILL.md frontmatter 读
    sk = os.path.join(SKILL_DIR, "SKILL.md")
    if os.path.exists(sk):
        try:
            with open(sk, "r", encoding="utf-8") as fh:
                head = fh.read(600)
            m = re.search(r"^version:\s*(\S+)", head, re.M)
            if m:
                return m.group(1).strip()
        except Exception:
            pass
    return "0.0.0"


def _ver_tuple(v):
    """'1.2.3' → (1, 2, 3)；无法解析的部分按 0 处理。"""
    parts = re.findall(r"\d+", str(v or ""))
    while len(parts) < 3:
        parts.append("0")
    return tuple(int(x) for x in parts[:3])


def remote_version(timeout=15):
    """取远端最新版本号；失败返回 None。"""
    try:
        req = Request(RAW_VERSION, headers={"User-Agent": UA})
        with urlopen(req, timeout=timeout) as resp:
            v = resp.read().decode("utf-8", "replace").strip()
        return v or None
    except HTTPError as e:
        _err("无法读取远端 VERSION（HTTP %s）：%s" % (e.code, RAW_VERSION))
    except URLError as e:
        _err("网络不可达：%s" % (e.reason))
    except Exception as e:
        _err("读取远端版本失败：%s" % e)
    return None


def is_newer(remote, local):
    """远端是否比本地新。"""
    return _ver_tuple(remote) > _ver_tuple(local)


# ------------------------------------------------------------------ 下载

def download(url, dest, timeout=120):
    """下载到 dest；返回 True 表示成功。"""
    try:
        req = Request(url, headers={"User-Agent": UA})
        with urlopen(req, timeout=timeout) as resp, open(dest, "wb") as fh:
            shutil.copyfileobj(resp, fh)
        return True
    except HTTPError as e:
        _err("下载失败 HTTP %s：%s" % (e.code, url))
    except URLError as e:
        _err("下载失败（网络）：%s" % e.reason)
    except Exception as e:
        _err("下载失败：%s" % e)
    return False


def extract_skill(zip_path, target_dir):
    """从 zip 解出技能目录到 target_dir。

    GitHub 的 zip 解压后顶层是 ``<repo>-main/``，需剥离这一层。
    返回解出的技能根目录路径；失败返回 None。
    """
    try:
        with zipfile.ZipFile(zip_path) as z:
            names = z.namelist()
            roots = set(n.split("/")[0] for n in names if n and not n.startswith("__MACOSX"))
            if len(roots) != 1:
                _err("压缩包结构异常，顶层目录：%s" % ", ".join(sorted(roots)))
                return None
            root = roots.pop()
            for n in names:
                if n.startswith("__MACOSX"):
                    continue
                rel = n[len(root):].lstrip("/")
                if not rel:
                    continue
                dest = os.path.join(target_dir, rel.replace("/", os.sep))
                # 阻断路径穿越（zip slip）
                if os.path.abspath(dest).startswith(os.path.abspath(target_dir) + os.sep) \
                        or os.path.abspath(dest) == os.path.abspath(target_dir):
                    pass
                else:
                    continue
                if n.endswith("/"):
                    if not os.path.isdir(dest):
                        os.makedirs(dest)
                    continue
                d = os.path.dirname(dest)
                if d and not os.path.isdir(d):
                    os.makedirs(d)
                with z.open(n) as src, open(dest, "wb") as out:
                    shutil.copyfileobj(src, out)
    except zipfile.BadZipFile:
        _err("下载的文件不是有效的 zip 包")
        return None
    except Exception as e:
        _err("解压失败：%s" % e)
        return None
    return target_dir


# ------------------------------------------------------------------ 替换

def _read_config():
    """读取本地配置（升级时保留）。"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return None
    return None


def apply_update(new_dir):
    """用 new_dir 的内容替换当前技能目录。

    先把旧目录移到 .update_backup，成功后删除备份；失败则回滚。
    返回 True 表示成功。
    """
    # 1) 备份配置
    cfg = _read_config()
    # 2) 备份旧目录
    if os.path.isdir(BACKUP_DIR):
        shutil.rmtree(BACKUP_DIR, ignore_errors=True)
    parent = os.path.dirname(SKILL_DIR)
    try:
        if os.path.isdir(SKILL_DIR):
            shutil.move(SKILL_DIR, BACKUP_DIR)
    except Exception as e:
        _err("备份旧版本失败：%s" % e)
        return False
    try:
        # 3) 复制新版本
        shutil.copytree(new_dir, SKILL_DIR,
                        ignore=shutil.ignore_patterns(*IGNORE_DIRS, *IGNORE_FILES))
        # 4) 恢复配置
        if cfg:
            cfg_path = os.path.join(SKILL_DIR, "scripts", "config.json")
            d = os.path.dirname(cfg_path)
            if not os.path.isdir(d):
                os.makedirs(d)
            with open(cfg_path, "w", encoding="utf-8") as fh:
                json.dump(cfg, fh, ensure_ascii=False, indent=2)
            _ok("已保留本地归档目录配置")
        return True
    except Exception as e:
        _err("替换失败：%s" % e)
        # 回滚
        try:
            if os.path.isdir(SKILL_DIR):
                shutil.rmtree(SKILL_DIR, ignore_errors=True)
            if os.path.isdir(BACKUP_DIR):
                shutil.move(BACKUP_DIR, SKILL_DIR)
            _ok("已回滚到原版本")
        except Exception as e2:
            _err("回滚也失败：%s" % e2)
        return False


def rollback():
    """回滚到 .update_backup 中的旧版本。"""
    if not os.path.isdir(BACKUP_DIR):
        _warn("没有可回滚的备份（.update_backup 不存在）")
        return 1
    cfg = _read_config()
    tmp_old = os.path.join(tempfile.gettempdir(), "gcr_rollback_%d" % os.getpid())
    try:
        shutil.move(SKILL_DIR, tmp_old)
        shutil.move(BACKUP_DIR, SKILL_DIR)
    except Exception as e:
        _err("回滚失败：%s" % e)
        # 尽量恢复原状
        if os.path.isdir(tmp_old) and not os.path.isdir(SKILL_DIR):
            shutil.move(tmp_old, SKILL_DIR)
        return 1
    if cfg:
        try:
            cfg_path = os.path.join(SKILL_DIR, "scripts", "config.json")
            with open(cfg_path, "w", encoding="utf-8") as fh:
                json.dump(cfg, fh, ensure_ascii=False, indent=2)
        except Exception:
            pass
    shutil.rmtree(tmp_old, ignore_errors=True)
    _ok("已回滚到上一个版本")
    return 0


# ------------------------------------------------------------------ 主流程

def do_check():
    """检查是否有新版本。返回 (remote, local, has_new)。"""
    local = local_version()
    _info("本地版本: %s" % local)
    remote = remote_version()
    if remote is None:
        _err("无法获取远端版本，检查中止")
        return None, local, False
    _info("远端版本: %s" % remote)
    has_new = is_newer(remote, local)
    if has_new:
        _out("[有更新] %s → %s" % (local, remote))
    else:
        _ok("已是最新版本")
    return remote, local, has_new


def do_update():
    """检查并更新。返回退出码。"""
    remote, local, has_new = do_check()
    if remote is None:
        return 2
    if not has_new:
        return 0

    _info("开始下载新版本 %s …" % remote)
    tmp = tempfile.mkdtemp(prefix="gcr_update_")
    try:
        zp = os.path.join(tmp, "skill.zip")
        if not download(CODE_ZIP, zp):
            return 2
        new_dir = os.path.join(tmp, "new")
        os.makedirs(new_dir, exist_ok=True)
        got = extract_skill(zp, new_dir)
        if not got:
            return 2
        # 校验新版本完整性
        for must in ("SKILL.md", "scripts/gas_change_core.py",
                     "scripts/material_reason.py", "scripts/fill_change_reason.py"):
            if not os.path.exists(os.path.join(got, must.replace("/", os.sep))):
                _err("新版本缺少必需文件：%s，更新中止" % must)
                return 2
        _ok("新版本校验通过")
        if apply_update(got):
            _ok("更新完成：%s → %s" % (local, remote))
            _out("")
            _out("如新版本有功能变化，建议重新运行一次 doctor 确认环境：")
            _out("    python scripts/fill_change_reason.py doctor")
            return 0
        return 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main(argv):
    check_only = "--check" in argv
    do_rb = "--rollback" in argv
    if do_rb:
        return rollback()
    if check_only:
        remote, local, has_new = do_check()
        return 0 if remote else 2
    return do_update()


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        _err("已中断")
        sys.exit(130)
