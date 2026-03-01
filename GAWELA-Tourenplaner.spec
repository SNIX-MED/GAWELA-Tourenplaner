# -*- mode: python ; coding: utf-8 -*-

import os

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


def collect_runtime_tree(source_dir, target_root):
    collected = []
    if not source_dir or not os.path.isdir(source_dir):
        return collected

    source_dir = os.path.abspath(source_dir)
    for current_root, _dirs, files in os.walk(source_dir):
        relative_root = os.path.relpath(current_root, source_dir)
        for filename in files:
            source_path = os.path.join(current_root, filename)
            if relative_root == ".":
                target_path = os.path.join(target_root, filename)
            else:
                target_path = os.path.join(target_root, relative_root, filename)
            collected.append((source_path, target_path))
    return collected


def find_webview2_runtime_dir(*roots):
    for root in roots:
        if not root:
            continue
        root = os.path.abspath(root)
        if not os.path.isdir(root):
            continue

        direct_executable = os.path.join(root, "msedgewebview2.exe")
        if os.path.exists(direct_executable):
            return root

        child_dirs = []
        try:
            child_dirs = [
                entry.path
                for entry in os.scandir(root)
                if entry.is_dir()
            ]
        except OSError:
            child_dirs = []

        def rank(path):
            name = os.path.basename(path)
            numeric_parts = [int(part) if part.isdigit() else -1 for part in name.split(".")]
            return tuple(numeric_parts), name.lower()

        for candidate in sorted(child_dirs, key=rank, reverse=True):
            executable = os.path.join(candidate, "msedgewebview2.exe")
            if os.path.exists(executable):
                return candidate
    return None


pywebview_datas = collect_data_files("webview")
pywebview_hiddenimports = collect_submodules("webview")

project_root = os.path.abspath(".")
local_runtime_dir = os.path.join(project_root, "assets", "webview2")
runtime_source_dir = find_webview2_runtime_dir(
    os.environ.get("GAWELA_WEBVIEW2_RUNTIME_DIR"),
    local_runtime_dir,
    r"C:\Program Files (x86)\Microsoft\EdgeWebView\Application",
    r"C:\Program Files\Microsoft\EdgeWebView\Application",
)
fixed_runtime_datas = []
if runtime_source_dir:
    runtime_source_dir = os.path.abspath(runtime_source_dir)
    local_runtime_match = find_webview2_runtime_dir(local_runtime_dir)
    local_runtime_match = os.path.abspath(local_runtime_match) if local_runtime_match else None
    if runtime_source_dir != local_runtime_match:
        fixed_runtime_datas = collect_runtime_tree(runtime_source_dir, os.path.join("assets", "webview2"))


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets'), *pywebview_datas, *fixed_runtime_datas],
    hiddenimports=pywebview_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='GAWELA-Tourenplaner',
    icon='assets\\Applogo.ico',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GAWELA-Tourenplaner',
)
