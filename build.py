import os
import sys
import shutil
import subprocess
import importlib
import textwrap

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MAIN_SCRIPT = os.path.join(SCRIPT_DIR, "send_email_gui.py")
REQUIREMENTS_FILE = os.path.join(SCRIPT_DIR, "requirements.txt")

REQUIRED_PACKAGES = [
    {"import_name": "openpyxl", "pip_name": "openpyxl", "min_version": "3.1.0"},
    {"import_name": "customtkinter", "pip_name": "customtkinter", "min_version": "5.2.0"},
]

BUILD_TOOLS = [
    {"import_name": "PyInstaller", "pip_name": "pyinstaller", "min_version": "6.0.0"},
]

CLEAN_DIRS = ["build", "__pycache__"]
CLEAN_PATTERNS = ["*.spec", "*.pyc"]


def version_tuple(v):
    parts = []
    for p in v.split("."):
        try:
            parts.append(int(p))
        except ValueError:
            break
    return tuple(parts)


def check_package(pkg):
    try:
        mod = importlib.import_module(pkg["import_name"])
        installed_version = getattr(mod, "__version__", "0.0.0")
        min_ver = version_tuple(pkg["min_version"])
        installed_ver = version_tuple(installed_version)
        if installed_ver >= min_ver:
            return True, installed_version
        return False, installed_version
    except ImportError:
        return False, None


def install_package(pkg, error_code=1):
    spec = f"{pkg['pip_name']}=={pkg['min_version']}"
    print(f"  [安装] pip install {spec}")
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", spec, "--quiet"],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode != 0:
            print(f"  [错误] 安装 {pkg['pip_name']} 失败:")
            print(f"         {result.stderr.strip()}")
            sys.exit(error_code)
        print(f"  [成功] {pkg['pip_name']} 安装完成")
        return True
    except subprocess.TimeoutExpired:
        print(f"  [错误] 安装 {pkg['pip_name']} 超时 (120秒)")
        sys.exit(error_code)
    except Exception as e:
        print(f"  [错误] 安装 {pkg['pip_name']} 异常: {e}")
        sys.exit(error_code)


def phase_check_dependencies():
    print("=" * 60)
    print("  阶段 1/3: 依赖检查与自动安装")
    print("=" * 60)

    if not os.path.isfile(MAIN_SCRIPT):
        print(f"[错误] 主程序文件不存在: {MAIN_SCRIPT}")
        sys.exit(1)

    all_packages = REQUIRED_PACKAGES + BUILD_TOOLS
    failed = []

    for pkg in all_packages:
        ok, ver = check_package(pkg)
        if ok:
            print(f"  [通过] {pkg['pip_name']} {ver} (需要 >= {pkg['min_version']})")
        else:
            if ver:
                print(f"  [不兼容] {pkg['pip_name']} {ver} (需要 >= {pkg['min_version']})")
            else:
                print(f"  [缺失] {pkg['pip_name']} (需要 >= {pkg['min_version']})")
            failed.append(pkg)

    if failed:
        print(f"\n  共 {len(failed)} 个依赖需要安装/更新")
        for pkg in failed:
            install_package(pkg)
        print("\n  重新验证...")
        for pkg in failed:
            ok, ver = check_package(pkg)
            if ok:
                print(f"  [通过] {pkg['pip_name']} {ver}")
            else:
                print(f"  [失败] {pkg['pip_name']} 无法满足要求，构建终止")
                sys.exit(1)

    print("\n  所有依赖检查通过\n")


def phase_package_exe():
    print("=" * 60)
    print("  阶段 2/3: 打包为 EXE")
    print("=" * 60)

    try:
        import PyInstaller.__main__
    except ImportError:
        print("[错误] PyInstaller 未安装，无法执行打包")
        sys.exit(1)

    ctk_path = os.path.dirname(importlib.import_module("customtkinter").__file__)

    pyinstaller_args = [
        MAIN_SCRIPT,
        "--name", "send_email_gui",
        "--onefile",
        "--noconsole",
        "--clean",
        "--hidden-import", "openpyxl",
        "--hidden-import", "customtkinter",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.scrolledtext",
        "--hidden-import", "tkinter.ttk",
        "--collect-data", "customtkinter",
        "--collect-submodules", "customtkinter",
        "--distpath", os.path.join(SCRIPT_DIR, "dist"),
        "--workpath", os.path.join(SCRIPT_DIR, "build"),
        "--specpath", SCRIPT_DIR,
        "--log-level", "WARN",
    ]

    print(f"  [打包] PyInstaller 开始构建...")
    print(f"  [信息] CustomTkinter 路径: {ctk_path}")
    print(f"  [信息] 输出目录: {os.path.join(SCRIPT_DIR, 'dist')}")
    print()

    try:
        PyInstaller.__main__.run(pyinstaller_args)
    except SystemExit as e:
        if e.code != 0:
            print(f"\n[错误] PyInstaller 打包失败 (退出码: {e.code})")
            sys.exit(1)
    except Exception as e:
        print(f"\n[错误] PyInstaller 打包异常: {e}")
        sys.exit(1)

    exe_path = os.path.join(SCRIPT_DIR, "dist", "send_email_gui.exe")
    if not os.path.isfile(exe_path):
        print("[错误] 打包完成后未找到 EXE 文件")
        sys.exit(1)

    size_mb = os.path.getsize(exe_path) / (1024 * 1024)
    print(f"\n  [成功] EXE 已生成: {exe_path}")
    print(f"  [信息] 文件大小: {size_mb:.1f} MB\n")


def phase_cleanup():
    print("=" * 60)
    print("  阶段 3/3: 清理构建临时文件")
    print("=" * 60)

    cleaned = 0

    for dirname in CLEAN_DIRS:
        dir_path = os.path.join(SCRIPT_DIR, dirname)
        if os.path.isdir(dir_path):
            try:
                shutil.rmtree(dir_path)
                print(f"  [已删除] 目录: {dirname}/")
                cleaned += 1
            except Exception as e:
                print(f"  [警告] 无法删除 {dirname}/: {e}")

    dist_dir = os.path.join(SCRIPT_DIR, "dist")
    if os.path.isdir(dist_dir):
        for fn in os.listdir(dist_dir):
            fp = os.path.join(dist_dir, fn)
            if os.path.isfile(fp) and not fn.endswith(".exe"):
                try:
                    os.remove(fp)
                    print(f"  [已删除] dist/{fn}")
                    cleaned += 1
                except Exception:
                    pass

    for pattern in CLEAN_PATTERNS:
        import glob
        for f in glob.glob(os.path.join(SCRIPT_DIR, pattern)):
            try:
                os.remove(f)
                print(f"  [已删除] 文件: {os.path.basename(f)}")
                cleaned += 1
            except Exception as e:
                print(f"  [警告] 无法删除 {os.path.basename(f)}: {e}")

    pip_cache = os.path.join(SCRIPT_DIR, "pip-delete-this-directory.txt")
    if os.path.isfile(pip_cache):
        try:
            os.remove(pip_cache)
            cleaned += 1
        except Exception:
            pass

    print(f"\n  [完成] 共清理 {cleaned} 个临时文件/目录\n")


def verify_final_artifacts():
    print("=" * 60)
    print("  最终验证")
    print("=" * 60)

    exe_path = os.path.join(SCRIPT_DIR, "dist", "send_email_gui.exe")
    checks = [
        ("EXE 文件", os.path.isfile(exe_path)),
        ("build/ 目录已清理", not os.path.isdir(os.path.join(SCRIPT_DIR, "build"))),
        ("__pycache__/ 已清理", not os.path.isdir(os.path.join(SCRIPT_DIR, "__pycache__"))),
        ("主程序文件完好", os.path.isfile(MAIN_SCRIPT)),
        ("requirements.txt 完好", os.path.isfile(REQUIREMENTS_FILE)),
    ]

    all_pass = True
    for name, passed in checks:
        status = "[通过]" if passed else "[未通过]"
        print(f"  {status} {name}")
        if not passed:
            all_pass = False

    if all_pass:
        print("\n  构建流程全部完成!")
    else:
        print("\n  部分检查未通过，请检查上方输出")
    print("=" * 60)


def main():
    print()
    phase_check_dependencies()
    phase_package_exe()
    phase_cleanup()
    verify_final_artifacts()


if __name__ == "__main__":
    main()
