import argparse
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
ENTRYPOINT = ROOT / "compare_tool.py"
APP_NAME = "PyCompareStudio"


def run(command: list[str]) -> None:
    print("$ " + " ".join(command))
    subprocess.run(command, cwd=ROOT, check=True)


def ensure_pyinstaller(auto_install: bool) -> None:
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        if not auto_install:
            raise SystemExit(
                "PyInstaller is not installed. Run: python -m pip install pyinstaller"
            )
        run([sys.executable, "-m", "pip", "install", "pyinstaller"])


def clean_build_outputs() -> None:
    for name in ("build", "dist"):
        path = ROOT / name
        if path.exists():
            shutil.rmtree(path)
    for spec in ROOT.glob("*.spec"):
        spec.unlink()


def build(args: argparse.Namespace) -> None:
    if not ENTRYPOINT.exists():
        raise SystemExit(f"Missing entrypoint: {ENTRYPOINT}")

    if args.clean:
        clean_build_outputs()

    ensure_pyinstaller(args.install_pyinstaller)

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--name",
        args.name,
        "--noconfirm",
        "--clean",
    ]

    if args.onefile:
        command.append("--onefile")
    else:
        command.append("--onedir")

    if args.console:
        command.append("--console")
    else:
        command.append("--windowed")

    hidden_imports = [
        "tkinter",
        "tkinter.ttk",
        "tkinter.scrolledtext",
        "PIL.Image",
        "PIL.ImageChops",
        "PIL.ImageTk",
        "openpyxl",
        "py7zr",
    ]
    for module in hidden_imports:
        command.extend(["--hidden-import", module])

    command.append(str(ENTRYPOINT))
    run(command)

    system = platform.system().lower()
    binary_name = args.name + (".exe" if system == "windows" and args.onefile else "")
    if args.onefile:
        output = ROOT / "dist" / binary_name
    else:
        output = ROOT / "dist" / args.name
    print()
    print(f"Build complete for {platform.system()} {platform.machine()}:")
    print(output)
    if platform.system().lower() != args.target.lower() and args.target != "current":
        print()
        print("Note: PyInstaller builds for the current OS only.")
        print("Run this same script on the target OS to create that platform's executable.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build PyCompare Studio executable with PyInstaller."
    )
    parser.add_argument("--name", default=APP_NAME, help="Executable/app folder name.")
    parser.add_argument(
        "--target",
        choices=["current", "windows", "linux"],
        default="current",
        help="Documentation hint only; build runs on the current OS.",
    )
    parser.add_argument(
        "--onedir",
        action="store_true",
        help="Create a dist folder instead of a single executable.",
    )
    parser.add_argument(
        "--console",
        action="store_true",
        help="Keep a console window visible for logs/debugging.",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Delete previous build, dist, and spec files before building.",
    )
    parser.add_argument(
        "--no-install",
        dest="install_pyinstaller",
        action="store_false",
        help="Do not auto-install PyInstaller if it is missing.",
    )
    parser.set_defaults(install_pyinstaller=True)
    args = parser.parse_args()
    args.onefile = not args.onedir
    return args


if __name__ == "__main__":
    build(parse_args())
