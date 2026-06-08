import argparse
import os
import platform
import shlex
import shutil
import stat
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
APP_SCRIPT = ROOT / "compare_tool.py"
APP_NAME = "PyCompare Studio"
WINDOWS_LEGACY_KEYS = [
    r"Software\Classes\*\shell\PyCompareStudioSelectLeft",
    r"Software\Classes\*\shell\PyCompareStudioCompareToLeft",
    r"Software\Classes\*\shell\PyCompareStudioCompareSelected",
    r"Software\Classes\Directory\shell\PyCompareStudioSelectLeft",
    r"Software\Classes\Directory\shell\PyCompareStudioCompareToLeft",
    r"Software\Classes\Directory\shell\PyCompareStudioCompareSelected",
]
WINDOWS_KEYS = [
    r"Software\Classes\AllFilesystemObjects\shell\PyCompareStudio",
    r"Software\Classes\PyCompareStudio.ContextMenu",
    r"Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PyCompareStudio.Compare",
    r"Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PyCompareStudio.CompareToLeft",
    r"Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PyCompareStudio.SelectLeft",
]


def run(command: list[str]) -> None:
    print("$ " + " ".join(command))
    subprocess.run(command, check=True)


def windows_pythonw() -> str:
    python = Path(sys.executable)
    candidate = python.with_name("pythonw.exe")
    return str(candidate if candidate.exists() else python)


def app_command(app_path: Path, args: list[str]) -> list[str]:
    if app_path.suffix.lower() == ".py":
        if os.name == "nt":
            return [windows_pythonw(), str(app_path), *args]
        return [sys.executable, str(app_path), *args]
    return [str(app_path), *args]


def command_string(app_path: Path, args: list[str]) -> str:
    command = app_command(app_path, args)
    if os.name == "nt":
        parts = []
        for item in command:
            if item == "%1":
                parts.append('"%1"')
            elif item == "%*":
                parts.append("%*")
            else:
                parts.append(subprocess.list2cmdline([item]))
        return " ".join(parts)
    return shlex.join(command)


def validate_app_path(app_path: Path) -> None:
    if not app_path.exists():
        raise SystemExit(f"App path does not exist: {app_path}")


def install_windows(app_path: Path) -> None:
    import winreg

    validate_app_path(app_path)
    uninstall_windows_registry_keys()
    parent_key = r"Software\Classes\AllFilesystemObjects\shell\PyCompareStudio"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, parent_key) as key:
        winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "PyCompareStudio")
        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, str(app_path))
        winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, "")
        winreg.SetValueEx(key, "MultiSelectModel", 0, winreg.REG_SZ, "Player")

    entries = [
        (
            parent_key + r"\shell\Compare",
            "Compare",
            ["--compare-selected", "%*"],
        ),
        (
            parent_key + r"\shell\CompareToLeft",
            "Compare to selected left",
            ["--compare-to-left", "%1"],
        ),
        (
            parent_key + r"\shell\SelectLeft",
            "Select left file to compare",
            ["--select-left", "%1"],
        ),
    ]
    for key_path, label, command_args in entries:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, label)
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, str(app_path))
            winreg.SetValueEx(key, "MultiSelectModel", 0, winreg.REG_SZ, "Player")
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r"\command") as command_key:
            winreg.SetValueEx(command_key, "", 0, winreg.REG_SZ, command_string(app_path, command_args))

    install_windows_sendto(app_path)
    print("Windows context submenu installed for current user.")
    print("For bulk selection, 'Send to > PyCompare Studio Compare' is also installed as a reliable fallback.")


def install_windows_sendto(app_path: Path) -> None:
    sendto = Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "SendTo"
    sendto.mkdir(parents=True, exist_ok=True)
    target = sendto / "PyCompare Studio Compare.cmd"
    command = subprocess.list2cmdline(app_command(app_path, ["--compare-selected", "%*"]))
    target.write_text(f"@echo off\r\n{command}\r\n", encoding="utf-8")


def uninstall_windows() -> None:
    uninstall_windows_registry_keys()
    sendto = Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "SendTo" / "PyCompare Studio Compare.cmd"
    if sendto.exists():
        sendto.unlink()
    print("Windows context menu removed for current user.")


def uninstall_windows_registry_keys() -> None:
    import winreg

    for key_path in WINDOWS_KEYS + WINDOWS_LEGACY_KEYS:
        delete_registry_tree(winreg.HKEY_CURRENT_USER, key_path)


def delete_registry_tree(root, key_path: str) -> None:
    import winreg

    try:
        with winreg.OpenKey(root, key_path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                delete_registry_tree(root, key_path + "\\" + child)
    except FileNotFoundError:
        return
    try:
        winreg.DeleteKeyEx(root, key_path)
    except OSError:
        try:
            winreg.DeleteKey(root, key_path)
        except OSError:
            pass


def shell_script(command: str) -> str:
    return f"""#!/usr/bin/env bash
set -e
{command}
"""


def install_linux(app_path: Path) -> None:
    validate_app_path(app_path)
    scripts_dir = Path.home() / ".local" / "share" / "nautilus" / "scripts" / APP_NAME
    scripts_dir.mkdir(parents=True, exist_ok=True)
    select_left_cmd = command_string(app_path, ["--select-left"])
    compare_left_cmd = command_string(app_path, ["--compare-to-left"])
    compare_selected_cmd = command_string(app_path, ["--compare-selected"])
    first_selected = 'first_path="${NAUTILUS_SCRIPT_SELECTED_FILE_PATHS%%$\'\\n\'*}"'
    write_executable(
        scripts_dir / "Select Left File for Compare",
        shell_script(f"{first_selected}\n{select_left_cmd} \"$first_path\""),
    )
    write_executable(
        scripts_dir / "Compare to selected left file",
        shell_script(f"{first_selected}\n{compare_left_cmd} \"$first_path\""),
    )
    write_executable(
        scripts_dir / "Compare",
        shell_script(
            f"mapfile -t paths <<< \"$NAUTILUS_SCRIPT_SELECTED_FILE_PATHS\"\n"
            f"{compare_selected_cmd} \"${{paths[@]}}\""
        ),
    )
    install_dolphin_service(app_path)
    print("Linux context menu helpers installed.")
    print("Nautilus: right click > Scripts > PyCompare Studio.")
    print("Dolphin: right click selected files > Actions > PyCompare Studio.")


def write_executable(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8", newline="\n")
    path.chmod(path.stat().st_mode | stat.S_IXUSR)


def install_dolphin_service(app_path: Path) -> None:
    service_dir = Path.home() / ".local" / "share" / "kio" / "servicemenus"
    service_dir.mkdir(parents=True, exist_ok=True)
    service = service_dir / "pycompare-studio.desktop"
    select_cmd = command_string(app_path, ["--select-left", "%f"])
    compare_left_cmd = command_string(app_path, ["--compare-to-left", "%f"])
    compare_selected_cmd = command_string(app_path, ["--compare-selected", "%F"])
    service.write_text(
        f"""[Desktop Entry]
Type=Service
MimeType=all/all;inode/directory;
X-KDE-ServiceTypes=KonqPopupMenu/Plugin
X-KDE-Priority=TopLevel
Actions=selectLeft;compareToLeft;compareSelected;

[Desktop Action selectLeft]
Name=Select Left File for Compare
Exec={select_cmd}

[Desktop Action compareToLeft]
Name=Compare to selected left file
Exec={compare_left_cmd}

[Desktop Action compareSelected]
Name=Compare
Exec={compare_selected_cmd}
""",
        encoding="utf-8",
    )


def uninstall_linux() -> None:
    scripts_dir = Path.home() / ".local" / "share" / "nautilus" / "scripts" / APP_NAME
    if scripts_dir.exists():
        shutil.rmtree(scripts_dir)
    service = Path.home() / ".local" / "share" / "kio" / "servicemenus" / "pycompare-studio.desktop"
    if service.exists():
        service.unlink()
    print("Linux context menu helpers removed.")


def install(args: argparse.Namespace) -> None:
    app_path = Path(args.app_path).expanduser().resolve()
    system = platform.system().lower()
    if system == "windows":
        install_windows(app_path)
    elif system == "linux":
        install_linux(app_path)
    else:
        raise SystemExit(f"Unsupported OS for context menu install: {platform.system()}")


def uninstall() -> None:
    system = platform.system().lower()
    if system == "windows":
        uninstall_windows()
    elif system == "linux":
        uninstall_linux()
    else:
        raise SystemExit(f"Unsupported OS for context menu uninstall: {platform.system()}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Install or remove PyCompare Studio context menus.")
    parser.add_argument("action", choices=["install", "uninstall"])
    parser.add_argument(
        "--app-path",
        default=str(APP_SCRIPT),
        help="Path to compare_tool.py or the built PyCompareStudio executable.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    parsed = parse_args()
    if parsed.action == "install":
        install(parsed)
    else:
        uninstall()
