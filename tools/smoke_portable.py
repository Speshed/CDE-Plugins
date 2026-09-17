from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def _clean_environment() -> dict[str, str]:
    env = os.environ.copy()
    # Do not let the packaged application accidentally see/import the source tree.
    for key in (
        "PYTHONPATH",
        "PYTHONHOME",
        "SOD_MANAGER_ASSETS",
        "SOD_MANAGER_ROOT",
        "SOD_MANAGER_PARENT",
        "SOD_MANAGER_THEME",
    ):
        env.pop(key, None)
    env.update(
        {
            "CDE_SMOKE_TEST": "1",
            "QT_QPA_PLATFORM": "offscreen",
            "QTWEBENGINE_CHROMIUM_FLAGS": "--disable-gpu --no-sandbox",
        }
    )
    return env


def _run_mode(exe: Path, name: str, args: list[str], env: dict[str, str]) -> int:
    flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) if os.name == "nt" else 0
    process = subprocess.Popen(
        [str(exe), *args],
        cwd=str(exe.parent),
        env=env,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        close_fds=True,
        creationflags=flags,
    )
    try:
        output, _ = process.communicate(timeout=60)
    except subprocess.TimeoutExpired:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        else:
            process.kill()
        raise
    if output.strip():
        print(output.strip(), flush=True)
    print(f"{name}: exit={process.returncode}", flush=True)
    return int(process.returncode or 0)


def main() -> int:
    source_exe = Path(__file__).resolve().parent.parent / "release" / "Larix_CDE.exe"
    if not source_exe.is_file():
        print(f"[ERROR] EXE not found: {source_exe}")
        return 1

    modes = [
        ("menu", []),
        *((key, ["--plugin", key]) for key in ("larix", "vitrocad", "signal", "projectpoint")),
    ]
    env = _clean_environment()

    with tempfile.TemporaryDirectory(prefix="Larix CDE portable - ") as tmp:
        isolated_dir = Path(tmp) / "Проверка portable"
        isolated_dir.mkdir(parents=True)
        exe = isolated_dir / "Larix_CDE.exe"
        shutil.copy2(source_exe, exe)

        # The directory deliberately contains only the distributed EXE.
        extra = [item.name for item in isolated_dir.iterdir() if item.name != exe.name]
        if extra:
            print(f"[ERROR] Isolated distribution is not clean: {extra}")
            return 1

        print(f"Portable smoke path: {isolated_dir}", flush=True)
        for name, args in modes:
            try:
                return_code = _run_mode(exe, name, args, env)
            except subprocess.TimeoutExpired:
                print(f"[ERROR] {name}: timeout")
                return 1
            except OSError as exc:
                print(f"[ERROR] {name}: OS failed to start copied EXE: {exc}")
                return 1
            if return_code != 0:
                print(f"[ERROR] {name}: packaged mode failed in isolated folder")
                return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
