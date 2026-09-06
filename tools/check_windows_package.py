"""解压发行 ZIP 到新目录，移除 Python 环境路径后运行冻结程序自检。"""

from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile


def main() -> None:
    archive = Path(sys.argv[1]).resolve()
    root = Path(__file__).resolve().parents[1]
    folder = root / "outputs" / "EXE 验证" / datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    folder.mkdir(parents=True)
    with ZipFile(archive) as zipped:
        for name in zipped.namelist():
            if not (folder / name).resolve().is_relative_to(folder):
                raise ValueError("ZIP 包含越界路径")
        zipped.extractall(folder)
    environment = {key: value for key, value in os.environ.items()
                   if not key.upper().startswith(("PYTHON", "QT_", "PYSIDE"))
                   and key.upper() not in ("VIRTUAL_ENV", "MPLCONFIGDIR", "MPLBACKEND")}
    windows = Path(os.environ["SystemRoot"])
    environment["PATH"] = os.pathsep.join((str(windows / "System32"), str(windows)))
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = subprocess.SW_HIDE
    result_folder = folder / "check"
    process = subprocess.run([str(folder / "ExcelPPT" / "ExcelPPT.exe"), "--self-test", str(result_folder)],
                             cwd=folder, env=environment, timeout=60, startupinfo=startup,
                             capture_output=True)
    result_file = result_folder / "result.json"
    if not result_file.exists():
        raise RuntimeError(f"冻结程序未生成自检结果，退出码 {process.returncode}，错误输出：{process.stderr!r}")
    result = json.loads(result_file.read_text(encoding="utf-8"))
    result["archive"] = str(archive)
    with archive.open("rb") as stream:
        result["sha256"] = hashlib.file_digest(stream, "sha256").hexdigest()
    result["output_directory"] = str(result_folder)
    result["exit_code"] = process.returncode
    (folder / "verification.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if process.returncode or result.get("status") != "passed" or not result.get("frozen"):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
