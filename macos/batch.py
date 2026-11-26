"""Finder batch entry point. The selected LibreOffice supplies the interpreter."""
import ctypes
import errno
import os
import platform
import sys
import tempfile
from pathlib import Path

arch_file = Path(__file__).resolve().parent.parent / "arch"
if arch_file.is_file() and (arch_file.read_text().strip() != platform.machine()
                            or sys.version_info[:2] != (3, 13)):
    sys.exit("转换器与本机 LibreOffice 版本或芯片不匹配。请使用同一离线包中的安装文件。")

from convert import CORE, IMAGES, OUTPUT, convert, ext

MODES = ("pdf", "docx", "xlsx", "pptx", "docx-layout", "pptx-edit", "txt", "csv", "tsv",
         "odt", "ods", "odp", "rtf", "doc", "xls", "ppt", "png", "jpg", "webp", "bmp")


def supports(source, mode):
    kind = ext(source)
    if kind not in CORE and kind not in IMAGES and kind != "pdf":
        return False
    if mode == "docx-layout":
        return kind == "pdf"
    if mode == "pptx-edit":
        return kind not in IMAGES and CORE.get(kind) != "pptx"
    if kind in IMAGES and mode in ("txt", "csv", "tsv"):
        return False
    return mode != ("jpg" if kind == "jpeg" else kind) or mode in ("png", "jpg")


def publish(result, directory, stem, extension):
    # Darwin's exclusive rename protects both files and multi-page output folders.
    rename = ctypes.CDLL(None, use_errno=True).renamex_np
    rename.argtypes = (ctypes.c_char_p, ctypes.c_char_p, ctypes.c_uint)
    rename.restype = ctypes.c_int
    is_folder = result.is_dir()
    name = stem + ("_" + extension.upper() if is_folder else "")
    suffix = "" if is_folder else "." + extension
    index = 1
    while True:
        serial = "" if index == 1 else " (" + str(index) + ")"
        target = directory / (name + serial + suffix)
        if rename(os.fsencode(result), os.fsencode(target), 4) == 0:  # RENAME_EXCL
            return target
        error = ctypes.get_errno()
        if error not in (errno.EEXIST, errno.ENOTEMPTY):
            raise OSError(error, os.strerror(error), target)
        index += 1


def run(source, mode):
    if not source.is_file() or not supports(source, mode):
        raise ValueError("请选择支持的本地文件。")
    if getattr(source.stat(), "st_flags", 0) & 0x40000000:  # SF_DATALESS
        raise ValueError("请先将文件完整保存在本地。")
    fmt = OUTPUT.get(mode, mode)
    with tempfile.TemporaryDirectory(prefix=".local-convert-", dir=source.parent) as stage:
        with tempfile.TemporaryDirectory(prefix="local-convert-") as tmp:
            target = Path(stage) / (source.stem + "." + fmt)
            convert(source, mode, target, Path(tmp))
            pages = Path(stage) / "multiple"
            return publish(pages if pages.is_dir() else target, source.parent, source.stem, fmt)


if __name__ == "__main__":
    action = sys.argv[1]
    files = [Path(value).absolute() for value in sys.argv[2:]]
    if action == "--targets":
        print("\n".join(mode for mode in MODES if files and all(supports(f, mode) for f in files)))
    elif action in MODES:
        failures = []
        for source in files:
            try:
                run(source, action)
            except Exception as exc:
                failures.append(f"{source.name} — {exc}")
        output = sys.stderr if failures else sys.stdout
        print(f"完成 {len(files) - len(failures)} / {len(files)} 个文件。结果保存在各原文件旁。", file=output)
        if failures:
            print("\n".join(failures), file=output)
            sys.exit(1)
    else:
        sys.exit("未知的目标格式。")
