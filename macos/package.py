"""Build one offline Mac ZIP from a matching LibreOffice DMG and wheel directory."""
import argparse
from email.parser import BytesParser
from pathlib import Path, PurePosixPath
import stat
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile, ZipInfo


def put(archive, path, data, executable=False):
    info = ZipInfo("本地文件转换器/" + path)
    info.create_system = 3
    info.external_attr = (stat.S_IFREG | (0o755 if executable else 0o644)) << 16
    archive.writestr(info, data, compress_type=ZIP_DEFLATED)


def package(arch, dmg, wheel_directory, output):
    source = Path(__file__).resolve().parent
    repo = source.parent
    expected = set((source / "requirements.txt").read_text().lower().replace("_", "-").splitlines())
    wheels = sorted(wheel_directory.glob("*.whl"))
    with ZipFile(output, "x", compression=ZIP_DEFLATED) as archive:
        for name in ("安装转换器.command", "convert.sh", "Converter.applescript", "使用说明.txt"):
            put(archive, name, (source / name).read_text(encoding="utf-8").encode("utf-8"),
                name.endswith((".command", ".sh")))
        put(archive, "LICENSE", (repo / "LICENSE").read_bytes())
        put(archive, "arch", arch.encode())
        put(archive, "formats/batch.py", (source / "batch.py").read_bytes())
        put(archive, "formats/convert.py", (repo / "formats/convert.py").read_bytes())
        for wheel in wheels:
            platform_tag = wheel.stem.rsplit("-", 1)[1]
            if platform_tag != "any" and not platform_tag.endswith(("_universal2", "_" + arch)):
                raise ValueError("不匹配的芯片架构：" + wheel.name)
            with ZipFile(wheel) as distribution:
                meta = BytesParser().parsebytes(distribution.read(next(
                    n for n in distribution.namelist() if n.endswith(".dist-info/METADATA"))))
                expected.remove(meta["Name"].lower().replace("_", "-") + "==" + meta["Version"])
                for entry in distribution.infolist():
                    name = entry.filename
                    if name.endswith("/"):
                        continue
                    parts = PurePosixPath(name).parts
                    if ".." in parts or name.startswith("/"):
                        raise ValueError(name)
                    if parts[0].endswith(".data") and len(parts) > 2 and parts[1] in ("platlib", "purelib"):
                        name = "/".join(parts[2:])
                    put(archive, "formats/vendor/" + name, distribution.read(entry),
                        name.endswith((".so", ".dylib")))
        if expected:
            raise ValueError("缺少依赖：" + ", ".join(sorted(expected)))
        archive.write(dmg, "本地文件转换器/LibreOffice.dmg", compress_type=ZIP_STORED)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("arch", choices=("arm64", "x86_64"))
    parser.add_argument("dmg", type=Path)
    parser.add_argument("wheels", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    # Create a new file exclusively and remove partial output if packaging fails.
    if args.output.exists():
        parser.error("输出文件已存在")
    try:
        package(args.arch, args.dmg, args.wheels, args.output)
    except FileExistsError:
        raise
    except BaseException:
        args.output.unlink(missing_ok=True)
        raise
