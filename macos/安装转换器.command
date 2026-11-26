#!/bin/sh
set -eu
base="$(cd "$(dirname "$0")" && pwd)"
appdir="$HOME/Applications"
mkdir -p "$appdir"
stage="$(mktemp -d "$appdir/.local-converter.XXXXXXXX")"
trap 'rm -rf "$stage"' EXIT
app="$stage/本地文件转换器.app"
/usr/bin/osacompile -l AppleScript -o "$app" "$base/Converter.applescript"
/usr/bin/ditto "$base/formats" "$app/Contents/Resources/formats"
/bin/cp "$base/convert.sh" "$base/LICENSE" "$app/Contents/Resources/"
if [ -f "$base/arch" ]; then /bin/cp "$base/arch" "$app/Contents/Resources/"; fi
# Seal the local app after adding resources on Apple Silicon.
/usr/bin/codesign --force --sign - "$app"
destination="$appdir/本地文件转换器.app"
i=1
while [ -e "$destination" ]; do
    i=$((i + 1))
    destination="$appdir/本地文件转换器 ($i).app"
done
/bin/mv -n "$app" "$destination"
if [ -e "$app" ]; then echo "目标已存在，请重新运行。" >&2; exit 1; fi
/usr/bin/open -R "$destination"
echo "安装完成：$destination"
