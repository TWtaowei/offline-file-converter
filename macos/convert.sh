#!/bin/sh
set -eu
base="$(cd "$(dirname "$0")" && pwd)"
for office in "$HOME/Applications/LibreOffice.app" "/Applications/LibreOffice.app"; do
    if [ -x "$office/Contents/MacOS/soffice" ]; then
        export LOCAL_CONVERTER_OFFICE="$office"
        for python in "$office/Contents/Resources/python" "$office/Contents/MacOS/python" "$office/Contents/Frameworks/LibreOfficePython.framework/Versions/Current/bin/python3"; do
            if [ -x "$python" ]; then
                unset PYTHONHOME PYTHONPATH
                exec "$python" -s -B "$base/formats/batch.py" "$@"
            fi
        done
    fi
done
echo '请用离线包中的 DMG 将对应芯片的 LibreOffice.app 安装到「应用程序」，再打开转换器。' >&2
exit 1
