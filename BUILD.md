# 从源码构建 Windows x64 版本

仓库包含桌面界面、渲染程序和格式转换脚本。完整离线包另含 LibreOffice、Python 解释器及 Python 依赖。

1. 在 Windows x64 电脑上准备 .NET Framework 4.8 和 64 位 LibreOffice。本版本使用 LibreOffice 26.8.3.1，附带 CPython 3.13。将完整的 LibreOffice 安装目录复制到仓库的 `build\LibreOffice`，确认其中有 `program\soffice.com`、`program\python.exe` 和 `program\mergedlo.dll`。
2. 在仓库根目录打开 PowerShell，编译程序并复制运行所需的项目文件。

```powershell
$ErrorActionPreference = 'Stop'
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
& $csc /nologo /target:winexe /platform:x64 /optimize+ /codepage:65001 /reference:System.Drawing.dll /reference:System.Windows.Forms.dll '/out:build\本地文件转换.exe' '源码\Converter.cs'
if ($LASTEXITCODE -ne 0) { throw '桌面程序编译失败' }
& $csc /nologo /target:winexe /platform:x64 /optimize+ /codepage:65001 '/out:build\LibreOffice\program\LocalRender.exe' '源码\RenderPdf.cs'
if ($LASTEXITCODE -ne 0) { throw '渲染程序编译失败' }
New-Item -ItemType Directory -Force 'build\formats' | Out-Null
Copy-Item -LiteralPath 'formats\convert.py' -Destination 'build\formats\convert.py'
Copy-Item -LiteralPath '本地文件转换.exe.config','LICENSE' -Destination 'build'
```

3. 用带 pip 的 64 位 CPython 3.13 准备依赖。以下命令在构建电脑上获取与内置解释器匹配的包。

```powershell
py -3.13 -m pip install --only-binary=:all: --no-compile --target 'build\formats\vendor' -r requirements.txt
```

4. 双击 `build\本地文件转换.exe`，拖入文件并选择目标格式。构建完成后，整个 `build` 目录可以复制到离线电脑运行。
