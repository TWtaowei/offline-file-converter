# 从源码构建

## Windows x64

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

## Windows x86 兼容测试版

在另一个干净的构建目录内完成以下步骤。使用 [LibreOffice Portable 25.2.7（Windows 7/8.0 版本）](https://portableapps.com/node/80166) 的 32 位引擎，把其中完整的 `App\libreoffice` 复制为 `build\LibreOffice`。把 [CPython 3.8.10 的 Windows 32 位嵌入式包](https://www.python.org/downloads/release/python-3810/) 解压到 `build\Python38`。

沿用上面的两条 C# 编译命令，将两处 `/platform:x64` 改为 `/platform:x86`，编译器路径使用 `Framework`。复制相同的配置、源码许可和转换脚本。使用下面的依赖命令代替 x64 版本的第 3 步。

```powershell
py -m pip install --no-compile --only-binary=:all: --implementation cp --python-version 3.8 --abi cp38 --platform win32 --target 'build\formats\vendor' -r requirements-win32.txt
```

Windows 7 目标环境为已更新的 SP1，需 .NET Framework 4.8 和通用 C 运行库。x86 程序、解释器、原生库及转换引擎须为同一位数。兼容版使用旧版组件，发布时应保留测试状态。

## macOS

每种芯片分别使用 LibreOffice 26.8.3 的原版 macOS DMG，其内置解释器为 CPython 3.13。构建时在同架构的 Mac 上用 CPython 3.13 下载依赖；为支持 macOS 11，请选择部署目标为 11 或更早的可用 wheel。打包脚本可在任一有 Python 3.9 或更新版本的构建电脑上运行。

```sh
python3.13 -m pip download --only-binary=:all: --no-deps -r macos/requirements.txt -d wheels
python3 macos/package.py arm64 LibreOffice.dmg wheels converter-mac-arm64.zip
```

Intel 版本使用 Intel DMG 和 x86_64 依赖，把架构参数替换为 `x86_64`。分发 ZIP 包含安装脚本、运行依赖及原版 DMG。安装脚本在目标 Mac 上生成可接收 Finder 拖放的 AppleScript 应用。构建机下载依赖后，用户安装和转换都可离线完成。
