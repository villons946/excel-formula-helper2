# =====================================================================
# 公式助手 - 安装脚本
#   1. 复制 dist\公式助手.xlam 到 Excel 加载项目录
#   2. 把加载项目录加入 Excel 受信任位置（避免每次打开弹宏警告）
#   3. 注册并启用加载项
# 注意：安装前请先关闭所有 Excel 窗口。
# =====================================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'excel-common.ps1')
$xlam = Join-Path $root 'dist\公式助手.xlam'
if (-not (Test-Path -LiteralPath $xlam)) {
    Write-Error "未找到 $xlam ，请先运行 build\build.ps1"
}

$addInsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
New-Item -ItemType Directory -Force -Path $addInsDir | Out-Null
$dest = Join-Path $addInsDir '公式助手.xlam'
Copy-Item -LiteralPath $xlam -Destination $dest -Force
Write-Host "1/3 已复制加载宏 -> $dest"

# 探测 Office 版本号（16.0 = 2016/2019/2021/365）
$probe = New-Object -ComObject Excel.Application
$ver = $probe.Version
Stop-ExcelGracefully $probe '版本探测 Excel'

$tlKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations\LocationFormulaHelper"
New-Item -Path $tlKey -Force | Out-Null
Set-ItemProperty -Path $tlKey -Name 'Path' -Value ($addInsDir + '\')
Set-ItemProperty -Path $tlKey -Name 'AllowSubFolders' -Value 1 -Type DWord
Write-Host "2/3 已把 $addInsDir 加入 Excel $ver 受信任位置"

$excel = New-Object -ComObject Excel.Application
try {
    $excel.DisplayAlerts = $false
    $addin = $null
    try { $addin = $excel.AddIns.Item('公式助手') } catch {}
    if (-not $addin) { $addin = $excel.AddIns.Add($dest) }
    $addin.Installed = $true
    Write-Host '3/3 已注册并启用加载项【公式助手】'
}
finally {
    Stop-ExcelGracefully $excel '安装 Excel'
}

Write-Host ''
Write-Host '安装完成。重新打开 Excel 后，功能区首页右侧会出现【公式助手】选项卡。'
