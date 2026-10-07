# =====================================================================
# 公式助手 - 卸载脚本
# 注意：卸载前请先关闭所有 Excel 窗口。
# =====================================================================
$ErrorActionPreference = 'Continue'
$name = '公式助手'
. (Join-Path $PSScriptRoot 'excel-common.ps1')

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
try {
    foreach ($a in $excel.AddIns) {
        if ($a.Name -eq $name -or $a.Title -eq $name) {
            $a.Installed = $false
            Write-Host "已停用加载项：$name"
        }
    }
    try {
        $excel.AddIns.Remove($name) | Out-Null
        Write-Host "已移除加载项注册：$name"
    } catch {
        Write-Host "（未找到需要移除的注册项，或需手动在 Excel 加载项列表中取消勾选）"
    }
}
finally {
    Stop-ExcelGracefully $excel '卸载 Excel'
}

$dest = Join-Path $env:APPDATA 'Microsoft\AddIns\公式助手.xlam'
Remove-Item -LiteralPath $dest -Force -ErrorAction SilentlyContinue
if (-not (Test-Path -LiteralPath $dest)) { Write-Host '已删除加载宏文件' } else { Write-Host "文件删除失败（可能被占用）：$dest" }

# 兼容各 Office 版本（14.0/15.0/16.0）逐一清理受信任位置
foreach ($ver in @('14.0', '15.0', '16.0')) {
    $tlKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations\LocationFormulaHelper"
    Remove-Item -Path $tlKey -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host '卸载完成。'
