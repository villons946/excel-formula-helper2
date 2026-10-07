# =====================================================================
# 公式助手 - 启用/禁用一键开关
#
# 用法（在 build 目录下）：
#   powershell -ExecutionPolicy Bypass -File toggle.ps1
#   powershell -ExecutionPolicy Bypass -File toggle.ps1 -NoPause   # 脚本调用不暂停
#
# 行为：
#   - Excel 未运行：直接改注册表自启动槽位（下次启动 Excel 生效）
#   - Excel 运行中：通过 COM 对本运行的实例立即 启用/卸载 插件，
#     不关闭你的表格、不碰其他加载项；Excel 退出时会自动记住该状态
#
# 只影响「公式助手」是否随 Excel 加载，不删除任何文件。
# =====================================================================
param([switch]$NoPause)

$ErrorActionPreference = 'Continue'
$addinTitle = '公式助手'
$addinPath = Join-Path $env:APPDATA 'Microsoft\AddIns\公式助手.xlam'

function Test-OurSlot([string]$value) {
    $t = "$value".Trim().Trim([char]34)
    $base = [IO.Path]::GetFileName($t)
    return ($base.StartsWith($addinTitle) -and $base -match '\.xla')
}

# ---- Office 版本枚举（有 Excel 子键的都算） ----
function Get-OfficeVersions {
    $vers = @()
    foreach ($k in (Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue)) {
        $v = $k.PSChildName
        if ($v -match '^\d+\.\d$' -and (Test-Path ("HKCU:\Software\Microsoft\Office\$v\Excel"))) { $vers += $v }
    }
    return $vers
}

# ---- 状态判定：任一版本的 OPEN 槽位指向本插件 = 已启用 ----
function Get-EnabledState {
    foreach ($ver in (Get-OfficeVersions)) {
        $opt = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
        if (-not (Test-Path $opt)) { continue }
        foreach ($p in (Get-Item $opt).Property) {
            if ($p -match '^OPEN\d*$') {
                if (Test-OurSlot ((Get-ItemProperty $opt -ErrorAction SilentlyContinue).$p)) { return $true }
            }
        }
    }
    return $false
}

# ---- 注册表禁用：删除所有版本里本插件的槽位 ----
function Disable-ByRegistry {
    foreach ($ver in (Get-OfficeVersions)) {
        $opt = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
        if (-not (Test-Path $opt)) { continue }
        foreach ($p in @((Get-Item $opt).Property)) {
            if ($p -match '^OPEN\d*$') {
                if (Test-OurSlot ((Get-ItemProperty $opt -ErrorAction SilentlyContinue).$p)) {
                    Remove-ItemProperty -Path $opt -Name $p
                    Write-Host ("  已移除自启动槽位 $p (Office $ver)")
                }
            }
        }
    }
}

# ---- 注册表启用：写第一个空槽 + 确保受信任位置 ----
function Enable-ByRegistry {
    $done = $false
    foreach ($ver in (Get-OfficeVersions)) {
        $opt = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
        if (-not (Test-Path $opt)) { continue }
        $free = $null
        for ($n = 0; $n -lt 64 -and -not $free; $n++) {
            $name = if ($n -eq 0) { 'OPEN' } else { "OPEN$n" }
            if ((Get-Item $opt).Property -notcontains $name) { $free = $name }
        }
        if ($free) {
            New-ItemProperty -Path $opt -Name $free -Value ('"' + $addinPath + '"') -PropertyType String -Force | Out-Null
            $tl = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations\LocationFormulaHelper"
            New-Item -Path $tl -Force | Out-Null
            Set-ItemProperty -Path $tl -Name 'Path' -Value ((Split-Path $addinPath) + '\')
            Set-ItemProperty -Path $tl -Name 'AllowSubFolders' -Value 1 -Type DWord
            Write-Host ("  已写入自启动槽位 $free (Office $ver)")
            $done = $true
            break
        }
    }
    if (-not $done) { Write-Host '  未找到空闲的 OPEN 槽位（自启动项过多），请手动在 Excel 选项-加载项里启用' }
}

# ---- 主流程 ----
if (-not (Test-Path -LiteralPath $addinPath)) {
    Write-Host '插件文件不存在（可能已卸载）。如需重新安装请运行 build\install.py'
    if (-not $NoPause) { Read-Host '按回车退出' | Out-Null }
    exit 1
}

$excelProc = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
$wasEnabled = Get-EnabledState
Write-Output ("当前状态: " + $(if ($wasEnabled) { '已启用' } else { '已禁用' }))

if ($excelProc.Count -gt 0) {
    # Excel 运行中：对本运行实例立即生效
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        $target = $null
        foreach ($a in $excel.AddIns) {
            if ($a.Name -like "$addinTitle*") { $target = $a; break }
        }
        if ($target) {
            $target.Installed = -not $wasEnabled
            # 注册表槽位同步到位（Excel 退出时本也会按内存状态重写，这里显式保持一致）
            if ($wasEnabled) { Disable-ByRegistry } else { Enable-ByRegistry }
            if ($wasEnabled) {
                Write-Host '已在本运行的 Excel 中立即卸载「公式助手」（功能区选项卡随即消失）。'
                Write-Host '如需恢复，再次运行本脚本。'
            } else {
                Write-Host '已在本运行的 Excel 中立即加载「公式助手」（功能区出现选项卡）。'
            }
        } else {
            Write-Host '运行中的 Excel 里没有找到本插件的注册项（可能以其他方式加载）。'
        }
    } catch {
        Write-Host ('未能连接到运行中的 Excel（' + $_.Exception.Message + '）。')
        Write-Host '已改为修改注册表：关闭所有 Excel 后状态生效。'
        if ($wasEnabled) { Disable-ByRegistry } else { Enable-ByRegistry }
    }
} else {
    # Excel 未运行：注册表直写
    if ($wasEnabled) { Disable-ByRegistry } else { Enable-ByRegistry }
}

$newState = Get-EnabledState
Write-Output ("操作后状态: " + $(if ($newState) { '已启用' } else { '已禁用' }) +
    $(if ($excelProc.Count -gt 0) { '（运行中的 Excel 已即时生效）' } else { '（下次启动 Excel 生效）' }))
Write-Output '再次运行本脚本可切回。'

if (-not $NoPause) { Read-Host '按回车退出' | Out-Null }
