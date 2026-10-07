# =====================================================================
# 公式助手 - Excel COM 进程管理公共函数（build/install/uninstall/test 共用）
#
# 背景：Excel 一旦加载了加载项，经常无视 Quit()，留下一个没有窗口的僵尸
# 进程 —— 用户看不到任何窗口，但任务管理器里 EXCEL.EXE 还在，并且会阻塞
# Windows 关机。这里在 Quit() 后轮询等待"本次自己启动的那个 PID"退出；
# 超时则只精确清理该 PID，绝不误伤用户自己的 Excel 实例。
# （与 excelutil.py 的 quit_excel 行为对齐）
# =====================================================================

Add-Type -Namespace Native -Name Win32 -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@

function Get-ComExcelPid {
    param($ExcelApp)
    try {
        $child = 0
        [Native.Win32]::GetWindowThreadProcessId([IntPtr]$ExcelApp.Hwnd, [ref]$child) | Out-Null
        return [int]$child
    } catch { return 0 }
}

function Stop-ExcelGracefully {
    param($ExcelApp, [string]$Label = 'Excel', [int]$TimeoutSec = 15)
    $ownPid = Get-ComExcelPid $ExcelApp
    try { $ExcelApp.DisplayAlerts = $false } catch {}
    try { $ExcelApp.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApp) } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()

    if ($ownPid -le 0) { return }
    for ($i = 0; $i -lt ($TimeoutSec * 2); $i++) {
        Start-Sleep -Milliseconds 500
        if (-not (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { return }
    }
    Write-Host "  $Label 未响应 Quit()，强制结束本次启动的进程 (PID $ownPid)"
    Stop-Process -Id $ownPid -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
}
