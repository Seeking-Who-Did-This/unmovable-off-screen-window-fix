#Requires -Version 5.1
<#
.SYNOPSIS
    GUI tool to move a selected window to a clicked screen position.
    Uses a full-screen transparent overlay (like Snipping Tool) to capture
    the click location without passing the click through to other programs.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# Win32 interop — window enumeration and movement only (no hook needed)
# ---------------------------------------------------------------------------
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class Win32 {

    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)] public struct RECT  { public int Left, Top, Right, Bottom; }

    // ── DPI ───────────────────────────────────────────────────────────────
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
    public static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

    [DllImport("shcore.dll")]
    public static extern int GetDpiForMonitor(IntPtr hMonitor, int dpiType, out uint dpiX, out uint dpiY);

    // ── Monitors ──────────────────────────────────────────────────────────
    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);
    public const uint MONITOR_DEFAULTTONEAREST = 2;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct MONITORINFOEX {
        public int    cbSize;
        public RECT   rcMonitor;
        public RECT   rcWork;
        public uint   dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szDevice;
    }
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

    // ── Window enumeration ────────────────────────────────────────────────
    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern int  GetWindowTextLength(IntPtr hwnd);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder sb, int cap);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

    public class WindowInfo {
        public IntPtr Handle;
        public string Title;
        public uint   PID;
    }

    public static List<WindowInfo> GetVisibleWindows() {
        var list = new List<WindowInfo>();
        EnumWindows((hwnd, _) => {
            if (IsWindowVisible(hwnd)) {
                int len = GetWindowTextLength(hwnd);
                if (len > 0) {
                    var sb = new StringBuilder(len + 1);
                    GetWindowText(hwnd, sb, sb.Capacity);
                    string title = sb.ToString().Trim();
                    if (!string.IsNullOrEmpty(title)) {
                        uint pid = 0;
                        GetWindowThreadProcessId(hwnd, out pid);
                        list.Add(new WindowInfo { Handle = hwnd, Title = title, PID = pid });
                    }
                }
            }
            return true;
        }, IntPtr.Zero);
        return list;
    }

    // ── Window movement ───────────────────────────────────────────────────
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);
    public const uint SWP_NOSIZE     = 0x0001;
    public const uint SWP_NOZORDER   = 0x0004;
    public const uint SWP_SHOWWINDOW = 0x0040;

    [StructLayout(LayoutKind.Sequential)]
    public struct WINDOWPLACEMENT {
        public int   length, flags, showCmd;
        public POINT ptMinPosition, ptMaxPosition;
        public RECT  rcNormalPosition, rcDevice;
    }
    [DllImport("user32.dll")] public static extern bool GetWindowPlacement(IntPtr hwnd, ref WINDOWPLACEMENT lpwp);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);
    public const int SW_RESTORE = 9;

    public static string MoveWindow(IntPtr hwnd, int x, int y) {
        var wp = new WINDOWPLACEMENT();
        wp.length = Marshal.SizeOf(typeof(WINDOWPLACEMENT));
        GetWindowPlacement(hwnd, ref wp);
        if (wp.showCmd == 2) {
            ShowWindow(hwnd, SW_RESTORE);
            Thread.Sleep(250);
        }
        bool ok = SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
        var pt = new POINT { X = x, Y = y };
        IntPtr hMon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
        var mi = new MONITORINFOEX();
        mi.cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
        GetMonitorInfo(hMon, ref mi);
        uint dpiX = 0, dpiY = 0;
        GetDpiForMonitor(hMon, 0, out dpiX, out dpiY);
        int scale = (int)Math.Round(dpiX / 96.0 * 100);
        if (ok)
            return string.Format("Moved to ({0}, {1})  |  Monitor: {2}  |  Scale: {3}%", x, y, mi.szDevice, scale);
        else
            return "SetWindowPos failed. Try running as Administrator for elevated windows.";
    }

    // ── Virtual screen bounds (spans all monitors) ─────────────────────────
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int nIndex);
    public const int SM_XVIRTUALSCREEN  = 76;
    public const int SM_YVIRTUALSCREEN  = 77;
    public const int SM_CXVIRTUALSCREEN = 78;
    public const int SM_CYVIRTUALSCREEN = 79;
}
'@ -ErrorAction Stop

[Win32]::SetProcessDpiAwarenessContext([Win32]::DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2) | Out-Null

# ---------------------------------------------------------------------------
# Full-screen transparent overlay — captures the click, blocks everything else
# ---------------------------------------------------------------------------
function Show-PickOverlay {
    <#
    Returns a [System.Drawing.Point] with the clicked screen coordinates,
    or $null if the user pressed Escape to cancel.
    #>

    # Span the entire virtual desktop (all monitors)
    $vx = [Win32]::GetSystemMetrics([Win32]::SM_XVIRTUALSCREEN)
    $vy = [Win32]::GetSystemMetrics([Win32]::SM_YVIRTUALSCREEN)
    $vw = [Win32]::GetSystemMetrics([Win32]::SM_CXVIRTUALSCREEN)
    $vh = [Win32]::GetSystemMetrics([Win32]::SM_CYVIRTUALSCREEN)

    $result = $null

    $overlay = New-Object System.Windows.Forms.Form
    $overlay.FormBorderStyle  = [System.Windows.Forms.FormBorderStyle]::None
    $overlay.TopMost          = $true
    $overlay.ShowInTaskbar    = $false
    $overlay.StartPosition    = [System.Windows.Forms.FormStartPosition]::Manual
    $overlay.Bounds           = New-Object System.Drawing.Rectangle($vx, $vy, $vw, $vh)
    $overlay.BackColor        = [System.Drawing.Color]::Black
    $overlay.Opacity          = 0.01        # near-invisible but still intercepts clicks
    $overlay.Cursor           = [System.Windows.Forms.Cursors]::Cross

    # ESC cancels
    $overlay.KeyPreview = $true
    $overlay.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $script:result = $null
            $overlay.Close()
        }
    })

    # Left-click captures coordinates and closes
    $overlay.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            # MouseDown coords are relative to the overlay; add virtual screen origin
            $script:result = New-Object System.Drawing.Point(
                ($vx + $e.X),
                ($vy + $e.Y)
            )
            $overlay.Close()
        }
    })

    # Draw a subtle instruction hint centred on the primary screen
    $overlay.Add_Paint({
        param($s, $e)
        $hint    = "Click to place window   •   Esc to cancel"
        $fnt     = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
        $sz      = $e.Graphics.MeasureString($hint, $fnt)

        # Centre relative to the overlay (which starts at vx,vy)
        $cx = ($vw - $sz.Width)  / 2
        $cy = ($vh - $sz.Height) / 2 - 60   # slightly above centre

        # Shadow
        $e.Graphics.DrawString($hint, $fnt,
            [System.Drawing.Brushes]::Black,
            ($cx + 2), ($cy + 2))
        # Text
        $e.Graphics.DrawString($hint, $fnt,
            [System.Drawing.Brushes]::White,
            $cx, $cy)
        $fnt.Dispose()
    })

    $overlay.ShowDialog() | Out-Null
    $overlay.Dispose()
    return $script:result
}

# ---------------------------------------------------------------------------
# Layout constants
# ---------------------------------------------------------------------------
$formW   = 540
$pad     = 30
$innerW  = $formW - ($pad * 2)
$headerH = 80
$yDrop   = $headerH + 30
$yCombo  = $yDrop   + 30
$yMove   = $yCombo  + 44
$yStatus = $yMove   + 64
$statusH = 80
$formH   = $yStatus + $statusH + 80

# ---------------------------------------------------------------------------
# Colours & fonts
# ---------------------------------------------------------------------------
$colBg      = [System.Drawing.Color]::FromArgb(30,  30,  30)
$colPanel   = [System.Drawing.Color]::FromArgb(42,  42,  42)
$colAccent  = [System.Drawing.Color]::FromArgb(0,  120, 212)
$colHover   = [System.Drawing.Color]::FromArgb(0,  100, 180)
$colWait    = [System.Drawing.Color]::FromArgb(70,  70,  70)
$colText    = [System.Drawing.Color]::FromArgb(240, 240, 240)
$colSubText = [System.Drawing.Color]::FromArgb(155, 155, 155)
$colSuccess = [System.Drawing.Color]::FromArgb(100, 210, 120)
$colError   = [System.Drawing.Color]::FromArgb(230,  80,  80)
$colBorder  = [System.Drawing.Color]::FromArgb(65,  65,  65)

$fontTitle  = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$fontCap    = New-Object System.Drawing.Font("Segoe UI",  7, [System.Drawing.FontStyle]::Bold)
$fontBody   = New-Object System.Drawing.Font("Segoe UI",  9)
$fontBtn    = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$fontStatus = New-Object System.Drawing.Font("Segoe UI",  9)

# ---------------------------------------------------------------------------
# Main form
# ---------------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Window Mover"
$form.ClientSize      = New-Object System.Drawing.Size($formW, $formH)
$form.MinimumSize     = New-Object System.Drawing.Size($formW, $formH)
$form.MaximumSize     = New-Object System.Drawing.Size($formW, $formH)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $colBg
$form.ForeColor       = $colText
$form.Font            = $fontBody
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false

# ── Header ────────────────────────────────────────────────────────────────
$pnlHeader           = New-Object System.Windows.Forms.Panel
$pnlHeader.BackColor = $colPanel
$pnlHeader.Size      = New-Object System.Drawing.Size($formW, $headerH)
$pnlHeader.Location  = New-Object System.Drawing.Point(0, 0)

$lblTitle            = New-Object System.Windows.Forms.Label
$lblTitle.Text       = "Window Mover"
$lblTitle.Font       = $fontTitle
$lblTitle.ForeColor  = $colText
$lblTitle.AutoSize   = $true
$lblTitle.Location   = New-Object System.Drawing.Point($pad, 5)

$lblSub              = New-Object System.Windows.Forms.Label
$lblSub.Text         = "Select a window, then click to place it"
$lblSub.Font         = $fontBody
$lblSub.ForeColor    = $colSubText
$lblSub.AutoSize     = $true
$lblSub.Location     = New-Object System.Drawing.Point($pad, 40)

$pnlHeader.Controls.AddRange(@($lblTitle, $lblSub))

# ── Dropdown label ────────────────────────────────────────────────────────
$lblDropCap           = New-Object System.Windows.Forms.Label
$lblDropCap.Text      = "SELECT WINDOW"
$lblDropCap.Font      = $fontCap
$lblDropCap.ForeColor = $colSubText
$lblDropCap.AutoSize  = $true
$lblDropCap.Location  = New-Object System.Drawing.Point($pad, $yDrop)

# ── Combo + Refresh ───────────────────────────────────────────────────────
$refreshW   = 38
$refreshGap = 8
$comboW     = $innerW - $refreshW - $refreshGap

$combo               = New-Object System.Windows.Forms.ComboBox
$combo.DropDownStyle = "DropDownList"
$combo.FlatStyle     = "Flat"
$combo.BackColor     = $colPanel
$combo.ForeColor     = $colText
$combo.Font          = $fontBody
$combo.Size          = New-Object System.Drawing.Size($comboW, 28)
$combo.Location      = New-Object System.Drawing.Point($pad, $yCombo)

$btnRefresh          = New-Object System.Windows.Forms.Button
$btnRefresh.Text     = "↺"
$btnRefresh.Font     = New-Object System.Drawing.Font("Segoe UI", 12)
$btnRefresh.FlatStyle = "Flat"
$btnRefresh.FlatAppearance.BorderColor        = $colBorder
$btnRefresh.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(55, 55, 55)
$btnRefresh.BackColor = $colPanel
$btnRefresh.ForeColor = $colSubText
$btnRefresh.Size      = New-Object System.Drawing.Size($refreshW, 28)
$btnRefresh.Location  = New-Object System.Drawing.Point(($pad + $comboW + $refreshGap), $yCombo)
$btnRefresh.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnRefresh.TabStop   = $false

# ── Move button ───────────────────────────────────────────────────────────
$btnMove             = New-Object System.Windows.Forms.Button
$btnMove.Text        = "Click Location to Move"
$btnMove.Font        = $fontBtn
$btnMove.FlatStyle   = "Flat"
$btnMove.FlatAppearance.BorderSize = 0
$btnMove.BackColor   = $colAccent
$btnMove.ForeColor   = [System.Drawing.Color]::White
$btnMove.Size        = New-Object System.Drawing.Size($innerW, 46)
$btnMove.Location    = New-Object System.Drawing.Point($pad, $yMove)
$btnMove.Cursor      = [System.Windows.Forms.Cursors]::Hand

# ── Status panel ──────────────────────────────────────────────────────────
$pnlStatus           = New-Object System.Windows.Forms.Panel
$pnlStatus.BackColor = $colPanel
$pnlStatus.Size      = New-Object System.Drawing.Size($innerW, $statusH)
$pnlStatus.Location  = New-Object System.Drawing.Point($pad, $yStatus)

$lblStatusCap           = New-Object System.Windows.Forms.Label
$lblStatusCap.Text      = "STATUS"
$lblStatusCap.Font      = $fontCap
$lblStatusCap.ForeColor = $colSubText
$lblStatusCap.AutoSize  = $true
$lblStatusCap.Location  = New-Object System.Drawing.Point(12, 10)

$lblStatus              = New-Object System.Windows.Forms.Label
$lblStatus.Text         = "Ready — select a window and click the button."
$lblStatus.Font         = $fontStatus
$lblStatus.ForeColor    = $colSubText
$lblStatus.AutoSize     = $false
$lblStatus.Size         = New-Object System.Drawing.Size(($innerW - 24), ($statusH - 32))
$lblStatus.Location     = New-Object System.Drawing.Point(12, 28)

$pnlStatus.Controls.AddRange(@($lblStatusCap, $lblStatus))

$form.Controls.AddRange(@(
    $pnlHeader,
    $lblDropCap, $combo, $btnRefresh,
    $btnMove,
    $pnlStatus
))

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Populate-Combo {
    $combo.BeginUpdate()
    $combo.Items.Clear()
    $raw = [Win32]::GetVisibleWindows()
    $script:windowItems = $raw | ForEach-Object {
        $n = try { (Get-Process -Id ([int]$_.PID) -EA SilentlyContinue).Name } catch { '?' }
        [PSCustomObject]@{
            Handle      = $_.Handle
            DisplayName = "$($_.Title)  [$n]"
            Title       = $_.Title
        }
    } | Sort-Object DisplayName
    foreach ($w in $script:windowItems) { $combo.Items.Add($w.DisplayName) | Out-Null }
    if ($combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 }
    $combo.EndUpdate()
}

function Set-Status($msg, $col) {
    $lblStatus.ForeColor = $col
    $lblStatus.Text      = $msg
}

function Set-Busy {
    $btnMove.Text       = "Waiting for click…"
    $btnMove.BackColor  = $colWait
    $btnMove.Enabled    = $false
    $btnRefresh.Enabled = $false
    $combo.Enabled      = $false
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Ready {
    $btnMove.Text       = "Click Location to Move"
    $btnMove.BackColor  = $colAccent
    $btnMove.Enabled    = $true
    $btnRefresh.Enabled = $true
    $combo.Enabled      = $true
}

$btnMove.Add_MouseEnter({ if ($btnMove.Enabled) { $btnMove.BackColor = $colHover } })
$btnMove.Add_MouseLeave({ if ($btnMove.Enabled) { $btnMove.BackColor = $colAccent } })

$btnRefresh.Add_Click({
    Populate-Combo
    Set-Status "Window list refreshed." $colSubText
})

# ---------------------------------------------------------------------------
# Move button — show overlay, capture click, move window
# ---------------------------------------------------------------------------
$btnMove.Add_Click({
    if ($combo.SelectedIndex -lt 0) {
        Set-Status "Please select a window first." $colError
        return
    }

    $sel = $script:windowItems[$combo.SelectedIndex]
    Set-Status "Overlay active — click anywhere to place '$($sel.Title)'.  Esc to cancel." $colSubText
    Set-Busy

    # Show the full-screen overlay (blocks until click or Esc)
    $clickPt = Show-PickOverlay

    if ($null -eq $clickPt) {
        Set-Status "Cancelled." $colSubText
    }
    else {
        try {
            $result = [Win32]::MoveWindow($sel.Handle, $clickPt.X, $clickPt.Y)
            $isErr  = $result -like "*failed*"
            Set-Status $result $(if ($isErr) { $colError } else { $colSuccess })
        }
        catch {
            Set-Status "Error: $_" $colError
        }
    }

    Set-Ready
    $form.Activate()
})

# ---------------------------------------------------------------------------
# Launch
# ---------------------------------------------------------------------------
$form.Add_Shown({ Populate-Combo })
[System.Windows.Forms.Application]::Run($form)
