Add-Type -AssemblyName System.Windows.Forms

# Set poop icon path (must be .ico file)
$iconPath = "C:\Temp\poop.ico"
$fartPath = "C:\Temp\fart-01.wav"

if (-not (Test-Path $iconPath) -or -not (Test-Path $fartPath)) {
    [System.Windows.Forms.MessageBox]::Show("Missing poop icon or fart file!")
    exit
}

# Load native methods for RegisterHotKey
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class HotkeyForm : Form {
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int WM_HOTKEY = 0x0312;
    private string _soundFile;

    public HotkeyForm(string soundFile) {
        _soundFile = soundFile;
        this.FormBorderStyle = FormBorderStyle.None;
        this.ShowInTaskbar = false;
        this.WindowState = FormWindowState.Minimized;
        this.Opacity = 0;
        this.Load += (s, e) => this.Hide();
    }

    protected override void OnHandleCreated(EventArgs e) {
        base.OnHandleCreated(e);
        for (int i = 65; i <= 90; i++) {
            RegisterHotKey(this.Handle, i, 0, (uint)i);
        }
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY) {
            var player = new System.Media.SoundPlayer(_soundFile);
            player.Play();
        }
        base.WndProc(ref m);
    }

    protected override void OnFormClosed(FormClosedEventArgs e) {
        for (int i = 65; i <= 90; i++) {
            UnregisterHotKey(this.Handle, i);
        }
        base.OnFormClosed(e);
    }
}
"@ -ReferencedAssemblies System.Windows.Forms,System.Drawing

# Create the tray icon
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
$trayIcon.Text = "💨 Fart Trap Active"
$trayIcon.Visible = $true

# Create exit menu
$exitItem = New-Object System.Windows.Forms.MenuItem "Exit", {
    $trayIcon.Visible = $false
    $form.Close()
    Stop-Process -Id $PID
}
$contextMenu = New-Object System.Windows.Forms.ContextMenu
$contextMenu.MenuItems.Add($exitItem)
$trayIcon.ContextMenu = $contextMenu

# Run the fart trap form
$form = New-Object HotkeyForm $fartPath
[System.Windows.Forms.Application]::Run($form)