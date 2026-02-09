<# :
    @echo off & Title MSI Tools
    if exist %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe   set "powershell=%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
    if exist %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe  set "powershell=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
    set args=%*
    if defined args set "args=%args:"=\"%"
    %powershell% -NoLogo -NoProfile -Ex Bypass -Window Hidden -Command ^
        "$batFile='%~f0'; $sb=[ScriptBlock]::Create([IO.File]::ReadAllText('%~f0'));& $sb @args" %args% -scriptPath '%~f0'
    exit /b
#>

param($scriptPath)

$script:Version = [version]"1.0"

# ── Taskbar identity : persistent icon + AppUserModelID + shortcut helper ──
$script:AppId       = "MSI-Tools.CustomTitleBar.1"
$script:AppDataDir  = [System.IO.Path]::Combine($env:APPDATA, "MSI_Tools")
$script:IcoPath     = [System.IO.Path]::Combine($script:AppDataDir, "MSI_Tools.ico")
$script:LnkName     = "MSI Tools.lnk"
$script:TaskbarPinDir  = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
$script:StartMenuDir   = [Environment]::GetFolderPath("Programs")
if (!([System.IO.Directory]::Exists($script:AppDataDir))) {
    [System.IO.Directory]::CreateDirectory($script:AppDataDir) | Out-Null
}
Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
public static class TaskBarHelper
{
    [DllImport("shell32.dll", SetLastError = true)]
    private static extern void SetCurrentProcessExplicitAppUserModelID(
        [MarshalAs(UnmanagedType.LPWStr)] string AppID
    );
    public static void SetAppId(string id)
    {
        SetCurrentProcessExplicitAppUserModelID(id);
    }
}
public static class ShortcutHelper
{
    [ComImport]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellLink
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, int fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
        void Resolve(IntPtr hwnd, int fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }
    [ComImport]
    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyStore
    {
        int GetCount(out uint cProps);
        int GetAt(uint iProp, out PROPERTYKEY pkey);
        int GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        int SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
        int Commit();
    }
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    private struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct PROPVARIANT
    {
        public ushort vt;
        public ushort wReserved1;
        public ushort wReserved2;
        public ushort wReserved3;
        public IntPtr pwszVal;
        public static PROPVARIANT FromString(string val)
        {
            var pv = new PROPVARIANT();
            pv.vt = 31; // VT_LPWSTR
            pv.pwszVal = Marshal.StringToCoTaskMemUni(val);
            return pv;
        }
    }
    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    private class ShellLink { }
    private static readonly PROPERTYKEY PKEY_AppUserModel_ID = new PROPERTYKEY
    {
        fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"),
        pid = 5
    };
    public static void CreateShortcut(string lnkPath, string targetPath, string arguments,
                                       string iconPath, string appId, string description)
    {
        CreateShortcut(lnkPath, targetPath, arguments, iconPath, appId, description, null);
    }
    public static void CreateShortcut(string lnkPath, string targetPath, string arguments,
                                       string iconPath, string appId, string description,
                                       string workingDirectory)
    {
        var link = (IShellLink)new ShellLink();
        link.SetPath(targetPath);
        link.SetArguments(arguments);
        link.SetIconLocation(iconPath, 0);
        link.SetDescription(description);
        if (!string.IsNullOrEmpty(workingDirectory))
            link.SetWorkingDirectory(workingDirectory);
        var store = (IPropertyStore)link;
        var pv = PROPVARIANT.FromString(appId);
        PROPERTYKEY key = PKEY_AppUserModel_ID;
        store.SetValue(ref key, ref pv);
        store.Commit();
        Marshal.FreeCoTaskMem(pv.pwszVal);
        var file = (IPersistFile)link;
        file.Save(lnkPath, true);
    }
}
'@
[TaskBarHelper]::SetAppId($script:AppId)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$currentPrincipal = [Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin          = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# ── App icon from base64
$iconBase64 = "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAIAAADYYG7QAAAABnRSTlMA/wAAAP+JwC+QAAANJklEQVR4nM1Za4xdV3X+vrX3OffemfG8PL62xx57/BjbsXFooVGBtmoRjwJVUEkhiFaFUiLalKaloDaohYLKsxKiFJpQCq2IqhTaqm8eBYSgChJFKSEE20nGHnue8bycGU9sz9xzzl5ff9zYTBzbaQNVWTo/jo7OPuu73/rW2muvS0H4YTL7/wZwucXvZzHBS/c/KKafPkPr0QAIMf++wQBPmyGC/UP7AkUCICxC4DjX83QJ8f+KvKfPUCAiAxFAmleUg7X1aGqdvVyHjODFyzJY+/7Jn70qQ0+pD0keRDeXRLpSUOvSWjJu2NjExmZXR8/4Q/dGsre5O++olUXZ2dXRahWrF9Y6G/mLDrzoy1/88hP8XtHZJTSbtu23yPnJMam89CbBZz3v+dPTj4BUEukOGgEL8+MPAkCIzcG9pCSQEK30Klh9ceLIuZWlzg097l7vaaZzSw6/DMAVAF1Cs3n3IXoy4ELSyswp9wJAMAwMH6ZXLnfRIIm0yhnooEGgXAQJGeFygkvLZx9bnMxjLri7slptYOvehemHr+D9aoCGRq5vFa0AuDthcs8ajVR5SmUbsEMGyOVGJjidZKRJFBNEM0ISQmttbXnupEtGOBSY7Tx4w8Sxb14xOFcWNUPWaq22f52CJUKGVlV4aonu8CSPFoukhenjK4uzg0PbfvLHb1iZn5mfGZ1/5GSsNQIhVQ5KaX52zCGBAvfse+amHfuK1aXLI3VR41cWtVIZQp7KAgAEQoiUAzCHgsWlmYm1tRV4ZZaLSvJoIckpuHswsxA27zwkb52ZORW9spDJ5fLxE0e27Lr+9MkHLkPTs3lXrZFxnFdgqI399KmjNBMCRAPpMkfI8r7ugUfGvtNqrbjgITi9kgySBEAyFxxoVWlu8ijJXXuuTw4JTjAGCAj8xde+/pK7t739XZuHr8vrdVq8OkMQQZcrgUQJwrk0O12tnqWZp8Ro2cD7ceMhIt16454/velQ9vrPomaBWTq20Dy8ffojzzWxdP3r5/8hhCh5qlLMYlWU9b7NX/063/YH7/7j9/4hyb5P7YtmM8ePbNu3366W9m0aB3bsIymH05cnjxepEiwarOs9vOmwE2bB5fjCkWLu9/NbvkxDcKTMQh6zzr6ez7ymrFqvuumVd3zoA6Cq5AEgTUCMoWfr3rzRUNkKIX/gW1/r69owOPKMxcnjV61D2/YeLtYKMgnxzMxoWbaMdqFIG97yVbQS2sq2yJBZVj32Zy/uvPUrAFjvJAwN1rua9U+8lI6F6VF3h+BQW2Exs1SWHmI9RAeqqgohANi887r56YeunGW79gxXZYtRTkusUlU5gyfvPPBRrFVS8iQwmAFMqcUNt32VIWO9AYseQ6z3dn34eXBYzE5OjAsQE8WLNR1illloJU+pgtGhN7/1nZ68Ue/i+vp7CVCt0dvTbMK1VlTT4w9tqAVPVKRWU6PvfdWrn0FFMRgdeV2Qxcxi7sEIIe/MP/IjGwb2ycLc+LGKCk6ZS4C7ezLECinAANAAUGR331C90bU0N/o4IIKdfVu7enqNMUa2ihJIABYnR4uyCgY+82Nx9tFzs2/PPDUOfLT4qf0mRzCG6JazllueKSnfe13X7Vu9dyvZOD3xHUtqofq9t/zu4Wc/+5Zffp2jImg0eZIhiQEEWG/kXZt2X1ieLc6vsJ1QjZ5N3d39pEAKkOAQlC2cPKIgqLTXfx6MmHxs+bOvajSs8YbPu5m5KYbQqAtRAfVDz8neMRJjKC6snZ0bl/Guv/m7b9zzH1WqQszPl8Xdn/w44UmQEMxUyQhG69u2N1ptYeqooMfDtGX3QU+J7X2yDSZgYWIsVatijFvfjRcfQsxJmnvKDMkQgq2tqbvHQoQxG97f/eGfSZUefWQ0VaW7/9Zbbz+/shRD9Eqf/qe/Pbe8qIp5FgXkjY7kyYvWll3XF61VGM9MjbZjFQEMDo+4J8PF3RkwEO433vRaIiuLlv3sj7pXSEkwmVlFZ7SqQi0HSSmeHq394xsq5/L0w1Uq3TWwZdsrfv7lRExe/MvnvrC6tOyeYhYHdowARiEYklCUqwjx0cmH1iV4m6HhA+1iK5jkRhfskYkH2+DCr35WySmBhhgYontC5dbIQ8g673lfdnZKKS1OH3elsih7Bzb/0mtuLhxwfPO/7j1y371mFiwO7NgPOUQGpxsAWViYOIZ1LZcByMwISzCaSckMEi2vBTAIRUoqWvBChCRl0Z0GhpjlAbU/f1E4c0oJZ2bGqrIqE3bu2vvqV99clFCqJqZmj97/bQkhrzeHRro6uu7+5MdoggdJ7lyYOCZofS00QaX76fFji9MnFQNAwUUECy6vpJYTj1b43JGfeNbW5TtfypVVGGiZNvTGO1+cb9nFWP/oB9/jKcFw0ytvftnPvUxCUrW8fP4rX/znpDLPsoHB3aFWO370Gy94wU+HGEnJaOb9O68zZusrzuV1qLN7sLO3A86QNWbG7idRJTdACA4Zlb3xSxazvLmj8cFnsH8o5h3jD91bi1kF/Notvy6jXEJyt7v+8uM0t9DYtG2PmBanTsjl7ttHDldV1VYrAdDOL82ef2zpe6K+tKEC4Ao7+/a7wb2sPIEWQqD7fQvlDc+9A684bPV63POcrg/s8uZw2UqnH74vy2KSfue33+w5vRDhq63y0399F+hm2cCOfZAThKHdsIomytotEARTvX9zUVRsUdCTtg6CkAlIxc2/8qYgeJXstq/dcPuXwksO2oXQs21fz0du8EbPmYmxs4+cyLJYlunWN9127vyqWsmUzl8oP3P3XUpViLXm9v1GgZYc/b3bKAIgZYiAKNAlByv1NHd8T9RPMMEQaTRm/3b3X0CgDA+eRjRT3nHPh+MnXphUPDr5cJHWEpQcv3Hbb5ZF4YCkyuLff/pTkGd5PrBtj8yVAMFdc/PjAgCrigKsSMiSYEEOamnq6FUAAYnJ3Z2+ZfezLECmgy/cE1Js3PX8jrOjVeFnJkY9KbN43/3HbnnjrUWrcncvW4jZXZ/8WEplCHHT0D4jCDODpHo092TGTYN73QwyySiDwWkhhHTR+xUaNC9LBKPbWrHCmJVF8fXbf2xoYDBs3Etqc//G+UlP8Hf80QdmTp2EGU1Jrpj91SfuBJRltYFt+6AgJkoukpidOp7F7EJRLC1NDTR2OUmXAUIyRIvrRLw+ywg0d1yXmAgBZnJZvry42OjuzvM8Ivvut76ysb+3SukXXvW67u46SToUfHBw+L3veSe8iiEbGDoAg7mDgsyNtGz++HdkrKSh3YfglRuVfHH6BJgHU0rlEwrjJWsOH5QSXaAAVbSUUmdfT56HMtnUqW9v7O93cXjk0NraY2OTU0pI1OD2Xe9/9zvpHkO2aegA6XSHIJkHRAuzJx5wo6BGo1tQghFanD4hSGpVqXhCYVxfhBqNTpEgmQKdcAfdEhbHRxdP3ke53OWYmz61cm7VXAgaGGi+/13vgDyEMLD9AOkUjACRPNVqnf95z78nVU4Fs42bBuFJ8pg1niyVx5Fc6oceJ2nn/nYwCRchxTOTD5Wq0D6kAkkI8BCy7buvu/HlL7njTz6ESpbZwM79dIJOI0SldO7C6vLCZCQA9vQ2855+yo0U2JnH8bEjVwR0uaiFzJCkCiSSG1WWRYCVqYhZ7fTc/JbmxiopuUMiJPdQi82d+5IkeICVLQ95KM6fP784BUJUoG0aOkAkp0ley7KJsaNPwdB6kjYPjSQSEoGF2Yly7TwdQwduSNWahWxq9H5SRjoYAMvz5uCIkACEvD439sDZ86udjYwwORJUa2zo37KdDqMnCwSBMtJOjx+/4vniCVtHG9Pc1HEAW3cdbJXVhcdWSLNoRm+5mxdVq1XvaFSAV1Wto6O5fQTu81NjVbnmlSxkziQnQAu2efjQwObtLiJITnOAEkNx9QnWE7Ks3Qm0L2a1ID89P0vS3WdOPBAtOzM1WutouBDkPf0bN20bAbAwO7FWXCBpedZSq57XOnq3NXcfaA7v91SgPakBRYoiBDBKIWZPEbLLrG9wb6QtzhxP7ka4y4gkmQUBIwcPnz1X0QShapUdXV0hs3ItpVSAoFwIwZJEEMkDzYNUOYKxkl9YObt6dv7KR8JrDACbO/aXHmbG7q/FAFog5HIqWty8Y38FBcDJAD5+DiSABJoqEObmi5Oj9TxrFWXf9pGQBSYCOjN3MrWKq/m9FqC+jTtiZ4MMCxMPrhVreVbr6R3s3LixahVCCrFmEEKW0ipkZkoOIxMQ5AuTY4JjnTS7+4d6mgMzo99uV82rOb0WIILdzeGslsdgZXK6CHqwzDjXHt1dJFIg4aALNLG1VizPjz95AtS+ufZQ9lpTWEEr8+OhKhZOT8VQ6+zecPbMbGtpeW78wUvaBxAbNbrEBA+WTLCzC6ef7HX9kmvYU8ypBeE0CM6Pf7f9pFg7e9k73koMAiIImCjZxXHs07Brhex/+gmQoACEWqCQiqSnZuL/ENAP1n7o/g36bzEsvIJNquS3AAAAAElFTkSuQmCC"
$iconBytes  = [System.Convert]::FromBase64String($iconBase64)
$iconStream = [System.IO.MemoryStream]::new($iconBytes)
$iconImage  = [System.Drawing.Image]::FromStream($iconStream)

$loadingForm = New-Object System.Windows.Forms.Form -Property @{Text="MSI Properties Viewer";Size=[System.Drawing.Size]::new(300, 160);StartPosition="CenterScreen";FormBorderStyle="FixedDialog";ControlBox=$false;Cursor=[System.Windows.Forms.Cursors]::WaitCursor}
$loadingForm.Show()
$loadingIcon = New-Object System.Windows.Forms.PictureBox -Property @{Location=[System.Drawing.Point]::new(118, 10);Size=[System.Drawing.Size]::new(48, 48);SizeMode=[System.Windows.Forms.PictureBoxSizeMode]::Zoom}
$loadingIcon.Image = $iconImage
$loadingForm.Controls.Add($loadingIcon)
$loadingLabel = New-Object System.Windows.Forms.Label -Property @{Location=[System.Drawing.Point]::new(10, 68);Size=[System.Drawing.Size]::new(260, 20);Text="Loading interface...";TextAlign=[System.Drawing.ContentAlignment]::MiddleCenter}
$loadingForm.Controls.Add($loadingLabel)
[System.Windows.Forms.Application]::DoEvents()
$launch_progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{Location=[System.Drawing.Point]::new(10, 92);Size=[System.Drawing.Size]::new(260, 20);Style="Continuous";Value=10}
$loadingForm.Controls.Add($launch_progressBar)
[System.Windows.Forms.Application]::DoEvents()

$script:LogDir  = [System.IO.Path]::Combine($env:TEMP, "MSI_Tools")
if (!([System.IO.Directory]::Exists($script:LogDir))) { [System.IO.Directory]::CreateDirectory($script:LogDir) | Out-Null }
$script:LogFile = [System.IO.Path]::Combine($script:LogDir, "MSI_Tools_$(Get-Date -Format 'yyyyMMdd').log")
if ([System.IO.File]::Exists($script:LogFile)) {
    [System.IO.File]::AppendAllText($script:LogFile, "`n`n`n------------------------------`n`n`n")
}
$logFiles = [System.IO.Directory]::GetFiles($script:LogDir, "*.log")
if ($logFiles.Count -gt 10) {
    $sorted = [System.Array]::CreateInstance([System.IO.FileInfo], $logFiles.Count)
    for ($i = 0; $i -lt $logFiles.Count; $i++) { $sorted[$i] = [System.IO.FileInfo]::new($logFiles[$i]) }
    [System.Array]::Sort($sorted, [System.Comparison[System.IO.FileInfo]]{ param($a, $b) $b.LastWriteTimeUtc.CompareTo($a.LastWriteTimeUtc) })
    for ($i = 10; $i -lt $sorted.Count; $i++) { [System.IO.File]::Delete($sorted[$i].FullName) }
}

function Write-Log {
    param([string]$Message, [ValidateSet('Info', 'Warning', 'Error', 'Debug')][string]$Level = 'Info')
    if ([string]::IsNullOrWhiteSpace($script:LogFile)) { return }
    $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    switch ($Level) {
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Debug'   { Write-Host $logMessage -ForegroundColor Gray }
        default   { Write-Host $logMessage -ForegroundColor White }
    }
    try { [System.IO.File]::AppendAllText($script:LogFile, "$logMessage`r`n") }
    catch { Write-Host "Failed to write to log : $_" -ForegroundColor Red }
}

Write-Log ("=" * 64)
Write-Log "MSI Tools started"
Write-Log "Version : $($script:Version)" ; Write-Log "PowerShell Version : $($PSVersionTable.PSVersion)"
Write-Log "User : $env:USERNAME"         ; Write-Log "Computer : $env:COMPUTERNAME"
Write-Log ("=" * 64)

if (-not $isAdmin) { Write-Log "Running without administrator privileges" -Level Warning }

$script:resizePending = $false
$script:lastWidth = @{}
$script:currentTabIndex = 0
$script:WinRMChecked = $false
$script:TrustedHostsCache = $null
$script:TrustedHostsToRemove = [System.Collections.ArrayList]::new()
$script:RemoteConnectionResults = @()
$script:SecurePassword = New-Object System.Security.SecureString
$script:PasswordPlaceholder = "Password"
$script:IsDomainJoined = -not [string]::IsNullOrEmpty($env:USERDNSDOMAIN)
$script:stopRequested = $false
$script:sortColumn = -1
$script:sortOrder = [System.Windows.Forms.SortOrder]::Ascending
$script:MSILoaded = $false
$script:RegistryCache = $null
$script:RegistryStopRequested = $false
$script:RegistrySortColumn = -1
$script:RegistrySortOrder = [System.Windows.Forms.SortOrder]::None
$script:fromBrowseButton = $false
$script:SmbSessionsToCleanup = [System.Collections.ArrayList]::new()
$script:WindowsInstallerCOM = $null
$script:MsiFileCache = [ordered]@{}  # Key=filePath, Value=@{ FileName; Results; SelectedListView }
$script:UserPinnedStartMenu = $false

Add-Type -ReferencedAssemblies System.Windows.Forms.dll, System.Drawing.dll -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Collections;
using System.Runtime.InteropServices;
public class CustomForm : Form
{
    private const int WM_NCHITTEST = 0x84;
    private const int HTLEFT = 10;
    private const int HTRIGHT = 11;
    private const int HTTOP = 12;
    private const int HTTOPLEFT = 13;
    private const int HTTOPRIGHT = 14;
    private const int HTBOTTOM = 15;
    private const int HTBOTTOMLEFT = 16;
    private const int HTBOTTOMRIGHT = 17;
    private const int GRIP = 6;
    public event EventHandler<Message> OnWindowMessage;
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_NCHITTEST)
        {
            base.WndProc(ref m);
            var pos = this.PointToClient(new System.Drawing.Point(
                (int)m.LParam & 0xFFFF,
                (int)m.LParam >> 16 & 0xFFFF
            ));
            int w = this.ClientSize.Width;
            int h = this.ClientSize.Height;
            bool top    = pos.Y <= GRIP;
            bool bottom = pos.Y >= h - GRIP;
            bool left   = pos.X <= GRIP;
            bool right  = pos.X >= w - GRIP;
            if      (top && left)   m.Result = (IntPtr)HTTOPLEFT;
            else if (top && right)  m.Result = (IntPtr)HTTOPRIGHT;
            else if (bottom && left)  m.Result = (IntPtr)HTBOTTOMLEFT;
            else if (bottom && right) m.Result = (IntPtr)HTBOTTOMRIGHT;
            else if (top)    m.Result = (IntPtr)HTTOP;
            else if (bottom) m.Result = (IntPtr)HTBOTTOM;
            else if (left)   m.Result = (IntPtr)HTLEFT;
            else if (right)  m.Result = (IntPtr)HTRIGHT;
            if (OnWindowMessage != null)
                OnWindowMessage(this, m);
            return;
        }
        base.WndProc(ref m);
        if (OnWindowMessage != null)
            OnWindowMessage(this, m);
    }
}
public class DragDropFix {
    [DllImport("shell32.dll")]
    public static extern void DragAcceptFiles(IntPtr hwnd, bool accept);
    [DllImport("shell32.dll")]
    public static extern uint DragQueryFile(IntPtr hDrop, uint iFile, [Out] System.Text.StringBuilder lpszFile, uint cch);
    [DllImport("shell32.dll")]
    public static extern void DragFinish(IntPtr hDrop);
    [DllImport("user32.dll")]
    public static extern bool ChangeWindowMessageFilterEx(IntPtr hwnd, uint msg, uint action, IntPtr pChangeFilterStruct);
    public static void Enable(IntPtr hwnd) {
        ChangeWindowMessageFilterEx(hwnd, 0x0233, 1, IntPtr.Zero); // WM_DROPFILES
        ChangeWindowMessageFilterEx(hwnd, 0x004A, 1, IntPtr.Zero); // WM_COPYDATA
        ChangeWindowMessageFilterEx(hwnd, 0x0049, 1, IntPtr.Zero); // WM_COPYGLOBALDATA
        DragAcceptFiles(hwnd, true);
    }
    public static string[] GetDroppedFiles(IntPtr hDrop) {
        uint count = DragQueryFile(hDrop, 0xFFFFFFFF, null, 0);
        string[] files = new string[count];
        for (uint i = 0; i < count; i++) {
            uint size = DragQueryFile(hDrop, i, null, 0) + 1;
            var sb = new System.Text.StringBuilder((int)size);
            DragQueryFile(hDrop, i, sb, size);
            files[i] = sb.ToString();
        }
        DragFinish(hDrop);
        return files;
    }
}
public class ListViewItemComparer : IComparer
{
    private int col;
    private SortOrder order;
    public ListViewItemComparer()
    {
        col = 0;
        order = SortOrder.Ascending;
    }
    public ListViewItemComparer(int column, SortOrder sortOrder)
    {
        col = column;
        order = sortOrder;
    }
    public int Compare(object x, object y)
    {
        int returnVal = -1;
        returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text,
                                   ((ListViewItem)y).SubItems[col].Text);
        if (order == SortOrder.Descending)
            returnVal *= -1;
        return returnVal;
    }
}
public class NativeMethods
{
    public const int SB_HORZ = 0; // Scroll horizontal
    public const int SB_VERT = 1;  // Scroll vertical
    public const int SW_HIDE = 0;
    public const int SW_SHOW = 5;
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetScrollPos(IntPtr hWnd, int nBar);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int SetScrollPos(IntPtr hWnd, int nBar, int nPos, bool bRedraw);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);
    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] public static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    public const int WM_VSCROLL = 0x0115;
    public const int WM_HSCROLL = 0x0114;
    public const int SB_THUMBPOSITION = 4;
    public const int SB_TOP = 6;
}
[StructLayout(LayoutKind.Sequential)]
public struct WTS_SESSION_INFO {
    public int SessionId;
    [MarshalAs(UnmanagedType.LPWStr)]
    public string pWinStationName;
    public int State;
}
public static class WtsApi32 {
    [DllImport("kernel32.dll")]
    public static extern int WTSGetActiveConsoleSessionId();
    [DllImport("wtsapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr WTSOpenServer(string pServerName);
    [DllImport("wtsapi32.dll")]
    public static extern void WTSCloseServer(IntPtr hServer);
    [DllImport("wtsapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool WTSQuerySessionInformation(
        IntPtr hServer,
        int sessionId,
        int wtsInfoClass,
        out IntPtr ppBuffer,
        out int pBytesReturned
    );
    [DllImport("wtsapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool WTSEnumerateSessions(
        IntPtr hServer,
        int Reserved,
        int Version,
        out IntPtr ppSessionInfo,
        out int pCount
    );
    [DllImport("wtsapi32.dll")]
    public static extern void WTSFreeMemory(IntPtr pMemory);
}
"@ -Language CSharp

#region Common UI

# Function that quick-create majority of controls
function New-Control {
    param(
        [Parameter(Mandatory = $false)][AllowNull()]     $container,
        [Parameter(Mandatory = $true)] [string]          $type,
        [Parameter(ValueFromRemainingArguments = $true)] $args
    )
    $text = $x = $y = $width = $height = $null
    $additionalProps = [System.Collections.Generic.List[string]]::new()
    if ($args -and $args.Count -gt 0) {
        $i = 0
        if ($args[$i] -is [string] -and $args[$i] -notmatch '^\w+=') { $text = $args[$i]  ;  $i++ }             # Check if first arg is text
        if (($args.Count - $i) -ge 4) {                                                                         # Check for 4 consecutive integers (x, y, width, height)
            $potentialNums = $args[$i..($i + 3)]
            $allInts       = $true
            foreach ($n in $potentialNums) { if ($n -isnot [int] -and -not ($n -is [string] -and $n -match '^\d+$')) { $allInts = $false  ;  break } }
            if      ($allInts)             { $x=[int]$potentialNums[0]  ;  $y=[int]$potentialNums[1]  ;  $width=[int]$potentialNums[2]  ;  $height=[int]$potentialNums[3]  ;  $i+=4 }
        }
        while ($i -lt $args.Count) { $additionalProps.Add($args[$i])  ;  $i++ }                                 # Remaining args are additional properties
    }
    $control = New-Object "System.Windows.Forms.$type"
    if ($null -ne $text -and $control.PSObject.Properties.Name -contains 'Text') { $control.Text = $text }      # Set text if provided and control supports it
    if ($null -ne $x) {                                                                                         # Set location and size if provided
        if ($control.PSObject.Properties.Name -contains 'Location') { $control.Location = [System.Drawing.Point]::new($x, $y) }
        if ($control.PSObject.Properties.Name -contains 'Size')     { $control.Size     = [System.Drawing.Size]::new($width, $height) }
    }
    if ($control -is [System.Windows.Forms.TextBox]) { $control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle }
    foreach ($prop in $additionalProps) {                                                                       # Process additional properties
        $eqIndex = $prop.IndexOf('=')
        if ($eqIndex -lt 1) { continue }
        $propName  = $prop.Substring(0, $eqIndex).Trim()
        $propValue = $prop.Substring($eqIndex + 1).Trim()
        if ($propValue -match '^New-Object|^\@\{|^\[.*\]::') {
            $control.$propName = Invoke-Expression $propValue
            continue
        }
        $enumMap = @{
            Dock          = [System.Windows.Forms.DockStyle]
            Orientation   = [System.Windows.Forms.Orientation]
            FlowDirection = [System.Windows.Forms.FlowDirection]
            BorderStyle   = [System.Windows.Forms.BorderStyle]
            SizeMode      = [System.Windows.Forms.PictureBoxSizeMode]
            TextAlign     = [System.Drawing.ContentAlignment]
            ScrollBars    = [System.Windows.Forms.ScrollBars]
            View          = [System.Windows.Forms.View]
            FlatStyle     = [System.Windows.Forms.FlatStyle]
            StartPosition = [System.Windows.Forms.FormStartPosition]
        }
        switch ($propName) {
            "Font" {
                $fontParts = $propValue -split ',\s*'
                $fontName  = $fontParts[0].Trim('"')
                $fontSize  = [float]$fontParts[1].Trim()
                if ($fontParts.Count -gt 2) { $fontStyle = [System.Drawing.FontStyle]($fontParts[2].Trim() -replace '\s','') }
                else                        { $fontStyle = [System.Drawing.FontStyle]::Regular  }
                $control.Font = [System.Drawing.Font]::new($fontName, $fontSize, $fontStyle)
            }
            { $_ -in @("Padding","Margin") } {
                $v = $propValue -split '\s+'
                if ($v.Count -eq 1) { $control.$propName = [System.Windows.Forms.Padding]::new([int]$v[0]) }
                else                { $control.$propName = [System.Windows.Forms.Padding]::new([int]$v[0], [int]$v[1], [int]$v[2], [int]$v[3]) }
            }
            { $_ -in @("ForeColor","BackColor","BorderColor","HoverColor") } {
                $colorValues = $propValue -split '\s+'
                if ($colorValues.Count -eq 1) { $color = [System.Drawing.ColorTranslator]::FromHtml($colorValues[0]) }
                else                          { $color = [System.Drawing.Color]::FromArgb([int]$colorValues[0], [int]$colorValues[1], [int]$colorValues[2]) }
                switch ($propName) {
                    "ForeColor"   { $control.ForeColor = $color }
                    "BackColor"   { $control.BackColor = $color }
                    "BorderColor" { $control.FlatAppearance.BorderColor = $color }
                    "HoverColor"  { $control.FlatAppearance.MouseOverBackColor = $color }
                }
            }
            "Anchor" {
                $combined = [System.Windows.Forms.AnchorStyles]::None
                foreach ($v in ($propValue -split ',\s*')) { $combined = $combined -bor [System.Windows.Forms.AnchorStyles]::$v }
                $control.Anchor = $combined
            }
            { $enumMap.ContainsKey($_) } { $control.$propName = $enumMap[$propName]::$propValue }
            default {
                if     ($propValue -eq    '$true')  { $control.$propName = $true }
                elseif ($propValue -eq    '$false') { $control.$propName = $false }
                elseif ($propValue -match '^\d+$')  { $control.$propName = [int]$propValue }
                else                                { $control.$propName = $propValue }
            }
        }
    }
    if ($null -ne $container) {
        if     ($container -is [System.Windows.Forms.StatusStrip] -or $container -is [System.Windows.Forms.ToolStrip] -or $container -is [System.Windows.Forms.ContextMenuStrip]){[void]$container.Items.Add($control)}
        elseif ($container -is [System.Windows.Forms.TabControl] -and $control   -is [System.Windows.Forms.TabPage])                                                             {[void]$container.TabPages.Add($control)}
        elseif ($container -is [System.Windows.Forms.SplitContainer])                                                                                                            {[void]$container.Panel1.Controls.Add($control)}
        else                                                                                                                                                                     {[void]$container.Controls.Add($control)}
    }
    return $control
}
Set-Alias -Name gen -Value New-Control

# ── Custom title bar with drag support ──
Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -Language CSharp -TypeDefinition @'
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
public class Win11TitleBar : Panel
{
    [DllImport("user32.dll")]
    private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();
    private const int WM_NCLBUTTONDOWN = 0xA1;
    private const int HT_CAPTION = 0x2;
    private const int WM_NCHITTEST = 0x84;
    private const int HTTRANSPARENT = -1;
    private const int GRIP = 6;
    public Win11TitleBar()
    {
        this.DoubleBuffered = true;
        this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
    }
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_NCHITTEST)
        {
            base.WndProc(ref m);
            var screenPt = new Point((int)m.LParam & 0xFFFF, (int)m.LParam >> 16 & 0xFFFF);
            var pos = this.PointToClient(screenPt);
            Form f = this.FindForm();
            if (f != null)
            {
                var formPos = f.PointToClient(screenPt);
                int w = f.ClientSize.Width;
                bool top   = pos.Y <= GRIP;
                bool left  = formPos.X <= GRIP;
                bool right = formPos.X >= w - GRIP;
                if (top || left || right)
                {
                    m.Result = (IntPtr)HTTRANSPARENT;
                    return;
                }
            }
            return;
        }
        base.WndProc(ref m);
    }
    public static void DragForm(Form form)
    {
        ReleaseCapture();
        SendMessage(form.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
    }
    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button == MouseButtons.Left)
            DragForm(this.FindForm());
    }
}
public class TitleBarButton : Label
{
    private const int WM_NCHITTEST = 0x84;
    private const int HTTRANSPARENT = -1;
    private const int GRIP = 6;
    public Color HoverBack { get; set; }
    public Color NormalBack { get; set; }
    public bool IsCloseButton { get; set; }
    public TitleBarButton()
    {
        this.TextAlign = ContentAlignment.MiddleCenter;
        this.NormalBack = Color.FromArgb(240, 240, 240);
        this.HoverBack = Color.FromArgb(218, 218, 218);
        this.BackColor = NormalBack;
        this.ForeColor = Color.FromArgb(50, 50, 50);
        this.Font = new Font("Segoe MDL2 Assets", 10f);
        this.Cursor = Cursors.Default;
        this.Size = new Size(46, 34);
    }
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_NCHITTEST)
        {
            base.WndProc(ref m);
            var screenPt = new Point((int)m.LParam & 0xFFFF, (int)m.LParam >> 16 & 0xFFFF);
            var pos = this.PointToClient(screenPt);
            Form f = this.FindForm();
            if (f != null && pos.Y <= GRIP)
            {
                m.Result = (IntPtr)HTTRANSPARENT;
                return;
            }
            return;
        }
        base.WndProc(ref m);
    }
    protected override void OnMouseEnter(EventArgs e)
    {
        base.OnMouseEnter(e);
        this.BackColor = IsCloseButton ? Color.FromArgb(196, 43, 28) : HoverBack;
        if (IsCloseButton) this.ForeColor = Color.White;
    }
    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        this.BackColor = NormalBack;
        this.ForeColor = Color.FromArgb(50, 50, 50);
    }
}
'@

# ── Rounded corners window ──
Add-Type -ReferencedAssemblies System.Windows.Forms -Language CSharp -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public static class DwmHelper
{
    [DllImport("dwmapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern long DwmSetWindowAttribute(IntPtr hwnd, uint dwAttribute, ref int pvAttribute, uint cbAttribute);
    public static void SetRoundedCorners(Form form)
    {
        int preference = 2;
        DwmSetWindowAttribute(form.Handle, 33, ref preference, sizeof(int));
    }
}
'@

$titleBarColor  = [System.Drawing.Color]::FromArgb(240, 240, 240)
$clientColor    = [System.Drawing.Color]::FromArgb(243, 243, 243)
$textColor      = [System.Drawing.Color]::FromArgb(30, 30, 30)
$titleBarHeight = 34

# ── Build Icon object for taskbar from the same source image ──
$iconBitmap = [System.Drawing.Bitmap]::new($iconImage, 32, 32)
$hIcon      = $iconBitmap.GetHicon()
$taskIcon   = [System.Drawing.Icon]::FromHandle($hIcon)

# ── Save persistent .ico for shortcut references ──
if (-not [System.IO.File]::Exists($script:IcoPath)) {
    try {
        $icoStream = [System.IO.FileStream]::new($script:IcoPath, [System.IO.FileMode]::Create)
        $taskIcon.Save($icoStream)
        $icoStream.Close()
    }
    catch { }
}

# ── Auto-register Start Menu shortcut for proper taskbar identity ──
if (-not [string]::IsNullOrWhiteSpace($ScriptPath) -and [System.IO.File]::Exists($script:IcoPath)) {
    $startMenuLnk = [System.IO.Path]::Combine($script:StartMenuDir, $script:LnkName)
    $needsUpdate = $true
    if ([System.IO.File]::Exists($startMenuLnk)) {
        try {
            $shell = New-Object -ComObject WScript.Shell
            $existing = $shell.CreateShortcut($startMenuLnk)
            $existingDesc = $existing.Description
            if ($existingDesc.Contains('[UserPinned]')) {
                $script:UserPinnedStartMenu = $true
                $needsUpdate = $false
            }
            elseif ($existing.TargetPath -eq $ScriptPath -and $existing.IconLocation -eq "$($script:IcoPath),0") {
                $needsUpdate = $false
            }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($existing)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
        }
        catch { }
    }
    if ($needsUpdate -and -not $script:UserPinnedStartMenu) {
        try {
            [ShortcutHelper]::CreateShortcut($startMenuLnk, $ScriptPath, "", $script:IcoPath, $script:AppId, "MSI Tools v$($script:Version)")
            Write-Log "Auto-registered Start Menu shortcut (temporary) : $startMenuLnk"
        }
        catch { Write-Log "Failed to auto-register Start Menu shortcut : $_" -Level Warning }
    }
}

# ── Main form ──
$screenWidth          = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight         = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
$form                 = New-Object CustomForm
$form.ShowInTaskbar   = $true
$form.Text            = "MSI Tools v$($script:version)"
$form.FormBorderStyle = 'None'
$form.BackColor       = $clientColor
$form.Icon            = $taskIcon
$form.ShowIcon        = $true
$tabControl           = gen $form       "TabControl" 0 $titleBarHeight $form.ClientSize.Width $titleBarHeight 'Anchor=Top,Bottom,Left,Right'
$tabPage1             = gen $tabControl "TabPage"    ".MSI Properties"
$tabPage2             = gen $tabControl "TabPage"    "Explore .MSI Files" 
$tabPage3             = gen $tabControl "TabPage"    "Explore Uninstall Keys"
$tabPage4             = gen $tabControl "TabPage"    "MSI Residues Cleanup"
# $tabControl.TabStop = $false

# ── Title bar ──
$titleBar = [Win11TitleBar]::new()
$titleBar.Dock      = 'Top'
$titleBar.Height    = $titleBarHeight
$titleBar.BackColor = $titleBarColor

# ── Title bar icon (20x20 display) ──
$iconBox = gen $titleBar "PictureBox" 8 ([int](($titleBarHeight - 20) / 2)) 20 20 "SizeMode=Zoom" "BackColor=240 240 240"
$iconBox.Image = $iconImage

# ── Title text ──
$titleLabel = gen $titleBar "Label" "MSI Tools v$($script:Version)" 34 0 200 $titleBarHeight "ForeColor=30 30 30" "Font=Arial, 10, Bold" "BackColor=240 240 240" "TextAlign=MiddleLeft"

# ── Forward drag from icon and label ──
$dragHandler = {
    param($s, $e)
    if ($e.Button -eq 'Left') { [Win11TitleBar]::DragForm($form) }
}
$iconBox.Add_MouseDown($dragHandler)
$titleLabel.Add_MouseDown($dragHandler)

# ── Title bar buttons (Dock=Right : first added = rightmost) ──
$btnMinimize = [TitleBarButton]::new()
$btnMinimize.Text = [char]0xE921
$btnMinimize.Dock = 'Right'
$btnMinimize.Add_Click({ $form.WindowState = 'Minimized' })

$btnMaximize = [TitleBarButton]::new()
$btnMaximize.Text = [char]0xE922
$btnMaximize.Dock = 'Right'
$btnMaximize.Add_Click({
    if ($form.WindowState -eq 'Maximized') {
        $form.WindowState = 'Normal'
        $btnMaximize.Text = [char]0xE922
    } else {
        $form.WindowState = 'Maximized'
        $btnMaximize.Text = [char]0xE923
    }
})

$btnClose = [TitleBarButton]::new()
$btnClose.Text          = [char]0xE8BB
$btnClose.IsCloseButton = $true
$btnClose.Dock          = 'Right'
$btnClose.Add_Click({ $form.Close() })

# ── Assemble title bar ──
$titleBar.Controls.Add($btnMinimize)
$titleBar.Controls.Add($btnMaximize)
$titleBar.Controls.Add($btnClose)
$titleBar.Controls.Add($iconBox)
$titleBar.Controls.Add($titleLabel)

# ── Double-click title bar to maximize/restore ──
$titleBar.Add_DoubleClick({
    if ($form.WindowState -eq 'Maximized') {
        $form.WindowState = 'Normal'
        $btnMaximize.Text = [char]0xE922
    } else {
        $form.WindowState = 'Maximized'
        $btnMaximize.Text = [char]0xE923
    }
})

# ── Subtle border ──
$form.Add_Paint({
    param($s, $e)
    $pen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(200, 200, 200), 1)
    $e.Graphics.DrawRectangle($pen, 0, 0, $s.ClientSize.Width - 1, $s.ClientSize.Height - 1)
    $pen.Dispose()
})

# ── About button (notch-style trapezoid in title bar) ──
$script:notchTopWidth    = 96
$script:notchBottomWidth = 64
$script:notchHeight      = 21
$notchInset = [int](($script:notchTopWidth - $script:notchBottomWidth) / 2)
$btnAbout = New-Object System.Windows.Forms.Panel
$btnAbout.Size      = [System.Drawing.Size]::new($script:notchTopWidth, $script:notchHeight)
$btnAbout.Location  = [System.Drawing.Point]::new([int](($form.ClientSize.Width - $script:notchTopWidth) / 2), 0)
$btnAbout.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$btnAbout.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnAbout.SetStyle([System.Windows.Forms.ControlStyles]::AllPaintingInWmPaint -bor [System.Windows.Forms.ControlStyles]::UserPaint -bor [System.Windows.Forms.ControlStyles]::OptimizedDoubleBuffer, $true)
# Trapezoid clipping region (wide at top, narrow at bottom)
$notchPath = [System.Drawing.Drawing2D.GraphicsPath]::new()
$notchPath.AddPolygon(@(
    [System.Drawing.Point]::new(0, 0),
    [System.Drawing.Point]::new($script:notchTopWidth, 0),
    [System.Drawing.Point]::new($script:notchTopWidth - $notchInset, $script:notchHeight),
    [System.Drawing.Point]::new($notchInset, $script:notchHeight)
))
$btnAbout.Region = [System.Drawing.Region]::new($notchPath)
# Paint handler for text and border
$btnAbout.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    # Trapezoid border (left slope, bottom, right slope)
    $borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(180, 180, 180), 1)
    $inset = [int](($s.Width - $script:notchBottomWidth) / 2)
    $g.DrawLine($borderPen, 0, 0, $inset, $s.Height - 1)
    $g.DrawLine($borderPen, $inset, $s.Height - 1, $s.Width - $inset, $s.Height - 1)
    $g.DrawLine($borderPen, $s.Width - $inset, $s.Height - 1, $s.Width, 0)
    $borderPen.Dispose()
    # Centered text
    $sf = [System.Drawing.StringFormat]::new()
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $font = [System.Drawing.Font]::new("Arial", 8)
    $rect = [System.Drawing.RectangleF]::new(0, 1, $s.Width, $s.Height)
    $g.DrawString("About", $font, [System.Drawing.Brushes]::Black, $rect, $sf)
    $font.Dispose()
    $sf.Dispose()
})
# Hover effects
$btnAbout.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(210, 210, 210); $this.Invalidate() })
$btnAbout.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228); $this.Invalidate() })
# Forward drag when clicking outside the text area
$btnAbout.Add_MouseDown({
    param($s, $e)
    if ($e.Button -eq 'Left' -and $e.Clicks -eq 1) {
        # Only trigger About on single click, drag is handled by title bar passthrough
    }
})
$btnAbout.Add_Click({
    $aboutForm = gen $null "Form" "About" 0 0 380 280 "StartPosition=CenterParent" "MaximizeBox=$false" "MinimizeBox=$false" "BackColor=243 243 243"
    $aboutForm.FormBorderStyle = 'FixedDialog'
    # Header panel with icon + title + pin button
    $headerPanel = gen $aboutForm "Panel" 0 0 380 80 "Dock=Top" "BackColor=255 255 255"
    $headerIcon  = gen $headerPanel "PictureBox" 20 8 64 64 "SizeMode=Zoom"
    $headerIcon.Image = $iconImage
    $headerTitle = gen $headerPanel "Label" "MSI Tools" 94 10 0 0 "Font=Arial, 18, Bold" "AutoSize=$true"
    $headerVer   = gen $headerPanel "Label" "v$($script:Version)" 94 45 0 0 "Font=Arial, 10" "ForeColor=120 120 120" "AutoSize=$true"
    $startMenuPinned = Test-AppPinned -Target StartMenu
    $btnPinStart = gen $headerPanel "Button" "" 0 0 130 26 "FlatStyle=Flat" "Font=Arial, 8" "Anchor=Top,Right"
    $btnPinStart.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
    $btnPinStart.Location = [System.Drawing.Point]::new($headerPanel.ClientSize.Width - $btnPinStart.Width - 10, $headerPanel.Height - $btnPinStart.Height - 8)
    $btnPinStart.Text = if ($script:UserPinnedStartMenu) { [char]0x2714 + " Start Menu" } else { "Keep in Start Menu" }
    $btnPinStart.Tag  = $script:UserPinnedStartMenu
    $btnPinStart.Add_Click({
        $startMenuLnk = [System.IO.Path]::Combine($script:StartMenuDir, $script:LnkName)
        if ($this.Tag) {
            try {
                if ([System.IO.File]::Exists($startMenuLnk)) { [System.IO.File]::Delete($startMenuLnk) }
                $script:UserPinnedStartMenu = $false
                $this.Tag  = $false
                $this.Text = "Keep in Start Menu"
                [ShortcutHelper]::CreateShortcut($startMenuLnk, $ScriptPath, "", $script:IcoPath, $script:AppId, "MSI Tools v$($script:Version)")
                Write-Log "User unpinned Start Menu shortcut (reverted to temporary)"
            }
            catch { Write-Log "Failed to revert Start Menu shortcut : $_" -Level Warning }
        }
        else {
            try {
                [ShortcutHelper]::CreateShortcut($startMenuLnk, $ScriptPath, "", $script:IcoPath, $script:AppId, "MSI Tools v$($script:Version) [UserPinned]")
                $script:UserPinnedStartMenu = $true
                $this.Tag  = $true
                $this.Text = [char]0x2714 + " Start Menu"
                Write-Log "User pinned Start Menu shortcut (permanent) : $startMenuLnk"
            }
            catch { Write-Log "Failed to create Start Menu shortcut : $_" -Level Warning }
        }
    })
    $lblMadeBy  = gen $aboutForm "Label"     "Made by Léo Gillet - Freenitial" 20 95 0 0 "Font=Arial, 10" "AutoSize=$true"
    $linkGitHub = gen $aboutForm "LinkLabel" "(GitHub)" ($lblMadeBy.Location.X + $lblMadeBy.PreferredWidth) 95 0 0 "Font=Arial, 10" "AutoSize=$true"
    $linkGitHub.LinkColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $linkGitHub.Add_LinkClicked({ [System.Diagnostics.Process]::Start("https://github.com/Freenitial/MSI_Tools") })
    $separator    = gen $aboutForm "Label"     "" 20 125 320 2 "BorderStyle=Fixed3D"
    $lblChangelog = gen $aboutForm "Label"     "Changelog :" 20 140 0 0 "Font=Arial, 10, Bold" "AutoSize=$true"
    $lblEntries   = gen $aboutForm "Label"     "$([char]0x2022)  v1.0 : First release" 25 165 0 0 "Font=Arial, 9" "AutoSize=$true"
    $aboutForm.ShowDialog($form) | Out-Null
    $aboutForm.Dispose()
})
$titleBar.Add_Resize({ $btnAbout.Location = [System.Drawing.Point]::new([int](($titleBar.Width - $script:notchTopWidth) / 2), 0) })
$titleBar.Controls.Add($btnAbout)
$btnAbout.BringToFront()

# ── Console visibility toggle ──
$chkConsole = gen $titleBar "CheckBox" "Show console" 0 ([int](($titleBarHeight - 18) / 2)) 101 18 "Font=Arial, 8" "BackColor=240 240 240" "FlatStyle=Flat"
$chkConsole.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$chkConsole.Location = [System.Drawing.Point]::new($btnMinimize.Left - $chkConsole.Width - 8, $chkConsole.Location.Y)
$chkConsole.Add_CheckedChanged({
    $hConsole = [NativeMethods]::GetConsoleWindow()
    if ($hConsole -ne [IntPtr]::Zero) {
        if ($this.Checked) { [NativeMethods]::ShowWindow($hConsole, [NativeMethods]::SW_SHOW) | Out-Null }
        else               { [NativeMethods]::ShowWindow($hConsole, [NativeMethods]::SW_HIDE) | Out-Null }
    }
})
$chkConsole.BringToFront()

$form.Controls.Add($titleBar)

$highlightBrush = [System.Drawing.SolidBrush]::new([System.Drawing.SystemColors]::Highlight)
$highlightTextBrush = [System.Drawing.SolidBrush]::new([System.Drawing.SystemColors]::HighlightText)

#region Common Helpers

function Get-MsiInfo {
    param (
        [string]$FilePath,
        [string[]]$Properties = @("ProductCode", "ProductVersion", "ProductName"),
        [switch]$Full
    )
    $result = @{}
    foreach ($p in $Properties) { $result[$p] = "None" }
    if ($Full) {
        $result["_AllProperties"]   = [System.Collections.Generic.List[PSCustomObject]]::new()
        $result["_Features"]        = [System.Collections.Generic.List[PSCustomObject]]::new()
        $result["_MultiValueProps"] = [ordered]@{}
        $result["_Icon"]            = $null
    }
    if (-not $FilePath.EndsWith(".msi", [System.StringComparison]::OrdinalIgnoreCase) -or
        -not ([System.IO.File]::Exists($FilePath))) {
        return ,$result
    }
    if ($null -eq $script:WindowsInstallerCOM) {
        $script:WindowsInstallerCOM = New-Object -ComObject WindowsInstaller.Installer
    }
    $database = $null
    try {
        $database = $script:WindowsInstallerCOM.OpenDatabase($FilePath, 0)
        if ($Full) {
            # ---------- Property table ----------
            $defaults = @{}
            $view = $null; $record = $null
            try {
                $view = $database.OpenView("SELECT Property, Value FROM Property")
                [void]$view.Execute()
                while ($true) {
                    $record = $view.Fetch()
                    if ($null -eq $record) { break }
                    $prop = $record.StringData(1); $val = $record.StringData(2)
                    $defaults[$prop] = $val
                    $result["_AllProperties"].Add([PSCustomObject]@{ Property = $prop; Value = $val })
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record); $record = $null
                }
            }
            catch { Write-Log "Get-MsiInfo : Error reading Property table from '$FilePath' : $_" -Level Warning }
            finally {
                if ($null -ne $record) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
                if ($null -ne $view)   { [void]$view.Close(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
            }
            foreach ($prop in $Properties) { $result[$prop] = if ($defaults.ContainsKey($prop)) { $defaults[$prop] } else { "None" } }
            # ---------- Feature table ----------
            $view = $null; $record = $null
            try {
                $view = $database.OpenView("SELECT Feature, Level, Title, Description FROM Feature")
                [void]$view.Execute()
                while ($true) {
                    $record = $view.Fetch()
                    if ($null -eq $record) { break }
                    $result["_Features"].Add([PSCustomObject]@{
                        Name        = $record.StringData(1)
                        Level       = $record.StringData(2)
                        Title       = $record.StringData(3)
                        Description = $record.StringData(4)
                    })
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record); $record = $null
                }
            }
            catch { Write-Log "Get-MsiInfo : Feature table not available in '$FilePath' : $_" -Level Debug }
            finally {
                if ($null -ne $record) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
                if ($null -ne $view)   { [void]$view.Close(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
            }
            # ---------- UI tables + boolean detection for associated properties ----------
            $uiProperties = @{}
            foreach ($table in @('ComboBox', 'ListBox', 'RadioButton')) {
                $view = $null; $record = $null
                try {
                    $view = $database.OpenView("SELECT Property, Value, Text FROM $table ORDER BY Property, ``Order``")
                    [void]$view.Execute()
                    while ($true) {
                        $record = $view.Fetch()
                        if ($null -eq $record) { break }
                        $prop = $record.StringData(1); $val = $record.StringData(2); $text = $record.StringData(3)
                        if (-not $uiProperties.ContainsKey($prop)) { $uiProperties[$prop] = [System.Collections.Generic.List[PSCustomObject]]::new() }
                        $uiProperties[$prop].Add([PSCustomObject]@{ Value = $val; Text = $text; Source = $table })
                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record); $record = $null
                    }
                }
                catch { <# Table may not exist #> }
                finally {
                    if ($null -ne $record) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
                    if ($null -ne $view)   { [void]$view.Close(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
                }
            }
            # ---------- CheckBox table ----------
            $view = $null; $record = $null
            try {
                $view = $database.OpenView("SELECT Property, Value FROM CheckBox")
                [void]$view.Execute()
                while ($true) {
                    $record = $view.Fetch()
                    if ($null -eq $record) { break }
                    $prop = $record.StringData(1); $val = $record.StringData(2)
                    if (-not $uiProperties.ContainsKey($prop) -and -not [string]::IsNullOrEmpty($val)) {
                        $uiProperties[$prop] = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $uiProperties[$prop].Add([PSCustomObject]@{ Value = $val; Text = ""; Source = "CheckBox" })
                        $uiProperties[$prop].Add([PSCustomObject]@{ Value = "";   Text = ""; Source = "CheckBox" })
                    }
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record); $record = $null
                }
            }
            catch { <# Table may not exist #> }
            finally {
                if ($null -ne $record) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
                if ($null -ne $view)   { [void]$view.Close(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
            }
            # Multi-value from UI tables (2+ options)
            foreach ($prop in ($uiProperties.Keys | Sort-Object)) {
                if ($uiProperties[$prop].Count -ge 2) {
                    $defaultVal = if ($defaults.ContainsKey($prop)) { $defaults[$prop] } else { "(not set)" }
                    $result["_MultiValueProps"][$prop] = @{ Default = $defaultVal; Options = $uiProperties[$prop] }
                }
            }
            # Boolean-like properties inferred from default values
            foreach ($propName in ($defaults.Keys | Sort-Object)) {
                if ($result["_MultiValueProps"].Contains($propName)) { continue }
                if ($uiProperties.ContainsKey($propName)) { continue }
                $val = $defaults[$propName]
                if ($val.Length -gt 5) { continue }
                $lower = $val.ToLowerInvariant()
                $other = switch ($lower) {
                    '0'     { '1' }
                    '1'     { '0' }
                    'true'  { 'False' }
                    'false' { 'True' }
                    'yes'   { 'No' }
                    'no'    { 'Yes' }
                    default { $null }
                }
                if ($null -eq $other) { continue }
                $result["_MultiValueProps"][$propName] = @{
                    Default = $val
                    Options = [System.Collections.Generic.List[PSCustomObject]]@(
                        [PSCustomObject]@{ Value = $val;   Text = ""; Source = "Boolean" },
                        [PSCustomObject]@{ Value = $other; Text = ""; Source = "Boolean" }
                    )
                }
            }
            # ---------- Icon extraction ----------
            $result["_Icon"] = $null
            $iconMethod = $null
            Write-Log "Starting icon extraction for '$FilePath'"
            $extractIconFromTable = {
                param([string]$iconName)
                Write-Log "Attempting icon extraction for entry '$iconName'"
                $iv = $null; $ir = $null
                try {
                    $safeName = $iconName.Replace("'", "''")
                    $iv = $database.OpenView("SELECT Data FROM Icon WHERE Name = '$safeName'")
                    [void]$iv.Execute()
                    $ir = $iv.Fetch()
                    if ($null -eq $ir) { Write-Log "No record found for icon '$iconName'" -Level Debug; return $null }
                    $sz = $ir.DataSize(1)
                    Write-Log "Icon '$iconName' : data size = $sz bytes"
                    if ($sz -le 0) { Write-Log "Icon '$iconName' : zero data size" -Level Debug; return $null }
                    $rawData = $null
                    try { $rawData = $ir.ReadStream(1, $sz, 1) }
                    catch { Write-Log "ReadStream failed for '$iconName' : $_" -Level Warning; return $null }
                    $rawType = if ($null -eq $rawData) { "null" } else { $rawData.GetType().FullName }
                    Write-Log "Icon '$iconName' : ReadStream returned type '$rawType'"
                    $bytes = $null
                    try {
                        if ($rawData -is [byte[]]) { $bytes = $rawData }
                        elseif ($rawData -is [string]) { $bytes = [System.Text.Encoding]::GetEncoding(28591).GetBytes($rawData) }
                        elseif ($rawData -is [System.Array]) {
                            $bytes = [byte[]]::new($rawData.Length)
                            for ($bi = 0; $bi -lt $rawData.Length; $bi++) { $bytes[$bi] = [byte]$rawData[$bi] }
                        }
                        else { $bytes = [byte[]]$rawData }
                    }
                    catch { Write-Log "Failed byte conversion for '$iconName' : $_" -Level Warning; return $null }
                    if ($null -eq $bytes -or $bytes.Length -eq 0) { Write-Log "Empty byte array for '$iconName'" -Level Warning; return $null }
                    Write-Log "Icon '$iconName' : $($bytes.Length) bytes, first 4 = [$($bytes[0]),$($bytes[1]),$($bytes[2]),$($bytes[3])]"
                    $isIco = ($bytes.Length -ge 4 -and $bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 1 -and $bytes[3] -eq 0)
                    $isPE  = ($bytes.Length -ge 2 -and $bytes[0] -eq 0x4D -and $bytes[1] -eq 0x5A)
                    Write-Log "Icon '$iconName' : isIco=$isIco, isPE=$isPE"
                    if ($isIco) {
                        try {
                            $ms  = [System.IO.MemoryStream]::new($bytes)
                            $ico = [System.Drawing.Icon]::new($ms, 256, 256)
                            $bmp = $ico.ToBitmap(); $ico.Dispose(); $ms.Dispose()
                            Write-Log "Successfully extracted ICO (256px) from '$iconName'"
                            return @{ Bitmap = $bmp; RawBytes = $bytes }
                        }
                        catch { Write-Log "ICO load failed for '$iconName' : $_" -Level Warning }
                    }
                    if ($isPE) {
                        $tmpExe = [System.IO.Path]::Combine($env:TEMP, "msi_icon_$([Guid]::NewGuid().ToString('N')).exe")
                        try {
                            [System.IO.File]::WriteAllBytes($tmpExe, $bytes)
                            $exeIcon = [System.Drawing.Icon]::ExtractAssociatedIcon($tmpExe)
                            if ($exeIcon) {
                                $bmp = $exeIcon.ToBitmap(); $exeIcon.Dispose()
                                Write-Log "Successfully extracted EXE icon from '$iconName'"
                                return @{ Bitmap = $bmp; RawBytes = $bytes }
                            }
                            else { Write-Log "ExtractAssociatedIcon returned null for '$iconName'" -Level Warning }
                        }
                        catch { Write-Log "EXE icon extraction failed for '$iconName' : $_" -Level Warning }
                        finally { if ([System.IO.File]::Exists($tmpExe)) { try { [System.IO.File]::Delete($tmpExe) } catch {} } }
                    }
                    if (-not $isIco) {
                        try {
                            $ms  = [System.IO.MemoryStream]::new($bytes)
                            $ico = [System.Drawing.Icon]::new($ms, 256, 256)
                            $bmp = $ico.ToBitmap(); $ico.Dispose(); $ms.Dispose()
                            Write-Log "Fallback ICO load (256px) succeeded for '$iconName'"
                            return @{ Bitmap = $bmp; RawBytes = $bytes }
                        }
                        catch { Write-Log "Fallback ICO load failed for '$iconName' : $_" -Level Debug }
                    }
                    try {
                        $ms  = [System.IO.MemoryStream]::new($bytes)
                        $img = [System.Drawing.Image]::FromStream($ms)
                        Write-Log "Loaded as generic image from '$iconName'"
                        return @{ Bitmap = $img; RawBytes = $bytes }
                    }
                    catch { Write-Log "Generic image load failed for '$iconName' : $_" -Level Debug }
                    Write-Log "No usable icon from '$iconName'" -Level Warning
                    return $null
                }
                catch { Write-Log "Error in extractIconFromTable for '$iconName' : $_" -Level Warning; return $null }
                finally {
                    if ($null -ne $ir) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ir) }
                    if ($null -ne $iv) { [void]$iv.Close(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($iv) }
                }
            }
            # Priority 1 : ARPPRODUCTICON property
            $arpIcon = if ($defaults.ContainsKey('ARPPRODUCTICON')) { $defaults['ARPPRODUCTICON'] } else { $null }
            Write-Log "ARPPRODUCTICON = '$arpIcon'"
            if ($arpIcon) {
                $result["_Icon"] = & $extractIconFromTable $arpIcon
                if ($result["_Icon"]) { $iconMethod = "ARPPRODUCTICON ('$arpIcon')" }
            }
            # Priority 2 : scan Icon table (EXE → ICO → any)
            if (-not $result["_Icon"]) {
                Write-Log "Scanning Icon table"
                $iv = $null; $ir = $null
                $allIconNames = [System.Collections.Generic.List[string]]::new()
                try {
                    $iv = $database.OpenView("SELECT Name FROM Icon")
                    [void]$iv.Execute()
                    while ($true) {
                        $ir = $iv.Fetch()
                        if ($null -eq $ir) { break }
                        $allIconNames.Add($ir.StringData(1))
                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ir); $ir = $null
                    }
                }
                catch { Write-Log "Icon table scan failed : $_" -Level Debug }
                finally {
                    if ($null -ne $ir) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ir) }
                    if ($null -ne $iv) { [void]$iv.Close(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($iv) }
                }
                Write-Log "Icon table contains $($allIconNames.Count) entries : $($allIconNames -join ', ')"
                foreach ($name in $allIconNames) {
                    if ($name.EndsWith(".exe", [System.StringComparison]::OrdinalIgnoreCase)) {
                        $result["_Icon"] = & $extractIconFromTable $name
                        if ($result["_Icon"]) { $iconMethod = "Icon table EXE entry ('$name')"; break }
                    }
                }
                if (-not $result["_Icon"]) {
                    foreach ($name in $allIconNames) {
                        if ($name.EndsWith(".ico", [System.StringComparison]::OrdinalIgnoreCase)) {
                            $result["_Icon"] = & $extractIconFromTable $name
                            if ($result["_Icon"]) { $iconMethod = "Icon table ICO entry ('$name')"; break }
                        }
                    }
                }
                if (-not $result["_Icon"]) {
                    foreach ($name in $allIconNames) {
                        $result["_Icon"] = & $extractIconFromTable $name
                        if ($result["_Icon"]) { $iconMethod = "Icon table generic entry ('$name')"; break }
                    }
                }
            }
            if ($result["_Icon"] -is [hashtable]) {
                $result["_IconBytes"] = $result["_Icon"].RawBytes
                $result["_Icon"]      = $result["_Icon"].Bitmap
                Write-Log "Icon extraction successful via : $iconMethod"
            }
            else {
                $result["_Icon"]      = $null
                $result["_IconBytes"] = $null
                Write-Log "No icon extracted, will use generic fallback" -Level Warning
            }
        }
        else {
            # Standard mode : specific properties only
            foreach ($prop in $Properties) {
                $view = $null; $record = $null
                try {
                    $safeProp = $prop.Replace("'", "''")
                    $view = $database.OpenView("SELECT Value FROM Property WHERE Property = '$safeProp'")
                    [void]$view.Execute()
                    $record = $view.Fetch()
                    if ($null -ne $record) { $result[$prop] = $record.StringData(1) }
                }
                catch { Write-Log "Get-MsiInfo : Error reading '$prop' from '$FilePath' : $_" -Level Warning }
                finally {
                    if ($null -ne $record) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
                    if ($null -ne $view)   { [void]$view.Close(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
                }
            }
        }
    }
    catch { Write-Log "Get-MsiInfo : Cannot open '$FilePath' : $_" -Level Warning }
    finally {
        if ($null -ne $database) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($database) }
    }
    return ,$result
}

function Test-AppPinned {
    param([ValidateSet('Taskbar','StartMenu')][string]$Target)
    if ($Target -eq 'StartMenu') {
        $lnk = [System.IO.Path]::Combine($script:StartMenuDir, $script:LnkName)
        return [System.IO.File]::Exists($lnk)
    }
    try {
        $regKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband", $false)
        if ($null -eq $regKey) { return $false }
        $favorites = $regKey.GetValue("Favorites", $null)
        $regKey.Close()
        if ($null -eq $favorites -or $favorites -isnot [byte[]]) { return $false }
        $searchBytes = [System.Text.Encoding]::Unicode.GetBytes($script:LnkName)
        $blob = [byte[]]$favorites
        for ($i = 0; $i -le $blob.Length - $searchBytes.Length; $i++) {
            $found = $true
            for ($j = 0; $j -lt $searchBytes.Length; $j++) {
                if ($blob[$i + $j] -ne $searchBytes[$j]) { $found = $false; break }
            }
            if ($found) { return $true }
        }
        return $false
    }
    catch {
        Write-Log "Test-AppPinned Taskbar registry check failed : $_" -Level Debug
        return $false
    }
}

function Set-AppPin {
    param(
        [ValidateSet('Taskbar','StartMenu')][string]$Target,
        [bool]$Pin
    )
    $dir = if ($Target -eq 'Taskbar') { $script:TaskbarPinDir } else { $script:StartMenuDir }
    $lnk = [System.IO.Path]::Combine($dir, $script:LnkName)
    if (-not $Pin) {
        if ([System.IO.File]::Exists($lnk)) {
            [System.IO.File]::Delete($lnk)
            Write-Log "Removed $Target shortcut : $lnk"
        }
        return
    }
    if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
        Write-Log "Cannot pin : script path unknown" -Level Warning
        [System.Windows.Forms.MessageBox]::Show("Cannot create shortcut : script path is unknown.", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    if (-not [System.IO.File]::Exists($lnk)) {
        try {
            $workDir = [System.IO.Path]::GetDirectoryName($ScriptPath)
            [ShortcutHelper]::CreateShortcut($lnk, $ScriptPath, "", $script:IcoPath, $script:AppId, "MSI Tools v$($script:Version)", $workDir)
            Write-Log "Created $Target shortcut : $lnk"
        }
        catch { Write-Log "Failed to create $Target shortcut : $_" -Level Warning }
    }
}

function Set-CustomPlaceholder {
    param(
        [System.Windows.Forms.TextBox]$TextBox,
        [string]$PlaceholderText
    )
    $TextBox.Tag       = @{ Text = $PlaceholderText; IsPlaceholder = $true }
    $TextBox.Text      = $PlaceholderText
    $TextBox.ForeColor = [System.Drawing.Color]::Gray
    $TextBox.Add_Enter({
        if ($this.Tag.IsPlaceholder) {
            $txt = $this.Tag.Text
            $this.Tag       = @{ Text = $txt; IsPlaceholder = $false }
            $this.Text      = ""
            $this.ForeColor = [System.Drawing.Color]::Black
        }
    })
    $TextBox.Add_Leave({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $txt            = $this.Tag.Text
            $this.Tag       = @{ Text = $txt; IsPlaceholder = $true }
            $this.Text      = $txt
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })
}

function Format-FileSize {
    param([long]$Bytes)
    if     ($Bytes -ge 1073741824) { return "{0:N2} GB" -f ($Bytes / 1073741824) }
    elseif ($Bytes -ge 1048576)    { return "{0:N2} MB" -f ($Bytes / 1048576) }
    else                           { return "{0:N2} KB" -f ($Bytes / 1024) }
}

Function HandleColumnClick {
    param ([System.Windows.Forms.ListView]$listViewparam, [System.Windows.Forms.ColumnClickEventArgs]$e)
    if ($e.Column -eq $script:sortColumn) {
        $script:sortOrder = if ($script:sortOrder -eq [System.Windows.Forms.SortOrder]::Ascending) {[System.Windows.Forms.SortOrder]::Descending} else {[System.Windows.Forms.SortOrder]::Ascending}
    } else {
        $script:sortColumn = $e.Column
        $script:sortOrder  = [System.Windows.Forms.SortOrder]::Ascending
    }
    $listViewparam.ListViewItemSorter = New-Object ListViewItemComparer($script:sortColumn, $script:sortOrder)
    $listViewparam.Sort()
}

function Add-ListViewSearchFilter {
    param(
        [System.Windows.Forms.TextBox]$SearchTextBox,
        [System.Windows.Forms.ListView]$ListView,
        [System.Collections.ArrayList]$AllItems,
        [scriptblock]$AdditionalFilter = $null
    )
    $SearchTextBox.Add_TextChanged({
        $currentText = $SearchTextBox.Text
        $placeholder = ""
        if ($SearchTextBox.Tag -is [hashtable] -and $SearchTextBox.Tag.IsPlaceholder) { $placeholder = $SearchTextBox.Tag.Text }
        if ($currentText -eq $placeholder) { $filterTerms = @() } 
        else {
            $filterTerms = @($currentText -split ';' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ -ne '' })
        }
        $ListView.BeginUpdate()
        $ListView.Items.Clear()
        foreach ($item in $AllItems) {
            $matchText = ($filterTerms.Count -eq 0)
            if (-not $matchText) {
                foreach ($subItem in $item.SubItems) {
                    $subText = $subItem.Text.ToLower()
                    foreach ($term in $filterTerms) {
                        if ($subText -like "*$term*") { $matchText = $true ; break }
                    }
                    if ($matchText) { break }
                }
            }
            if ($matchText) {
                $includeItem = if ($AdditionalFilter) { & $AdditionalFilter $item } else { $true }
                if ($includeItem) { [void]$ListView.Items.Add($item.Clone()) }
            }
        }
        $ListView.EndUpdate()
    }.GetNewClosure())
}

Function Update-ProgressBarWidth {
    param($statusStrip, $statusLabel, $stopButton, $progressBar)
    $availableWidth = $statusStrip.Width - ($statusLabel.Width + $stopButton.Width + 5)
    if ($availableWidth - 5 -lt 0) { $availableWidth = 0 }
    $progressBar.Width = $availableWidth
}

function ConfigureListViewContextMenu($listView) {
    $contextMenu     = [System.Windows.Forms.ContextMenuStrip]::new()
    $contextMenu.Tag = $listView
    $items           = $contextMenu.Items
    # ── Build column index map for this ListView ──
    $colMap          = @{}
    for ($i = 0; $i -lt $listView.Columns.Count; $i++) { $colMap[$listView.Columns[$i].Text] = $i }
    # ── Menu item creation helpers ──
    $addItem = { param($text, $action, $check) 
        $menuItems = [System.Windows.Forms.ToolStripMenuItem]::new($text)
        $menuItems.Add_Click($action)
        if ($check) { $menuItems.CheckOnClick = $true }
        [void]$items.Add($menuItems)
        return $menuItems
    }
    $addSep = { [void]$items.Add([System.Windows.Forms.ToolStripSeparator]::new()) }
    # =========================================== REGISTRY TAB 3 (top items)
    if ($colMap.ContainsKey("Registry Path")) {
        [void](& $addItem "Open Regedit here" {
            $lv = $this.GetCurrentParent().Tag
            if ($lv.SelectedItems.Count -ne 1) { return }
            $regPath = $lv.SelectedItems[0].Tag
            $computerName = $null
            $cm = @{}; for ($i = 0; $i -lt $lv.Columns.Count; $i++) { $cm[$lv.Columns[$i].Text] = $i }
            if ($cm.ContainsKey("Device")) {
                $deviceText = $lv.SelectedItems[0].SubItems[$cm["Device"]].Text
                if ($deviceText -and $deviceText -ne $env:COMPUTERNAME) {
                    foreach ($r in $script:RemoteConnectionResults) {
                        if ((Get-ComputerDisplayName -Computer $r.Computer) -eq $deviceText) { $computerName = $r.Computer ; break }
                    }
                    if (-not $computerName) { $computerName = ($deviceText -split '\s*\(')[0].Trim() }
                }
            }
            if ($computerName) { Open-RegeditHere -Path $regPath -ComputerName $computerName }
            else               { Open-RegeditHere -Path $regPath }
        })
        [void](& $addItem "Go to Uninstall Tab" {
            $lv = $this.GetCurrentParent().Tag
            $cm = @{}; for ($i = 0; $i -lt $lv.Columns.Count; $i++) { $cm[$lv.Columns[$i].Text] = $i }
            if ($lv.SelectedItems.Count -gt 0) {
                $selectedItem = $lv.SelectedItems[0]
                $displayName = if ($cm.ContainsKey("DisplayName")) { $selectedItem.SubItems[$cm["DisplayName"]].Text } else { "" }
                $guid        = if ($cm.ContainsKey("GUID"))        { $selectedItem.SubItems[$cm["GUID"]].Text }        else { "" }
                $tabControl.SelectedTab = $tabPage4
                if ($script:textBox_UninstallTitle)       { $script:textBox_UninstallTitle.Text = $displayName }
                if ($script:textBox_UninstallGUID)        { $script:textBox_UninstallGUID.Text = $guid }
                if ($script:button_UninstallSearch)       { $script:button_UninstallSearch.PerformClick() }
            }
        })
        [void](& $addSep)
    }
    # =========================================== COPY ITEMS (all tabs)
    $optName = & $addItem "OPTION: Include Name" $null $true
    [void](& $addItem "Copy Full Row" {
        $lv = $this.GetCurrentParent().Tag
        $sb = [System.Text.StringBuilder]::new()
        foreach ($item in $lv.SelectedItems) {
            $row = [System.Collections.Generic.List[string]]::new()
            foreach ($sub in $item.SubItems) { if (![string]::IsNullOrWhiteSpace($sub.Text) -and $sub.Text -ne "None") { [void]$row.Add($sub.Text) } }
            if      ($row.Count -gt 0)       { [void]$sb.AppendLine(($row -join "`t")) }
        }
        if ($sb.Length -gt 0) { [System.Windows.Forms.Clipboard]::SetText($sb.ToString()) }
    })
    # Per-column copy items
    foreach ($col in $listView.Columns) {
        if ($col.Text -eq [string][char]0x2198) { continue }
        if ($col.Text -eq "InstallDate")        { continue }
        $menuItems     = [System.Windows.Forms.ToolStripMenuItem]::new("Copy $($col.Text)")
        $menuItems.Tag = @{ OptName = $optName; ColName = $col.Text }
        $menuItems.Add_Click({
            $lv      = $this.GetCurrentParent().Tag
            $colName = $this.Tag.ColName
            $optNameRef = $this.Tag.OptName
            if ($lv.SelectedItems.Count -gt 0) {
                $cm = @{}; for ($i = 0; $i -lt $lv.Columns.Count; $i++) { $cm[$lv.Columns[$i].Text] = $i }
                $idx = if ($cm.ContainsKey($colName)) { $cm[$colName] } else { -1 }
                if ($idx -lt 0) { return }
                $sb = [System.Text.StringBuilder]::new()
                $nameIdx = if ($cm.ContainsKey("File Name")) { $cm["File Name"] } elseif ($cm.ContainsKey("DisplayName")) { $cm["DisplayName"] } else { -1 }
                foreach ($item in $lv.SelectedItems) {
                    if ($colName -eq "Registry Path") { $val = $item.Tag -replace '^HKLM:', 'HKLM\' -replace '^HKU:', 'HKU\' -replace '^HKCU:', 'HKCU\' }
                    else                              { $val = $item.SubItems[$idx].Text }
                    if ($optNameRef.Checked -and $nameIdx -ge 0) {
                        $n = $item.SubItems[$nameIdx].Text
                        if (![string]::IsNullOrWhiteSpace($n)) { $val = "$n - $val" }
                    }
                    [void]$sb.AppendLine($val)
                }
                if ($sb.Length -gt 0) { [System.Windows.Forms.Clipboard]::SetText($sb.ToString()) }
            }
        })
        [void]$items.Add($menuItems)
    }
    # =========================================== EXPORT (all tabs)
    [void](& $addSep)
    [void](& $addItem "Export"     { ListExport -listview $this.GetCurrentParent().Tag -all $false })
    [void](& $addItem "Export All" { ListExport -listview $this.GetCurrentParent().Tag -all $true })
    # =========================================== EXPLORE TAB 2 (file operations + shell items)
    if ($colMap.ContainsKey("Path")) {
        [void](& $addSep)
        $pIdx = $colMap["Path"]
        # ── File operation items ──
        [void](& $addItem "Copy File" ({
            $lv = $this.GetCurrentParent().Tag
            if ($lv.SelectedItems.Count -gt 0) {
                $coll = [System.Collections.Specialized.StringCollection]::new()
                foreach ($i in $lv.SelectedItems) { [void]$coll.Add($i.SubItems[$pIdx].Text) }
                [System.Windows.Forms.Clipboard]::SetFileDropList($coll)
            }
        }.GetNewClosure()))
        [void](& $addItem "Open Parent Folder" ({
            $lv = $this.GetCurrentParent().Tag
            if ($lv.SelectedItems.Count -gt 0) {
                foreach ($i in $lv.SelectedItems) { [System.Diagnostics.Process]::Start("explorer.exe", "/select,`"$($i.SubItems[$pIdx].Text)`"") }
            }
        }.GetNewClosure()))
        [void](& $addItem "Open File" ({
            $lv = $this.GetCurrentParent().Tag
            if ($lv.SelectedItems.Count -gt 0) {
                foreach ($i in $lv.SelectedItems) { [System.Diagnostics.Process]::Start($i.SubItems[$pIdx].Text) }
            }
        }.GetNewClosure()))
        # ── Shell / Remote session items ──
        # Separator before shell items (visibility controlled in Opening handler)
        $sepShell = [System.Windows.Forms.ToolStripSeparator]::new()
        $sepShell.Tag = "ShellSep"
        [void]$items.Add($sepShell)
        # CMD here : local paths only, single selection only
        $miCmd = & $addItem "CMD here" ({
            $lv = $this.GetCurrentParent().Tag
            if ($lv.SelectedItems.Count -eq 1) {
                [System.Diagnostics.Process]::Start("cmd.exe", "/k cd /d `"$([IO.Path]::GetDirectoryName($lv.SelectedItems[0].SubItems[$pIdx].Text))`"")
            }
        }.GetNewClosure())
        $miCmd.Tag = "CmdNonAdmin"
        # CMD Admin here : local paths only, single selection only
        $miCmdAdmin = & $addItem "CMD Admin here" ({
            $lv = $this.GetCurrentParent().Tag
            if ($lv.SelectedItems.Count -eq 1) {
                $psi = [System.Diagnostics.ProcessStartInfo]::new()
                $psi.FileName        = "cmd.exe"
                $psi.Arguments       = "/k cd /d `"$([IO.Path]::GetDirectoryName($lv.SelectedItems[0].SubItems[$pIdx].Text))`""
                $psi.Verb            = "runas"
                $psi.UseShellExecute = $true
                [System.Diagnostics.Process]::Start($psi)
            }
        }.GetNewClosure())
        $miCmdAdmin.Tag = "CmdHere"
        # PowerShell Admin here : local paths only, single selection only
        $miPsAdmin = & $addItem "PowerShell Admin here" ({
            $lv = $this.GetCurrentParent().Tag
            if ($lv.SelectedItems.Count -eq 1) {
                $dirPath = [IO.Path]::GetDirectoryName($lv.SelectedItems[0].SubItems[$pIdx].Text)
                $psi = [System.Diagnostics.ProcessStartInfo]::new()
                $psi.FileName        = "powershell.exe"
                $psi.Arguments       = "-NoExit -Command `"Set-Location -LiteralPath '$dirPath'`""
                $psi.Verb            = "runas"
                $psi.UseShellExecute = $true
                [System.Diagnostics.Process]::Start($psi)
            }
        }.GetNewClosure())
        $miPsAdmin.Tag = "CmdHere"
        # PSSession here : remote admin share paths (\\server\c$\...) only, single selection only
        $miPsSession = & $addItem "PSSession here" ({
            $lv = $this.GetCurrentParent().Tag
            if ($lv.SelectedItems.Count -eq 1) {
                $filePath = $lv.SelectedItems[0].SubItems[$pIdx].Text
                if ($filePath -match '^\\\\([^\\]+)\\([A-Za-z])\$(.*)') {
                    $server  = $matches[1]
                    $dirPath = [IO.Path]::GetDirectoryName("$($matches[2]):$($matches[3])")
                    Start-RemotePSSession -Server $server -RemotePath $dirPath
                }
            }
        }.GetNewClosure())
        $miPsSession.Tag = "PsSessionHere"
    }
    # =========================================== SEARCH MENU ITEMS (cross-tab navigation)
    if ($colMap.ContainsKey("GUID")) {
        $isExploreTab  = $colMap.ContainsKey("Path") -and -not $colMap.ContainsKey("Registry Path")
        $isRegistryTab = $colMap.ContainsKey("Registry Path")
        $searchSep     = [System.Windows.Forms.ToolStripSeparator]::new()
        $searchSep.Tag = "SearchSeparator"
        [void]$items.Add($searchSep)
        if ($isExploreTab) {
            $menuItem_ExploreUninstall     = [System.Windows.Forms.ToolStripMenuItem]::new("Search in Explore Uninstall Keys")
            $menuItem_ExploreUninstall.Tag = @{ ListView = $listView; MenuTag = "SearchItem" }
            $menuItem_ExploreUninstall.Add_Click({
                $lv = $this.Tag.ListView
                $cm = @{}; for ($i = 0; $i -lt $lv.Columns.Count; $i++) { $cm[$lv.Columns[$i].Text] = $i }
                $guidIdx = if ($cm.ContainsKey("GUID"))         { $cm["GUID"] }         else { -1 }
                $nameIdx = if ($cm.ContainsKey("Product Name")) { $cm["Product Name"] } else { -1 }
                if ($lv.SelectedItems.Count -eq 0) { return }
                $searchTerms = [System.Collections.Generic.List[string]]::new()
                foreach ($selItem in $lv.SelectedItems) {
                    if ($nameIdx -ge 0) {
                        $n = $selItem.SubItems[$nameIdx].Text
                        if ($n -and $n -ne "None" -and $n -ne "") { if (-not $searchTerms.Contains($n)) { $searchTerms.Add($n) } }
                    }
                    if ($guidIdx -ge 0) {
                        $g = $selItem.SubItems[$guidIdx].Text
                        if ($g -and $g -ne "None" -and $g -ne "") { if (-not $searchTerms.Contains($g)) { $searchTerms.Add($g) } }
                    }
                }
                if ($searchTerms.Count -eq 0) { return }
                if ($searchTextBox_Tab3.Tag -is [hashtable]) { $searchTextBox_Tab3.Tag = @{ Text = $searchTextBox_Tab3.Tag.Text; IsPlaceholder = $false } }
                $searchTextBox_Tab3.ForeColor = [System.Drawing.Color]::Black
                $searchTextBox_Tab3.Text = ($searchTerms -join ";")
                $tabControl.SelectedTab = $tabPage3
                $searchButton_Registry.PerformClick()
            })
            [void]$items.Add($menuItem_ExploreUninstall)
        }
        $menuItem_MSICleanup     = [System.Windows.Forms.ToolStripMenuItem]::new("Search in MSI Residues Cleanup")
        $menuItem_MSICleanup.Tag = @{ ListView = $listView; IsRegistryTab = $isRegistryTab; MenuTag = "SearchItem" }
        $menuItem_MSICleanup.Add_Click({
            $lv       = $this.Tag.ListView
            $isRegTab = $this.Tag.IsRegistryTab
            $cm = @{}; for ($i = 0; $i -lt $lv.Columns.Count; $i++) { $cm[$lv.Columns[$i].Text] = $i }
            $guidIdx = if ($cm.ContainsKey("GUID")) { $cm["GUID"] } else { -1 }
            if ($lv.SelectedItems.Count -eq 0) { return }
            $guids = [System.Collections.Generic.List[string]]::new()
            foreach ($selItem in $lv.SelectedItems) {
                if ($guidIdx -ge 0) {
                    $g = $selItem.SubItems[$guidIdx].Text
                    if ($g -and $g -ne "None" -and $g -ne "") { $guids.Add($g) }
                }
            }
            $titleTextBox_MSICleanupTab.Text = ""
            $guidTextBox_MSICleanupTab.Text  = ""
            if ($isRegTab) {
                $nameIdx = if ($cm.ContainsKey("DisplayName")) { $cm["DisplayName"] } else { -1 }
                if ($nameIdx -ge 0) {
                    $names = [System.Collections.Generic.List[string]]::new()
                    foreach ($selItem in $lv.SelectedItems) {
                        $n = $selItem.SubItems[$nameIdx].Text
                        if ($n -and $n -ne "") { $names.Add($n) }
                    }
                    if ($names.Count -gt 0) { $titleTextBox_MSICleanupTab.Text = ($names -join ";") }
                }
            }
            else {
                $nameIdx = if ($cm.ContainsKey("Product Name")) { $cm["Product Name"] } else { -1 }
                if ($nameIdx -ge 0) {
                    $names = [System.Collections.Generic.List[string]]::new()
                    foreach ($selItem in $lv.SelectedItems) {
                        $n = $selItem.SubItems[$nameIdx].Text
                        if ($n -and $n -ne "None" -and $n -ne "") { $names.Add($n) }
                    }
                    if ($names.Count -gt 0) { $titleTextBox_MSICleanupTab.Text = ($names -join ";") }
                }
            }
            if ($guids.Count -gt 0) { $guidTextBox_MSICleanupTab.Text = ($guids -join ";") }
            $tabControl.SelectedTab = $tabPage4
            $searchButton_MSICleanupTab.PerformClick()
        })
        [void]$items.Add($menuItem_MSICleanup)
        if ($isExploreTab) {
            $menuItem_ShowMsiProps     = [System.Windows.Forms.ToolStripMenuItem]::new("Show .MSI Properties")
            $menuItem_ShowMsiProps.Tag = @{ ListView = $listView; MenuTag = "SearchItem" }
            $menuItem_ShowMsiProps.Add_Click({
                $lv = $this.Tag.ListView
                $cm = @{}; for ($i = 0; $i -lt $lv.Columns.Count; $i++) { $cm[$lv.Columns[$i].Text] = $i }
                $pathIdx = if ($cm.ContainsKey("Path")) { $cm["Path"] } else { -1 }
                if ($lv.SelectedItems.Count -eq 0 -or $pathIdx -lt 0) { return }
                $paths = [System.Collections.Generic.List[string]]::new()
                foreach ($selItem in $lv.SelectedItems) {
                    $fp = $selItem.SubItems[$pathIdx].Text
                    if (![string]::IsNullOrWhiteSpace($fp)) { $paths.Add($fp) }
                }
                if ($paths.Count -gt 0) {
                    $script:fromBrowseButton = $true
                    $textBoxPath_Tab1.Text = ($paths -join ";")
                    $tabControl.SelectedTab = $tabPage1
                    $findGuidButton.PerformClick()
                }
            })
            [void]$items.Add($menuItem_ShowMsiProps)
            # ── Opening handler for search items visibility (Explore tab only) ──
            # Controls whether "Search in..." and "Show .MSI Properties" items are shown
            # based on whether the selected items have valid GUID or Product Name data
            $contextMenu.Add_Opening({
                $lv = $this.Tag
                $cm = @{}; for ($i = 0; $i -lt $lv.Columns.Count; $i++) { $cm[$lv.Columns[$i].Text] = $i }
                $guidIdx = if ($cm.ContainsKey("GUID"))         { $cm["GUID"] }         else { -1 }
                $nameIdx = if ($cm.ContainsKey("Product Name")) { $cm["Product Name"] } else { -1 }
                $hasValidSearchData = $false
                foreach ($selItem in $lv.SelectedItems) {
                    if ($hasValidSearchData) { break }
                    if ($guidIdx -ge 0) {
                        $g = $selItem.SubItems[$guidIdx].Text
                        if ($g -and $g -ne "None" -and $g -ne "") { $hasValidSearchData = $true }
                    }
                    if (-not $hasValidSearchData -and $nameIdx -ge 0) {
                        $n = $selItem.SubItems[$nameIdx].Text
                        if ($n -and $n -ne "None" -and $n -ne "") { $hasValidSearchData = $true }
                    }
                }
                foreach ($mi in $this.Items) {
                    if ($mi -is [System.Windows.Forms.ToolStripItem] -and $mi.Tag -is [hashtable] -and $mi.Tag.MenuTag -eq "SearchItem") { $mi.Visible = $hasValidSearchData }
                    if ($mi.Tag -eq "SearchSeparator") { $mi.Visible = $hasValidSearchData }
                }
            }.GetNewClosure())
        }
    }
    # =========================================== MAIN OPENING HANDLER
    # This handler runs every time the context menu is about to show.
    # It controls the global visibility of options based on selection state,
    # and specifically handles shell item visibility (CMD/PS/PSSession) based on:
    #   - Whether exactly 1 item is selected (multi-select = no shell items at all)
    #   - Whether the selected file path is local or remote
    #   - If remote: whether it's an admin share (\\server\c$\...) -> PSSession only
    $contextMenu.Add_Opening({
        param($s, $e)
        $lv              = $this.Tag
        $hasItems        = ($lv.Items.Count -gt 0)
        $hasSelection    = ($lv.SelectedItems.Count -gt 0)
        $singleSelection = ($lv.SelectedItems.Count -eq 1)
        # No items at all -> cancel the menu entirely
        if (-not $hasItems) { $e.Cancel = $true ; return }
        # No selection -> only show "Export All", hide everything else
        if (-not $hasSelection) {
            foreach ($mi in $this.Items) {
                if ($mi -is [System.Windows.Forms.ToolStripMenuItem] -and $mi.Text -eq "Export All") { $mi.Visible = $true }
                else { $mi.Visible = $false }
            }
            return
        }
        # Has selection -> show all non-search-managed items by default
        foreach ($mi in $this.Items) {
            if ($mi -is [System.Windows.Forms.ToolStripSeparator]) { $mi.Visible = $true ; continue }
            if ($mi -is [System.Windows.Forms.ToolStripItem]) {
                # Search items are managed by the other Opening handler above, skip them here
                $isSearchManaged = ($mi.Tag -is [hashtable] -and $mi.Tag.MenuTag -eq "SearchItem") -or ($mi.Tag -eq "SearchSeparator")
                if ($isSearchManaged) { continue }
                # "Open Regedit here" only makes sense with a single selection
                if ($mi.Text -eq "Open Regedit here") { $mi.Visible = $singleSelection }
                else { $mi.Visible = $true }
            }
        }
        # ── Shell items visibility (CMD / PowerShell / PSSession) ──
        # Only applies to ListViews that have a "Path" column (Explore tab)
        $cm2 = @{}; for ($ci = 0; $ci -lt $lv.Columns.Count; $ci++) { $cm2[$lv.Columns[$ci].Text] = $ci }
        if ($cm2.ContainsKey("Path")) {
            # Default: hide everything (covers multi-selection case)
            $showLocal     = $false
            $showPsSession = $false

            # Only show shell items when exactly 1 item is selected
            if ($singleSelection) {
                $firstPath    = $lv.SelectedItems[0].SubItems[$cm2["Path"]].Text
                $isAdminShare = $firstPath -match '^\\\\[^\\]+\\[A-Za-z]\$'

                # Local path  -> CMD here, CMD Admin here, PowerShell Admin here
                # Admin share -> PSSession here
                $showLocal     = (-not $firstPath.StartsWith("\\"))
                $showPsSession = $isAdminShare
            }
            $anyShellItem = $showLocal -or $showPsSession
            # Apply visibility to each shell item based on its Tag
            foreach ($mi in $this.Items) {
                switch ($mi.Tag) {
                    "CmdNonAdmin"   { $mi.Visible = $showLocal -and (-not $isAdmin) }
                    "CmdHere"       { $mi.Visible = $showLocal }
                    "PsSessionHere" { $mi.Visible = $showPsSession }
                    "ShellSep"      { $mi.Visible = $anyShellItem }
                }
            }
        }
    }.GetNewClosure())
    $listView.ContextMenuStrip = $contextMenu
}

Function Open-RegeditHere {
    param([string]$Path, [string]$ComputerName)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $isRemote    = -not [string]::IsNullOrWhiteSpace($ComputerName)
    $autoTimeout = 10000  # Timeout for automated operations (UI polling, connection, navigation)
    $userTimeout = 60000  # Extended timeout when waiting for user interaction (credential prompt)
    # Normalize PowerShell-style short hive names (HKLM:, HKCU:, etc.) to full registry hive paths
    $regPath = $Path -replace '^HKLM[:\\]*', 'HKEY_LOCAL_MACHINE\' `
                      -replace '^HKCU[:\\]*', 'HKEY_CURRENT_USER\' `
                      -replace '^HKU[:\\]*',  'HKEY_USERS\' `
                      -replace '^HKCR[:\\]*', 'HKEY_CLASSES_ROOT\' `
                      -replace '^HKCC[:\\]*', 'HKEY_CURRENT_CONFIG\'
    # Win32 interop class for regedit window manipulation (singleton, loaded once per session)
    if (-not ([System.Management.Automation.PSTypeName]'RegeditUI').Type) {
        Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class RegeditUI {
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "SendMessage")] public static extern IntPtr SendMessageStr(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string wnd);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);
    [DllImport("user32.dll")] public static extern IntPtr GetMenu(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetSubMenu(IntPtr hMenu, int nPos);
    [DllImport("user32.dll")] public static extern int GetMenuItemCount(IntPtr hMenu);
    [DllImport("user32.dll")] public static extern uint GetMenuItemID(IntPtr hMenu, int nPos);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetMenuString(IntPtr hMenu, uint uIDItem, StringBuilder sb, int nMax, uint uFlag);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP   = 0x0101;
    public const uint WM_SETTEXT = 0x000C;
    public const uint WM_COMMAND = 0x0111;
    public const uint BM_CLICK   = 0x00F5;
    public const uint MF_BYPOSITION = 0x0400;
    public const int VK_TAB    = 0x09;
    public const int VK_RETURN = 0x0D;
    public const int VK_HOME   = 0x24;
    public const int VK_LEFT   = 0x25;
    public const int VK_DOWN   = 0x28;
    public const int VK_DELETE = 0x2E;
    public static string GetText(IntPtr hWnd) {
        var sb = new StringBuilder(512); GetWindowText(hWnd, sb, 512); return sb.ToString();
    }
    public static string GetClass(IntPtr hWnd) {
        var sb = new StringBuilder(256); GetClassName(hWnd, sb, 256); return sb.ToString();
    }
    public static void SendKey(IntPtr hWnd, int vk) {
        PostMessage(hWnd, WM_KEYDOWN, (IntPtr)vk, IntPtr.Zero);
        System.Threading.Thread.Sleep(30);
        PostMessage(hWnd, WM_KEYUP, (IntPtr)vk, IntPtr.Zero);
    }
    public static void SetText(IntPtr hWnd, string text) {
        SendMessageStr(hWnd, WM_SETTEXT, IntPtr.Zero, text);
    }
    // Enumerate top-level windows to find the first #32770 dialog owned by a given PID
    public static IntPtr FindDialogByPid(int pid) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hWnd, lp) => {
            uint wndPid;
            GetWindowThreadProcessId(hWnd, out wndPid);
            if ((int)wndPid == pid && GetClass(hWnd) == "#32770") {
                found = hWnd; return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
    // Check if any visible window exists for a given process name (used for credential prompt detection)
    public static bool ProcessHasWindow(string processName) {
        bool found = false;
        System.Diagnostics.Process[] procs;
        try { procs = System.Diagnostics.Process.GetProcessesByName(processName); }
        catch { return false; }
        foreach (var p in procs) {
            if (p.MainWindowHandle != IntPtr.Zero) { found = true; }
            p.Dispose();
        }
        return found;
    }
    // Scan the File menu to find the "Connect Network Registry" command ID (works on EN/FR locales)
    public static uint FindConnectCommandId(IntPtr mainWnd) {
        IntPtr hMenu = GetMenu(mainWnd);
        if (hMenu == IntPtr.Zero) return 0;
        IntPtr fileMenu = GetSubMenu(hMenu, 0);
        if (fileMenu == IntPtr.Zero) return 0;
        int count = GetMenuItemCount(fileMenu);
        var sb = new StringBuilder(256);
        for (int i = 0; i < count; i++) {
            sb.Clear();
            GetMenuString(fileMenu, (uint)i, sb, 256, MF_BYPOSITION);
            string text = sb.ToString().ToLower();
            if (text.Contains("connect") || text.Contains("connexion") || text.Contains("network") || text.Contains("r\u00e9seau")) {
                return GetMenuItemID(fileMenu, i);
            }
        }
        return 0;
    }
    // Find the first child window matching a given class name
    public static IntPtr FindChildByClass(IntPtr parent, string className) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(parent, (hWnd, lp) => {
            if (GetClass(hWnd) == className) { found = hWnd; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
"@
    }
    # ── Launch regedit with /m (multi-instance mode) to guarantee a new window ──
    $launchTime = Get-Date
    Start-Process "regedit.exe" "/m"
    # Poll until we find a regedit window that was created after our launch timestamp
    $hWnd = [IntPtr]::Zero
    $sw   = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $autoTimeout) {
        foreach ($p in (Get-Process -Name regedit -ErrorAction SilentlyContinue)) {
            if ($p.MainWindowHandle -ne [IntPtr]::Zero -and $p.StartTime -ge $launchTime) {
                $hWnd = $p.MainWindowHandle ; break
            }
        }
        if ($hWnd -ne [IntPtr]::Zero) { break }
        Start-Sleep -Milliseconds 50
    }
    $sw.Stop()
    if ($hWnd -eq [IntPtr]::Zero) {
        Write-Host "[FAIL] Regedit not detected" -ForegroundColor Red ; return
    }
    $treeView   = [RegeditUI]::FindWindowEx($hWnd, [IntPtr]::Zero, "SysTreeView32", $null)
    $regeditPid = (Get-Process -Name regedit | Where-Object { $_.MainWindowHandle -eq $hWnd }).Id
    # ── Scriptblock : navigate regedit via address bar ──
    # Sequence : Tab -> Tab (focus address bar) -> Delete (enter edit mode) -> SetText -> Enter
    $navigateAddressBar = {
        param([IntPtr]$MainWnd, [string]$TargetPath, [int]$Timeout)
        # Two Tab keystrokes to move focus from TreeView to address bar
        [RegeditUI]::SendKey($MainWnd, [RegeditUI]::VK_TAB)
        Start-Sleep -Milliseconds 100
        [RegeditUI]::SendKey($MainWnd, [RegeditUI]::VK_TAB)
        Start-Sleep -Milliseconds 200
        # Delete triggers edit mode in the address bar, which spawns an Edit control
        [RegeditUI]::SendKey($MainWnd, [RegeditUI]::VK_DELETE)
        Start-Sleep -Milliseconds 100
        # Poll for the Edit control that appears when the address bar enters edit mode
        $edit = [IntPtr]::Zero
        $sw2  = [System.Diagnostics.Stopwatch]::StartNew()
        while ($sw2.ElapsedMilliseconds -lt $Timeout) {
            $edit = [RegeditUI]::FindChildByClass($MainWnd, "Edit")
            if ($edit -ne [IntPtr]::Zero) { break }
            Start-Sleep -Milliseconds 50
        }
        $sw2.Stop()
        if ($edit -eq [IntPtr]::Zero) {
            Write-Host "      [FAIL] Address bar Edit control not found" -ForegroundColor Red
            return $false
        }
        # Inject the target path into the Edit control and confirm with Enter
        [RegeditUI]::SetText($edit, $TargetPath)
        Start-Sleep -Milliseconds 50
        [RegeditUI]::SendKey($edit, [RegeditUI]::VK_RETURN)
        return $true
    }
    # ══════════════════════════════════════════════════════════════════════
    #  LOCAL MODE
    # ══════════════════════════════════════════════════════════════════════
    if (-not $isRemote) {
        # Check if any regedit instance was already running BEFORE our /m launch
        $existingRegedit = @(Get-Process -Name regedit -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -lt $launchTime })
        if ($existingRegedit.Count -eq 0) {
            # No pre-existing instance : we can use the LastKey registry trick
            # This makes regedit open directly at the target path on launch (instant, no UI automation)
            # Step 1 : kill the /m instance we launched (it was only needed to confirm regedit works)
            try { Stop-Process -Id $regeditPid -Force -ErrorAction SilentlyContinue } catch { }
            Start-Sleep -Milliseconds 200
            # Step 2 : write the target path into the LastKey registry value
            $regKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit'
            if (-not (Test-Path $regKeyPath)) { New-Item -Path $regKeyPath -Force | Out-Null }
            Set-ItemProperty -Path $regKeyPath -Name 'LastKey' -Value $regPath -Force
            # Step 3 : relaunch regedit normally (without /m) — it reads LastKey and opens there
            # Without /m, regedit would just focus an existing window instead of opening a new one,
            # which is why this trick only works when no other instance is running
            Start-Process "regedit.exe"
            return
        }
        # Pre-existing instance detected : we must keep the /m window and navigate via address bar
        # Small delay to let the window fully initialize its internal controls
        Start-Sleep -Milliseconds 500
        & $navigateAddressBar $hWnd $regPath $autoTimeout
        return
    }
    # ══════════════════════════════════════════════════════════════════════
    #  REMOTE MODE
    # ══════════════════════════════════════════════════════════════════════
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  Open-RegeditHere (Remote)" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  Path     : $regPath" -ForegroundColor Gray
    Write-Host "  Computer : $ComputerName" -ForegroundColor Gray
    Write-Host ""
    # Remote registry only supports HKLM and HKU hives
    if ($regPath -notmatch '^HKEY_LOCAL_MACHINE\\|^HKEY_USERS\\') {
        Write-Host "[FAIL] Remote registry only supports HKEY_LOCAL_MACHINE and HKEY_USERS" -ForegroundColor Red
        return
    }
    # ── Step 1 : Collapse the local tree to avoid visual clutter ──
    Write-Host "[1/5] Collapsing local tree..." -ForegroundColor Yellow
    [RegeditUI]::SendKey($treeView, [RegeditUI]::VK_HOME)
    Start-Sleep -Milliseconds 150
    [RegeditUI]::SendKey($treeView, [RegeditUI]::VK_LEFT)
    Write-Host "      Done" -ForegroundColor Green
    # ── Step 2 : Trigger "Connect Network Registry" via the File menu ──
    Write-Host "[2/5] Connect Network Registry..." -ForegroundColor Yellow
    $cmdId = [RegeditUI]::FindConnectCommandId($hWnd)
    if ($cmdId -eq 0) {
        Write-Host "      [FAIL] Menu command not found" -ForegroundColor Red ; return
    }
    # Foreground is required for WM_COMMAND to be processed by the menu system
    [RegeditUI]::SetForegroundWindow($hWnd) | Out-Null
    Start-Sleep -Milliseconds 200
    [RegeditUI]::PostMessage($hWnd, [RegeditUI]::WM_COMMAND, [IntPtr]$cmdId, [IntPtr]::Zero) | Out-Null
    # Poll for the connection dialog (identified as #32770 class owned by regedit PID)
    $dialog = [IntPtr]::Zero
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $autoTimeout) {
        $dialog = [RegeditUI]::FindDialogByPid($regeditPid)
        if ($dialog -ne [IntPtr]::Zero) { break }
        Start-Sleep -Milliseconds 150
    }
    $sw.Stop()
    if ($dialog -eq [IntPtr]::Zero) {
        Write-Host "      [FAIL] No dialog detected" -ForegroundColor Red ; return
    }
    Write-Host "      Dialog : $([RegeditUI]::GetText($dialog))" -ForegroundColor Green
    # ── Step 3 : Inject the computer name into the dialog and submit ──
    Write-Host "[3/5] Submitting : $ComputerName..." -ForegroundColor Yellow
    # The dialog uses a RICHEDIT50W control for the computer name input field
    $richEdit = [RegeditUI]::FindChildByClass($dialog, "RICHEDIT50W")
    if ($richEdit -eq [IntPtr]::Zero) {
        Write-Host "      [FAIL] RICHEDIT50W not found" -ForegroundColor Red ; return
    }
    [RegeditUI]::SetText($richEdit, $ComputerName)
    Start-Sleep -Milliseconds 50
    # Try to find and click the OK button ; fallback to Enter if the button isn't found
    $okBtn = [RegeditUI]::FindWindowEx($dialog, [IntPtr]::Zero, "Button", "OK")
    if ($okBtn -ne [IntPtr]::Zero) { [RegeditUI]::PostMessage($okBtn, [RegeditUI]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null }
    else                           { [RegeditUI]::SendKey($dialog,    [RegeditUI]::VK_RETURN) }
    Write-Host "      Submitted" -ForegroundColor Green
    # ── Step 4 : Wait for the remote connection to complete ──
    Write-Host "[4/5] Waiting for connection..." -ForegroundColor Yellow
    $waitingForUser = $false
    $currentTimeout = $autoTimeout
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $currentTimeout) {
        if (-not [RegeditUI]::IsWindow($dialog)) {
            # Dialog closed — check if regedit spawned an error dialog instead
            $errDlg = [RegeditUI]::FindDialogByPid($regeditPid)
            if ($errDlg -ne [IntPtr]::Zero) {
                Write-Host "      [FAIL] $([RegeditUI]::GetText($errDlg))" -ForegroundColor Red
                [RegeditUI]::SendKey($errDlg, [RegeditUI]::VK_RETURN)
                return
            }
            Write-Host "      Connected" -ForegroundColor Green
            break
        }
        # If Windows credential prompt appears (CredentialUIBroker.exe), switch to user timeout
        if (-not $waitingForUser -and [RegeditUI]::ProcessHasWindow("CredentialUIBroker")) {
            Write-Host "      Credential prompt detected, waiting for user ($($userTimeout / 1000)s)..." -ForegroundColor DarkYellow
            $waitingForUser = $true
            $currentTimeout = $userTimeout
            $sw.Restart()
        }
        Start-Sleep -Milliseconds 100
    }
    $sw.Stop()
    if ([RegeditUI]::IsWindow($dialog)) {
        Write-Host "      [FAIL] Timeout" -ForegroundColor Red ; return
    }
    # ── Step 5 : Navigate to the target path on the remote registry ──
    Write-Host "[5/5] Navigating to remote path..." -ForegroundColor Yellow
    Start-Sleep -Milliseconds 300
    # Move down to select the remote computer node (just connected, appears below local Computer)
    [RegeditUI]::SendKey($treeView, [RegeditUI]::VK_DOWN)
    Start-Sleep -Milliseconds 200
    # Navigate via address bar using the remote path format : \\ComputerName\HKEY_...\...
    $remotePath = "$ComputerName\$regPath"
    $result = & $navigateAddressBar $hWnd $remotePath $autoTimeout
    if ($result -ne $false) {
        # Brief wait to detect potential error dialogs (e.g. access denied on the remote key)
        Start-Sleep -Milliseconds 300
        $errDlg = [RegeditUI]::FindDialogByPid($regeditPid)
        if ($errDlg -ne [IntPtr]::Zero) {
            Write-Host "      [WARN] $([RegeditUI]::GetText($errDlg))" -ForegroundColor DarkYellow
            [RegeditUI]::SendKey($errDlg, [RegeditUI]::VK_RETURN)
        }
        else {
            Write-Host ""
            Write-Host "[DONE] $ComputerName -> $regPath" -ForegroundColor Green
        }
    }
}

function Open-FileSystemPath {
    param(
        [string]$Path,
        [string]$ComputerName = "",
        [ValidateSet('Explore', 'Select', 'Properties')][string]$Action = 'Explore'
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $isRemote   = (-not [string]::IsNullOrWhiteSpace($ComputerName) -and $ComputerName -ne $env:COMPUTERNAME)
    $targetPath = $Path
    if ($isRemote) {
        $targetPath = $Path -replace '^([A-Za-z]):\\', "\\$ComputerName\`$1`$\"
        Write-Log "Open-FileSystemPath : Remote target $targetPath (computer : $ComputerName)"
        if (-not ([System.IO.File]::Exists($targetPath) -or [System.IO.Directory]::Exists($targetPath))) {
            Write-Log "Open-FileSystemPath : Path not found : $targetPath" -Level Warning
            [System.Windows.Forms.MessageBox]::Show("Path not found :`n$targetPath", "Path Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }
    else {
        if (-not ([System.IO.File]::Exists($targetPath) -or [System.IO.Directory]::Exists($targetPath))) {
            Write-Log "Open-FileSystemPath : Local path not found : $targetPath" -Level Warning
            [System.Windows.Forms.MessageBox]::Show("Path not found :`n$targetPath", "Path Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }
    $fileExists = [System.IO.File]::Exists($targetPath)
    Write-Log "Open-FileSystemPath : Action=$Action, isFile=$fileExists : $targetPath"
    switch ($Action) {
        'Select' {
            if ($fileExists) { [System.Diagnostics.Process]::Start("explorer.exe", "/select,`"$targetPath`"") }
            else             { [System.Diagnostics.Process]::Start("explorer.exe", "`"$targetPath`"") }
        }
        'Explore' {
            [System.Diagnostics.Process]::Start("explorer.exe", "`"$targetPath`"")
        }
        'Properties' {
            try {
                $shell  = New-Object -ComObject Shell.Application
                $parent = $shell.NameSpace([System.IO.Path]::GetDirectoryName($targetPath))
                $item   = if ($parent) { $parent.ParseName([System.IO.Path]::GetFileName($targetPath)) }
                if ($item) { $item.InvokeVerb("Properties") }
                else { Write-Log "Open-FileSystemPath : Shell.Application could not resolve $targetPath" -Level Warning }
            }
            catch { Write-Log "Open-FileSystemPath : Properties error : $($_.Exception.Message)" -Level Warning }
        }
    }
}

#region Network functions

# P/Invoke for Windows Network and Credential Management
if (-not ([System.Management.Automation.PSTypeName]'NetSession.Native').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
namespace NetSession {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct NETRESOURCE {
        public int    dwScope;
        public int    dwType;
        public int    dwDisplayType;
        public int    dwUsage;
        public string lpLocalName;
        public string lpRemoteName;
        public string lpComment;
        public string lpProvider;
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct CREDENTIAL {
        public int    Flags;
        public int    Type;
        public string TargetName;
        public string Comment;
        public long   LastWritten;
        public int    CredentialBlobSize;
        public IntPtr CredentialBlob;
        public int    Persist;
        public int    AttributeCount;
        public IntPtr Attributes;
        public string TargetAlias;
        public string UserName;
    }
    public static class Native {
        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        public static extern int WNetAddConnection2(ref NETRESOURCE netResource, string password, string username, int flags);
        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        public static extern int WNetCancelConnection2(string name, int flags, bool force);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredWrite(ref CREDENTIAL credential, int flags);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredDelete(string targetName, int type, int flags);
        public static int ConnectSmb(string target, string username, string password) {
            NETRESOURCE nr = new NETRESOURCE();
            nr.dwType       = 0;
            nr.lpRemoteName = @"\\" + target + @"\IPC$";
            nr.lpLocalName  = null;
            return WNetAddConnection2(ref nr, password, username, 4);
        }
        public static int DisconnectSmb(string target) {
            return WNetCancelConnection2(@"\\" + target + @"\IPC$", 0, true);
        }
        public static bool SaveDomainCredential(string target, string username, string password) {
            byte[] passwordBytes = Encoding.Unicode.GetBytes(password);
            IntPtr passwordPtr = Marshal.AllocHGlobal(passwordBytes.Length);
            try {
                Marshal.Copy(passwordBytes, 0, passwordPtr, passwordBytes.Length);
                CREDENTIAL cred = new CREDENTIAL();
                cred.Type              = 2; // CRED_TYPE_DOMAIN_PASSWORD
                cred.TargetName        = target;
                cred.UserName          = username;
                cred.CredentialBlob     = passwordPtr;
                cred.CredentialBlobSize = passwordBytes.Length;
                cred.Persist           = 2; // CRED_PERSIST_LOCAL_MACHINE
                return CredWrite(ref cred, 0);
            }
            finally {
                Marshal.FreeHGlobal(passwordPtr);
            }
        }
        public static bool DeleteDomainCredential(string target) {
            return CredDelete(target, 2, 0);
        }
    }
}
'@
}

function Reset-SecurePassword {
    if ($script:SecurePassword) { $script:SecurePassword.Dispose() }
    $script:SecurePassword = New-Object System.Security.SecureString
}
function Set-PasswordPlaceholderState {
    param([bool]$ShowPlaceholder)
    if ($ShowPlaceholder) {
        $textBox_CredentialPwd.UseSystemPasswordChar = $false
        $textBox_CredentialPwd.Text      = $script:PasswordPlaceholder
        $textBox_CredentialPwd.ForeColor = [System.Drawing.Color]::Gray
        $textBox_CredentialPwd.Tag       = @{ IsPlaceholder = $true }
    }
    else {
        $textBox_CredentialPwd.UseSystemPasswordChar = $true
        $textBox_CredentialPwd.ForeColor = [System.Drawing.Color]::Black
        $textBox_CredentialPwd.Tag       = @{ IsPlaceholder = $false }
    }
}

function Test-WinRMReady {
    if ($script:WinRMChecked) { return $true }
    try {
        $svc = Get-Service -Name WinRM -ErrorAction Stop
        if ($svc.Status -eq 'Running') {
            $script:WinRMChecked = $true
            Write-Log "WinRM service is running"
            return $true
        }
    }
    catch {
        Write-Log "WinRM service not found : $_" -Level Error
        return $false
    }
    Write-Log "WinRM service is not running" -Level Warning
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "WinRM is stopped.`nSet your connection to private/domain and use ""winrm quickconfig"" ?",
        "WinRM Configuration Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Log "User declined WinRM configuration" -Level Warning
        return $false
    }
    # Switch any Public network profile to Private
    try {
        $profiles = Get-NetConnectionProfile -ErrorAction Stop
        foreach ($p in $profiles) {
            if ($p.NetworkCategory -eq 'Public') {
                Write-Log "Switching network '$($p.Name)' (InterfaceIndex $($p.InterfaceIndex)) from Public to Private"
                Set-NetConnectionProfile -InterfaceIndex $p.InterfaceIndex -NetworkCategory Private -ErrorAction Stop
                Write-Log "Network '$($p.Name)' set to Private"
            }
        }
    }
    catch {
        Write-Log "Failed to update network profile : $_" -Level Error
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to set network profile to Private :`n$_`n`nWinRM quickconfig may fail.",
            "Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
    # Run winrm quickconfig non-interactively
    try {
        Write-Log "Running winrm quickconfig -force"
        $output = & winrm quickconfig -force 2>&1
        $outputText = ($output | Out-String).Trim()
        Write-Log "winrm quickconfig output : $outputText"
    }
    catch {
        Write-Log "winrm quickconfig failed : $_" -Level Error
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to configure WinRM :`n$_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    Start-Sleep -Milliseconds 800
    # Verify WinRM is now running
    try {
        $svc = Get-Service -Name WinRM -ErrorAction Stop
        if ($svc.Status -eq 'Running') {
            $script:WinRMChecked = $true
            Write-Log "WinRM successfully configured and running"
            return $true
        }
        Write-Log "WinRM status after quickconfig : $($svc.Status)" -Level Warning
    }
    catch {
        Write-Log "Failed to verify WinRM status : $_" -Level Error
    }
    [System.Windows.Forms.MessageBox]::Show(
        "WinRM is still not running after configuration.`nPlease run 'winrm quickconfig' manually in an elevated prompt.",
        "WinRM Configuration Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    return $false
}

function Test-RemoteConnections {
    param(
        [Parameter(Mandatory = $true)]   [string[]]$Computers,
        [System.Management.Automation.PSCredential]$Credential = $null
    )
    # WinRM readiness check (once per session)
    if (-not (Test-WinRMReady)) {
        Write-Log "WinRM not ready, aborting connection tests" -Level Warning
        return @()
    }
    Write-Log "Testing connections to $($Computers.Count) computer(s)"
    $runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, [Math]::Min($Computers.Count, 10))
    $runspacePool.Open()
    $jobs = [System.Collections.ArrayList]::new()
    $scriptBlock = {
        param($computer, [System.Management.Automation.PSCredential]$cred, $needsTrustedHosts)
        $result = [PSCustomObject]@{
            Computer                = $computer
            Success                 = $false
            Status                  = "Unknown"
            Details                 = ""
            RawError                = ""
            IsWinRMFailure          = $false
            IsRemoteRegistryFailure = $false
            IsTrustedHostsIssue     = $false
            ConsoleUser             = ""
            ConsoleUserFullName     = ""
        }
        $isIP = $computer -match '^\d{1,3}(\.\d{1,3}){3}$'
        # Pre-check TrustedHosts for IP addresses
        if ($needsTrustedHosts) {
            $result.IsWinRMFailure      = $true
            $result.IsTrustedHostsIssue = $true
            $result.Status              = "TrustedHosts Required"
            $result.Details             = "IP address must be added to TrustedHosts"
            return $result
        }
        # TCP test
        $tcpClient = New-Object Net.Sockets.TcpClient
        try {
            $asyncResult = $tcpClient.BeginConnect($computer, 445, $null, $null)
            if (-not $asyncResult.AsyncWaitHandle.WaitOne(150, $false)) {
                $tcpClient.Close()
                $result.Status  = "Unreachable"
                $result.Details = "TCP port 445 connection failed"
                return $result
            }
            $tcpClient.EndConnect($asyncResult) | Out-Null
            $tcpClient.Close()
        }
        catch {
            $tcpClient.Close()
            $result.Status  = "Unreachable"
            $result.Details = "TCP connection error"
            return $result
        }
        # WinRM test
        try {
            $testParams = @{ ComputerName = $computer; ErrorAction = 'Stop' }
            if ($cred) { $testParams.Credential = $cred ; $testParams.Authentication = 'Negotiate' }
            Test-WSMan @testParams | Out-Null
        }
        catch {
            $rawError = $_.Exception.Message
            $result.RawError = $rawError
            # Detect Access Denied through exception chain
            $isAccessDenied = $false
            $currentEx = $_.Exception
            while ($currentEx -and -not $isAccessDenied) {
                if ($currentEx.PSObject.Properties['ErrorCode'] -and $currentEx.ErrorCode -eq 5)                  { $isAccessDenied = $true }
                if ($currentEx.GetType().Name -eq 'PSRemotingTransportException' -and $currentEx.ErrorCode -eq 5) { $isAccessDenied = $true }
                if ($currentEx.PSObject.Properties['ErrorRecord'] -and $currentEx.ErrorRecord.Exception) {
                    $embedded = $currentEx.ErrorRecord.Exception
                    if ($embedded.PSObject.Properties['ErrorCode'] -and $embedded.ErrorCode -eq 5)                { $isAccessDenied = $true }
                }
                # HRESULT 0x80070005 = E_ACCESSDENIED
                if ($currentEx.HResult -eq [int]0x80070005)                                                       { $isAccessDenied = $true }
                $currentEx = $currentEx.InnerException
            }
            if ($isAccessDenied) {
                $result.Status  = "Access Denied"
                $result.Details = "Credentials insufficient or rejected by target"
                return $result
            }
            $result.IsWinRMFailure = $true
            $result.Status         = "WinRM Failed"
            if ($rawError -match 'TrustedHosts' -or $rawError -match '2150858981') { $result.IsTrustedHostsIssue = $true }
            if ($rawError -match '<f:Message[^>]*>(.+?)</f:Message>')              { $result.Details = $matches[1] }
            else                                                                   { $result.Details = $rawError }
            return $result
        }
        # Check Remote Registry service status
        $serviceState     = $null
        $serviceStartType = $null
        # sc.exe method (fast, no WinRM overhead, but cannot use PSCredential)
        if (-not $cred) {
            try {
                $scOutput = & sc.exe "\\$computer" query RemoteRegistry 2>&1
                $scString = $scOutput -join "`n"
                if ($scString -match 'STATE\s+:\s+\d+\s+(\w+)') {
                    $serviceState = $matches[1]
                    Write-Information "RemoteRegistry service state on ${computer} via sc.exe : $serviceState"
                }
                else {
                    Write-Information "sc.exe query failed for ${computer} : $scString"
                }
            }
            catch {
                Write-Information "sc.exe exception for ${computer} : $_"
            }
        }
        # Invoke-Command method (supports PSCredential)
        if ($null -eq $serviceState) {
            try {
                Write-Information "Checking RemoteRegistry on $computer via Invoke-Command"
                $svcParams = @{
                    ComputerName = $computer
                    ScriptBlock  = {
                        $svc = Get-Service -Name RemoteRegistry -ErrorAction Stop
                        $wmi = Get-WmiObject Win32_Service -Filter "Name='RemoteRegistry'" -ErrorAction SilentlyContinue
                        return @{
                            State     = $svc.Status.ToString()
                            StartType = if ($wmi) { $wmi.StartMode } else { 'Unknown' }
                        }
                    }
                    ErrorAction = 'Stop'
                }
                if ($cred) { $svcParams.Credential = $cred }
                $svcInfo = Invoke-Command @svcParams
                $serviceState = switch ($svcInfo.State) {
                    'Running' { 'RUNNING' }
                    'Stopped' { 'STOPPED' }
                    default   { $svcInfo.State.ToUpper() }
                }
                $serviceStartType = $svcInfo.StartType
                Write-Information "RemoteRegistry on ${computer} : $serviceState (StartType : $serviceStartType)"
            }
            catch {
                Write-Information "Invoke-Command service check failed for ${computer} : $_"
                $isAccessDenied = $false
                $currentEx = $_.Exception
                while ($currentEx -and -not $isAccessDenied) {
                    if ($currentEx.GetType().Name -eq 'PSRemotingTransportException' -and $currentEx.ErrorCode -eq 5)   { $isAccessDenied = $true }
                    if ($currentEx.PSObject.Properties['ErrorRecord'] -and $currentEx.ErrorRecord.Exception) {
                        $embedded = $currentEx.ErrorRecord.Exception
                        if ($embedded.GetType().Name -eq 'PSRemotingTransportException' -and $embedded.ErrorCode -eq 5) { $isAccessDenied = $true }
                    }
                    $currentEx = $currentEx.InnerException
                }
                if ($isAccessDenied) {
                    $result.Status  = "Access Denied"
                    $result.Details = "Credentials insufficient or rejected by target"
                    return $result
                }
            }
        }
        # Act on service state if we managed to retrieve it
        if ($null -ne $serviceState -and $serviceState -ne 'RUNNING') {
            if ($null -eq $serviceStartType) {
                try {
                    $scConfigOutput = & sc.exe "\\$computer" qc RemoteRegistry 2>&1
                    $scConfigString = $scConfigOutput -join "`n"
                    if ($scConfigString -match 'START_TYPE\s+:\s+\d+\s+(\w+)') { $serviceStartType = $matches[1] }
                }
                catch { }
            }
            if ($null -eq $serviceStartType) { $serviceStartType = 'UNKNOWN' }
            Write-Information "RemoteRegistry not running on ${computer}, start type : $serviceStartType"
            if ($serviceStartType -match 'Disabled|DISABLED') {
                $result.IsRemoteRegistryFailure = $true
                $result.Status  = "RemoteRegistry Disabled"
                $result.Details = "Remote Registry service is disabled on target"
                Write-Information "RemoteRegistry is disabled on $computer"
                return $result
            }
            elseif ($serviceStartType -match 'Manual') {
                Write-Information "RemoteRegistry service state is Manual on ${computer}, it should auto-start on registry access"
            }
            else {
                Write-Information "Attempting to start RemoteRegistry on $computer"
                try {
                    $startParams = @{
                        ComputerName = $computer
                        ScriptBlock  = {
                            try {
                                Set-Service -Name RemoteRegistry -StartupType Manual -ErrorAction Stop
                                Start-Service -Name RemoteRegistry -ErrorAction Stop
                                Start-Sleep -Milliseconds 500
                                $svc = Get-Service -Name RemoteRegistry -ErrorAction Stop
                                return @{ Success = ($svc.Status -eq 'Running'); State = $svc.Status.ToString() }
                            }
                            catch { return @{ Success = $false; State = $_.Exception.Message } }
                        }
                        ErrorAction = 'Stop'
                    }
                    if ($cred) { $startParams.Credential = $cred }
                    $startResult = Invoke-Command @startParams
                    if ($startResult.Success) {
                        Write-Information "RemoteRegistry successfully started on $computer"
                    }
                    else {
                        $result.IsRemoteRegistryFailure = $true
                        $result.Status  = "RemoteRegistry Start Failed"
                        $result.Details = "Could not start Remote Registry service : $($startResult.State)"
                        Write-Information "Failed to start RemoteRegistry on ${computer} : $($startResult.State)"
                        return $result
                    }
                }
                catch {
                    $result.IsRemoteRegistryFailure = $true
                    $result.Status  = "RemoteRegistry Start Failed"
                    $result.Details = "Remote start attempt failed : $($_.Exception.Message)"
                    Write-Information "Invoke-Command start failed for ${computer} : $_"
                    return $result
                }
            }
        }
        # Registry access test via Invoke-Command
        try {
            $invokeParams = @{
                ComputerName = $computer
                ScriptBlock  = {
                    try {
                        $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey("SOFTWARE", $false)
                        if ($key) { $key.Close(); return @{ Success = $true } }
                        return @{ Success = $false; Error = "Could not open SOFTWARE key" }
                    }
                    catch { return @{ Success = $false; Error = $_.Exception.Message } }
                }
                ErrorAction = 'Stop'
            }
            if ($cred) { $invokeParams.Credential = $cred }
            $testResult = Invoke-Command @invokeParams
            if ($testResult.Success) {
                $result.Success = $true
                $result.Status  = "Success"
                $result.Details = "All checks passed"
            }
            else {
                $result.Status  = "Registry Access Denied"
                $result.Details = $testResult.Error
            }
        }
        catch {
            $rawError = $_.Exception.Message
            $result.RawError = $rawError
            $isAccessDenied = $false
            $currentEx = $_.Exception
            while ($currentEx -and -not $isAccessDenied) {
                if ($currentEx.GetType().Name -eq 'PSRemotingTransportException' -and $currentEx.ErrorCode -eq 5)   { $isAccessDenied = $true }
                if ($currentEx.PSObject.Properties['ErrorRecord'] -and $currentEx.ErrorRecord.Exception) {
                    $embedded = $currentEx.ErrorRecord.Exception
                    if ($embedded.GetType().Name -eq 'PSRemotingTransportException' -and $embedded.ErrorCode -eq 5) { $isAccessDenied = $true }
                }
                $currentEx = $currentEx.InnerException
            }
            if ($isAccessDenied) {
                $result.Status  = "Access Denied"
                $result.Details = "Credentials insufficient or rejected by target"
            }
            else {
                $result.IsRemoteRegistryFailure = $true
                $result.Status                  = "Registry Error"
                $result.Details                 = $rawError
            }
        }
        return $result
    }
    $trustedHosts = @(Get-TrustedHostsList)
    $trustedHostsAllowAll = $trustedHosts -contains "*"
    foreach ($computer in $Computers) {
        if ([string]::IsNullOrWhiteSpace($computer)) { continue }
        Write-Log "Queuing connection test for : $computer"
        $isIP = $computer -match '^\d{1,3}(\.\d{1,3}){3}$'
        $needsTrustedHosts = $false
        if ($isIP -and -not $trustedHostsAllowAll -and $trustedHosts -notcontains $computer) {
            $needsTrustedHosts = $true
            Write-Log "IP $computer not in TrustedHosts, will require addition" -Level Debug
        }
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $runspacePool
        [void]$ps.AddScript($scriptBlock)
        [void]$ps.AddArgument($computer)
        [void]$ps.AddArgument($Credential)
        [void]$ps.AddArgument($needsTrustedHosts)
        [void]$jobs.Add(@{
            PowerShell = $ps
            Handle     = $ps.BeginInvoke()
            Computer   = $computer
        })
    }
    $results = [System.Collections.ArrayList]::new()
    foreach ($job in $jobs) {
        try {
            $jobResult = $job.PowerShell.EndInvoke($job.Handle)
            # Drain information stream from runspace into Write-Log
            foreach ($infoRecord in $job.PowerShell.Streams.Information) { Write-Log "$($infoRecord.MessageData)" -Level Debug }
            if ($jobResult) { 
                [void]$results.Add($jobResult)
                if ($jobResult.Success) { Write-Log "Connection to $($jobResult.Computer) : OK" }
                else { Write-Log "Connection to $($jobResult.Computer) : FAILED - $($jobResult.Status) - $($jobResult.Details)" -Level Warning }
            }
        }
        catch   { Write-Log "Job failed for $($job.Computer) : $($_.Exception.Message)" -Level Error }
        finally { $job.PowerShell.Dispose() }
    }
    $runspacePool.Close()
    $runspacePool.Dispose()
    # Establish SMB sessions and store credentials for explorer.exe access
    if ($Credential) {
        $smbUser = $Credential.UserName
        $smbPass = $Credential.GetNetworkCredential().Password
        foreach ($r in $results) {
            if (-not $r.Success) { continue }
            $computer = $r.Computer
            # SMB session for PowerShell process (elevated logon session)
            $smbResult = [NetSession.Native]::ConnectSmb($computer, $smbUser, $smbPass)
            if ($smbResult -eq 1219) {
                [NetSession.Native]::DisconnectSmb($computer) | Out-Null
                $smbResult = [NetSession.Native]::ConnectSmb($computer, $smbUser, $smbPass)
            }
            # Domain credential for explorer.exe (non-elevated logon session)
            $credSaved = [NetSession.Native]::SaveDomainCredential($computer, $smbUser, $smbPass)
            Write-Log "SMB session for $computer : smb=$smbResult, cred=$credSaved"
            if ($smbResult -eq 0 -or $credSaved) {
                [void]$script:SmbSessionsToCleanup.Add($computer)
            }
        }
    }
    # Resolve hostnames for IP addresses (all results, including failures)
    foreach ($result in $results) {
        if ($result.Computer -match '^\d{1,3}(\.\d{1,3}){3}$') {
            try {
                $dnsEntry = [System.Net.Dns]::GetHostEntry($result.Computer)
                if ($dnsEntry -and $dnsEntry.HostName) {
                    $result | Add-Member -NotePropertyName Hostname -NotePropertyValue ($dnsEntry.HostName.Split('.')[0]) -Force
                }
            }
            catch { Write-Log "DNS hostname lookup failed for $($result.Computer) : $_" -Level Debug }
        }
    }
    # Retrieve console user info for successful connections
    foreach ($result in $results) {
        if (-not $result.Success) { continue }
        try {
            $userInfo = Get-RemoteConsoleUser -ComputerName $result.Computer -Credential $Credential
            $result | Add-Member -NotePropertyName ConsoleUser         -NotePropertyValue $userInfo.User     -Force
            $result | Add-Member -NotePropertyName ConsoleUserFullName -NotePropertyValue $userInfo.FullName -Force
            $result | Add-Member -NotePropertyName ConsoleUserSID      -NotePropertyValue $userInfo.SID      -Force
            $result | Add-Member -NotePropertyName Hostname            -NotePropertyValue $userInfo.Hostname -Force
        }
        catch {
            Write-Log "Failed to get console user for $($result.Computer) : $_" -Level Debug
        }
    }
    $successCount = @($results | Where-Object { $_.Success }).Count
    $failCount    = $results.Count - $successCount
    Write-Log "Connection tests completed : $successCount success / $failCount failed"
    $script:RemoteConnectionResults = $results
    return $results
}

function Get-ComputerDisplayName {
    param([Parameter(Mandatory = $true)][string]$Computer)
    $isIP = $Computer -match '^\d{1,3}(\.\d{1,3}){3}$'
    if (-not $isIP) { return $Computer }
    $connectionResult = $script:RemoteConnectionResults | Where-Object { $_.Computer -eq $Computer } | Select-Object -First 1
    if ($connectionResult -and $connectionResult.Hostname) { return "$Computer ($($connectionResult.Hostname))" }
    return $Computer
}

function Get-RemoteConsoleUser {
    param([Parameter(Mandatory = $true)][string]$ComputerName, [System.Management.Automation.PSCredential]$Credential = $null)
    $result = @{
        User     = ""
        FullName = ""
        SID      = ""
        Hostname = ""
    }
    $isIP = $ComputerName -match '^\d{1,3}(\.\d{1,3}){3}$'
    if ($isIP) {
        try {
            $dnsResult = [System.Net.Dns]::GetHostEntry($ComputerName)
            if ($dnsResult -and $dnsResult.HostName) { $result.Hostname = $dnsResult.HostName.Split('.')[0] }
        }
        catch { Write-Log "DNS hostname lookup failed for $ComputerName : $_" -Level Debug }
    }
    else { $result.Hostname = $ComputerName }
    $isRemote = $ComputerName -ne $env:COMPUTERNAME -and -not [string]::IsNullOrWhiteSpace($ComputerName)
    $useWinRM = $isRemote -and $Credential
    # WTS native API path (local, or remote without explicit credentials)
    if (-not $useWinRM) {
        $hServer = [IntPtr]::Zero
        $ppSessionInfo = [IntPtr]::Zero
        try {
            if ($isRemote) { $hServer = [WtsApi32]::WTSOpenServer($ComputerName) }
            else           { $hServer = [WtsApi32]::WTSOpenServer("") }
            if ($hServer -eq [IntPtr]::Zero) {
                Write-Log "WTSOpenServer failed for $ComputerName" -Level Debug
                if ($isRemote) { $useWinRM = $true }
                else           { return $result }
            }
            if ($hServer -ne [IntPtr]::Zero) {
                $sessionCount = 0
                if (-not [WtsApi32]::WTSEnumerateSessions($hServer, 0, 1, [ref]$ppSessionInfo, [ref]$sessionCount)) {
                    Write-Log "WTSEnumerateSessions failed for $ComputerName" -Level Debug
                    if ($isRemote) { $useWinRM = $true }
                }
                else {
                    $structSize = [Runtime.InteropServices.Marshal]::SizeOf([type][WTS_SESSION_INFO])
                    $currentPtr = $ppSessionInfo
                    $consoleSessionId = -1
                    for ($i = 0; $i -lt $sessionCount; $i++) {
                        $sessionInfo = [Runtime.InteropServices.Marshal]::PtrToStructure($currentPtr, [type][WTS_SESSION_INFO])
                        if ($sessionInfo.pWinStationName -eq "Console" -and $sessionInfo.State -eq 0) {
                            $consoleSessionId = $sessionInfo.SessionId
                            break
                        }
                        $currentPtr = [IntPtr]::Add($currentPtr, $structSize)
                    }
                    [WtsApi32]::WTSFreeMemory($ppSessionInfo)
                    $ppSessionInfo = [IntPtr]::Zero
                    if ($consoleSessionId -ge 0) {
                        $pBuffer = [IntPtr]::Zero ; $bytes=0 ; $userName=$null ; $domain=$null
                        if ([WtsApi32]::WTSQuerySessionInformation($hServer, $consoleSessionId, 5, [ref]$pBuffer, [ref]$bytes)) {
                            $userName = [Runtime.InteropServices.Marshal]::PtrToStringUni($pBuffer)
                            [WtsApi32]::WTSFreeMemory($pBuffer)
                        }
                        $pBuffer = [IntPtr]::Zero ; $bytes=0
                        if ([WtsApi32]::WTSQuerySessionInformation($hServer, $consoleSessionId, 7, [ref]$pBuffer, [ref]$bytes)) {
                            $domain = [Runtime.InteropServices.Marshal]::PtrToStringUni($pBuffer)
                            [WtsApi32]::WTSFreeMemory($pBuffer)
                        }
                        if ($userName) {
                            $result.User = $userName
                            $accountName = if ($domain) { "$domain\$userName" } else { $userName }
                            try {
                                $ntAccount = New-Object System.Security.Principal.NTAccount($accountName)
                                $result.SID = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                                Write-Log "Resolved SID for $accountName : $($result.SID)" -Level Debug
                            }
                            catch {
                                Write-Log "Failed to resolve SID for $accountName : $_" -Level Debug
                            }
                            if ($script:IsDomainJoined) {
                                try {
                                    $searcher = [adsisearcher]"(samaccountname=$userName)"
                                    $searcher.PropertiesToLoad.Add("displayname") | Out-Null
                                    $adResult = $searcher.FindOne()
                                    if ($adResult -and $adResult.Properties["displayname"]) {
                                        $result.FullName = $adResult.Properties["displayname"][0]
                                    }
                                }
                                catch { Write-Log "AD lookup failed for $userName : $_" -Level Debug }
                            }
                        }
                    }
                    else { Write-Log "No active console session on $ComputerName" -Level Debug }
                }
            }
        }
        catch {
            Write-Log "WTS API error for $ComputerName : $_" -Level Debug
            if ($isRemote) { $useWinRM = $true }
        }
        finally {
            if ($ppSessionInfo -ne [IntPtr]::Zero) { [WtsApi32]::WTSFreeMemory($ppSessionInfo) }
            if ($hServer -ne [IntPtr]::Zero)       { [WtsApi32]::WTSCloseServer($hServer) }
        }
    }
    # WinRM path (remote with credentials, or WTS fallback)
    if ($useWinRM) {
        Write-Log "Retrieving console user on $ComputerName via Invoke-Command" -Level Debug
        try {
            $invokeParams = @{
                ComputerName = $ComputerName
                ScriptBlock  = {
                    $info = [PSCustomObject]@{ User = ""; SID = ""; FullName = "" }
                    try {
                        $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
                        if ($cs.UserName) {
                            $parts = $cs.UserName -split '\\'
                            $info.User = $parts[-1]
                            try {
                                $ntAccount = New-Object System.Security.Principal.NTAccount($cs.UserName)
                                $info.SID  = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                            }
                            catch { }
                            try {
                                if ($env:USERDNSDOMAIN) {
                                    $searcher = [adsisearcher]"(samaccountname=$($info.User))"
                                    $searcher.PropertiesToLoad.Add("displayname") | Out-Null
                                    $found = $searcher.FindOne()
                                    if ($found -and $found.Properties["displayname"]) {
                                        $info.FullName = $found.Properties["displayname"][0]
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                    return $info
                }
                ErrorAction = 'Stop'
            }
            if ($Credential) { $invokeParams.Credential = $Credential }
            $remoteInfo = Invoke-Command @invokeParams
            if ($remoteInfo.User) {
                $result.User     = $remoteInfo.User
                $result.SID      = $remoteInfo.SID
                $result.FullName = $remoteInfo.FullName
                Write-Log "Console user on $ComputerName : $($remoteInfo.User)" -Level Debug
            }
            else {
                Write-Log "No logged-on user found on $ComputerName" -Level Debug
            }
        }
        catch {
            Write-Log "Invoke-Command failed for console user on $ComputerName : $_" -Level Debug
        }
    }
    return $result
}


function Show-MultipleFailuresDialog {
    param([array]$Results, [System.Management.Automation.PSCredential]$Credential)
    $failedResults = @($Results | Where-Object { -not $_.Success })
    $trustedHostsCandidates = @($failedResults | Where-Object { $_.IsTrustedHostsIssue })
    $remoteRegistryCandidates = @($failedResults | Where-Object { $_.IsRemoteRegistryFailure })
    $hasTrustedHostsCandidates = $trustedHostsCandidates.Count -gt 0
    $hasRemoteRegistryCandidates = $remoteRegistryCandidates.Count -gt 0
    $hasRepairableCandidates = $hasTrustedHostsCandidates -or $hasRemoteRegistryCandidates
    $repairButtonText = ""
    if ($hasTrustedHostsCandidates -and $hasRemoteRegistryCandidates) {
        $repairButtonText = "Repairs : Add missing devices to TrustedHosts / Enable Remote Registry"
    }
    elseif ($hasTrustedHostsCandidates) {
        $repairButtonText = "Repair : Add Missing Devices to TrustedHosts"
    }
    elseif ($hasRemoteRegistryCandidates) {
        $repairButtonText = "Repair : Enable Remote Registry"
    }
    $detailForm                 = New-Object System.Windows.Forms.Form
    $detailForm.Text            = "Connection Status Details"
    $detailForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $detailForm.MinimizeBox     = $false
    $detailForm.MaximizeBox     = $false
    $detailForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dataGridView                              = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location                     = New-Object System.Drawing.Point(10, 10)
    $dataGridView.Size                         = New-Object System.Drawing.Size(814, 260)
    $dataGridView.AllowUserToAddRows           = $false
    $dataGridView.AllowUserToDeleteRows        = $false
    $dataGridView.AllowUserToResizeRows        = $false
    $dataGridView.AllowUserToResizeColumns     = $false
    $dataGridView.ReadOnly                     = $true
    $dataGridView.MultiSelect                  = $false
    $dataGridView.RowHeadersVisible            = $false
    $dataGridView.ColumnHeadersHeightSizeMode  = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $dataGridView.AutoSizeRowsMode             = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
    $dataGridView.DefaultCellStyle.WrapMode    = [System.Windows.Forms.DataGridViewTriState]::True
    $dataGridView.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    $dataGridView.BackgroundColor              = [System.Drawing.SystemColors]::Window
    $dataGridView.BorderStyle                  = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dataGridView.DefaultCellStyle.SelectionBackColor = $dataGridView.DefaultCellStyle.BackColor
    $dataGridView.DefaultCellStyle.SelectionForeColor = $dataGridView.DefaultCellStyle.ForeColor
    $dataGridView.EnableHeadersVisualStyles    = $false
    [void]$dataGridView.Columns.Add("Computer", "Computer")
    [void]$dataGridView.Columns.Add("Status", "Status")
    [void]$dataGridView.Columns.Add("User", "User")
    [void]$dataGridView.Columns.Add("FullName", "Full Name")
    [void]$dataGridView.Columns.Add("Details", "Details")
    $dataGridView.Columns["Computer"].Width  = 120
    $dataGridView.Columns["Status"].Width    = 100
    $dataGridView.Columns["User"].Width      = 100
    $dataGridView.Columns["FullName"].Width  = 150
    $dataGridView.Columns["Details"].Width   = 320
    foreach ($result in $Results) {
        $rowIndex = $dataGridView.Rows.Add(
            (Get-ComputerDisplayName -Computer $result.Computer),
            $result.Status,
            $result.ConsoleUser,
            $result.ConsoleUserFullName,
            $result.Details
        )
        $row      = $dataGridView.Rows[$rowIndex]
        $row.Tag  = $result
        $rowColor = [System.Drawing.Color]::Black
        if ($result.Success) { $rowColor = [System.Drawing.Color]::DarkGreen } 
        else { 
            $isRepairable = ($trustedHostsCandidates | Where-Object { $_.Computer -eq $result.Computer }) -or
                            ($remoteRegistryCandidates | Where-Object { $_.Computer -eq $result.Computer })
            if ($isRepairable) { $rowColor = [System.Drawing.Color]::DarkOrange }
            else               { $rowColor = [System.Drawing.Color]::DarkRed }
        }
        $row.DefaultCellStyle.ForeColor          = $rowColor
        $row.DefaultCellStyle.SelectionForeColor = $rowColor
        $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.SystemColors]::Window
    }
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $dataGridView.Add_MouseDown({
        param($s, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
        $hitTest = $dataGridView.HitTest($e.X, $e.Y)
        if ($hitTest.RowIndex -lt 0) { return }
        $contextMenu.Items.Clear()
        $clickedRow = $dataGridView.Rows[$hitTest.RowIndex]
        foreach ($col in $dataGridView.Columns) {
            $cellValue = $clickedRow.Cells[$col.Index].Value
            $menuItem      = New-Object System.Windows.Forms.ToolStripMenuItem
            $menuItem.Text = "Copy $($col.HeaderText)"
            if ([string]::IsNullOrWhiteSpace($cellValue)) { $menuItem.Enabled = $false }
            else {
                $menuItem.Tag = $cellValue
                $menuItem.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($this.Tag.ToString()) })
            }
            [void]$contextMenu.Items.Add($menuItem)
        }
        $contextMenu.Show($dataGridView, $e.Location)
    })
    $detailForm.Controls.Add($dataGridView)
    $yOffset = 280
    if ($hasRepairableCandidates) {
        $btnRepair          = New-Object System.Windows.Forms.Button
        $btnRepair.Text     = $repairButtonText
        $btnRepair.Location = New-Object System.Drawing.Point(10, $yOffset)
        $btnRepair.Size     = New-Object System.Drawing.Size(814, 35)
        $btnRepair.Add_Click({
            $computersToRetry = [System.Collections.ArrayList]::new()
            foreach ($candidate in $trustedHostsCandidates) {
                $computer = $candidate.Computer
                if (Add-HostToTrustedHosts -Computer $computer) {
                    [void]$script:TrustedHostsToRemove.Add($computer)
                    [void]$computersToRetry.Add($computer)
                }
            }
            foreach ($candidate in $remoteRegistryCandidates) {
                $computer = $candidate.Computer
                $enableResult = Enable-RemoteRegistry -Computer $computer -Credential $Credential -Persistent $true
                if ($enableResult.Success) {
                    if ($computersToRetry -notcontains $computer) { [void]$computersToRetry.Add($computer) }
                }
            }
            if ($computersToRetry.Count -gt 0) {
                $detailForm.Tag = @{ Action = "Retry"; Computers = $computersToRetry }
                $detailForm.Close()
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("No repairs could be applied.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        })
        $detailForm.Controls.Add($btnRepair)
        $yOffset += 40
    }
    $btnClose          = New-Object System.Windows.Forms.Button
    $btnClose.Text     = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(10, $yOffset)
    $btnClose.Size     = New-Object System.Drawing.Size(814, 35)
    $btnClose.Add_Click({ $detailForm.Close() })
    $detailForm.Controls.Add($btnClose)
    $yOffset += 45
    $detailForm.ClientSize = New-Object System.Drawing.Size(834, $yOffset)
    $detailForm.ShowDialog($form) | Out-Null
    $dialogResult = $detailForm.Tag
    $detailForm.Dispose()
    return $dialogResult
}

function Get-TargetComputersFromPanel {
    $targetText = $textBox_TargetDevice.Text.Trim()
    if ($textBox_TargetDevice.Tag.IsPlaceholder -or [string]::IsNullOrWhiteSpace($targetText)) { return @() }
    $computers = $targetText -split '[,;\s]+' | Where-Object { $_ }
    return @($computers)
}

function Get-CredentialFromPanel {
    if ($textBox_CredentialID.Tag.IsPlaceholder -or $textBox_CredentialPwd.Tag.IsPlaceholder) { return $null }
    $username = $textBox_CredentialID.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($username) -or $script:SecurePassword.Length -eq 0) { return $null }
    try {
        $credential = New-Object System.Management.Automation.PSCredential($username, $script:SecurePassword.Copy())
        Write-Log "Credential created for user : $username"
        return $credential
    }
    catch {
        Write-Log "Failed to create credential : $_" -Level Error
        return $null
    }
}

function Update-ConnectionStatusDisplay {
    param([array]$Results, [System.Management.Automation.PSCredential]$Credential = $null)
    if ($Results.Count -eq 0) {
        $button_ConnectionStatus.Visible = $false
        return
    }
    $failedResults  = @($Results | Where-Object { -not $_.Success })
    $successResults = @($Results | Where-Object { $_.Success })
    $okCount = $successResults.Count
    $koCount = $failedResults.Count
    Write-Log "Connection status : $okCount OK  /  $koCount KO"
    $button_ConnectionStatus.Visible = $true
    $button_ConnectionStatus.Tag     = $Results
    if ($Results.Count -eq 1) {
        $result = $Results[0]
        $button_ConnectionStatus.Text = $result.Status
        if ($result.Success) {
            $button_ConnectionStatus.ForeColor = [System.Drawing.Color]::DarkGreen
            return
        }
        $button_ConnectionStatus.ForeColor = [System.Drawing.Color]::DarkRed
    }
    else {
        $button_ConnectionStatus.Text = "$okCount OK  /  $koCount KO"
        if     ($koCount -eq 0) { $button_ConnectionStatus.ForeColor = [System.Drawing.Color]::DarkGreen }
        elseif ($okCount -eq 0) { $button_ConnectionStatus.ForeColor = [System.Drawing.Color]::DarkRed }
        else                    { $button_ConnectionStatus.ForeColor = [System.Drawing.Color]::DarkOrange }
    }
    if ($koCount -gt 0) {
        $hasRepairableFail = $false
        foreach ($f in $failedResults) {
            if ($f.IsTrustedHostsIssue -or $f.IsRemoteRegistryFailure) { $hasRepairableFail = $true; break }
        }
        $showDialog = ($Results.Count -gt 1) -or $hasRepairableFail
        if (-not $showDialog) { return }
        $dialogResult = Show-MultipleFailuresDialog -Results $Results -Credential $Credential
        if ($dialogResult -and $dialogResult.Action -eq "Retry") {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            [System.Windows.Forms.Application]::DoEvents()
            try {
                Write-Log "Flushing DNS cache before retry..."
                Clear-DnsClientCache -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 300
                $retryResults = Test-RemoteConnections -Computers $dialogResult.Computers -Credential $Credential
                $updatedResults = [System.Collections.ArrayList]::new()
                foreach ($r in $Results) {
                    $retried = $retryResults | Where-Object { $_.Computer -eq $r.Computer }
                    if ($retried) { [void]$updatedResults.Add($retried) }
                    else          { [void]$updatedResults.Add($r) }
                }
                $script:RemoteConnectionResults = $updatedResults
                Update-ConnectionStatusDisplay -Results $updatedResults -Credential $Credential
            }
            finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
        }
    }
}

function Update-RemoteTargetPanelVisibility {
    $isTab3or4 = ($tabControl.SelectedTab -eq $tabPage3) -or ($tabControl.SelectedTab -eq $tabPage4)
    $panel_RemoteTarget.Visible   = $true
    $textBox_TargetDevice.Visible = $isTab3or4
    $button_Connect.Visible       = $isTab3or4
    if (-not $isAdmin) {
        $flowPanel_RemoteTarget.Visible = $false
        $button_RestartAsAdmin.Visible  = $true
        $button_RestartAsAdmin.Dock     = [System.Windows.Forms.DockStyle]::Right
    }
    else {
        $flowPanel_RemoteTarget.Visible = $true
        $button_RestartAsAdmin.Visible  = $false
    }
}

function Get-TrustedHostsList {
    param([switch]$ForceRefresh)
    if (-not $ForceRefresh -and $null -ne $script:TrustedHostsCache) {
        return $script:TrustedHostsCache
    }
    # Skip WSMan access if WinRM is not running
    try {
        $svc = Get-Service -Name WinRM -ErrorAction Stop
        if ($svc.Status -ne 'Running') {
            Write-Log "WinRM not running, skipping TrustedHosts lookup" -Level Debug
            return @()
        }
    }
    catch {
        Write-Log "Cannot check WinRM service : $_" -Level Debug
        return @()
    }
    try {
        $current = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value
        Write-Log "TrustedHosts raw value : '$current'" -Level Debug
        if ([string]::IsNullOrWhiteSpace($current)) { $script:TrustedHostsCache = @() }
        elseif ($current -eq "*") { script:TrustedHostsCache = @("*") }
        else { $script:TrustedHostsCache = @($current -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
        return $script:TrustedHostsCache
    }
    catch {
        Write-Log "Failed to get TrustedHosts : $_" -Level Error
        return @()
    }
}

function Add-HostToTrustedHosts {
    param([string]$Computer)
    try {
        $trustedHosts = @(Get-TrustedHostsList)
        if ($trustedHosts -contains "*") {
            Write-Log "TrustedHosts already allows all (*)"
            return $true
        }
        if ($trustedHosts -contains $Computer) {
            Write-Log "Host $Computer already in TrustedHosts"
            return $true
        }
        if ($trustedHosts.Count -eq 0 -or ($trustedHosts.Count -eq 1 -and [string]::IsNullOrWhiteSpace($trustedHosts[0]))) {
            $newValue = $Computer
        }
        else {
            $newValue = (@($trustedHosts) + $Computer) -join ','
        }
        Write-Log "Setting TrustedHosts to : '$newValue'" -Level Debug
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newValue -Force -ErrorAction Stop
        $script:TrustedHostsCache = $null
        $verifyList = @(Get-TrustedHostsList -ForceRefresh)
        if ($verifyList -contains $Computer) {
            Write-Log "Verified : $Computer successfully added to TrustedHosts"
            return $true
        }
        else {
            Write-Log "FAILED to verify $Computer in TrustedHosts after add. Current list : $($verifyList -join ', ')" -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Failed to add $Computer to TrustedHosts : $_" -Level Error
        $script:TrustedHostsCache = $null
        return $false
    }
}

function Remove-HostFromTrustedHosts {
    param([string]$Computer)
    try {
        $trustedHosts = @(Get-TrustedHostsList)
        if ($trustedHosts -contains "*") {
            Write-Log "TrustedHosts is set to *, cannot remove individual host"
            return $false
        }
        $newList = @($trustedHosts | Where-Object { $_ -ne $Computer })
        $newValue = $newList -join ','
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newValue -Force -ErrorAction Stop
        $script:TrustedHostsCache = $null
        Write-Log "Removed $Computer from TrustedHosts"
        return $true
    }
    catch {
        Write-Log "Failed to remove $Computer from TrustedHosts : $_" -Level Error
        $script:TrustedHostsCache = $null
        return $false
    }
}

function Enable-RemoteRegistry {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [System.Management.Automation.PSCredential]$Credential = $null,
        [bool]$Persistent = $false
    )
    Write-Log "Enabling Remote Registry on $Computer (Persistent : $Persistent)"
    $scriptBlock = {
        param($persistent)
        $result = @{
            Success           = $false
            OriginalStartType = $null
            Status            = ""
            Details           = ""
        }
        try {
            $service = Get-Service -Name RemoteRegistry -ErrorAction Stop
            $result.OriginalStartType = (Get-WmiObject Win32_Service -Filter "Name='RemoteRegistry'").StartMode
            Write-Log "Original StartType : $($result.OriginalStartType)"
            if ($service.Status -eq 'Running') {
                $result.Success = $true
                $result.Status  = "Already Running"
                $result.Details = "Service was already running"
                return $result
            }
            $startType = if ($persistent) { 'Automatic' } else { 'Manual' }
            Set-Service -Name RemoteRegistry -StartupType $startType -ErrorAction Stop
            Start-Service -Name RemoteRegistry -ErrorAction Stop
            Start-Sleep -Milliseconds 500
            $service = Get-Service -Name RemoteRegistry -ErrorAction Stop
            if ($service.Status -eq 'Running') {
                $result.Success = $true
                $result.Status  = "Running"
                $result.Details = "Service started successfully"
            }
            else {
                $result.Status  = "Cannot start RemoteRegistry"
                $result.Details = "Service status : $($service.Status)"
            }
        }
        catch {
            $result.Status  = "Error"
            $result.Details = $_.Exception.Message
        }
        return $result
    }
    try {
        if ($Computer -eq $env:COMPUTERNAME -or [string]::IsNullOrWhiteSpace($Computer)) {
            $result = & $scriptBlock $Persistent
        }
        else {
            $invokeParams = @{
                ComputerName = $Computer
                ScriptBlock  = $scriptBlock
                ArgumentList = @($Persistent)
                ErrorAction  = 'Stop'
            }
            if ($Credential) { $invokeParams.Credential = $Credential }
            $result = Invoke-Command @invokeParams
        }
        if ($result.Success) { Write-Log "Remote Registry enabled on $Computer : $($result.Status)" }
        else                 { Write-Log "Failed to enable Remote Registry on $Computer : $($result.Details)" -Level Warning }
        return $result
    }
    catch {
        Write-Log "Failed to enable Remote Registry on $Computer : $_" -Level Error
        return @{
            Success           = $false
            OriginalStartType = $null
            Status            = "Error"
            Details           = $_.Exception.Message
        }
    }
}

function Disable-RemoteRegistry {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [string]$OriginalStartType = "Disabled",
        [System.Management.Automation.PSCredential]$Credential = $null
    )
    Write-Log "Restoring Remote Registry on $Computer to : $OriginalStartType"
    $scriptBlock = {
        param($originalStartType)
        try {
            Stop-Service -Name RemoteRegistry -Force -ErrorAction SilentlyContinue
            $startType = switch ($originalStartType) {
                "Auto"     { "Automatic" }
                "Manual"   { "Manual" }
                "Disabled" { "Disabled" }
                default    { "Disabled" }
            }
            Set-Service -Name RemoteRegistry -StartupType $startType -ErrorAction Stop
            return $true
        }
        catch { return $false }
    }
    try {
        if ($Computer -eq $env:COMPUTERNAME -or [string]::IsNullOrWhiteSpace($Computer)) {
            $result = & $scriptBlock $OriginalStartType
        }
        else {
            $invokeParams = @{
                ComputerName = $Computer
                ScriptBlock  = $scriptBlock
                ArgumentList = @($OriginalStartType)
                ErrorAction  = 'Stop'
            }
            if ($Credential) { $invokeParams.Credential = $Credential }
            $result = Invoke-Command @invokeParams
        }
        if ($result) { Write-Log "Remote Registry restored on $Computer" }
        else         { Write-Log "Failed to restore Remote Registry on $Computer" -Level Warning }
        return $result
    }
    catch {
        Write-Log "Failed to restore Remote Registry on $Computer : $_" -Level Error
        return $false
    }
}

function Remove-ScheduledTrustedHosts {
    if ($script:TrustedHostsToRemove.Count -eq 0) { return }
    Write-Log "Cleaning up $($script:TrustedHostsToRemove.Count) host(s) from TrustedHosts"
    foreach ($h in $script:TrustedHostsToRemove) { Remove-HostFromTrustedHosts -Computer $h }
    $script:TrustedHostsToRemove.Clear()
}

function Invoke-RemoteDataCollection {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [hashtable]$ArgumentList = @{},
        [System.Management.Automation.PSCredential]$Credential = $null,
        [int]$TimeoutSeconds = 1000
    )
    if ([string]::IsNullOrWhiteSpace($ComputerName) -or $ComputerName -eq $env:COMPUTERNAME) {
        return & $ScriptBlock $ArgumentList
    }
    $invokeParams = @{
        ComputerName = $ComputerName
        ScriptBlock  = $ScriptBlock
        ArgumentList = @($ArgumentList)
        ErrorAction  = 'Stop'
        AsJob        = $true
    }
    if ($Credential) { $invokeParams.Credential = $Credential }
    try {
        $job = Invoke-Command @invokeParams
        $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
        if ($completed) {
            $result = Receive-Job -Job $job
            Remove-Job -Job $job -Force
            return $result
        }
        else {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force
            Write-Log "Remote execution timeout on $ComputerName after ${TimeoutSeconds}s" -Level Warning
            return $null
        }
    }
    catch {
        Write-Log "Remote execution failed on $ComputerName : $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Test-RemotePathAccess {
    param([string[]]$Paths)
    $computers = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $Paths) {
        if ($p -match '^\\\\([^\\]+)\\') {
            $comp = $matches[1]
            if (-not $computers.Contains($comp)) { [void]$computers.Add($comp) }
        }
    }
    if ($computers.Count -eq 0) { return $true }
    # Auto-fill Target Devices textbox (invisible on tab1/tab2 but functional)
    $placeholder = $textBox_TargetDevice.Tag.Text
    $textBox_TargetDevice.Tag       = @{ Text = $placeholder; IsPlaceholder = $false }
    $textBox_TargetDevice.ForeColor = [System.Drawing.Color]::Black
    $textBox_TargetDevice.Text      = ($computers -join ',')
    $credential = Get-CredentialFromPanel
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $connectionResults = Test-RemoteConnections -Computers $computers.ToArray() -Credential $credential
        Update-ConnectionStatusDisplay -Results $connectionResults -Credential $credential
        $connectedComputers = @($script:RemoteConnectionResults | Where-Object { $_.Success } | ForEach-Object { $_.Computer })
        if ($connectedComputers.Count -eq 0) {
            Write-Log "No remote computers are accessible for UNC path" -Level Warning
            return $false
        }
        Write-Log "Remote path access : $($connectedComputers.Count) computer(s) connected"
        return $true
    }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
}

function Start-RemotePSSession {
    param([string]$Server, [string]$RemotePath)
    $cmdParts = [System.Collections.Generic.List[string]]::new()
    $cred = Get-CredentialFromPanel
    if ($cred) {
        $encPwd = $cred.Password | ConvertFrom-SecureString
        $cmdParts.Add("`$p = ConvertTo-SecureString '$encPwd'")
        $cmdParts.Add("`$c = New-Object PSCredential('$($cred.UserName)',`$p)")
        $cmdParts.Add("`$s = New-PSSession '$Server' -Credential `$c")
    } else {
        $cmdParts.Add("`$s = New-PSSession '$Server'")
    }
    $cmdParts.Add("Invoke-Command `$s { Set-Location '$RemotePath' }")
    $cmdParts.Add("Enter-PSSession `$s")
    $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes(($cmdParts -join '; ')))
    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName        = "powershell.exe"
    $psi.Arguments       = "-NoExit -EncodedCommand $encoded"
    $psi.Verb            = "runas"
    $psi.UseShellExecute = $true
    [System.Diagnostics.Process]::Start($psi)
}

#region Network UI

$panel_RemoteTarget      = gen $form                   "Panel"   ($form.ClientSize.Width - 702) ($titleBarHeight + 3) 700 19 'Visible=$false'  'Anchor=Top,Right'          'BackColor=Transparent'   
$flowPanel_RemoteTarget  = gen $panel_RemoteTarget     "FlowLayoutPanel"                      0 0 0   0  'Dock=Right'      'FlowDirection=RightToLeft' 'BackColor=Transparent' 'WrapContents=$false' 'AutoSize=$true' 'Padding=0'
$button_Connect          = gen $flowPanel_RemoteTarget "Button"          "Test"               0 0 55  19 'Margin=5 0 1 0'                                                      'FlatStyle=System'
$textBox_CredentialPwd   = gen $flowPanel_RemoteTarget "TextBox"                              0 0 120 19 'Margin=5 0 0 0'       
$textBox_CredentialID    = gen $flowPanel_RemoteTarget "TextBox"                              0 0 120 19 'Margin=5 0 0 0'       
$textBox_TargetDevice    = gen $flowPanel_RemoteTarget "TextBox"                              0 0 180 19 'Margin=5 0 0 0'       
$button_ConnectionStatus = gen $flowPanel_RemoteTarget "Button"                               0 0 0 0    'Margin=10 0 0 0'                             'Visible=$false'        'FlatStyle=Flat'      'AutoSize=$true' 
$button_RestartAsAdmin   = gen $panel_RemoteTarget     "Button"          "Restart as Admin"   0 0 110 19 'Dock=Right'      'ForeColor=DarkRed'         'Visible=$false'        'FlatStyle=Flat'

$panel_RemoteTarget.BringToFront()
$textBox_TargetDevice.TabIndex = 0  ;  $textBox_CredentialID.TabIndex = 1  ;  $textBox_CredentialPwd.TabIndex = 2  ;  $button_Connect.TabIndex = 3

$textBox_TargetDevice.Add_KeyPress({ param($s, $e); if ($e.KeyChar -eq [char]13)                              { $e.Handled = $true } })
$textBox_TargetDevice.Add_KeyDown({
    param($s, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $this.SelectAll() ; $e.Handled = $true ; $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $button_Connect.PerformClick() ; $e.Handled = $true ; $e.SuppressKeyPress = $true
    }
    elseif ($e.Control -and $e.Shift -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        $e.Handled = $true ; $e.SuppressKeyPress = $true
        if (-not [System.Windows.Forms.Clipboard]::ContainsText()) { return }
        $clipText = [System.Windows.Forms.Clipboard]::GetText()
        $lines = @($clipText -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
        if ($lines.Count -eq 0) { return }
        if ($this.Tag -is [hashtable] -and $this.Tag.IsPlaceholder) {
            $this.Tag = @{ Text = $this.Tag.Text; IsPlaceholder = $false }
            $this.ForeColor = [System.Drawing.Color]::Black ; $this.Text = ""
        }
        $this.Text = $lines[0].Trim()
        $targets = @($textBox_CredentialID, $textBox_CredentialPwd)
        $tabCount = 0
        for ($i = 1; $i -lt $lines.Count -and $tabCount -lt 2; $i++) {
            $lineText = $lines[$i].Trim()
            $nextCtrl = $targets[$tabCount]
            if ($nextCtrl -eq $textBox_CredentialPwd) {
                Reset-SecurePassword
                $textBox_CredentialPwd.Tag = @{ IsPlaceholder = $false }
                $textBox_CredentialPwd.UseSystemPasswordChar = $true
                $textBox_CredentialPwd.ForeColor = [System.Drawing.Color]::Black
                foreach ($ch in $lineText.ToCharArray()) { $script:SecurePassword.AppendChar($ch) }
                $textBox_CredentialPwd.Text = [string]::new([char]0x25CF, $lineText.Length)
            }
            else {
                if ($nextCtrl.Tag -is [hashtable] -and $nextCtrl.Tag.IsPlaceholder) {
                    $nextCtrl.Tag = @{ Text = $nextCtrl.Tag.Text; IsPlaceholder = $false }
                    $nextCtrl.ForeColor = [System.Drawing.Color]::Black
                }
                $nextCtrl.Text = $lineText
            }
            $tabCount++
        }
    }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        $e.Handled = $true ; $e.SuppressKeyPress = $true
        if (-not [System.Windows.Forms.Clipboard]::ContainsText()) { return }
        $clipText = [System.Windows.Forms.Clipboard]::GetText()
        if ($this.Tag -is [hashtable] -and $this.Tag.IsPlaceholder) {
            $this.Tag = @{ Text = $this.Tag.Text; IsPlaceholder = $false }
            $this.ForeColor = [System.Drawing.Color]::Black ; $this.Text = ""
        }
        $hasNewlines    = $clipText -match "`r?`n"
        $hasSpecialSeps = $clipText -match '[,;]'
        if ($hasNewlines -or (-not $hasSpecialSeps -and ($clipText -match '\s'))) {
            $parts = @($clipText -split "`r?`n|\s+" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
            $this.Text = $parts -join ';'
        }
        else {
            $this.SelectedText = $clipText
        }
    }
})
$textBox_TargetDevice.Add_TextChanged({
    if ($textBox_TargetDevice.Tag.IsPlaceholder)                           { return }
    if ([string]::IsNullOrWhiteSpace($textBox_TargetDevice.Text))          { $button_ConnectionStatus.Visible = $false }
})
$textBox_CredentialID.Add_KeyDown({
    param($s, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $this.SelectAll() ; $e.Handled = $true ; $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $textBox_CredentialPwd.Focus() ; $e.Handled = $true ; $e.SuppressKeyPress = $true
    }
    elseif ($e.Control -and $e.Shift -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        $e.Handled = $true ; $e.SuppressKeyPress = $true
        if (-not [System.Windows.Forms.Clipboard]::ContainsText()) { return }
        $clipText = [System.Windows.Forms.Clipboard]::GetText()
        $lines = @($clipText -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
        if ($lines.Count -eq 0) { return }
        if ($this.Tag -is [hashtable] -and $this.Tag.IsPlaceholder) {
            $this.Tag = @{ Text = $this.Tag.Text; IsPlaceholder = $false }
            $this.ForeColor = [System.Drawing.Color]::Black
        }
        $this.Text = $lines[0].Trim()
        if ($lines.Count -gt 1) {
            $lineText = $lines[1].Trim()
            Reset-SecurePassword
            $textBox_CredentialPwd.Tag = @{ IsPlaceholder = $false }
            $textBox_CredentialPwd.UseSystemPasswordChar = $true
            $textBox_CredentialPwd.ForeColor = [System.Drawing.Color]::Black
            foreach ($ch in $lineText.ToCharArray()) { $script:SecurePassword.AppendChar($ch) }
            $textBox_CredentialPwd.Text = [string]::new([char]0x25CF, $lineText.Length)
        }
    }
})
$textBox_CredentialPwd.Add_Enter({
    if ($this.Tag.IsPlaceholder) {
        $this.Tag = @{ IsPlaceholder = $false }
        $this.UseSystemPasswordChar = $true
        $this.ForeColor = [System.Drawing.Color]::Black
        $this.Text = ""
        Reset-SecurePassword
    }
})
$textBox_CredentialPwd.Add_Leave({
    if ([string]::IsNullOrEmpty($this.Text)) { Reset-SecurePassword ; Set-PasswordPlaceholderState -ShowPlaceholder $true }
})
$textBox_CredentialPwd.Add_PreviewKeyDown({ param($s, $e) ; if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Tab -and $e.Shift) { $e.IsInputKey = $true } })
$textBox_CredentialPwd.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Tab -and $e.Shift) { $textBox_CredentialID.Focus() ; $e.Handled = $true ; $e.SuppressKeyPress = $true ; return }
    if ($this.Tag.IsPlaceholder) { return }
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $button_Connect.PerformClick() ; $e.Handled = $true ; $e.SuppressKeyPress = $true ; return }
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
        if ($this.SelectionLength -gt 0) { Reset-SecurePassword ; $this.Text = "" }
        $e.Handled = $true ; $e.SuppressKeyPress = $true ; return
    }
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $this.SelectAll() ; $e.Handled = $true ; $e.SuppressKeyPress = $true ; return }
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        $e.Handled = $true ; $e.SuppressKeyPress = $true
        if ([System.Windows.Forms.Clipboard]::ContainsText()) {
            Reset-SecurePassword
            $this.Text = ""
            $clip = [System.Windows.Forms.Clipboard]::GetText()
            foreach ($ch in $clip.ToCharArray()) { $script:SecurePassword.AppendChar($ch) }
            $this.AppendText([string]::new([char]0x25CF, $clip.Length))
        }
        return
    }
})
$textBox_CredentialPwd.Add_KeyPress({
    param($s, $e)
    if ($this.Tag.IsPlaceholder -or $e.KeyChar -eq [char]13) { $e.Handled = $true ; return }
    if ($e.KeyChar -eq [char]8) {
        # Backspace with selection = clear selection
        if ($this.SelectionLength -gt 0) {
            Reset-SecurePassword
            $this.Text = ""
        }
        elseif ($script:SecurePassword.Length -gt 0) {
            $script:SecurePassword.RemoveAt($script:SecurePassword.Length - 1)
        }
        return
    }
    if (-not [char]::IsControl($e.KeyChar)) {
        $e.Handled = $true
        # Selection active = replace all content
        if ($this.SelectionLength -gt 0) {
            Reset-SecurePassword
            $this.Text = ""
        }
        $script:SecurePassword.AppendChar($e.KeyChar)
        $this.AppendText([string][char]0x25CF)
    }
})
$button_RestartAsAdmin.Add_Click({
    $psi                 = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName        = $ScriptPath
    $psi.Verb            = "runas"
    $psi.UseShellExecute = $true
    [void][System.Diagnostics.Process]::Start($psi)
    $form.Close()
})
$button_Connect.Add_Click({
    Write-Log "Test button clicked"
    $targetComputers = Get-TargetComputersFromPanel
    if ($targetComputers.Count -eq 0) {
        $button_ConnectionStatus.Visible   = $true
        $button_ConnectionStatus.Text      = "Local mode"
        $button_ConnectionStatus.ForeColor = [System.Drawing.Color]::DarkBlue
        $button_ConnectionStatus.Tag       = @()
        return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $button_Connect.Enabled = $false
    $button_Connect.Text = "Testing..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $credential = Get-CredentialFromPanel
        $results = Test-RemoteConnections -Computers $targetComputers -Credential $credential
        Update-ConnectionStatusDisplay -Results $results -Credential $credential
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $button_Connect.Enabled = $true
        $button_Connect.Text = "Test"
    }
})
$button_ConnectionStatus.Add_Click({
    $results = $this.Tag
    if (-not $results -or $results.Count -eq 0) { return }
    $credential = Get-CredentialFromPanel
    $dialogResult = Show-MultipleFailuresDialog -Results $results -Credential $credential
    if ($dialogResult -and $dialogResult.Action -eq "Retry") {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $retryResults = Test-RemoteConnections -Computers $dialogResult.Computers -Credential $credential
            $updatedResults = [System.Collections.ArrayList]::new()
            foreach ($r in $results) {
                $retried = $retryResults | Where-Object { $_.Computer -eq $r.Computer }
                if ($retried) { [void]$updatedResults.Add($retried) }
                else          { [void]$updatedResults.Add($r) }
            }
            $script:RemoteConnectionResults = $updatedResults
            $this.Tag = $updatedResults
            $okCount = @($updatedResults | Where-Object { $_.Success }).Count
            $koCount = $updatedResults.Count - $okCount
            if ($updatedResults.Count -eq 1) {
                $this.Text = $updatedResults[0].Status
                if ($updatedResults[0].Success) { $this.ForeColor = [System.Drawing.Color]::DarkGreen }
                else                            { $this.ForeColor = [System.Drawing.Color]::DarkRed }
            }
            else {
                $this.Text = "$okCount OK  /  $koCount KO"
                if     ($koCount -eq 0) { $this.ForeColor = [System.Drawing.Color]::DarkGreen }
                elseif ($okCount -eq 0) { $this.ForeColor = [System.Drawing.Color]::DarkRed }
                else                    { $this.ForeColor = [System.Drawing.Color]::DarkOrange }
            }
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
})


$launch_progressBar.Value = 15



#region  Tab 1: Drop MSI

$script:MsiSelectorFontBold = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$script:MsiSelectorFont     = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$script:MsiSelectorRowH     = 22
$script:MsiSelectorMaxRows  = 3

$labels      = @()
$textBoxes   = @()
$copyButtons = @()
$properties = [ordered]@{
    "GUID"             = "ProductCode"
    "Product Name"     = "ProductName"
    "Product Version"  = "ProductVersion"
    "Manufacturer"     = "Manufacturer"
    "Upgrade Code"     = "UpgradeCode"
}

# MSI Selector panel (top, hidden initially)
$msiSelectorOuterPanel       = gen $tabPage1            "Panel"           "" 5 5 0 0 'Visible=$false'
$msiSelectorPanel            = gen $msiSelectorOuterPanel "Panel"         "" 0 0 0 0 'Dock=Fill' 'BackColor=240 240 245' 'BorderStyle=FixedSingle'
$msiSelectorLabel            = gen $msiSelectorPanel    "Label"           "Loaded MSI :" 10 8 90 20 'Font=Segoe UI, 9, Bold'
$resetMsiButton              = gen $msiSelectorPanel    "Button"          "Reset" 0 5 55 22
$msiSelectorScrollPanel      = gen $msiSelectorPanel    "Panel"           "" 105 3 0 0 'AutoScroll=$true'
$msiSelectorFlowPanel        = gen $msiSelectorScrollPanel "FlowLayoutPanel" "" 0 0 0 0 'FlowDirection=TopDown' 'WrapContents=$false' 'AutoSize=$true' 'AutoSizeMode=GrowAndShrink'
$msiSelectorScrollPanel.HorizontalScroll.Maximum = 0
$msiSelectorFlowPanel.Add_MouseWheel({
    param($s, $e)
    $panel = $msiSelectorScrollPanel
    $scrollAmount = $script:MsiSelectorRowH + 3
    $currentY = -$msiSelectorFlowPanel.Location.Y
    $newY = $currentY - [Math]::Sign($e.Delta) * $scrollAmount
    $maxScroll = [Math]::Max(0, $msiSelectorFlowPanel.Height - $panel.ClientSize.Height)
    $newY = [Math]::Max(0, [Math]::Min($newY, $maxScroll))
    $msiSelectorFlowPanel.Location = [System.Drawing.Point]::new(0, -$newY)
    ($e -as [System.Windows.Forms.HandledMouseEventArgs]).Handled = $true
})

$yPos = 35
foreach ($key in $properties.Keys) {
    $label        = gen $tabPage1 "Label" "$key :" 10 $yPos 0 0 'AutoSize=$true'
    $labels      += $label
    $textBox      = gen $tabPage1 "TextBox" "" 101 $yPos 50 25 'ReadOnly=$true'
    $textBoxes   += $textBox
    $copyButton   = gen $tabPage1 "Button" "COPY" 0 0 60 25 'Enabled=$false'
    $copyButtons += $copyButton
    $copyButton.Add_Click({
        $buttonIndex = $copyButtons.IndexOf($this)
        if (![string]::IsNullOrWhiteSpace($textBoxes[$buttonIndex].Text)) { [System.Windows.Forms.Clipboard]::SetText($textBoxes[$buttonIndex].Text) }
    })
    $yPos += 30
}

$borderTop =        gen $tabPage1 "Panel" "" 0 0 0 1   'Dock=Top'    'BackColor=Gray'
$separatorLine =    gen $tabPage1 "Panel" "" 0 ($yPos + 30) $tabPage1.ClientSize.Width 1 'BorderStyle=None' 'BackColor=Gray' 'AutoSize=$true'
$labelMsiPath =     gen $tabPage1 "Label" "MSI Path:" 10 $yPos 0 0 'AutoSize=$true'
$yPos += 25
$labelDropMessage = gen $tabPage1 "Label" "Drop MSI here, or write path below" 0 0 0 0 'AutoSize=$true' 'Font=Arial,20'
$textBoxPath_Tab1 =      gen $tabPage1 "TextBox" "" 10 $yPos 260 25
$findGuidButton =   gen $tabPage1 "Button" "FIND GUID" 0 0 85 25
$browseButton   =   gen $tabPage1 "Button" "BROWSE" 0 0 85 25
$pictureBoxIcon  =  gen $tabPage1 "PictureBox" "" 0 0 64 64 'SizeMode=StretchImage' 'Visible=$false'
$iconButton_Tab1 =  gen $tabPage1 "Button" "" 0 0 70 70 'Visible=$false' 'FlatStyle=Flat' 'BackColor=Transparent'
$iconButton_Tab1.FlatAppearance.BorderColor = [System.Drawing.Color]::LightGray
$iconButton_Tab1.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(230, 230, 240)
$iconButton_Tab1.Add_Click({
    if (-not $this.Tag) { return }
    $originalBitmap = $this.Tag
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $itemFile = [System.Windows.Forms.ToolStripMenuItem]::new("Copy as File (48x48)")
    $itemFile.Tag = $originalBitmap
    $itemFile.Add_Click({
        try {
            $resized = New-Object System.Drawing.Bitmap($this.Tag, 48, 48)
            $currentPath = $null
            foreach ($ctrl in $msiSelectorFlowPanel.Controls) {
                $innerRadio = $ctrl.Controls | Where-Object { $_ -is [System.Windows.Forms.RadioButton] } | Select-Object -First 1
                if ($innerRadio -and $innerRadio.Checked) { $currentPath = $innerRadio.Tag; break }
            }
            $baseName = $null
            if ($currentPath -and $script:MsiFileCache.Contains($currentPath)) {
                $pn = $script:MsiFileCache[$currentPath].ProductName
                if ($pn -and $pn -ne "None") { $baseName = $pn -replace '\s+', '_' }
            }
            if (-not $baseName) { $baseName = [System.IO.Path]::GetFileNameWithoutExtension($currentPath) -replace '\s+', '_' }
            $tmpPng = [System.IO.Path]::Combine($env:TEMP, "$baseName.png")
            $resized.Save($tmpPng, [System.Drawing.Imaging.ImageFormat]::Png)
            $resized.Dispose()
            $files = [System.Collections.Specialized.StringCollection]::new()
            [void]$files.Add($tmpPng)
            [System.Windows.Forms.Clipboard]::SetFileDropList($files)
            Write-Log "Icon copied to clipboard as 48x48 file : $tmpPng"
        }
        catch { Write-Log "Failed to copy icon as file : $_" -Level Warning }
    })
    $itemRaw = [System.Windows.Forms.ToolStripMenuItem]::new("Copy Raw Image")
    $itemRaw.Tag = $originalBitmap
    $itemRaw.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetImage($this.Tag)
            Write-Log "Icon copied to clipboard as raw image ($($this.Tag.Width)x$($this.Tag.Height))"
        }
        catch { Write-Log "Failed to copy raw image : $_" -Level Warning }
    })
    [void]$menu.Items.Add($itemFile)
    [void]$menu.Items.Add($itemRaw)
    $menu.Show($this, [System.Drawing.Point]::new(0, $this.Height))
})
$iconButton_Tab1.Add_MouseUp({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) { $s.PerformClick() }
})

$labelFileName   =  gen $tabPage1 "Label" "" 0 0 0 0 'AutoSize=$true' 'Font=Arial,16' 'Visible=$false'
$labelFileSize   =  gen $tabPage1 "Label" "" 0 0 0 0 'AutoSize=$true' 'Font=Segoe UI, 8' 'ForeColor=Gray' 'Visible=$false'

# Vertical separator between left and right zones
$vertSep_Tab1 = gen $tabPage1 "Panel" "" 0 0 1 0 'BackColor=Gray' 'Visible=$false'

# Right panel for listviews
$rightPanel_Tab1       = gen $tabPage1 "Panel" "" 0 0 0 0 'Visible=$false'

$tabControl_Tab1Detail   = gen $rightPanel_Tab1 "TabControl" 0 0 0 0 'Dock=Fill'
$tabPageAllProps_Tab1    = gen $tabControl_Tab1Detail "TabPage" "All Properties"       'Padding=10'
$tabPageDetected_Tab1    = gen $tabControl_Tab1Detail "TabPage" "Detected Properties"  'Padding=10'
$tabPageFeatures_Tab1    = gen $tabControl_Tab1Detail "TabPage" "Features"             'Padding=10'
$listView_Tab1Features   = gen $tabPageFeatures_Tab1 "ListView" "" 0 0 0 0 'Dock=Fill' 'View=Details' 'FullRowSelect=$true' 'GridLines=$true' 'HideSelection=$false'
$listView_Tab1Props      = gen $tabPageAllProps_Tab1 "ListView" "" 0 0 0 0 'Dock=Fill' 'View=Details' 'FullRowSelect=$true' 'GridLines=$true' 'HideSelection=$false'
$allListViewItems_Tab1Props  = New-Object System.Collections.ArrayList
$panelFilter_Tab1Props       = gen $tabPageAllProps_Tab1   "Panel"   "" 0 0 0 20 'Dock=Top'
$searchTextBox_Tab1Props     = gen $panelFilter_Tab1Props  "TextBox" "" 0 0 0 0  'Dock=Fill'
$searchLabel_Tab1Props       = gen $panelFilter_Tab1Props  "Label"   "Filter:" 0 0 0 0 'Dock=Left' 'Autosize=$true'
Add-ListViewSearchFilter -SearchTextBox $searchTextBox_Tab1Props -ListView $listView_Tab1Props -AllItems $allListViewItems_Tab1Props
$searchTextBox_Tab1Props.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $searchTextBox_Tab1Props.SelectAll()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    }
})

# Detected properties panel
$detectedPropsOuterPanel   = gen $tabPageDetected_Tab1      "Panel"  ""                          0 0 0 0  'Dock=Fill'
$detectedPropsScrollPanel  = gen $detectedPropsOuterPanel   "Panel"  ""                          0 0 0 0  'Dock=Fill'   'AutoScroll=$true'
$detectedPropsCopyBtnPanel = gen $detectedPropsOuterPanel   "Panel"  ""                          0 0 0 25 'Dock=Bottom' 'Padding=10 0 10 0'
$detectedPropsCopyBtn      = gen $detectedPropsCopyBtnPanel "Button" "Copy checked as Arguments" 0 0 0 25 'Dock=Fill'
$detectedPropsFlowPanel    = gen $detectedPropsScrollPanel  "FlowLayoutPanel" ""                 0 0 0 0  'FlowDirection=TopDown' 'WrapContents=$false' 'AutoSize=$true' 'AutoSizeMode=GrowAndShrink'
$detectedPropsScrollPanel.Add_Resize({
    $newW = $detectedPropsScrollPanel.ClientSize.Width - 4
    if ($newW -lt 100) { $newW = 100 }
    $detectedPropsFlowPanel.MaximumSize = [System.Drawing.Size]::new($newW, 0)
    $innerW = $newW - 20
    foreach ($propPanel in $detectedPropsFlowPanel.Controls) {
        if ($propPanel -isnot [System.Windows.Forms.Panel]) { continue }
        $propPanel.Width = $innerW
        foreach ($ctrl in $propPanel.Controls) {
            if ($ctrl -is [System.Windows.Forms.FlowLayoutPanel]) { $ctrl.Width = $innerW - 4 }
        }
    }
})

foreach ($colName in @("Property", "Value")) {
    $ch = New-Object System.Windows.Forms.ColumnHeader; $ch.Text = $colName; [void]$listView_Tab1Props.Columns.Add($ch)
}
foreach ($colName in @("Name", "Default", "UI Title", "Description")) {
    $ch = New-Object System.Windows.Forms.ColumnHeader; $ch.Text = $colName; [void]$listView_Tab1Features.Columns.Add($ch)
}
foreach ($lv in @($listView_Tab1Props, $listView_Tab1Features)) {
    $lv.Add_KeyDown({
        if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $this.BeginUpdate()
            foreach ($item in $this.Items) { $item.Selected = $true }
            $this.EndUpdate()
            $_.Handled          = $true
            $_.SuppressKeyPress = $true
        }
    })
}

$listView_Tab1Props.Add_ColumnClick({ HandleColumnClick -listViewparam $listView_Tab1Props -e $_ })
$listView_Tab1Features.Add_ColumnClick({ HandleColumnClick -listViewparam $listView_Tab1Features -e $_ })

$contextMenu_Tab1Props = New-Object System.Windows.Forms.ContextMenuStrip
$menuCopyArg_Props = [System.Windows.Forms.ToolStripMenuItem]::new("Copy as Argument")
$menuCopyArg_Props.Add_Click({
    $sb = [System.Text.StringBuilder]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $listView_Tab1Props.SelectedItems) {
        $prop = $item.SubItems[0].Text
        $val  = $item.SubItems[1].Text
        $parts.Add("$prop=$val")
    }
    if ($parts.Count -gt 0) { [System.Windows.Forms.Clipboard]::SetText($parts -join " ") }
})
[void]$contextMenu_Tab1Props.Items.Add($menuCopyArg_Props)
$listView_Tab1Props.ContextMenuStrip = $contextMenu_Tab1Props

$contextMenu_Tab1Features = New-Object System.Windows.Forms.ContextMenuStrip
$menuCopyArg_Features = [System.Windows.Forms.ToolStripMenuItem]::new("Copy as Argument")
$menuCopyArg_Features.Add_Click({
    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $listView_Tab1Features.SelectedItems) { $names.Add($item.SubItems[0].Text) }
    if ($names.Count -gt 0) { [System.Windows.Forms.Clipboard]::SetText("ADDLOCAL=$($names -join ',')") }
})
[void]$contextMenu_Tab1Features.Items.Add($menuCopyArg_Features)
$listView_Tab1Features.ContextMenuStrip = $contextMenu_Tab1Features

# Generic MSI icon fallback
$GenericMsiIconBase64 = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAABepSURBVHhe7Zp3cJNntsbZ2d2bBikkgdAhQCD03iEQnLaQQArJZlNJsiXZm80GklBCApgeQg0BTDE2BtNsbIMtW5LVrC6rF6tYlmTLKu42LezdGT133vf7VDGB+0cyd2d4Z575JNme0e88zznnZYYuXe6cO+fOuXPunDvnzkk8kUjkruuRyLCO/yCR7/tzikQig1M5Oz16d+PQQ4eyLdvSt2HL+i3YTLRuM6P1rKLvk7QJmzrT2uT3G9dtxMa1t9IGqg03KB0bvu1c6Ulaf4O2btzy78OZJ4oAdE1lTjrp36Tv/ei7nZifX4T5eUWYf64Q888WUD19pgBPnz6Pp0+dx7zcfEYn8zH3RD7m5uThKaLjjOZkn8OcrHOYc+wcZh87i9mZZzH76FnMOnIGM4kOn8HMQ6cxI4PVwVOYfoAoF9P252Laj7mYuo/VDycxde9JTCHacwJTdp/A5F0nMGlXDibtzMGkHTmY+P1xRtuzMfG7bEzYlo3xW7Ooxm3Nxvj1Gfhg2dfYsvG7JanMSefdl149vpgnxIyf/o1ZTe2Y2dgWmdnQhhnhNkwPtWJ6sBXTAq2Y5m/BtLoWTK1txlRfM6Z4mzHF04zJNU2Y7G7CpOpGTHI1YqKzERMdDZhQ1YAJtjDGW8MYbwlhnDmEsaYQxhqDGKNnpQtitDaA0ZoARqnrMUpVj5HKeoxQ+DFC7scImR9PVtThSUkdhotrMUxUi2FCH4YJfHii3IuhPG9kKM+LIWUeDCn1YDDHg8ElNXi8uAYDuEE8s2Yn3kx7+qNU5qTz3uIlmYvLhAQcs70BzPLUY1ZNPWa56zGzuh4zXH7McPox3VGH6fY6TKuqwzRbLaZZazHV7MNUcy2mGH2YYvBhst6LyTovJmm9mFTpwUSNBxPVNZio8mCCsgbjFW6Ml7sxTlaNcdJqjKuoxlhJNcaKXBgjdFKNLndiFJ/IgVE8B0aW2TGy1I4RRJwqPFlShSeLqzD8gg3DilgVWPHEeSueyLdiaJ4FQ89ZMKjQjWfX7MKbafOWpjInHVqA0nLqOAGf6a6PzKz2M/BOP2aw4NOj4BZGFN7Eghu8FD4FPELgJyjdDLzcjfEydzK4uBpjRC6MFrowWhCFd2AklxWFZkTBLzLgFL6Q0RPnrZEn8i0gIuBDz1ow5KwFA89X45mvSQJuUYB3Fi/JXFRSjhnBVgKPmazj1PVEcOI4Ba/FFAJuJOAJruu8CY7XRCaoCHwCOFEM3EXBqeNRcCLiOAEvizpux4gSO+M4gS+yRoYTxwl4zHULdX3IOQZ8yGlTZPApU2RAfjWeWb0Db6Y9dYsCLHo186WLfEwPtMbjTsBp3JPBGXgm8hSeBZ+kjcc9Bh6LOwvOwo8h8EJnJO446zqNe4LrDHiEghPXKbg1FvehLDhxnYAPPm1mlGvC47km9D/nQtqq2ynA4lczX7zAowNuhoNE3s+6XpfQ5z4GnESdiIWmcWcjHwdnXKfgUhacQHfmOulzNu7UcRY+Ke5RcCbuCX1uZhw/w8KfMjHwJ4wYlGOI9DvtRNrK2ynAolczFxbxMLW2hYJH4RnXfZF43Nk+jzleE5moiTs+gTpeEwMfJ3XfGPdyRyTmOhv3ESTubJ+PiIJftGF4oTUSBR9GwInrBJyIgpsxhIIzevxkFN6IQceN6JfrwPwV399mAQq4mOJtoo7HIk/BvREGPB53xnHa52TARcZH4x6d7FHXSdyj0z3RcZ4TI4lifV7FgCe4Tid7UtytzICL9vmZhLgngucYMfC4EQOzjOh70o75X91WAV7JXHCeS/c57fNOHY9O9xpMVCfGPXGtuZPjzkx21nFHhDgejfvIUkfSSrsh7gXRuCf3eSzubJ8/ftKIQSdM1PFB2UYMJMoyYECmHn1y7Hj6y+14/bYKkF9GLzHEcWafeyIx+EoPJrFrLR53Bj6pzyUEPLrWkvt8JNcRYaY7AWejHt3nFL4qeZ8TxcDNjOOnyHQ3s+DJcSfQFPyYAf0z9eh/RI/e2VWY98V3ty7AW4teyfxDXhm9wdGVlnqZYdYaE3cWnFlr7D6Pun4DOOlzZ/JaSwJn93lsrdmiAy5C11pqn6eCk7hHHT+mJ65H+h8l8Dr0O6xDr2NVmLf8NgvwwrlSTLA3xAZc4i0uNe7j2YvMOImLgsducQRckHqLS3GdxJ3e4uhOp44n3eLoRYY6HokPODbuBJzouAEDsxkRx0ncY+BH9Oh3SIe+GTo8dtSGectutwBnOPTefuM+d2OCvJoZcBVOjBM7ImNFDowV2DGmvAqjifhVGM1jNIpnxyiuHaMoeMJkTwSP9non+5zC031uYsCjjkfBWfio4wNijusZUXgt+h7UoucRG+Z+vu32CvD8KQ7GW8KJcWfBXRgncWCsyI7R5TYM51owhGPEoIsGDCjSYUChDgOL9Bh8wYBhF00YUWzGiGILRhRbMaLEFh9wiXHv5N6eNNmjjp80kZhHmMnOun5MH6Gus+D9D+soNFWGjoL3OVCJPvsr0SPDiqf+ufV2CrA487lTJRhnCmEiHXBx8DFCG4aWmTCgUIPheTLMzK/A4iIp3iuW4aNiGT4sluHNIinSzoox5qQY/U7I0Oe0GkPyDRheYMLwQguGFxElRp11nPY5ucxE+5y9yCT0OV1p7JBjXWfhieMMeNTxPge06LNfiz4/VqL3Pg16HLRgDi3ArNsoQG4JxhoCcXCRDYNLjRiSL8eLhRJsFalRZrLB7vUiEAqiuakBrS1NaG1uRENDGF6/HxqHC8cVevy1qAJjswXomS3HoDM6DMsz4ol8EwWPrzUWnLqeEHfiehSeHXCJ8BQ82ucHtREC3peCV6I3Ba9E7x806LVXg0f3mzHns9sswLMnijGm0o9xYjuG88wYdF6BNy5IkF9pRH2gHpfaW9Ha0Y5wexuCra0ItLagvoVRoLUVobY2NHe049KldrS3NMHgqsaGUjlGHy1HzywFhpzWYegZI+1xdsjF437SGO/z6Fojfc7u81jcWfDokOvLRp1xPAqujvTao8ZjuzV4ZJ8Zc/6xBa/Pu0UB/vTi4sxnsi9glMKLQSV6TDknRqZci4ZQAO0d7RTY39z8s6pPUKC5hRbjSnsr9HYnPjgrQo+DwsiAHA2G5Oox5JQRQzqJOx1yqY4fZfs82usk7sRxIjbqvX/QRIjjRI/tVuOxXWr03KXGw3tNmP3pbRTgjy8uypx/tAB9uFYsLhDD6HDh6uUOCl7X1HRb8t9ELR3taG0MYRdPjv4/8tHnmBqDc3QUOukiQ+JOB1wUnIl7dMhRx2N9Xok+xPGo63tY8N1q9NypQs+davT4XoXuu02Y/d+bb12ANxYsypy49wTe5qoRrPfj2rUrNNKBlhYKUdfYSFXLPjuT/yYiP6tvbUbblQZky5UYuJeL3kdUeDxbSx0fRCZ7ouMsfNxxAq5LGXAEvJI6zsBrGNcJ/PcqCt9juwoP7TRi1t833XoGzJk7P/PjfUcRamnB9es/obm9A41tbQi1tCLQ3MzANzTcVHU/p6YwPC1B2K7WIfA/fmRJVOi1m4e+h1UYmKklrscvM7TPE8AT406GHHWdiTrT52zcd6rRc0ccvMd3Sjy6TYmHvjdg1ieb8NqtCvDBR3876PF4cOWna2jp6EDbpUtobm9HAxluZNA1NVFQX5gB9rGvo89oEWJP9rPaRgbeeqUWlis+1LQH0dHRgI0XKvDgznL0P6TGgCNaDCB9Hl1rbNSTpnsSeDzujxFwNu4UfDsD/uhWBR7ZosSD2w2Y+fEmvHaLFvit3misvPSvfyHY3IzG1lZahNaEIpDP60kKwmH4OhH5/AY1hlDTEoD1so/Cu1sC8DWGEGhpQjjox6IjXDy4W4L+GWr0P1QZubHPE+Peues9oq6zjj+6VYlHtijwyGYFHt6kwP1b9Zjxt414LW36zQsgEok+brp8Gb5QCHXhMIUl0LQIHR1oamtDmC0C6WkC5w2FkkT+NkkNIdQ0J8LXw9cQRG04BF84hObWZgh0JvTdXope+xTod0CDfgmO39jnzFqj4DtUCXGn4BEC/whxnYLLqbpvkOH+zbqfL8Dy7cvvM1mtgXBbG9x+PzyBAC1CoKmJKUJ7e7wILa0IkoFIYk/Ag8GYfEQhVuEgaprrYbnshfmKF9Ut9fA2JPyc/m4IrY1B/Pl4Obp9J0LfH1Xou19zS8eTwBNdJ+Cs6w9vJPBydF8vQ7dNWkz/ywa8Nu8mBVCWy96/Gr4Of20DPP561JAi1AdQFwoh0NiIhtYWpgjt7WhqbUWYXHqamhBsbkJDcxjhxgAaGoMINZHYh+ANB+BuqoflkhfmywTeD29DAN4QWyzyDASoGpsbwVEb0WMzB4/tlqPPPnVyn0d7nQWn8AkDjunzKLgc3Sm4DN3TZXhovQwPrZOh6wYtpv85/eYFsKvs5fAC14xX0VbVjmBNA3y1AXhJEYJMEcLkUtPWRgtBZkJHezPMLjfOyGw4yDPjMN+EUo0NvoAPde1BFt6D6mY/POEAvMEAPIEgoyB5kvcB1IZC8Pt9SNtXige2SdB7jxK96UUm4TJD4GMDTpUEHot71HECvo7Rg0RrpbhvvRbTblaAgwcP9rJbXVdb/e1otbfhqukqrhuu4bL5MpqcLQj4wvAHQgiEGxBuaqat0NHWhFyJDZ+etmFprgtv5lTjj9kOvJNlxhenKiGptsB13QdnUx08oXp4A/Xw1N9cLU0hrD4nxr3pfPTaKUevXar4Pt8Rn+5M3BVMn8dcV8TB10sjjOtSCv7gt1I8sKYC966txNSPblIAPp+/kLhcU1sHr9cPvyeIBlcTLlk78JPhGq4ZrqLD1oFGTwvCwSZ0tDQhS2jFh6eq8c7pWryQ6cHcDDdmH3Di2UN2LDlmw7tHVVA7q+BvCsJT74fH76dtReYLeaaqoTGEXJEGD67joMd2GR7boWTAt6sicfBU16N9nhj3BPBvGPj7v67APd9oMPXD9Z0XoEIuXxFsaYHd44HT44XL64PbWwuvx496dwjN9hZcMV2mqfh31XVo1R789bSdws877MOiLA9NwCtZTkz7oQpT9ljw4iETVp9WIhD0MfB1dXAnqJo8a2upqmtr4Q8GINGZ0W9DCR7eWoGe2xV0wDF9rkhZa/HpngzOxD0KTuFXV6DbKgnu/lqNqR+s73wLKNTqfeSWZ3O7I1U1NbCzctZ44KrxosZTi9qaeoRcDfiXvR2HL9jx3ukaPHPEi9dyvHg1uwaDNlgxfocVLx+1YexOM57aZ8TrGUooTDbUBvyo9vkoqMvn61Refx30Vhue3FyCBzaK0WObHD2I49G1RuIen+yRJPCUuBPHqQj8SjG6rhDj7lUqTFm6rvMEqCorj5GJbquuJkWIqcrthp0VKYbb40OduwbpBWa8caIGT2XU4E85LnT9xoouy63o8rkV0/ZYMfdHI8bv0GHBfjXOSw3wB+rg9Png9JJ0xRV9T541dbWwOuwYs6UY3dJFeHSrPOZ63HGmz7uny5Pi/gCNewr4Kgm6rZRQ+K5fiXHXChUmv08KMPXGAig0mmPkimt1uVJUHXttc7loQaqrHVh73ojXjldj1gEXXs9y4L41bAGWWTFjrw1zftBjzPZKLNyvRp5EB1+dFw7aXp4YfDURSQUrT10dbA4HxmwuRrd1Ijy6SZbS59Ehx073WJ+zUf9agvtXSxLAJej6pRj3EX0hwn99pcLk99bi5c4KIFcqf/CHw7A4nbDeQm63Hd9f0OKNbAem77PjuUN2LMiw0RaYsMuOxZlVGLpFg+m7tXjjkArFEjmUOi2kGg2VrLIScq0WCp0OSoMBaqMRGpMJRpsVCq0OT24oRre1IjyyQZoAnuA4Cx8FfyDJcTG6EfCvxOj6pYiC37tchHuXCfH7L5SY9O5NClBRUbHCHwrB4nDA6nDQJ5XdAbPdTmWqqqKyuxzgyLR465gRC49WYcxOC5760YyFh81I22/E8G2VGL5VjecP6PDlCRlMJh1kGg3EKhVESiWVUKGAQC5HuUwGvkwKnlQKsVyGfJ4Ifb8pxgNrRXg4XcrEnYAnTXfWcdb1bqzr1PEVxHECLsZ9FFyEez8X4p5/CvG7ZQpMevfbzgtA1iCJZBTUSGSzwRCVNeG1zQaHw4pdRSq8etSMPxyyYMJOA4Zv02LYVg0m7KjE8wf0+FOGHAqdEdYqG5Q6HSrUaoiVSioJ+4xJoYC6UoWMIgG6rSjGQ2vF6L5eyl5kOltrLPjK5D6ncV8uwn0UnIX/TIC7/yHEb/+pwIR31uKVuVM+TOXvkpub21uj010jEddbLDAQWYmsncpcVQW9XoPvzgrx1hENXszQ45n9RDq8dFCDT7Lk0Jqr0NIQgt3lgt5khkKrhUSpgkgu71QGrRL/yBahy0oRuqWr0G1jJb2+dk3X4r71lbh3nQb3rtXgnm81uGeNmq61u1ercNcqFe5aqcJdK5S0z//rSyWN+++XKfC7zxUU/LefKdDlMw2mvb0ai+dPey+Vnx65Uikkw0hnMkFvNsekY8W8t8BoJb2qBVckgqRCiIKycuzKEyD9tAhbz4pxskwKi9mAWp8HgVAI9fX1qHI6oTOaoCCtoFBAKJVBKJXGJZNBrpRj8d/XYOSbX2L6n9fTf7hM+0s6vb5O+ygdUz9aTy8yRFM+YLV0HRWZ7pPfX0uHHOlzqne+pZr4zreY8PZaTH1rNeb/YXHtggVP90llp6dcJHrf7fNBazJBazRCZzQyr6PvSWEsFsjUanDKy6lKBULwhAKIxeWQiPmQSUVQa5RQVWrp7zurqxEk93xSBIcDlQYDZGQWyGQor6iISayQ40JpGZYtX4Glb75+lOzq29NUKtLXRImvE0U+fyVt+jtpaTN6p3LHztl9+7oq1Gq/zemE1mCgIl84+poAkS9fzOOhhIjPj4v9jMPno1wigZg4qlbTIjpdLgSDQVoEm92OSr0eUpUKQgIvFoMnFqNCIUfm8VzsPJCFH44c//n/x/dLHqFQ+Inb64XGYIBGp4NGr6ciCahQKnGxrIwoUszl4mbiCgTgi8UQSaW0YKR4NxSBDEWyCSQSqlIeH9v3HsK+oyexbU/GB6nf61c7S5Ys+a1UJlM73W6otVpaBOKYRC7HBQ6HihahtDRZCZ9xeDxwy8vBEwohlEggVSig1evhcDoRIEXw+2GrqoJGq0WFXI4KmQyZOWfw/Y+ZOJRzDrsPZr2d+r1+1VNYWDhCq9NdNttstAAkzoUlJSgqLkYRh4Mi8rqkhBajqKQkQl/H36OkrAxlfD5K2UKUi0QUlLhuTymCwWjA+Qsc7DmUg/2Zp5B19iJ2789emPqdfvXD4XBeNprNkKtUOF9UhIILFyKFpADFxSi8eBEFFy8i8X2iLnI4FJ7D5dInKQZfKKRO0yI4HAiGgggE6sEXiHAg6wwyjp/DoZw8HDl5HnklArFYZcrky/QJ0maWK4yZIrVlH0+ifS71+/4ip6ysbClfIIgQELYIsSdV4uuEn5GicMrKaBKiIoXgCwSQSKWo1Grh83lhd1ajXK6HQGGEQGlEucIAnkwHnb0WNl8jrJ4GWDxhmGtCMBG5gzDXhGFwBSA3udekft9f5JSWli4QCIUNQrEY+QUFyC8sjOn8TUQKU8zhxFVaimJOKUpKSyEUiaDV6XAw4zA+Xb4SX61Jx4pv0ukzqi9Wr8PyVWuplq36FstWMvqcaMU3WLV2Mz757MuGLl26/D71+/4ih8vlDioXCIoFIhFKuVzknT+Pc/n59JkqWqSCAlwoLsbFkhIq8pr8nUwuB49fjlVr1uLp5xbgqfnPY27a85j3zAuYF33ehtKeX4gJU6Y3/2oFiB6BQLCEy+cry3g88MrLKVhefj7O5uXh7LlzceXlobCoCJzSUghI70ul4PH5rTweb+9fPv770hFjJy2d9+wLS9NeeGlp2gsL/896duHLS2fMnruoS5cuv0n9jr/KEYlEcwlMCYdjKSoqul7C4YBLepysPjL8SktBPistK2vg8/lcgUDwKY/Hu/kN7D/1APiNVCp9XCgUzudyuW/zeLylpVzu+1wu92WRSDRJoVB0T/2bO+fOuXPunP9v538BYpSk+sZH/jYAAAAASUVORK5CYII="
$GenericMsiIconBytes  = [System.Convert]::FromBase64String($GenericMsiIconBase64)
$GenericMsiIconStream = [System.IO.MemoryStream]::new($GenericMsiIconBytes)
$GenericMsiIcon       = [System.Drawing.Image]::FromStream($GenericMsiIconStream)
$pictureBoxIcon.Image = $GenericMsiIcon

function MeasureTextWidth {
    param ([string]$text, [System.Drawing.Font]$font)
    $bitmap   = New-Object System.Drawing.Bitmap(1, 1)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $sizeF    = $graphics.MeasureString($text, $font)
    $graphics.Dispose()
    $bitmap.Dispose()
    return [int]$sizeF.Width
}

function Show-Tab1ListView {
    param([int]$Index)
    $tabControl_Tab1Detail.SelectedIndex = $Index
    if ($Index -eq 0) { AdjustListViewColumns -listView $listView_Tab1Props }
    if ($Index -eq 2) { AdjustListViewColumns -listView $listView_Tab1Features }
    $currentPath = $null
    foreach ($ctrl in $msiSelectorFlowPanel.Controls) {
        $innerRadio = $ctrl.Controls | Where-Object { $_ -is [System.Windows.Forms.RadioButton] } | Select-Object -First 1
        if ($innerRadio -and $innerRadio.Checked) { $currentPath = $innerRadio.Tag; break }
    }
    if ($currentPath -and $script:MsiFileCache.Contains($currentPath)) { $script:MsiFileCache[$currentPath].SelectedListView = $Index }
}

function Update-Tab1FromCache {
    param([string]$FilePath)
    if (-not $script:MsiFileCache.Contains($FilePath)) { return }
    $cached     = $script:MsiFileCache[$FilePath]
    $msiResults = $cached.Results
    # Property fields
    $i = 0
    foreach ($key in $properties.Keys) {
        $value = $msiResults[$properties[$key]]
        if ($value -and $value -ne "None") { $textBoxes[$i].Text = $value.Trim(); $copyButtons[$i].Enabled = $true }
        else { $textBoxes[$i].Text = ""; $copyButtons[$i].Enabled = $false }
        $i++
    }
    $script:fromBrowseButton = $true
    $textBoxPath_Tab1.Text = $FilePath
    # Properties ListView
    $listView_Tab1Props.BeginUpdate()
    $listView_Tab1Props.Items.Clear()
    foreach ($prop in $msiResults["_AllProperties"]) {
        $item = New-Object System.Windows.Forms.ListViewItem($prop.Property)
        [void]$item.SubItems.Add($prop.Value)
        [void]$listView_Tab1Props.Items.Add($item)
    }
    $listView_Tab1Props.EndUpdate()
    $allListViewItems_Tab1Props.Clear()
    foreach ($item in $listView_Tab1Props.Items) { [void]$allListViewItems_Tab1Props.Add($item.Clone()) }
    if ($searchTextBox_Tab1Props.Text -ne '') { $searchTextBox_Tab1Props.Text = '' }
    # Features ListView
    $listView_Tab1Features.BeginUpdate()
    $listView_Tab1Features.Items.Clear()
    foreach ($feat in $msiResults["_Features"]) {
        $levelInt    = 0
        $defaultText = if ([int]::TryParse($feat.Level, [ref]$levelInt)) { if ($levelInt -gt 0) { "Yes" } else { "No" } } else { $feat.Level }
        $item = New-Object System.Windows.Forms.ListViewItem($feat.Name)
        [void]$item.SubItems.Add($defaultText)
        [void]$item.SubItems.Add($feat.Title)
        [void]$item.SubItems.Add($feat.Description)
        [void]$listView_Tab1Features.Items.Add($item)
    }
    $listView_Tab1Features.EndUpdate()
    # Detected Properties panel
    $detectedPropsFlowPanel.Controls.Clear()
    $fontBold    = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $fontRegular = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $panelWidth  = $detectedPropsScrollPanel.ClientSize.Width - 4
    if ($panelWidth -lt 200) { $panelWidth = 200 }
    $detectedPropsFlowPanel.MaximumSize = [System.Drawing.Size]::new($panelWidth, 0)
    $innerW = $panelWidth - 20
    foreach ($prop in $msiResults["_MultiValueProps"].Keys) {
        $entry   = $msiResults["_MultiValueProps"][$prop]
        $default = $entry.Default
        $source  = if ($entry.Options.Count -gt 0) { $entry.Options[0].Source } else { "" }
        # Container for one property
        $propPanel        = New-Object System.Windows.Forms.Panel
        $propPanel.Width  = $innerW
        $propPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
        # Label
        $propLabel          = New-Object System.Windows.Forms.Label
        $propLabel.Text     = "$prop ($source) :"
        $propLabel.Font     = $fontBold
        $propLabel.AutoSize = $true
        $propLabel.Location = [System.Drawing.Point]::new(2, 2)
        $propPanel.Controls.Add($propLabel)
        # Flow panel for value buttons
        $valuesFlow               = New-Object System.Windows.Forms.FlowLayoutPanel
        $valuesFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
        $valuesFlow.WrapContents  = $true
        $valuesFlow.AutoSize      = $true
        $valuesFlow.Location      = [System.Drawing.Point]::new(2, 22)
        $valuesFlow.Width         = $innerW - 4
        foreach ($option in $entry.Options) {
            $isDefault  = ($option.Value -ceq $default)
            $btnText = if ([string]::IsNullOrEmpty($option.Value)) { '(empty)' } else { $option.Value }
            # Wrapper panel for checkbox + button
            $wrapper        = New-Object System.Windows.Forms.Panel
            $wrapper.Height = 26
            $wrapper.Margin = New-Object System.Windows.Forms.Padding(0, 1, 10, 1)
            $cb             = New-Object System.Windows.Forms.CheckBox
            $cb.Location    = [System.Drawing.Point]::new(0, 4)
            $cb.Size        = [System.Drawing.Size]::new(16, 18)
            $cb.Margin      = [System.Windows.Forms.Padding]::new(0)
            $cb.Tag         = $valuesFlow
            $cb.Add_CheckedChanged({
                if (-not $this.Checked) { return }
                foreach ($sibling in $this.Tag.Controls) {
                    if ($sibling -isnot [System.Windows.Forms.Panel]) { continue }
                    foreach ($inner in $sibling.Controls) {
                        if ($inner -is [System.Windows.Forms.CheckBox] -and $inner -ne $this) { $inner.Checked = $false }
                    }
                }
            }.GetNewClosure())
            $cb.Checked     = $false
            $btn            = New-Object System.Windows.Forms.Button
            $btn.Text       = $btnText
            $btn.Font       = if ($isDefault) { $fontBold } else { $fontRegular }
            $btn.AutoSize   = $false
            $btn.Location   = [System.Drawing.Point]::new(16, 0)
            $btn.Height     = 25
            $btn.Tag        = "$prop=$($option.Value)"
            $btn.FlatStyle  = [System.Windows.Forms.FlatStyle]::System
            $btnWidth       = [System.Windows.Forms.TextRenderer]::MeasureText($btnText, $btn.Font).Width + 20
            $btn.Width      = $btnWidth
            $btn.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($this.Tag) })
            $wrapper.Controls.AddRange(@($cb, $btn))
            $wrapper.Width  = 16 + $btnWidth
            $wrapper.Margin = New-Object System.Windows.Forms.Padding(0, 1, 14, 1)
            $valuesFlow.Controls.Add($wrapper)
        }
        # Calculate panel height
        $valuesFlow.PerformLayout()
        $propPanel.Height = 22 + $valuesFlow.PreferredSize.Height + 4
        $propPanel.Controls.Add($valuesFlow)
        $detectedPropsFlowPanel.Controls.Add($propPanel)
    }
    # Restore tab selection
    $selectedLV = $cached.SelectedListView
    $tabControl_Tab1Detail.SelectedIndex = $selectedLV
    # Icon
    if ($cached.Icon -and $cached.IconBytes) {
        $btnW = $iconButton_Tab1.Width - 6
        $btnH = $iconButton_Tab1.Height - 6
        $resized = [System.Drawing.Bitmap]::new($cached.Icon, $btnW, $btnH)
        $iconButton_Tab1.Image   = $resized
        $iconButton_Tab1.Tag     = $cached.Icon
        $iconButton_Tab1.Visible = $true
        $pictureBoxIcon.Visible  = $false
    }
    else {
        $pictureBoxIcon.Image    = $GenericMsiIcon
        $pictureBoxIcon.Visible  = $true
        $iconButton_Tab1.Visible = $false
    }
    # Display state
    $script:MSILoaded         = $true
    $labelFileName.Text       = $cached.FileName
    $fileSizeBytes            = 0
    $fileSizeBytes            = 0
    try { $fileSizeBytes      = ([System.IO.FileInfo]::new($FilePath)).Length } catch {}
    $labelFileSize.Text       = Format-FileSize -Bytes $fileSizeBytes
    $labelFileSize.Visible    = $true
    $labelDropMessage.Visible = $false
    if (-not $iconButton_Tab1.Visible) { $pictureBoxIcon.Visible = $true }
    $labelFileName.Visible    = $true
    $labelFileSize.Visible    = $true
    $rightPanel_Tab1.Visible  = $true
    $findGuidButton.Text      = "GUID FOUND"
    $findGuidButton.Enabled   = $false
    Update-ButtonPositions
    if (-not $script:resizePending) {
        foreach ($lv in @($listView_Tab1Props, $listView_Tab1Features, $listView_Tab1MultiVal)) {
            if ($lv.Visible) { AdjustListViewColumns -listView $lv }
        }
    }
}


function Add-MsiToSelectorPanel {
    param([string]$FilePath)
    # Check if row already exists
    $existingRow = $null
    foreach ($ctrl in $msiSelectorFlowPanel.Controls) {
        $innerRadio = $ctrl.Controls | Where-Object { $_ -is [System.Windows.Forms.RadioButton] } | Select-Object -First 1
        if ($innerRadio -and $innerRadio.Tag -eq $FilePath) { $existingRow = $ctrl; break }
    }
    if ($existingRow) { $msiSelectorFlowPanel.Controls.Remove($existingRow) }
    $cached     = $script:MsiFileCache[$FilePath]
    $prodName   = $cached.ProductName
    $isNone     = ($prodName -eq "None")
    $pathString = $FilePath
    # Build row panel
    $rowPanel             = New-Object System.Windows.Forms.Panel
    $rowPanel.Height      = $script:MsiSelectorRowH
    $rowPanel.Margin      = New-Object System.Windows.Forms.Padding(0, 0, 0, 3)
    $radio                = New-Object System.Windows.Forms.RadioButton
    $radio.Text           = ""
    $radio.Location       = [System.Drawing.Point]::new(0, 2)
    $radio.Size           = [System.Drawing.Size]::new(18, 20)
    $radio.Tag            = $FilePath
    $lblName              = New-Object System.Windows.Forms.Label
    $lblName.Text         = $prodName
    $lblName.Font         = if ($isNone) { $script:MsiSelectorFont } else { $script:MsiSelectorFontBold }
    $lblName.AutoSize     = $true
    $lblName.Location     = [System.Drawing.Point]::new(20, 4)
    $lblName.ForeColor    = if ($isNone) { [System.Drawing.Color]::Gray } else { [System.Drawing.Color]::Black }
    $lblArrow             = New-Object System.Windows.Forms.Label
    $lblArrow.Text        = " -> "
    $lblArrow.Font        = $script:MsiSelectorFont
    $lblArrow.AutoSize    = $true
    $lblArrow.ForeColor   = [System.Drawing.Color]::Gray
    $lblPath              = New-Object System.Windows.Forms.Label
    $lblPath.Text         = $pathString
    $lblPath.Font         = $script:MsiSelectorFont
    $lblPath.AutoSize     = $false
    $lblPath.AutoEllipsis = $true
    $lblPath.Height       = 20
    $lblPath.ForeColor    = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $rowPanel.Controls.AddRange(@($radio, $lblName, $lblArrow, $lblPath))
    $rowPanel.Tag = @{ Radio = $radio; LblName = $lblName; LblArrow = $lblArrow; LblPath = $lblPath }
    # Position labels dynamically
    $layoutRow = {
        param($rp)
        $t      = $rp.Tag
        $nameW  = [System.Windows.Forms.TextRenderer]::MeasureText($t.LblName.Text, $t.LblName.Font).Width + 2
        $arrowW = [System.Windows.Forms.TextRenderer]::MeasureText($t.LblArrow.Text, $t.LblArrow.Font).Width + 2
        $t.LblName.Location  = [System.Drawing.Point]::new(20, 4)
        $t.LblArrow.Location = [System.Drawing.Point]::new(20 + $nameW, 4)
        $pathLeft  = 20 + $nameW + $arrowW
        $pathWidth = $rp.Width - $pathLeft - 5
        if ($pathWidth -lt 20) { $pathWidth = 20 }
        $t.LblPath.Location = [System.Drawing.Point]::new($pathLeft, 4)
        $t.LblPath.Width    = $pathWidth
    }
    & $layoutRow $rowPanel
    # Click anywhere on row selects the radio
    $clickHandler = { $radio.Checked = $true }.GetNewClosure()
    $rowPanel.Add_Click($clickHandler)
    $lblName.Add_Click($clickHandler)
    $lblArrow.Add_Click($clickHandler)
    $lblPath.Add_Click($clickHandler)
    # Radio group management (manual since radios are in separate containers)
    $radio.Add_CheckedChanged({
        param($s, $e)
        if ($s.Checked) {
            foreach ($ctrl in $msiSelectorFlowPanel.Controls) {
                $otherRadio = $ctrl.Controls | Where-Object { $_ -is [System.Windows.Forms.RadioButton] } | Select-Object -First 1
                if ($otherRadio -and $otherRadio -ne $s) { $otherRadio.Checked = $false }
            }
            Update-Tab1FromCache -FilePath $s.Tag
        }
    })
    $radio.Checked = $true
    $msiSelectorFlowPanel.Controls.Add($rowPanel)
    $msiSelectorOuterPanel.Visible = $true
    Update-ButtonPositions
}


function Update-LabelAndIcon {
    param ([string]$filePath)
    $script:MSILoaded         = $true
    $labelFileName.Text       = [System.IO.Path]::GetFileName($filePath)
    $labelDropMessage.Visible = $false
    $pictureBoxIcon.Visible   = $true
    $iconButton_Tab1.Visible  = $false
    $labelFileName.Visible    = $true
    try {
        $labelFileSize.Text    = Format-FileSize -Bytes ([System.IO.FileInfo]::new($filePath)).Length
        $labelFileSize.Visible = $true
    } catch {}
    Update-ButtonPositions | Out-Null
}


function Reset-Tab1MsiState {
    $script:MsiFileCache.Clear()
    $msiSelectorFlowPanel.Controls.Clear()
    $msiSelectorOuterPanel.Visible = $false
    $rightPanel_Tab1.Visible       = $false
    $vertSep_Tab1.Visible          = $false
    $listView_Tab1Props.Items.Clear()
    $allListViewItems_Tab1Props.Clear()
    if ($searchTextBox_Tab1Props.Text -ne '') { $searchTextBox_Tab1Props.Text = '' }
    $listView_Tab1Features.Items.Clear()
    $detectedPropsFlowPanel.Controls.Clear()
    $tabControl_Tab1Detail.SelectedIndex = 1
    $pictureBoxIcon.Visible   = $false
    $pictureBoxIcon.Image     = $GenericMsiIcon
    $iconButton_Tab1.Visible  = $false
    $iconButton_Tab1.Image    = $null
    $labelFileName.Visible    = $false
    $labelFileSize.Visible    = $false
    $labelDropMessage.Visible = $true
    $script:MSILoaded = $false
    foreach ($tb in $textBoxes)   { $tb.Text = "" }
    foreach ($cb in $copyButtons) { $cb.Enabled = $false }
    $textBoxPath_Tab1.Text       = ""
    $findGuidButton.Text    = "FIND GUID"
    $findGuidButton.Enabled = $false
    Update-ButtonPositions
}


function Invoke-MsiLoad {
    param([string]$MsiPath)
    $msiPath = $msiPath.Trim('"')
    if (-not ([System.IO.File]::Exists($msiPath))) {
        $findGuidButton.Enabled = $true
        $findGuidButton.Text    = "FIND GUID"
        return
    }
    $isMsi = $msiPath.EndsWith(".msi", [System.StringComparison]::OrdinalIgnoreCase)
    if (-not $isMsi) {
        Update-LabelAndIcon -filePath $msiPath
        $rightPanel_Tab1.Visible = $false
        return
    }
    if ($script:MsiFileCache.Contains($msiPath)) { $script:MsiFileCache.Remove($msiPath) }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $propertyArray = @($properties.Values)
    $results = Get-MsiInfo -FilePath $msiPath -Properties $propertyArray -Full
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    if ($results["ProductCode"] -eq "None" -or [string]::IsNullOrWhiteSpace($results["ProductCode"])) {
        $findGuidButton.Enabled = $true
        $findGuidButton.Text    = "FIND GUID"
        return
    }
    $prodName = $results["ProductName"]
    if ([string]::IsNullOrWhiteSpace($prodName) -or $prodName -eq "None") { $prodName = "None" }
    $script:MsiFileCache[$msiPath] = @{
        FileName         = [System.IO.Path]::GetFileName($msiPath)
        ProductName      = $prodName
        Results          = $results
        SelectedListView = 0
        Icon             = $results["_Icon"]
        IconBytes        = $results["_IconBytes"]
    }
    Add-MsiToSelectorPanel -FilePath $msiPath
}

function Update-ButtonPositions {
    $formWidth  = [int]$tabPage1.ClientSize.Width
    $formHeight = [int]$tabPage1.ClientSize.Height
    # MSI Selector panel layout
    $selectorBottom = 13
    if ($msiSelectorOuterPanel.Visible) {
        if (-not $script:resizePending) {
        $msiSelectorOuterPanel.Location = [System.Drawing.Point]::new(5, 10)
        $msiSelectorOuterPanel.Width    = $formWidth - 10
        $selectorInnerWidth             = $msiSelectorOuterPanel.Width - 4
        $msiSelectorPanel.Width         = $selectorInnerWidth
        $scrollPanelLeft                = 105
        $scrollPanelWidth               = $selectorInnerWidth - $scrollPanelLeft - 70
        $msiSelectorScrollPanel.Location = [System.Drawing.Point]::new($scrollPanelLeft, 3)
        $msiSelectorScrollPanel.Width    = $scrollPanelWidth
        $resetMsiButton.Location         = [System.Drawing.Point]::new(($selectorInnerWidth - 65), 5)
        # Layout each row
        if (-not $script:resizePending) {
            foreach ($ctrl in $msiSelectorFlowPanel.Controls) {
                if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Tag -is [hashtable]) {
                    $ctrl.Width = $scrollPanelWidth - 25
                    $t = $ctrl.Tag
                    $nameW  = [System.Windows.Forms.TextRenderer]::MeasureText($t.LblName.Text, $t.LblName.Font).Width + 2
                    $arrowW = [System.Windows.Forms.TextRenderer]::MeasureText($t.LblArrow.Text, $t.LblArrow.Font).Width + 2
                    $t.LblName.Location  = [System.Drawing.Point]::new(20, 4)
                    $t.LblArrow.Location = [System.Drawing.Point]::new(20 + $nameW, 4)
                    $pathLeft  = 20 + $nameW + $arrowW
                    $pathWidth = $ctrl.Width - $pathLeft - 5
                    if ($pathWidth -lt 20) { $pathWidth = 20 }
                    $t.LblPath.Location = [System.Drawing.Point]::new($pathLeft, 4)
                    $t.LblPath.Width    = $pathWidth
                }
            }
        }
        # Scroll height limit
        $radioCount = $msiSelectorFlowPanel.Controls.Count
        $rowStep    = $script:MsiSelectorRowH + 1
        $maxH       = $script:MsiSelectorMaxRows * $rowStep
        $neededH    = $radioCount * $rowStep
        $scrollH    = [Math]::Min($neededH, $maxH) + 2
        $msiSelectorScrollPanel.Height = $scrollH
        $panelH = $scrollH + 4
        if ($panelH -lt 28) { $panelH = 28 }
        $msiSelectorPanel.Height       = $panelH
        $msiSelectorOuterPanel.Height  = $panelH + 4
        }
        $selectorBottom = $msiSelectorOuterPanel.Location.Y + $msiSelectorOuterPanel.Height + 5
    }
    # Property fields
    $textBoxWidth = $formWidth - 177
    foreach ($textBox in $textBoxes) { $textBox.Width = $textBoxWidth }
    $separatorLine.Width = $formWidth
    $propsBaseY = $selectorBottom + 5
    for ($i = 0; $i -lt $labels.Count; $i++) {
        $y = $propsBaseY + ($i * 30)
        $labels[$i].Location    = [System.Drawing.Point]::new(10, $y)
        $textBoxes[$i].Location = [System.Drawing.Point]::new(101, $y)
    }
    for ($i = 0; $i -lt $copyButtons.Count; $i++) {
        $copyButtons[$i].Location = [System.Drawing.Point]::new(($formWidth - 71), ($textBoxes[$i].Location.Y - 4))
    }
    $separatorLineY = $propsBaseY + ($labels.Count * 30) + 5
    $separatorLine.Location = [System.Drawing.Point]::new(0, $separatorLineY)
    # Bottom elements
    $buttonSpacing       = [int](($formWidth - 2 * $findGuidButton.Width) / 3)
    $browseButtonX       = $buttonSpacing
    $findGuidButtonX     = 2 * $buttonSpacing + $browseButton.Width
    $buttonY             = $formHeight - 40
    $browseButton.Location   = [System.Drawing.Point]::new($browseButtonX, $buttonY)
    $findGuidButton.Location = [System.Drawing.Point]::new($findGuidButtonX, $buttonY)
    $textBoxPath_Tab1.Location    = [System.Drawing.Point]::new(10, ($buttonY - 30))
    $textBoxPath_Tab1.Width       = $formWidth - 20
    $labelMsiPathY           = ($textBoxPath_Tab1.Location.Y - 18)
    $labelMsiPath.Location   = [System.Drawing.Point]::new(10, $labelMsiPathY)
    # Central zone boundaries
    $centralTop    = $separatorLineY + 1
    $centralBottom = $labelMsiPathY - 2
    $centralHeight = [Math]::Max($centralBottom - $centralTop, 50)
    # Right panel
    if (-not $script:resizePending -and $rightPanel_Tab1.Visible -and $script:MSILoaded) {
        $leftZoneWidth = 420
        $vertSep_Tab1.Location = [System.Drawing.Point]::new($leftZoneWidth, $centralTop)
        $vertSep_Tab1.Size     = [System.Drawing.Size]::new(1, $centralHeight)
        $vertSep_Tab1.Visible  = $true
        $rightPanel_Tab1.Location = [System.Drawing.Point]::new($leftZoneWidth + 2, $centralTop)
        $rightPanel_Tab1.Size     = [System.Drawing.Size]::new($formWidth - $leftZoneWidth - 2, $centralHeight)
    } else {
        $vertSep_Tab1.Visible = $false
    }
    # Icon, filename, filesize (stacked vertically, centered in left zone)
    $zoneWidth = if ($rightPanel_Tab1.Visible -and $script:MSILoaded) { 420 } else { $formWidth }
    $useIconButton = $iconButton_Tab1.Visible
    $iconCtrl      = if ($useIconButton) { $iconButton_Tab1 } else { $pictureBoxIcon }
    if (($pictureBoxIcon.Visible -or $iconButton_Tab1.Visible) -and $labelFileName.Visible) {
        $iconW = $iconCtrl.Width
        $iconH = $iconCtrl.Height
        # Truncate filename if needed
        $maxLabelWidth = $zoneWidth - 40
        $currentPath   = $textBoxPath_Tab1.Text
        $originalText  = if ([string]::IsNullOrEmpty($currentPath)) { $labelFileName.Text } else { [System.IO.Path]::GetFileName($currentPath) }
        $shortenedText = $originalText
        while ((MeasureTextWidth -text $shortenedText -font $labelFileName.Font) -gt $maxLabelWidth -and $shortenedText.Length -gt 1) {
            $shortenedText = $shortenedText.Substring(0, $shortenedText.Length - 1)
        }
        if ($shortenedText.Length -lt $originalText.Length) { $shortenedText += "..." }
        $labelFileName.Text = $shortenedText
        $textW  = MeasureTextWidth -text $labelFileName.Text -font $labelFileName.Font
        $textH  = $labelFileName.Height
        $sizeW  = if ($labelFileSize.Visible) { MeasureTextWidth -text $labelFileSize.Text -font $labelFileSize.Font } else { 0 }
        $sizeH  = if ($labelFileSize.Visible) { $labelFileSize.Height } else { 0 }
        $totalH = $iconH + 4 + $textH + $(if ($sizeH -gt 0) { 2 + $sizeH } else { 0 })
        $posY   = $centralTop + [int](($centralHeight - $totalH) / 2)
        $iconX  = [int](($zoneWidth - $iconW) / 2)
        $iconCtrl.Location = [System.Drawing.Point]::new($iconX, $posY)
        $nameY = $posY + $iconH + 4
        $textX = [int](($zoneWidth - $textW) / 2)
        $labelFileName.Location = [System.Drawing.Point]::new($textX, $nameY)
        if ($labelFileSize.Visible) {
            $sizeX = [int](($zoneWidth - $sizeW) / 2)
            $labelFileSize.Location = [System.Drawing.Point]::new($sizeX, ($nameY + $textH + 2))
        }
    } else {
        $labelDropMessageX = [int](($zoneWidth - $labelDropMessage.Width) / 2)
        $labelDropMessageY = $centralTop + [int](($centralHeight - $labelDropMessage.Height) / 2)
        $labelDropMessage.Location = [System.Drawing.Point]::new($labelDropMessageX, $labelDropMessageY)
    }
}

$tabPage1.Add_Resize({ Update-ButtonPositions })

$launch_progressBar.Value = 20

$textBoxPath_Tab1.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $findGuidButton.PerformClick()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    } elseif ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $textBoxPath_Tab1.SelectAll()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    }
})
$textBoxPath_Tab1.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) { $_.Handled = $true }
})

$browseButton.Add_Click({
    $openFileDialog             = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter      = "Supported files (*.msi;*.msix;*.mst;*.msp)|*.msi;*.msix;*.mst;*.msp|All files (*.*)|*.*"
    $openFileDialog.Title       = "Select one or more files"
    $openFileDialog.Multiselect = $true
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $validFiles = @($openFileDialog.FileNames | Where-Object { $_ -match "\.(msi|msix|mst|msp)$" })
        if ($validFiles.Count -gt 0) {
            $script:fromBrowseButton = $true
            $textBoxPath_Tab1.Text = $validFiles -join ";"
            foreach ($file in $validFiles) { Invoke-MsiLoad -MsiPath $file }
        }
    }
})

$findGuidButton.Add_Click({
    $paths = @($textBoxPath_Tab1.Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    if (-not (Test-RemotePathAccess -Paths $paths)) { return }
    foreach ($p in $paths) { Invoke-MsiLoad -MsiPath $p }
})

$resetMsiButton.Add_Click({ Reset-Tab1MsiState })

$textBoxPath_Tab1.Add_TextChanged({
    $script:fromBrowseButton = $false
    if ($textBoxPath_Tab1.Text.Trim() -eq "") {
        $findGuidButton.Enabled = $false
        $findGuidButton.Text    = "FIND GUID"
    } else {
        $findGuidButton.Enabled = $true
        $findGuidButton.Text    = "FIND GUID"
    }
})

$tabControl_Tab1Detail.Add_SelectedIndexChanged({
    $idx = $tabControl_Tab1Detail.SelectedIndex
    if ($idx -eq 0) { AdjustListViewColumns -listView $listView_Tab1Props }
    if ($idx -eq 2) { AdjustListViewColumns -listView $listView_Tab1Features }
    $currentPath = $null
    foreach ($ctrl in $msiSelectorFlowPanel.Controls) {
        $innerRadio = $ctrl.Controls | Where-Object { $_ -is [System.Windows.Forms.RadioButton] } | Select-Object -First 1
        if ($innerRadio -and $innerRadio.Checked) { $currentPath = $innerRadio.Tag; break }
    }
    if ($currentPath -and $script:MsiFileCache.Contains($currentPath)) { $script:MsiFileCache[$currentPath].SelectedListView = $idx }
})

$detectedPropsCopyBtn.Add_Click({
    $args_list = [System.Collections.Generic.List[string]]::new()
    foreach ($propPanel in $detectedPropsFlowPanel.Controls) {
        if ($propPanel -isnot [System.Windows.Forms.Panel]) { continue }
        foreach ($ctrl in $propPanel.Controls) {
            if ($ctrl -is [System.Windows.Forms.FlowLayoutPanel]) {
                foreach ($wrapper in $ctrl.Controls) {
                    if ($wrapper -isnot [System.Windows.Forms.Panel]) { continue }
                    $cb  = $null; $btn = $null
                    foreach ($inner in $wrapper.Controls) {
                        if ($inner -is [System.Windows.Forms.CheckBox]) { $cb = $inner }
                        if ($inner -is [System.Windows.Forms.Button])   { $btn = $inner }
                    }
                    if ($cb -and $cb.Checked -and $btn -and $btn.Tag) { $args_list.Add($btn.Tag) }
                }
            }
        }
    }
    if ($args_list.Count -gt 0) { [System.Windows.Forms.Clipboard]::SetText($args_list -join " ") }
})

$launch_progressBar.Value = 25


#region Tab 2 : Folders

$splitContainer2 = gen $tabPage2 "SplitContainer" "" 0 0 0 0 'Dock=Fill' 'SplitterDistance=250' 'Orientation=Vertical' 'Panel1MinSize=100'
$splitContainer2.FixedPanel = [System.Windows.Forms.FixedPanel]::None


# Panels for the left side
$borderTop =          gen $tabPage2                "Panel"     ""                         0 0 0 1    'Dock=Top'    'BackColor=Gray' 
$panelLeftMain =      gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 0    'Dock=Fill'
$treeView =           gen $panelLeftMain           "TreeView"  ""                         0 0 0 0    'Dock=Fill'   'CheckBoxes=$true'
$panelLeftCtrls =     gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 20   'Dock=Top'
$panelLeftCtrls_B =   gen $panelLeftCtrls          "Panel"     ""                         0 0 0 20   'Dock=Bottom'
$panelLeftOptions =   gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 20   'Dock=Top'
$sep1 =               gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 10   'Dock=Top'
$sep2 =               gen $splitContainer2.Panel1  "Panel"     ""                         0 0 1 0    'Dock=Right'  'BackColor=Gray'
$sortComboBoxLabel =  gen $panelLeftOptions        "Label"     "Sorting:"                 0 0 43 0   'Dock=Right'
$sortComboBox =       gen $panelLeftOptions        "ComboBox"  ""                         0 0 44 20  'Dock=Right'
    
$openBtn =            gen $panelLeftOptions        "Button"    "Open selected folder"     0 0 130 0  'Dock=Left'
$refreshBtn =         gen $panelLeftOptions        "Button"    "Refresh selected folder"  0 0 130 0  'Dock=Left'
$textBoxPath_Tab2 =   gen $panelLeftCtrls_B        "TextBox"   ""                         0 0 0 0    'Dock=Fill' 
$gotoButton =         gen $panelLeftCtrls_B        "Button"    "Goto"                     0 0 43 0   'Dock=Right'
$pathTextBoxLabel =   gen $panelLeftCtrls_B        "Label"     "Path: "                   0 0 0 0    'Dock=Left'   'Autosize=$true'
$panelLeftBottom =    gen $splitContainer2.Panel1  "Panel"     ""                         0 0 0 20   'Dock=Bottom'
   
$sep5               = gen $panelLeftBottom         "Panel"     ""                         0 0 1 0    'Dock=Right'  'BackColor=Gray' 
$ScanSelectedMSIBtn = gen $panelLeftBottom         "Button"    "Scan selected folder"     0 0 71 0   'Dock=Left'
$ScanCheckedMSIBtn  = gen $panelLeftBottom         "Button"    "Scan checked folders"     0 0 75 0   'Dock=Left'
$recursionComboBox  = gen $panelLeftBottom         "ComboBox"  ""                         0 0 36 20  'Dock=Left'
$recursionLabel     = gen $panelLeftBottom         "Label"     "Recursion:"               0 0 58 0   'Dock=Left'


$launch_progressBar.Value = 30

$recursionComboBox.Items.AddRange(@("No", "1", "2", "3", "4", "5", "All"))
$recursionComboBox.SelectedIndex = 6  # default = All
$sortComboBox.Items.AddRange(@("A-Z", "Z-A", "Old", "New"))
$sortComboBox.SelectedIndex = 0
$sortComboBox.Add_SelectedIndexChanged({ refreshTreeViewFolder })


$textBoxPath_Tab2.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $gotoButton.PerformClick()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    } elseif ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $textBoxPath_Tab2.SelectAll()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    }
})
$textBoxPath_Tab2.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) { $_.Handled = $true }
})


# Panels for the right side
$panelRightMain =     gen $splitContainer2.Panel2  "Panel"       ""          0 0 0 0  'Dock=Fill'
$panelRightCtrls =    gen $splitContainer2.Panel2  "Panel"       ""          0 0 0 33 'Dock=Top'
$panelRightCtrls_T =  gen $panelRightCtrls         "Panel"       ""          0 0 0 20 'Dock=Fill'
$showMspCheckbox =    gen $panelRightCtrls_T       "CheckBox"    "Show MSP"  0 0 0 0  'Dock=Right'  'Autosize=$true' 'Checked=$false'
$showMstCheckbox =    gen $panelRightCtrls_T       "CheckBox"    "Show MST"  0 0 0 0  'Dock=Right'  'Autosize=$true' 'Checked=$false'
$panelRightCtrls_B =  gen $panelRightCtrls         "Panel"       ""          0 0 0 20 'Dock=Bottom'
$searchTextBox_Tab2 = gen $panelRightCtrls_B       "TextBox"     ""          0 0 0 0  'Dock=Fill'
$searchTextBoxLabel = gen $panelRightCtrls_B       "Label"       "Filter:"   0 0 0 0  'Dock=Left'   'Autosize=$true'
$listView_Explore =   gen $panelRightMain          "ListView"    ""          0 0 0 0  'Dock=Fill'   'View=Details'   'FullRowSelect=$true' 'GridLines=$true'  'AllowColumnReorder=$true'  'HideSelection=$false' 
$sep3 =               gen $splitContainer2.Panel2  "Panel"       ""          0 0 1 0  'Dock=Left'   'BackColor=Gray' 
$progressPanel_Tab2 =      gen $splitContainer2.Panel2  "Panel"       ""          0 0 0 20 'Dock=Bottom'

$Tab2_statusStrip  = gen $progressPanel_Tab2 "StatusStrip" "" 0 0 0 0 'Dock=Bottom' 'SizingGrip=$false'
$Tab2_statusLabel  = gen $Tab2_statusStrip "ToolStripStatusLabel" "0 items" 0 0 0 0 'Font=Consolas, 8'
$Tab2_progressBar  = gen $Tab2_statusStrip "ToolStripProgressBar" "" 0 0 0 0 'AutoSize=$false'
$Tab2_progressBar.Control.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$Tab2_stopButton   = gen $Tab2_statusStrip "ToolStripButton" "STOP" 0 0 0 0 'Enabled=$false' 'BackColor=LightGray'
$Tab2_statusStrip.add_SizeChanged({ Update-ProgressBarWidth $Tab2_statusStrip $Tab2_statusLabel $Tab2_stopButton $Tab2_progressBar })
$sep6              = gen $progressPanel_Tab2 "Panel" "" 0 0 1 0 'Dock=Left' 'BackColor=Gray'

$allListViewItemsExplore = New-Object System.Collections.ArrayList

$listView_Explore.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $this.BeginUpdate()
        foreach ($item in $this.Items) { $item.Selected = $true }
        $this.EndUpdate()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    }
})

$columns_listView_Explore = @("File Name", "Product Name", "GUID", "Version", "Path", "Weight", "Modified")
foreach ($col in $columns_listView_Explore) {
    $columnHeader      = New-Object System.Windows.Forms.ColumnHeader
    $columnHeader.Text = $col
    [void]$listView_Explore.Columns.Add($columnHeader)
}

$searchTextBox_Tab2.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $searchTextBox_Tab2.SelectAll()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    }
})
$searchTextBox_Tab2.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) { $_.Handled = $true }
})


$launch_progressBar.Value = 35


# Event handler for custom drawing of nodes
$treeView.HideSelection = $false
$treeView.DrawMode = [System.Windows.Forms.TreeViewDrawMode]::OwnerDrawText  # Enable custom drawing
$treeView.Add_DrawNode({
    param($s,$e)
    if($e.Node.IsSelected){
        $bounds=$e.Bounds
        $e.Graphics.FillRectangle($highlightBrush,$bounds)
        $e.Graphics.DrawString($e.Node.Text,$treeView.Font,$highlightTextBrush,$bounds.X,$bounds.Y+1)
        $e.DrawDefault=$false
    }else{ $e.DrawDefault=$true }
})

function FilterListViewItems {
    param(
        [System.Windows.Forms.ListView]$listView,     [string]$Mode, [bool]$showMsp = $false, [bool]$showMst = $false,
        [System.Collections.ArrayList]$registryPaths, [System.Collections.ArrayList]$allListViewItems
    )
    $visibleItems = @()
    switch ($Mode) {
        "Explore" {
            $listview_NameIndex = -1
            for ($i = 0; $i -lt $listView.Columns.Count; $i++) { if ($listView.Columns[$i].Text -in @("File Name", "DisplayName")) { $listview_NameIndex = $i ; break } }
            if ($listview_NameIndex -eq -1) {
                [System.Windows.Forms.MessageBox]::Show("Column 'File Name' or 'DisplayName' not found in the ListView", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            } 
            foreach ($item in $allListViewItems) {
                $listview_Name = $item.SubItems[$listview_NameIndex].Text.ToLower()
                $isMsp         = $listview_Name.EndsWith(".msp", [System.StringComparison]::OrdinalIgnoreCase)
                $isMst         = $listview_Name.EndsWith(".mst", [System.StringComparison]::OrdinalIgnoreCase)
                if (($showMsp -or -not $isMsp) -and ($showMst -or -not $isMst)) { $visibleItems += $item.Clone() }
            }
        }
        "Registry" {
            $pathColumnIndex = -1
            for ($i = 0; $i -lt $listView.Columns.Count; $i++) { if ($listView.Columns[$i].Text -eq "Registry Path") { $pathColumnIndex = $i ; break } }
            if ($pathColumnIndex -eq -1) {
                [System.Windows.Forms.MessageBox]::Show("'Registry Path' Column not found", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            foreach ($item in $allListViewItems) {
                $itemPath = $item.SubItems[$pathColumnIndex].Text
                $matchFound = $false
                foreach ($rp in $registryPaths) { if ($itemPath -like "$rp*") { $matchFound = $true ; break } }
                if ($matchFound) { $visibleItems += $item.Clone() }
            }
        }
        default {
            [System.Windows.Forms.MessageBox]::Show("Unknown Mode. Code using 'Explore' or 'Registry'", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
    $listView.Items.Clear()
    foreach ($item in $visibleItems) { [void]$listView.Items.Add($item) }
    $Tab2_statusLabel.Text = "$($listView_Explore.Items.Count) items"
    Update-ProgressBarWidth $Tab2_statusStrip $Tab2_statusLabel $Tab2_stopButton $Tab2_progressBar
}

$showMspCheckbox.Add_CheckedChanged({ FilterListViewItems -Mode Explore -listView $listView_Explore -showMsp $showMspCheckbox.Checked -showMst $showMstCheckbox.Checked -allListViewItems $allListViewItemsExplore })
$showMstCheckbox.Add_CheckedChanged({ FilterListViewItems -Mode Explore -listView $listView_Explore -showMsp $showMspCheckbox.Checked -showMst $showMstCheckbox.Checked -allListViewItems $allListViewItemsExplore })

function AdjustListViewColumns {
    param([System.Windows.Forms.ListView]$listView)
    $ScanCheckedMSIBtn.Width = ($splitContainer2.Panel1.Width - $recursionLabel.Width - $recursionComboBox.Width) / 2
    $ScanSelectedMSIBtn.Width     = ($splitContainer2.Panel1.Width - $recursionLabel.Width - $recursionComboBox.Width) / 2
    $totalWidth             = $listView.ClientSize.Width
    $columnsInfo            = @()
    for ($i = 0; $i -lt $listView.Columns.Count; $i++) {
        $maxWidth    = 0
        $headerText  = $listView.Columns[$i].Text
        $headerWidth = [System.Windows.Forms.TextRenderer]::MeasureText($headerText, $listView.Font).Width + 8
        foreach ($item in $listView.Items) {
            if ($item.SubItems.Count -gt $i) {
                $textWidth = [System.Windows.Forms.TextRenderer]::MeasureText($item.SubItems[$i].Text, $listView.Font).Width
                if ($textWidth -gt $maxWidth) { $maxWidth = $textWidth }
            }
        }
        $maxWidth   = [Math]::Max($maxWidth + 8, $headerWidth)
        $fixedWidth = switch ($headerText) { 
            "GUID"                    { 250 } 
            "Modified"                { 80 } 
            "InstallDate"             { 65 } 
            "Version"                 { 75 } 
            "DisplayVersion"          { 75 } 
            $([string][char]0x2198)   { 20 }
            default                   { $null } 
        }
        $columnsInfo += [PSCustomObject]@{
            Index      = $i
            MaxWidth   = $maxWidth
            MinWidth   = switch ($headerText) { "File Name" { 180 } ; "DisplayName" { 180 } ; default { 0 } }
            FixedWidth = $fixedWidth  ;  Width = 0  ;  IsMaxed = $false
        }
    }
    do {
        $remaining      = $columnsInfo.Where({ -not $_.IsMaxed -and -not $_.FixedWidth })
        $maxedWidth     = ($columnsInfo.Where({ $_.IsMaxed })    | Measure-Object -Property Width -Sum).Sum
        $fixedWidth     = ($columnsInfo.Where({ $_.FixedWidth }) | Measure-Object -Property FixedWidth -Sum).Sum
        $remainingWidth = $totalWidth - $maxedWidth - $fixedWidth  ;  $widthChanged = $false
        foreach ($col in $remaining) {
            $newWidth = [int]($remainingWidth / $remaining.Count) - 4
            $newWidth = [Math]::Max($newWidth, $col.MinWidth)  # Apply minimal width
            if ($newWidth -ge $col.MaxWidth) {
                $col.Width    = $col.MaxWidth
                $col.IsMaxed  = $true
                $widthChanged = $true
            } else {
                $col.Width = $newWidth
            }
        }
    } while ($widthChanged)
    foreach ($col in $columnsInfo) { if ($col.FixedWidth) { $listView.Columns[$col.Index].Width = $col.FixedWidth } else { $listView.Columns[$col.Index].Width = $col.Width } }
}


$launch_progressBar.Value = 40


# Initialize the TreeView
$rootNode      = New-Object System.Windows.Forms.TreeNode
$rootNode.Text = "This Device"
$rootNode.Tag  = "This Device"
$treeView.Nodes.Add($rootNode) | Out-Null

# Add "Fast Access" root node
$fastAccessNode      = New-Object System.Windows.Forms.TreeNode
$fastAccessNode.Text = "Fast Access"
$fastAccessNode.Tag  = "Fast Access"
$treeView.Nodes.Add($fastAccessNode) | Out-Null

# Populate "Fast Access" node with Quick Access folders
$shell = New-Object -ComObject Shell.Application
$quickAccess = $shell.Namespace('shell:::{679f85CB-0220-4080-B29B-5540CC05AAB6}')
$quickAccessItems = $quickAccess.Items()


foreach ($item in $quickAccessItems) {
    if ($item.IsFolder) {
        $node      = New-Object System.Windows.Forms.TreeNode
        $node.Text = $item.Name  ;  $node.Tag = $item.Path
        $node.Nodes.Add([System.Windows.Forms.TreeNode]::new()) | Out-Null  # Add a dummy child node to enable expansion
        $fastAccessNode.Nodes.Add($node) | Out-Null
    }
}

$drives = Get-PSDrive -PSProvider 'FileSystem'
foreach ($drive in $drives) {
    $driveNode      = New-Object System.Windows.Forms.TreeNode
    $driveNode.Text = $drive.Name + " (" + $drive.Root + ")"  ;  $driveNode.Tag = $drive.Root
    $driveNode.Nodes.Add([System.Windows.Forms.TreeNode]::new()) | Out-Null
    $rootNode.Nodes.Add($driveNode) | Out-Null
}
$rootNode.Expand()


$launch_progressBar.Value = 45


function SortDirectories {
    param ([System.IO.DirectoryInfo[]]$dirs, [string]$sortOption)
    if (-not $dirs -or $dirs.Count -eq 0) { return $dirs }
    switch ($sortOption) {
        "A-Z" { [Array]::Sort($dirs, [Comparison[System.IO.DirectoryInfo]]{ param($a, $b); [string]::Compare($a.Name, $b.Name, [StringComparison]::OrdinalIgnoreCase) }) }
        "Z-A" { [Array]::Sort($dirs, [Comparison[System.IO.DirectoryInfo]]{ param($a, $b); [string]::Compare($b.Name, $a.Name, [StringComparison]::OrdinalIgnoreCase) }) }
        "Old" { [Array]::Sort($dirs, [Comparison[System.IO.DirectoryInfo]]{ param($a, $b); [DateTime]::Compare($a.LastWriteTime, $b.LastWriteTime) }) }
        "New" { [Array]::Sort($dirs, [Comparison[System.IO.DirectoryInfo]]{ param($a, $b); [DateTime]::Compare($b.LastWriteTime, $a.LastWriteTime) }) }
    }
    return $dirs
}

function PopulateTree {
    param ([System.Windows.Forms.TreeNode]$parentNode, [string]$path)
    try {
        # Handle network shares
        if ($path.StartsWith("\\")) {
            $start = $false
            cmd /c "net view $path" 2>&1 | ForEach-Object {
                if (!$start) { $start = $_ -match "^-{5,}"  ;  return }
                if ($_ -match "^(.+?)\s{2,}") {
                    $shareName      = $matches[1].Trim()
                    $shareNode      = New-Object System.Windows.Forms.TreeNode
                    $shareNode.Text = $shareName  ;  $shareNode.Tag = Join-Path $parentNode.Tag $shareName
                    $shareNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())  # Add empty child to allow expansion
                    $parentNode.Nodes.Add($shareNode)
                }
            }
        }
        # Get directories
        $dirInfo         = [System.IO.DirectoryInfo]::new($path)
        $dirs            = $dirInfo.GetDirectories()
        $excludedFolders = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $null            = $excludedFolders.Add('$RECYCLE.BIN')               ;  $null = $excludedFolders.Add('$WinREAgent')
        $null            = $excludedFolders.Add('System Volume Information')  ;  $null = $excludedFolders.Add('Recovery')
        $filteredDirs    = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()
        foreach            ($dir in $dirs) { if (-not $excludedFolders.Contains($dir.Name)) { $filteredDirs.Add($dir) } }
        $dirs            = $filteredDirs.ToArray()
        $dirs            = SortDirectories -dirs $dirs -sortOption $sortComboBox.Text
        $currentUserSid  = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        $writeRight      = [System.Security.AccessControl.FileSystemRights]::Write
        # Batch add with BeginUpdate
        $treeView = $parentNode.TreeView
        if ($treeView) { $treeView.BeginUpdate() }
        try {
            $nodeBatch    = [System.Collections.Generic.List[System.Windows.Forms.TreeNode]]::new()
            $batchCounter = 0
            foreach ($dir in $dirs) {
                $node = [System.Windows.Forms.TreeNode]::new()
                $node.Text     = $dir.Name  ;  $node.Tag = $dir.FullName
                $hasReadAccess = $true      ;  $hasWriteAccess = $true  ;  $hasSubDirectories = $false
                # Check for subdirectories
                try {
                    $enumerator = [System.IO.Directory]::EnumerateDirectories($dir.FullName).GetEnumerator()
                    if ($enumerator.MoveNext()) { $hasSubDirectories = $true }
                    $enumerator.Dispose()
                } 
                catch { $hasReadAccess = $false }
                if ($hasReadAccess) {
                    # Check write access
                    try {
                        $writeAllow  = $false  ;  $writeDeny = $false
                        $acl         = [System.IO.Directory]::GetAccessControl($dir.FullName)
                        $accessRules = $acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])
                        foreach ($rule in $accessRules) {
                            if ($rule.IdentityReference -eq $currentUserSid) {
                                $hasWriteRight = ($rule.FileSystemRights -band $writeRight) -ne 0
                                if ($hasWriteRight) {
                                    if     ($rule.AccessControlType -eq 'Allow') { $writeAllow = $true }
                                    elseif ($rule.AccessControlType -eq 'Deny')  { $writeDeny  = $true }
                                }
                            }
                        }
                        if ($writeDeny -and -not $writeAllow) { $hasWriteAccess = $false }
                    } 
                    catch { $hasWriteAccess = $false }
                    if (-not $hasWriteAccess) { $node.ForeColor = [System.Drawing.Color]::Orange }
                    if ($hasSubDirectories)   { $node.Nodes.Add([System.Windows.Forms.TreeNode]::new()) }
                } 
                else { $node.ForeColor = [System.Drawing.Color]::Red }
                $nodeBatch.Add($node)
                $batchCounter++
                # Add in batches of 10
                if ($batchCounter -ge 10) {
                    $parentNode.Nodes.AddRange($nodeBatch.ToArray())
                    $nodeBatch.Clear()
                    $batchCounter = 0
                }
            }
            # Add remaining nodes
            if ($nodeBatch.Count -gt 0) { $parentNode.Nodes.AddRange($nodeBatch.ToArray()) }
        } 
        finally { if ($treeView) { $treeView.EndUpdate() } }
    } catch { Write-Warning "Error accessing path $path : $_" }
}

function Expand-TreeViewPath {
    param([System.Windows.Forms.TreeView]$treeView, [string]$path)
    # Initial path validation
    $path = $path.TrimEnd('\')
    if (-not [System.IO.Directory]::Exists($path)) {
        # Fallback check for network paths that Directory.Exists might miss
        if ($path -match "^\\\\") {
            try {
                $testPath = $path.Split([char]'\') | Where-Object { $_ } | Select-Object -First 2
                $serverTest = "\\$($testPath -join '\')"
                if (-not [System.IO.Directory]::Exists($serverTest)) {
                    Show-NonBlockingMessage -message "Path not found" -title "Error" -timeout 2
                    return
                }
            } catch {
                Show-NonBlockingMessage -message "Path not found" -title "Error" -timeout 2
                return
            }
        } else {
            Show-NonBlockingMessage -message "Path not found" -title "Error" -timeout 2
            return
        }
    }
    # Handle network paths
    if ($path -match "^\\\\([^\\]+)\\?(.*)") {
        $serverName = $matches[1]
        $specificPath = $matches[2]
        # Find or create "Network" node
        $networkRootNode = $null
        foreach ($node in $treeView.Nodes) { if ($node.Text -eq "Network") { $networkRootNode = $node; break } }
        if (-not $networkRootNode) {
            $networkRootNode = [System.Windows.Forms.TreeNode]::new()
            $networkRootNode.Text = "Network"  ;  $networkRootNode.Tag = "Network"
            $treeView.Nodes.Add($networkRootNode)
        }
        # Find or create server node
        $serverNode = $null
        foreach ($node in $networkRootNode.Nodes) { if ($node.Text -eq $serverName) { $serverNode = $node; break } }
        if (-not $serverNode) {
            $serverNode = [System.Windows.Forms.TreeNode]::new()
            $serverNode.Text = $serverName  ;  $serverNode.Tag = "\\$serverName"
            $networkRootNode.Nodes.Add($serverNode)
        }
        # Load server shares if not already loaded
        if ($serverNode.Nodes.Count -eq 0) {
            $serverNode.Nodes.Clear()
            $shares = @()
            $start = $false
            $netViewOutput = cmd /c "net view \\$serverName" 2>&1
            foreach ($line in $netViewOutput) {
                if (-not $start)                 { if ($line -match "^-{5,}") { $start = $true }  ;  continue  }
                if ($line -match "^(.+?)\s{2,}") { $shares += $matches[1].Trim() }
            }
            # Sort shares
            $sortOption = $sortComboBox.Text
            switch ($sortOption) {
                "A-Z" { [Array]::Sort($shares) }
                "Z-A" { [Array]::Sort($shares); [Array]::Reverse($shares) }
            }
            # Batch add nodes
            $treeView.BeginUpdate()
            try {
                $nodeBatch = [System.Collections.Generic.List[System.Windows.Forms.TreeNode]]::new()
                $batchCounter = 0
                foreach ($shareName in $shares) {
                    $childNode = [System.Windows.Forms.TreeNode]::new()
                    $childNode.Text = $shareName  ;  $childNode.Tag = "\\$serverName\$shareName"
                    $childNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
                    $nodeBatch.Add($childNode)
                    $batchCounter++
                    if ($batchCounter -ge 10) {
                        $serverNode.Nodes.AddRange($nodeBatch.ToArray())
                        $nodeBatch.Clear()
                        $batchCounter = 0
                    }
                }
                if ($nodeBatch.Count -gt 0) { $serverNode.Nodes.AddRange($nodeBatch.ToArray()) }
            } 
            finally { $treeView.EndUpdate() }
        }
        $networkRootNode.Expand()
        # Navigate to specific path if provided
        if ($specificPath) {
            $segments = $specificPath.Split([char]'\') | Where-Object { $_ }
            $currentNode = $serverNode
            foreach ($segment in $segments) {
                # Find existing child node
                $childNode = $null
                foreach ($node in $currentNode.Nodes) { if ($node.Text -eq $segment) { $childNode = $node; break } }
                if (-not $childNode) {
                    $childNode = [System.Windows.Forms.TreeNode]::new()
                    $childNode.Text = $segment  ;  $childNode.Tag = [System.IO.Path]::Combine($currentNode.Tag, $segment)
                    $childNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
                    $currentNode.Nodes.Add($childNode)
                }
                $currentNode = $childNode
                # Load subdirectories if needed
                if ($currentNode.Nodes.Count -eq 1 -and -not $currentNode.Nodes[0].Tag) {
                    $currentNode.Nodes.Clear()
                    try {
                        $dirInfo = [System.IO.DirectoryInfo]::new($currentNode.Tag)
                        $dirs    = $dirInfo.GetDirectories()
                        $dirs    = SortDirectories -dirs $dirs -sortOption $sortComboBox.Text
                        $treeView.BeginUpdate()
                        try {
                            $nodeBatch = [System.Collections.Generic.List[System.Windows.Forms.TreeNode]]::new()
                            $batchCounter = 0
                            foreach ($dir in $dirs) {
                                $subChildNode = [System.Windows.Forms.TreeNode]::new()
                                $subChildNode.Text = $dir.Name  ;  $subChildNode.Tag = $dir.FullName
                                $subChildNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
                                $nodeBatch.Add($subChildNode)
                                $batchCounter++
                                if ($batchCounter -ge 10) {
                                    $currentNode.Nodes.AddRange($nodeBatch.ToArray())
                                    $nodeBatch.Clear()
                                    $batchCounter = 0
                                }
                            }
                            if ($nodeBatch.Count -gt 0) { $currentNode.Nodes.AddRange($nodeBatch.ToArray()) }
                        } 
                        finally { $treeView.EndUpdate() }
                    } 
                    catch { Write-Warning "Unable to access directories under $($currentNode.Tag)" }
                }
            }
            $treeView.SelectedNode = $currentNode    ;  $currentNode.EnsureVisible()
        } 
        else { $treeView.SelectedNode = $serverNode  ;  $serverNode.EnsureVisible() }
        return
    }
    # Handle local paths
    $path        = [System.IO.Path]::GetFullPath($path).TrimEnd('\')
    $segments    = $path.Split([char]'\') | Where-Object { $_ }
    $driveRoot   = "$($segments[0][0]):\"
    $currentNode = $null
    foreach ($node in $treeView.Nodes[0].Nodes) { if ($node.Tag -eq $driveRoot) { $currentNode = $node; break } }
    if (-not $currentNode) {
        [System.Windows.Forms.MessageBox]::Show("Drive not found in treeview.", "Error", 
                                                [System.Windows.Forms.MessageBoxButtons]::OK, 
                                                [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # Build hierarchy for remaining segments
    for ($i = 1; $i -lt $segments.Count; $i++) {
        $segment = $segments[$i]
        $found = $false
        # Load subdirectories if node hasn't been expanded
        if ($currentNode.Nodes.Count -eq 1 -and -not $currentNode.Nodes[0].Tag) {
            $currentNode.Nodes.Clear()
            try {
                $dirInfo = [System.IO.DirectoryInfo]::new($currentNode.Tag)
                $dirs = $dirInfo.GetDirectories()
                $dirs = SortDirectories -dirs $dirs -sortOption $sortComboBox.Text
                $treeView.BeginUpdate()
                try {
                    $nodeBatch = [System.Collections.Generic.List[System.Windows.Forms.TreeNode]]::new()
                    $batchCounter = 0
                    foreach ($dir in $dirs) {
                        $childNode = [System.Windows.Forms.TreeNode]::new()
                        $childNode.Text = $dir.Name  ;  $childNode.Tag = $dir.FullName
                        $childNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
                        $nodeBatch.Add($childNode)
                        $batchCounter++
                        if ($batchCounter -ge 10) {
                            $currentNode.Nodes.AddRange($nodeBatch.ToArray())
                            $nodeBatch.Clear()
                            $batchCounter = 0
                        }
                    }
                    if ($nodeBatch.Count -gt 0) { $currentNode.Nodes.AddRange($nodeBatch.ToArray()) }
                } finally {
                    $treeView.EndUpdate()
                }
            } catch {
                Write-Warning "Directory inaccessible : $($currentNode.Tag)"
            }
        }
        # Search for segment in loaded children
        foreach ($childNode in $currentNode.Nodes) {
            if ($childNode.Text -eq $segment) { $currentNode = $childNode  ;  $found = $true  ;  break }
        }
        # Create node if not found
        if (-not $found) {
            $newNode = [System.Windows.Forms.TreeNode]::new()
            $newNode.Text = $segment  ;  $newNode.Tag = [System.IO.Path]::Combine($currentNode.Tag, $segment)
            $newNode.Nodes.Add([System.Windows.Forms.TreeNode]::new())
            $currentNode.Nodes.Add($newNode)
            $currentNode = $newNode
        }
    }
    $treeView.SelectedNode = $currentNode
    $currentNode.EnsureVisible()
}


$launch_progressBar.Value = 50


function OpenTreeViewSelectedFolder {
    $selectedNode = $treeView.SelectedNode
    if ($null -ne $selectedNode) {
        if ($selectedNode.Text -eq "This Device" -or $selectedNode.Text -eq "Fast Access") {
            [System.Windows.Forms.MessageBox]::Show("Cannot open special node: $($selectedNode.Text).", "Error", 
                                                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                                                    [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        if ($selectedNode.Tag -is [string] -and $selectedNode.Tag.StartsWith("\\")) {
            Start-Process -FilePath "explorer.exe" -ArgumentList "`"$($selectedNode.Tag)`""
            return
        }
        if ($selectedNode.Tag -is [string] -and ([System.IO.Directory]::Exists($selectedNode.Tag))) {
            Start-Process -FilePath "explorer.exe" -ArgumentList "`"$($selectedNode.Tag)`""
        } else {
            [System.Windows.Forms.MessageBox]::Show("Path does not exist: $($selectedNode.Tag)", "Error", 
                                                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                                                    [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a node.", "Error", 
                                                [System.Windows.Forms.MessageBoxButtons]::OK, 
                                                [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


$openBtn.Add_Click({ param($s, $e) ; OpenTreeViewSelectedFolder })


function refreshTreeViewFolder {
    $selectedNode = $treeView.SelectedNode
    if ($null -ne $selectedNode) {
        $scrollPosVert = [NativeMethods]::GetScrollPos($treeView.Handle, [NativeMethods]::SB_VERT)
        try {
            if (-not [string]::IsNullOrEmpty($selectedNode.Tag) -and $null -ne $selectedNode.Parent) {
                $selectedNode.Nodes.Clear()
                PopulateTree -parentNode $selectedNode -path $selectedNode.Tag
                $selectedNode.Expand()
            } 
            else { [System.Windows.Forms.MessageBox]::Show("Cannot refresh because node is invalid.") }
        } 
        catch { [System.Windows.Forms.MessageBox]::Show("Error refreshing tree node: $_") } 
        finally {
            [NativeMethods]::SetScrollPos($treeView.Handle, [NativeMethods]::SB_VERT, $scrollPosVert, $true)
            [NativeMethods]::SendMessage($treeView.Handle, [NativeMethods]::WM_VSCROLL, ([NativeMethods]::SB_THUMBPOSITION -bor ($scrollPosVert -shl 16)), 0)
        }
    } else { [System.Windows.Forms.MessageBox]::Show("Please select a node to refresh.") }
}


$refreshBtn.Add_Click({ refreshTreeViewFolder })


$launch_progressBar.Value = 55


function Get-FilesRecursive {
    param(
        [string[]]$paths,      [int]$depth,     $progressBar,          [ref]$allItems,       [ref]$progressCounter,
        [System.Diagnostics.Stopwatch]$stopwatch, [regex]$compiledRegex, [switch]$foldersOnly
    )
    # Initialize on first call only
    if (-not $allItems)        { $allItems = [ref]([System.Collections.Generic.List[object]]::new()) }
    if (-not $progressCounter) { $progressCounter = [ref]0 }
    if (-not $stopwatch)       { $stopwatch = [System.Diagnostics.Stopwatch]::StartNew() }
    if (-not $compiledRegex)   {
        $compiledRegex = [regex]::new('\.(msi|msix|mst|msp)$', 
                         [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled)
    }
    $excludedRootDirs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    [void]$excludedRootDirs.Add('$RECYCLE.BIN')
    [void]$excludedRootDirs.Add('$WinREAgent')
    [void]$excludedRootDirs.Add('System Volume Information')
    [void]$excludedRootDirs.Add('Recovery')
    if ($depth -eq 0) { return $allItems.Value }
    foreach ($path in $paths) {
        if ($script:stopRequested)                                         { break }
        if (-not [System.IO.Directory]::Exists($path))                     { continue }
        try {
            if ($foldersOnly) {
                $dirs = [System.IO.Directory]::GetDirectories($path)
                foreach ($dirPath in $dirs) {
                    if ($script:stopRequested)                             { break }
                    $dirName   = [System.IO.Path]::GetFileName($dirPath)
                    $parentDir = [System.IO.Path]::GetDirectoryName($dirPath)
                    $isRoot    = ($parentDir -match '^[A-Za-z]:\\?$') -or ($parentDir -match '^\\\\[^\\]+\\[^\\]+\\?$')
                    if ($isRoot -and $excludedRootDirs.Contains($dirName)) { continue }
                    $allItems.Value.Add([System.IO.DirectoryInfo]::new($dirPath))
                    $progressCounter.Value++
                    if ($progressCounter.Value % 50 -eq 0) {
                        $Tab2_statusLabel.Text = "Scanning... $($progressCounter.Value) folders"
                        Update-ProgressBarWidth $Tab2_statusStrip $Tab2_statusLabel $Tab2_stopButton $progressBar
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    Get-FilesRecursive $dirPath ($depth - 1) $progressBar $allItems $progressCounter $stopwatch $compiledRegex -foldersOnly 
                }
            } else {
                # Process files in current directory
                $allFiles = [System.IO.Directory]::GetFiles($path)
                foreach ($filePath in $allFiles) {
                    if ($script:stopRequested)                             { break }
                    if ($compiledRegex.IsMatch($filePath)) { $allItems.Value.Add([System.IO.FileInfo]::new($filePath)) }
                }
                # Process subdirectories
                $dirs = [System.IO.Directory]::GetDirectories($path)
                foreach ($dirPath in $dirs) {
                    if ($script:stopRequested)                             { break }
                    $dirName   = [System.IO.Path]::GetFileName($dirPath)
                    $parentDir = [System.IO.Path]::GetDirectoryName($dirPath)
                    $isRoot    = ($parentDir -match '^[A-Za-z]:\\?$') -or ($parentDir -match '^\\\\[^\\]+\\[^\\]+\\?$')
                    if ($isRoot -and $excludedRootDirs.Contains($dirName)) { continue }
                    $progressCounter.Value++
                    if ($progressCounter.Value % 50 -eq 0) {
                        $Tab2_statusLabel.Text = "Scanning... $($progressCounter.Value) folders"
                        Update-ProgressBarWidth $Tab2_statusStrip $Tab2_statusLabel $Tab2_stopButton $progressBar
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    Get-FilesRecursive $dirPath ($depth - 1) $progressBar $allItems $progressCounter $stopwatch $compiledRegex
                }
            }
        }
        catch { Write-Warning "Error accessing path : $path - $_" }
    }
    return $allItems.Value
}


# Convert ComboBox selection to depth
function Get-RecursionDepth { switch ($recursionComboBox.SelectedItem) { "No" { return 1 } ; "All" { return [int]::MaxValue } ; default { return [int]$recursionComboBox.SelectedItem + 1 } } }

$treeView.Add_BeforeExpand({ 
    $node = $_.Node
    if ($node.Nodes.Count -eq 1 -and -not $node.Nodes[0].Tag) {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        $node.Nodes.Clear()
        PopulateTree $node $node.Tag 
        $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor 
    }
})

function Get-CheckedNodes {
    param([System.Windows.Forms.TreeNodeCollection]$nodes)
    foreach ($node in $nodes) {
        if ($node.Checked) { $node }
        if ($node.Nodes.Count -gt 0) { Get-CheckedNodes -nodes $node.Nodes }
    }
}


$launch_progressBar.Value = 60


function Complete-Listview {
    param([bool]$multiSearch = $false)
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $script:stopRequested = $false
    $listView_Explore.Items.Clear()
    $allListViewItemsExplore.Clear()
    $recursionDepth = Get-RecursionDepth
    $finalPathsToSearch = [System.Collections.Generic.List[string]]::new()  # Initialize the list to store paths
    $allItems = [System.Collections.Generic.List[Object]]::new()  # Ensure $allItems is initialized before any calls
    $allItemsRef = [ref]$allItems
    if (-not $multiSearch) {  # Single search mode
        $selectedNode = $treeView.SelectedNode
        if ($null -eq $selectedNode) {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Select a node before search.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $textBoxPath_Tab2.Text = $selectedNode.Tag
        if (-not [string]::IsNullOrEmpty($selectedNode.Tag)) { $finalPathsToSearch.Add($selectedNode.Tag) } else { Write-Warning "Selected node has no valid tag: $($selectedNode.Text)" }
    } else {  # Multi-search mode
        $checkedNodes = Get-CheckedNodes -nodes $treeView.Nodes
        if (-not $checkedNodes) {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Check some nodes before search", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        foreach ($node in $checkedNodes) { if (-not [string]::IsNullOrEmpty($node.Tag)) { $finalPathsToSearch.Add($node.Tag) } else { Write-Warning "Node has no valid tag: $($node.Text)" } }
    }
    if ($finalPathsToSearch.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No valid paths to search", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }
    $Tab2_stopButton.Enabled = $true
    [System.Windows.Forms.Application]::DoEvents()
    $pathsArray = $finalPathsToSearch.ToArray()
    $items      = Get-FilesRecursive -paths $pathsArray -depth $recursionDepth -allItems $allItemsRef -progressBar $Tab2_progressBar
    $allFiles   = $allItems | Where-Object { -not $_.PSIsContainer }
    if ($allFiles.Count -gt 0) {
        $totalFiles  = $allFiles.Count
        $currentFile = 0
        $stopwatch   = [System.Diagnostics.Stopwatch]::StartNew()  # Global stopwatch for total elapsed time
        $lastUpdate  = [System.Diagnostics.Stopwatch]::StartNew()  # Timer for 100ms updates
        foreach ($file in $allFiles) {
            if ($script:stopRequested) { break }
            $currentFile++
            $fullPath    = $file.FullName
            $sizeBytes   = $file.Length
            $sizeMB      = $sizeMB = Format-FileSize -Bytes $sizeBytes
            $modified    = $file.LastWriteTime.ToString()
            $msiInfo     = Get-MsiInfo -FilePath $file.FullName
            $dataToAdd   = @{ "File Name"=$file.Name ; "Product Name"=$msiInfo["ProductName"] ; "GUID"=$msiInfo["ProductCode"] ; "Version"=$msiInfo["ProductVersion"] ; "Path"=$fullPath ; "Weight"=$sizeMB ; "Modified"=$modified }
            $columnOrder = $listView_Explore.Columns | ForEach-Object { $_.Text }
            $subItems    = @()
            foreach ($columnName in $columnOrder) { if ($dataToAdd.ContainsKey($columnName)) { $subItems += [string]$dataToAdd[$columnName] } else { $subItems += "" } }
            $firstItemText = $subItems[0]  # Create ListView with first sub-item
            $item = New-Object System.Windows.Forms.ListViewItem ($firstItemText)
            for ($i = 1; $i -lt $subItems.Count; $i++) { $item.SubItems.Add($subItems[$i]) | Out-Null }  # Add remaining sub-elements
            $item.Tag = $file.FullName  # Full path in tag
            $listView_Explore.Items.Add($item)
            $allListViewItemsExplore.Add($item)
            if ($lastUpdate.ElapsedMilliseconds -ge 100) { # Update progress and status every 100ms
                $lastUpdate.Restart()
                $Tab2_progressBar.Value       = [Math]::Min(($currentFile / $totalFiles) * 100, 100)
                $elapsedTime             = $stopwatch.Elapsed.TotalSeconds
                $itemsProcessedPerSecond = if ($elapsedTime -gt 0) { [Math]::Max($currentFile / $elapsedTime, 0.01) } else { 1 }
                $remainingFiles          = $totalFiles - $currentFile
                $estimatedRemainingTime  = [Math]::Ceiling($remainingFiles / $itemsProcessedPerSecond)
                $timeRemainingText       =  if ($estimatedRemainingTime -gt 0) { "$([TimeSpan]::FromSeconds($estimatedRemainingTime).ToString('hh\:mm\:ss')) remaining" } 
                                            else { "Calculating..." }
                $Tab2_statusLabel.Text   = "$($listView_Explore.Items.Count) items - $timeRemainingText"
                Update-ProgressBarWidth $Tab2_statusStrip $Tab2_statusLabel $Tab2_stopButton $Tab2_progressBar
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    }
    FilterListViewItems -Mode Explore -listView $listView_Explore -showMsp $showMspCheckbox.Checked -showMst $showMstCheckbox.Checked -allListViewItems $allListViewItemsExplore
    AdjustListViewColumns -listView $listView_Explore
    $Tab2_progressBar.Value  = 0
    $Tab2_stopButton.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

$ScanCheckedMSIBtn.Add_Click({ Complete-Listview -multiSearch $true })
$ScanSelectedMSIBtn.Add_Click({ Complete-Listview -multiSearch $false })

$gotoButton.Add_Click({
    $path = $textBoxPath_Tab2.Text.Trim()
    if (-not (Test-RemotePathAccess -Paths @($path))) { return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    Expand-TreeViewPath -treeView $treeView -path $path
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

Add-ListViewSearchFilter -SearchTextBox $searchTextBox_Tab2 -ListView $listView_Explore -AllItems $allListViewItemsExplore -AdditionalFilter {
    param($item)
    $nameIndex = 0
    for ($i = 0; $i -lt $listView_Explore.Columns.Count; $i++) { if ($listView_Explore.Columns[$i].Text -eq "File Name") { $nameIndex = $i ; break } }
    $name  = $item.SubItems[$nameIndex].Text.ToLower()
    $isMsp = $name.EndsWith(".msp")
    $isMst = $name.EndsWith(".mst")
    return ($showMspCheckbox.Checked -or -not $isMsp) -and ($showMstCheckbox.Checked -or -not $isMst)
}

$Tab2_stopButton.Add_Click({ $script:stopRequested = $true })

$listView_Explore.Add_ColumnClick({ HandleColumnClick -listViewparam $listView_Explore -e $_ })

function ConfigureTreeViewContextMenu($treeView) {
    $treeView.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            $node = $s.HitTest($e.X, $e.Y).Node
            if ($node) { $s.SelectedNode = $node }
        }
    })
    $contextMenu = [System.Windows.Forms.ContextMenuStrip]::new()
    $contextMenu.Tag = $treeView
    $treeView.ContextMenuStrip = $contextMenu
    $addItem = { param($text, $action) 
        $i = [System.Windows.Forms.ToolStripMenuItem]::new($text)
        $i.Add_Click($action)
        [void]$contextMenu.Items.Add($i)
        return $i
    }
    [void](& $addItem "Copy Folder Path" {
        $tv = $this.GetCurrentParent().Tag
        $node = $tv.SelectedNode
        if ($node) {
            $tagVal = $node.Tag
            $txt = if ($tagVal.IndexOf(' ') -ge 0) { "`"$tagVal`"" } else { $tagVal }
            [System.Windows.Forms.Clipboard]::SetText($txt)
        }
    })
    [void](& $addItem "Open Folder" { OpenTreeViewSelectedFolder })
    [void](& $addItem "Open Parent Folder" {
        $tv   = $this.GetCurrentParent().Tag
        $node = $tv.SelectedNode
        if ($node -and ([System.IO.Directory]::Exists($node.Tag) -or [System.IO.File]::Exists($node.Tag))) {
            $parent = Split-Path -Path $node.Tag -Parent
            if ([System.IO.Directory]::Exists($parent)) { [System.Diagnostics.Process]::Start("explorer.exe", "/select,`"$parent`"") }
            else                                        { [System.Windows.Forms.MessageBox]::Show("Cannot open parent folder : Path does not exist.", "Error", 
                                                          [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
        }
    })
    [void](& $addItem "Refresh Folder" { refreshTreeViewFolder })
    [void](& $addItem "Scan Folder" {
        $tv = $this.GetCurrentParent().Tag
        if ($tv.SelectedNode -and ([System.IO.Directory]::Exists($tv.SelectedNode.Tag))) { Complete-Listview -multiSearch $false }
    })
    # ── Shell / Remote session items ──
    $sepShell = [System.Windows.Forms.ToolStripSeparator]::new()
    $sepShell.Tag = "ShellSep"
    [void]$contextMenu.Items.Add($sepShell)
    $miCmd = & $addItem "CMD here" {
        $tv = $this.GetCurrentParent().Tag
        if ($tv.SelectedNode -and ([System.IO.Directory]::Exists($tv.SelectedNode.Tag))) { [System.Diagnostics.Process]::Start("cmd.exe", "/k cd /d `"$($tv.SelectedNode.Tag)`"") }
        else { [System.Windows.Forms.MessageBox]::Show("Cannot open CMD here : Invalid path.", "Error", 
               [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    }
    $miCmd.Tag = "CmdHere"
    $miCmdAdmin = & $addItem "CMD Admin here" {
        $tv = $this.GetCurrentParent().Tag
        if ($tv.SelectedNode -and ([System.IO.Directory]::Exists($tv.SelectedNode.Tag))) { Start-Process -FilePath "cmd.exe" -ArgumentList "/k cd /d `"$($tv.SelectedNode.Tag)`"" -Verb RunAs }
        else { [System.Windows.Forms.MessageBox]::Show("Cannot open CMD Admin here : Invalid path.", "Error", 
               [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    }
    $miCmdAdmin.Tag = "CmdAdminHere"
    $miPsAdmin = & $addItem "PowerShell Admin here" {
        $tv = $this.GetCurrentParent().Tag
        if ($tv.SelectedNode -and ([System.IO.Directory]::Exists($tv.SelectedNode.Tag))) {
            $psi = [System.Diagnostics.ProcessStartInfo]::new()
            $psi.FileName        = "powershell.exe"
            $psi.Arguments       = "-NoExit -Command `"Set-Location -LiteralPath '$($tv.SelectedNode.Tag)'`""
            $psi.Verb            = "runas"
            $psi.UseShellExecute = $true
            [System.Diagnostics.Process]::Start($psi)
        }
        else { [System.Windows.Forms.MessageBox]::Show("Cannot open PowerShell here : Invalid path.", "Error", 
               [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    }
    $miPsAdmin.Tag = "PsAdminHere"
    $miPsSession = & $addItem "PSSession here" {
        $tv = $this.GetCurrentParent().Tag
        if ($tv.SelectedNode) {
            $nodePath = $tv.SelectedNode.Tag
            if ($nodePath -match '^\\\\([^\\]+)\\([A-Za-z])\$(.*)') {
                $server  = $matches[1]
                $dirPath = "$($matches[2]):$($matches[3])".TrimEnd('\')
                Start-RemotePSSession -Server $server -RemotePath $dirPath
            }
        }
    }
    $miPsSession.Tag = "PsSessionHere"
    # ── Visibility logic ──
    $contextMenu.Add_Opening({
        param($s, $e)
        $tv   = $this.Tag
        $node = $tv.SelectedNode
        if (-not $node) { $e.Cancel = $true; return }
        $path          = if ($node.Tag -is [string]) { $node.Tag } else { "" }
        $isNetwork     = $path.StartsWith("\\")
        $isAdminShare  = $path -match '^\\\\[^\\]+\\[A-Za-z]\$'
        $showCmd       = (-not $isNetwork) -and (-not $isAdmin)
        $showPsAdmin   = (-not $isNetwork) -or ($isNetwork -and -not $isAdminShare)
        $showPsSession = ($isNetwork -and $isAdminShare)
        $anyShellItem  = $showCmd -or $showPsAdmin -or $showPsSession
        foreach ($mi in $this.Items) {
            switch ($mi.Tag) {
                "CmdHere"      { $mi.Visible = $showCmd }
                "CmdAdminHere" { $mi.Visible = $showCmd }
                "PsAdminHere"  { $mi.Visible = $showPsAdmin }
                "PsSessionHere"{ $mi.Visible = $showPsSession }
                "ShellSep"     { $mi.Visible = $anyShellItem }
            }
        }
    }.GetNewClosure())
}
ConfigureTreeViewContextMenu -treeView $treeView


$launch_progressBar.Value = 70


#region Tab 3 : Registry

$panelMain = New-Object System.Windows.Forms.Panel
$panelMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabPage3.Controls.Add($panelMain)

$sep2_tab3                = gen $tabPage3               "Panel"      0 0 1 0    'Dock=Fill' 'BackColor=Gray'
$sep1_tab3                = gen $tabPage3               "Panel"      0 0 0 10   'Dock=Top'
$borderTop_tab3           = gen $tabPage3               "Panel"      0 0 0 1    'Dock=Top'  'BackColor=Gray' 
$listView_Registry        = gen $panelMain              "ListView"   0 0 0 0    'Dock=Fill' 'View=Details' 'FullRowSelect=$true' 'GridLines=$true' 'AllowColumnReorder=$true' 'HideSelection=$false'
$panelCtrls_Registry      = gen $panelMain              "Panel"      0 0 0 60   'Dock=Top'
        
$subPanelCtrls_SearchOpts           = gen $panelCtrls_Registry       "Panel"    0 0 250 40 'Dock=Left'
$subPanelCtrls_SearchOpts_line1     = gen $subPanelCtrls_SearchOpts  "Panel"    0 0 0 20   'Dock=Top'
$subPanelCtrls_SearchOpts_line2     = gen $subPanelCtrls_SearchOpts  "Panel"    0 0 0 20   'Dock=Bottom'
$checkbox_ShowInstallSource         = gen $subPanelCtrls_SearchOpts_line1 "CheckBox" "Show InstallSource"                0 0 140 0 'Dock=Left' 'Checked=$false'
$checkbox_QuietUninstallIfAvailable = gen $subPanelCtrls_SearchOpts_line2 "CheckBox" "QuietUninstallString if available" 0 0 200 0 'Dock=Left' 'Checked=$true'

$sep6_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"   0 0 25 0   'Dock=Left'
$sep5_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"   0 0 1 0    'Dock=Left' 'BackColor=Gray'
$sep4_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"   0 0 25 0   'Dock=Left'
$subPanelCtrls_HKCU       = gen $panelCtrls_Registry    "Panel"   0 0 450 40 'Dock=Left'
$subPanelCtrls_HKCU_line1 = gen $subPanelCtrls_HKCU     "Panel"   0 0 0 20   'Dock=Top'
$subPanelCtrls_HKCU_line2 = gen $subPanelCtrls_HKCU     "Panel"   0 0 0 20   'Dock=Bottom'

$sep3_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"   0 0 25 0   'Dock=Left'
$sep2_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"   0 0 1 0    'Dock=Left' 'BackColor=Gray'
$sep1_panelCtrls_Registry = gen $panelCtrls_Registry    "Panel"   0 0 25 0   'Dock=Left'
$subPanelCtrls_HKLM       = gen $panelCtrls_Registry    "Panel"   0 0 450 40 'Dock=Left'
$subPanelCtrls_HKLM_line1 = gen $subPanelCtrls_HKLM     "Panel"   0 0 0 20   'Dock=Top'
$subPanelCtrls_HKLM_line2 = gen $subPanelCtrls_HKLM     "Panel"   0 0 0 20   'Dock=Bottom'
$subPanelCtrls_Filter     = gen $panelCtrls_Registry    "Panel"   0 0 0 20   'Dock=Bottom'

$HKLM32_btn               = gen $subPanelCtrls_HKLM_line2  "Button"   "Show"      0 0 0 0 'Dock=Right' 'Autosize=$true'
$checkbox_HKLM32          = gen $subPanelCtrls_HKLM_line2  "CheckBox" "HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" 0 0 410 0 'Dock=Left' 'checked=$true'
$HKLM64_btn               = gen $subPanelCtrls_HKLM_line1  "Button"   "Show"      0 0 0 0 'Dock=Right'  'Autosize=$true'
$checkbox_HKLM64          = gen $subPanelCtrls_HKLM_line1  "CheckBox" "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall"             0 0 330 0 'Dock=Left' 'checked=$true'

$HKCU32_btn               = gen $subPanelCtrls_HKCU_line2  "Button"   "Show"      0 0 0 0 'Dock=Right' 'Autosize=$true'
$checkbox_HKCU32          = gen $subPanelCtrls_HKCU_line2  "CheckBox" "HKCU\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"  0 0 410 0 'Dock=Left' 'checked=$true'
$HKCU64_btn               = gen $subPanelCtrls_HKCU_line1  "Button"   "Show"      0 0 0 0 'Dock=Right'  'Autosize=$true'
$checkbox_HKCU64          = gen $subPanelCtrls_HKCU_line1  "CheckBox" "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall"              0 0 330 0 'Dock=Left' 'checked=$true'

$searchTextBox_Tab3       = gen $subPanelCtrls_Filter "TextBox"  ""                          0 0 0 20  'Dock=Fill' 

$searchButton_Registry     = gen $subPanelCtrls_Filter "Button"   "Search"                    0 0 55 0  'Dock=Right'
$checkbox_RestrictToFilter = gen $subPanelCtrls_Filter "CheckBox" "Restrict Search to:" 0 0 0 0 'Dock=Left' 'Checked=$true'  'Autosize=$true'

$Tab3_progressPanel             = gen $panelMain            "Panel"                            0 0 0 20  'Dock=Bottom'
$Tab3_statusStrip               = gen $Tab3_progressPanel   "StatusStrip"                      0 0 0 0   'Dock=Bottom' 'SizingGrip=$false'
$Tab3_statusLabel               = gen $Tab3_statusStrip     "ToolStripStatusLabel" "0 items"   0 0 0 0   'Font=Consolas, 8'
$Tab3_progressBar               = gen $Tab3_statusStrip     "ToolStripProgressBar"             0 0 0 0   'AutoSize=$false'
$Tab3_progressBar.Control.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous 
$Tab3_stopButton                = gen $Tab3_statusStrip     "ToolStripButton"      "STOP"      0 0 0 0   'Enabled=$false' 'BackColor=LightGray'
$sep7_Tab3                      = gen $Tab3_progressPanel   "Panel"                            0 0 1 0   'Dock=Left' 'BackColor=Gray'

foreach ($ctrl in @($checkbox_HKLM64, $checkbox_HKLM32, $checkbox_HKCU64, $checkbox_HKCU32,
                     $HKLM64_btn, $HKLM32_btn, $HKCU64_btn, $HKCU32_btn,
                     $checkbox_ShowInstallSource, $checkbox_QuietUninstallIfAvailable,
                     $checkbox_RestrictToFilter)) {
    $ctrl.TabStop = $false
}
$searchTextBox_Tab3.TabIndex   = 0
$searchButton_Registry.TabIndex = 1

$listView_Registry.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $this.BeginUpdate()
        foreach ($item in $this.Items) { $item.Selected = $true }
        $this.EndUpdate()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    }
})
$columnsRegistry = @($([string][char]0x2198), "DisplayName", "GUID", "UninstallString", "InstallDate", "DisplayVersion", "Registry Path")
$allListViewItems_Registry = New-Object System.Collections.ArrayList
foreach ($col in $columnsRegistry) {
    $columnHeader = New-Object System.Windows.Forms.ColumnHeader
    $columnHeader.Text = $col
    [void]$listView_Registry.Columns.Add($columnHeader)
}
foreach ($btn in @($HKLM32_btn, $HKLM64_btn, $HKCU32_btn, $HKCU64_btn)) {
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 1
}

$Tab3_statusStrip.add_SizeChanged({ Update-ProgressBarWidth $Tab3_statusStrip $Tab3_statusLabel $Tab3_stopButton $Tab3_progressBar })
$Tab3_stopButton.Add_Click({ $script:RegistryStopRequested = $true })

$searchTextBox_Tab3.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $searchButton_Registry.PerformClick()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    } 
    elseif ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $searchTextBox_Tab3.SelectAll()
        $_.Handled          = $true
        $_.SuppressKeyPress = $true
    }
})
$searchTextBox_Tab3.Add_KeyPress({ if ($_.KeyChar -eq [char]13) { $_.Handled = $true } })
$searchTextBox_Tab3.Add_TextChanged({
    if ($null -eq $script:RegistryCache -or $script:RegistryCache.Count -eq 0) { return }
    Update-RegistryListViewFilter
})

$checkbox_ShowInstallSource.Add_CheckedChanged({ Update-RegistryListViewFilter })
$checkbox_QuietUninstallIfAvailable.Add_CheckedChanged({ Update-RegistryListViewFilter })

$launch_progressBar.Value = 75

function Format-RegistryPathDisplay {
    param([string]$RegistryPath)
    if ([string]::IsNullOrEmpty($RegistryPath)) { return $null }
    $keyName     = $RegistryPath.Split('\')[-1]
    $isWow32     = $RegistryPath.IndexOf('Wow6432Node', 0, [StringComparison]::OrdinalIgnoreCase) -ge 0
    switch -Exact ($RegistryPath.Substring(0, [Math]::Min(5, $RegistryPath.Length))) {
        'HKLM:' { if ($isWow32) { return "Machine - x32 - $keyName" } else { return "Machine - x64 - $keyName" } }
        'HKCU:' { if ($isWow32) { return "User - x32 - $keyName" }    else { return "User - x64 - $keyName" } }
        'HKU: ' { if ($isWow32) { return "User - x32 - $keyName" }    else { return "User - x64 - $keyName" } }
        default {
            if ($RegistryPath.Length -ge 4 -and $RegistryPath.Substring(0, 4) -eq 'HKU:') {
                if ($isWow32)   { return "User - x32 - $keyName" }    else { return "User - x64 - $keyName" }
            }
        }
    }
    return $keyName
}

function Update-DeviceColumnVisibility {
    param([System.Windows.Forms.ListView]$ListView, [bool]$ShowDevice)
    $deviceColumnExists = $false
    $deviceColumnIndex = -1
    $sortColumnIndex = -1
    for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
        if ($ListView.Columns[$i].Text -eq "Device")              { $deviceColumnExists = $true ; $deviceColumnIndex = $i }
        if ($ListView.Columns[$i].Text -eq $([string][char]0x2198)) { $sortColumnIndex = $i }
    }
    $insertIndex = if ($sortColumnIndex -ge 0) { $sortColumnIndex + 1 } else { 0 }
    if ($ShowDevice -and -not $deviceColumnExists) {
        $deviceColumn = New-Object System.Windows.Forms.ColumnHeader
        $deviceColumn.Text = "Device"
        $ListView.Columns.Insert($insertIndex, $deviceColumn)
        Write-Log "Added Device column to ListView"
    }
    elseif (-not $ShowDevice -and $deviceColumnExists) {
        $ListView.Columns.RemoveAt($deviceColumnIndex)
        Write-Log "Removed Device column from ListView"
    }
}

function Update-InstallSourceColumnVisibility {
    param(
        [System.Windows.Forms.ListView]$ListView,
        [bool]$ShowInstallSource
    )
    $installSourceColumnExists = $false
    $installSourceColumnIndex = -1
    $guidColumnIndex = -1
    for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
        if ($ListView.Columns[$i].Text -eq "InstallSource") { $installSourceColumnExists = $true ; $installSourceColumnIndex = $i }
        if ($ListView.Columns[$i].Text -eq "GUID")          { $guidColumnIndex = $i }
    }
    $insertIndex = if ($guidColumnIndex -ge 0) { $guidColumnIndex + 1 } else { 2 }
    if ($ShowInstallSource -and -not $installSourceColumnExists) {
        $installSourceColumn = New-Object System.Windows.Forms.ColumnHeader
        $installSourceColumn.Text = "InstallSource"
        $ListView.Columns.Insert($insertIndex, $installSourceColumn)
        Write-Log "Added InstallSource column to ListView"
    }
    elseif (-not $ShowInstallSource -and $installSourceColumnExists) {
        $ListView.Columns.RemoveAt($installSourceColumnIndex)
        Write-Log "Removed InstallSource column from ListView"
    }
}

Function PopulateRegistryListView {
    param(
        [System.Windows.Forms.ListView]$listViewparam,
        [System.Collections.ArrayList]$registryPaths,
        [System.Collections.ArrayList]$allListViewItems,
        [string]$FilterText = "",
        [bool]$RestrictToFilter = $false,
        [bool]$UseQuietUninstall = $false,
        [bool]$ShowInstallSource = $false,
        [string[]]$TargetComputers = @(),
        [System.Management.Automation.PSCredential]$Credential = $null
    )
    Write-Log "PopulateRegistryListView called : Filter='$FilterText', RestrictToFilter=$RestrictToFilter, UseQuietUninstall=$UseQuietUninstall, ShowInstallSource=$ShowInstallSource, Targets=$($TargetComputers -join ',')"
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:RegistryStopRequested = $false
    $Tab3_stopButton.Enabled = $true
    [System.Windows.Forms.Application]::DoEvents()
    $listViewparam.Items.Clear()
    $allListViewItems.Clear()
    $computersToQuery  = if ($TargetComputers.Count -gt 0) { $TargetComputers } else { @("") }
    $allEntriesForSort = [System.Collections.ArrayList]::new()
    $deviceDisplayNames = @{}
    foreach ($computerName in $computersToQuery) {
        if ([string]::IsNullOrWhiteSpace($computerName)) { continue }
        $displayName = Get-ComputerDisplayName -Computer $computerName
        $deviceDisplayNames[$computerName] = $displayName
    }
    $collectionScriptBlock = {
        param($params)
        $registryPaths     = $params.RegistryPaths
        $filterText        = $params.FilterText
        $restrictToFilter  = $params.RestrictToFilter
        $deviceLabel       = $params.DeviceLabel
        $loggedOnUserSID   = $params.LoggedOnUserSID
        $isRemote          = $params.IsRemote
        $results = [System.Collections.ArrayList]::new()
        $hasTextFilter = $restrictToFilter -and -not [string]::IsNullOrWhiteSpace($filterText)
        $filterTerms = @()
        if ($hasTextFilter) {
            $filterTerms = @($filterText -split ';' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ -ne '' })
            $hasTextFilter = ($filterTerms.Count -gt 0)
        }
        foreach ($registryPath in $registryPaths) {
            $rootKey, $subPath = $registryPath -split ":", 2
            $subPath           = $subPath.TrimStart('\')
            $isHKU             = ($rootKey -eq "HKU")
            $baseKey           = $null
            $effectiveSubPath  = $subPath
            $storedPathPrefix  = $null
            if ($isHKU) {
                if ($isRemote) {
                    if ([string]::IsNullOrWhiteSpace($loggedOnUserSID)) { continue }
                    $baseKey = [Microsoft.Win32.RegistryHive]::Users
                    $effectiveSubPath = "$loggedOnUserSID\$subPath"
                    $storedPathPrefix = "HKU:$loggedOnUserSID\$subPath"
                }
                else {
                    $baseKey = [Microsoft.Win32.RegistryHive]::CurrentUser
                    $effectiveSubPath = $subPath
                    $storedPathPrefix = "HKCU:$subPath"
                }
            }
            else {
                $baseKey = switch ($rootKey) {
                    "HKLM" { [Microsoft.Win32.RegistryHive]::LocalMachine }
                    "HKCU" { [Microsoft.Win32.RegistryHive]::CurrentUser }
                    default { $null }
                }
                $storedPathPrefix = "$rootKey`:$subPath"
            }
            if ($null -eq $baseKey) { continue }
            try {
                $registryKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($baseKey, [Microsoft.Win32.RegistryView]::Default).OpenSubKey($effectiveSubPath)
                if ($null -eq $registryKey) { continue }
                foreach ($subKeyName in $registryKey.GetSubKeyNames()) {
                    $subKey = $registryKey.OpenSubKey($subKeyName)
                    if ($null -eq $subKey) { continue }
                    $displayName = $subKey.GetValue("DisplayName")
                    if ([string]::IsNullOrEmpty($displayName)) { 
                        $subKey.Close()
                        continue 
                    }
                    if ($hasTextFilter) {
                        $matchFound = $false
                        $tempUninstall = $subKey.GetValue("UninstallString")
                        foreach ($term in $filterTerms) {
                            if ($matchFound) { break }
                            if ($displayName.ToLower() -like "*$term*") { $matchFound = $true }
                            elseif ($tempUninstall -and $tempUninstall.ToLower() -like "*$term*") { $matchFound = $true }
                        }
                        if (-not $matchFound) { 
                            $subKey.Close()
                            continue 
                        }
                    }
                    $rawUninstallString   = if ($hasTextFilter) { $tempUninstall } else { $subKey.GetValue("UninstallString") }
                    $quietUninstallString = $subKey.GetValue("QuietUninstallString")
                    $installSource        = $subKey.GetValue("InstallSource")
                    $installDate          = $subKey.GetValue("InstallDate")
                    $displayVersion       = $subKey.GetValue("DisplayVersion")
                    $subKey.Close()
                    $guid = $null
                    if ($rawUninstallString -match '\{[0-9A-Fa-f-]+\}') { $guid = $matches[0].ToUpperInvariant() }
                    [void]$results.Add([PSCustomObject]@{
                        Device               = $deviceLabel
                        DisplayName          = $displayName
                        GUID                 = $guid
                        InstallSource        = $installSource
                        UninstallString      = $rawUninstallString
                        QuietUninstallString = $quietUninstallString
                        InstallDate          = $installDate
                        DisplayVersion       = $displayVersion
                        RegistryPath         = "$storedPathPrefix\$subKeyName"
                        HasGUID              = (-not [string]::IsNullOrWhiteSpace($guid))
                    })
                }
                $registryKey.Close()
            }
            catch { }
        }
        return $results
    }
    $totalComputers = $computersToQuery.Count
    $currentComputer = 0
    foreach ($computerName in $computersToQuery) {
        if ($script:RegistryStopRequested) { 
            Write-Log "Registry search stopped by user"
            break 
        }
        $currentComputer++
        $isRemote    = -not [string]::IsNullOrWhiteSpace($computerName)
        $deviceLabel =  if ($isRemote) { 
                            if ($deviceDisplayNames.ContainsKey($computerName)) { $deviceDisplayNames[$computerName] } else { $computerName }
                        } else { $env:COMPUTERNAME }
        $Tab3_statusLabel.Text = "Querying $deviceLabel ($currentComputer/$totalComputers)..."
        $Tab3_progressBar.Value = [Math]::Min(($currentComputer / $totalComputers) * 100, 100)
        Update-ProgressBarWidth $Tab3_statusStrip $Tab3_statusLabel $Tab3_stopButton $Tab3_progressBar
        [System.Windows.Forms.Application]::DoEvents()
        Write-Log "Querying registry on : $deviceLabel (Remote=$isRemote)"
        $loggedOnUserSID = $null
        $hkuPaths = @($registryPaths | Where-Object { $_.StartsWith("HKU:") })
        if ($hkuPaths.Count -gt 0 -and $isRemote) {
            $connectionResult = $script:RemoteConnectionResults | Where-Object { $_.Computer -eq $computerName } | Select-Object -First 1
            if ($connectionResult -and $connectionResult.ConsoleUserSID) {
                $loggedOnUserSID = $connectionResult.ConsoleUserSID
                Write-Log "Using cached SID for $computerName : $loggedOnUserSID" -Level Debug
            }
            else {
                $userInfo = Get-RemoteConsoleUser -ComputerName $computerName -Credential $Credential
                $loggedOnUserSID = $userInfo.SID
                if ($loggedOnUserSID) { Write-Log "Found console user SID on $computerName : $loggedOnUserSID" -Level Debug }
                else                  { Write-Log "No console user on $computerName, skipping HKU paths" -Level Debug }
            }
        }
        $collectionParams = @{
            RegistryPaths     = @($registryPaths)
            FilterText        = $FilterText
            RestrictToFilter  = $RestrictToFilter
            UseQuietUninstall = $UseQuietUninstall
            DeviceLabel       = $deviceLabel
            LoggedOnUserSID   = $loggedOnUserSID
            IsRemote          = $isRemote
        }
        $entries = if ($isRemote) {
            Invoke-RemoteDataCollection -ComputerName $computerName -ScriptBlock $collectionScriptBlock -ArgumentList $collectionParams -Credential $Credential -TimeoutSeconds 120
        }
        else {
            & $collectionScriptBlock $collectionParams
        }
        if ($entries) {
            foreach ($entry in $entries) { [void]$allEntriesForSort.Add($entry) }
        }
    }
    Write-Log "Total entries found before sorting : $($allEntriesForSort.Count)"
    $script:RegistrySortColumn = -1
    $script:RegistrySortOrder = [System.Windows.Forms.SortOrder]::None
    $sortedEntries = if ($computersToQuery.Count -gt 1) {
        $deviceOrder = @{}
        for ($i = 0; $i -lt $computersToQuery.Count; $i++) {
            $deviceName = if ([string]::IsNullOrWhiteSpace($computersToQuery[$i])) { $env:COMPUTERNAME } 
                          else { 
                              if ($deviceDisplayNames.ContainsKey($computersToQuery[$i])) { $deviceDisplayNames[$computersToQuery[$i]] } 
                              else { $computersToQuery[$i] }
                          }
            $deviceOrder[$deviceName] = $i
        }
        $allEntriesForSort | Sort-Object @{Expression = { $deviceOrder[$_.Device] }}, @{Expression = { -not $_.HasGUID }}, @{Expression = { $_.DisplayName }}
    }
    else {
        $allEntriesForSort | Sort-Object @{Expression = { -not $_.HasGUID }}, @{Expression = { $_.DisplayName }}
    }
    $showDeviceColumn = ($computersToQuery.Count -gt 1 -or ($computersToQuery.Count -eq 1 -and -not [string]::IsNullOrWhiteSpace($computersToQuery[0])))
    $script:RegistryCache = [System.Collections.ArrayList]::new()
    $sortIndex = 0
    foreach ($entry in $sortedEntries) {
        $sortIndex++
        $formattedPath = Format-RegistryPathDisplay -RegistryPath $entry.RegistryPath
        $cacheEntry = [PSCustomObject]@{
            SortIndex            = $sortIndex
            Device               = $entry.Device
            DisplayName          = $entry.DisplayName
            GUID                 = $entry.GUID
            InstallSource        = $entry.InstallSource
            UninstallString      = $entry.UninstallString
            QuietUninstallString = $entry.QuietUninstallString
            InstallDate          = $entry.InstallDate
            DisplayVersion       = $entry.DisplayVersion
            RegistryPath         = $entry.RegistryPath
            FormattedPath        = $formattedPath
            HasGUID              = $entry.HasGUID
        }
        [void]$script:RegistryCache.Add($cacheEntry)
    }
    Update-RegistryListViewFilter
    $Tab3_progressBar.Value = 0
    $Tab3_stopButton.Enabled = $false
    $Tab3_statusLabel.Text = "$($listViewparam.Items.Count) items"
    Update-ProgressBarWidth $Tab3_statusStrip $Tab3_statusLabel $Tab3_stopButton $Tab3_progressBar
    Write-Log "ListView populated with $($listViewparam.Items.Count) items"
    $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor
}

Function update-RegistryListViewFromCache {
    param(
        [System.Windows.Forms.ListView]$listViewparam,
        [System.Collections.ArrayList]$allListViewItems,
        [bool]$ShowDevice = $false,
        [bool]$ShowInstallSource = $false,
        [bool]$UseQuietUninstall = $false,
        [System.Collections.ArrayList]$SourceData = $null
    )
    if ($null -eq $SourceData) { 
        if ($null -eq $script:RegistryCache -or $script:RegistryCache.Count -eq 0) { return }
        $SourceData = $script:RegistryCache
    }
    $listViewparam.BeginUpdate()
    $listViewparam.Items.Clear()
    $allListViewItems.Clear()
    Update-DeviceColumnVisibility -ListView $listViewparam -ShowDevice $ShowDevice
    Update-InstallSourceColumnVisibility -ListView $listViewparam -ShowInstallSource $ShowInstallSource
    $sortedCache = if ($script:RegistrySortColumn -ge 0 -and $script:RegistrySortOrder -ne [System.Windows.Forms.SortOrder]::None) {
        $colText = $listViewparam.Columns[$script:RegistrySortColumn].Text
        $propName = switch ($colText) {
            $([string][char]0x21A1) { "SortIndex" }
            "Device"                { "Device" }
            "DisplayName"           { "DisplayName" }
            "GUID"                  { "GUID" }
            "InstallSource"         { "InstallSource" }
            "UninstallString"       { "UninstallString" }
            "InstallDate"           { "InstallDate" }
            "DisplayVersion"        { "DisplayVersion" }
            "Registry Path"         { "FormattedPath" }
            default                 { "SortIndex" }
        }
        if ($script:RegistrySortOrder -eq [System.Windows.Forms.SortOrder]::Ascending) {
            $SourceData | Sort-Object -Property $propName
        }
        else {
            $SourceData | Sort-Object -Property $propName -Descending
        }
    }
    else {
        $SourceData | Sort-Object -Property SortIndex
    }
    foreach ($entry in $sortedCache) {
        $dataToAdd = [ordered]@{}
        $dataToAdd[$([string][char]0x21A1)] = "" 
        if ($ShowDevice) { $dataToAdd["Device"] = $entry.Device }
        $dataToAdd["DisplayName"]     = $entry.DisplayName
        $dataToAdd["GUID"]            = $entry.GUID
        if ($ShowInstallSource) { $dataToAdd["InstallSource"] = $entry.InstallSource }
        $effectiveUninstall = if ($UseQuietUninstall -and -not [string]::IsNullOrWhiteSpace($entry.QuietUninstallString)) {
            $entry.QuietUninstallString
        } else {
            $entry.UninstallString
        }
        $dataToAdd["UninstallString"] = $effectiveUninstall
        $dataToAdd["InstallDate"]     = $entry.InstallDate
        $dataToAdd["DisplayVersion"]  = $entry.DisplayVersion
        $dataToAdd["Registry Path"]   = $entry.FormattedPath
        $item     = New-Object System.Windows.Forms.ListViewItem
        $item.Tag = $entry.RegistryPath
        $sortedKeys = $listViewparam.Columns | Select-Object -ExpandProperty Text
        $firstKey   = $sortedKeys[0]
        $item.Text  = if ($dataToAdd.Contains($firstKey)) { $dataToAdd[$firstKey] } else { "" }
        foreach ($key in $sortedKeys | Where-Object { $_ -ne $firstKey }) {
            $value = if ($dataToAdd.Contains($key)) { $dataToAdd[$key] } else { "" }
            if ($null -eq $value) { $value = "" }
            $item.SubItems.Add($value)
        }
        [void]$listViewparam.Items.Add($item)
        [void]$allListViewItems.Add($item)
    }
    $listViewparam.EndUpdate()
    $Tab3_statusLabel.Text = "$($listViewparam.Items.Count) items"
    Update-ProgressBarWidth $Tab3_statusStrip $Tab3_statusLabel $Tab3_stopButton $Tab3_progressBar
}


$launch_progressBar.Value = 80


function Update-RegistryListViewFilter {
    if ($null -eq $script:RegistryCache -or $script:RegistryCache.Count -eq 0) { return }
    $filterTerms = @()
    if ($searchTextBox_Tab3.Tag -is [hashtable] -and -not $searchTextBox_Tab3.Tag.IsPlaceholder) {
        $filterTerms = @($searchTextBox_Tab3.Text -split ';' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ -ne '' })
    }
    $hasTextFilter = ($filterTerms.Count -gt 0)
    $filteredCache = [System.Collections.ArrayList]::new()
    foreach ($entry in $script:RegistryCache) {
        $path = $entry.RegistryPath
        $includeByHive = $false
        if ($path.StartsWith("HKLM:", [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($checkbox_HKLM64.Checked -and $path -like "HKLM:Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -and $path -notlike "*Wow6432Node*") {
                $includeByHive = $true
            }
            elseif ($checkbox_HKLM32.Checked -and $path -like "HKLM:Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*") {
                $includeByHive = $true
            }
        }
        elseif ($path.StartsWith("HKCU:", [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($checkbox_HKCU64.Checked -and $path -like "HKCU:Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -and $path -notlike "*Wow6432Node*") {
                $includeByHive = $true
            }
            elseif ($checkbox_HKCU32.Checked -and $path -like "HKCU:Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*") {
                $includeByHive = $true
            }
        }
        elseif ($path.StartsWith("HKU:", [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($checkbox_HKCU64.Checked -and $path -like "HKU:*\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -and $path -notlike "*Wow6432Node*") {
                $includeByHive = $true
            }
            elseif ($checkbox_HKCU32.Checked -and $path -like "HKU:*\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*") {
                $includeByHive = $true
            }
        }
        if (-not $includeByHive) { continue }
        if ($hasTextFilter) {
            $matchText = $false
            $fieldsToCheck = @(
                $entry.DisplayName, $entry.GUID, $entry.UninstallString,
                $entry.DisplayVersion, $entry.InstallDate, $entry.InstallSource,
                $entry.FormattedPath, $entry.Device
            )
            foreach ($term in $filterTerms) {
                if ($matchText) { break }
                foreach ($field in $fieldsToCheck) {
                    if ($field -and $field.ToLower() -like "*$term*") { $matchText = $true ; break }
                }
            }
            if (-not $matchText) { continue }
        }
        [void]$filteredCache.Add($entry)
    }
    $showDevice = $false
    for ($i = 0; $i -lt $listView_Registry.Columns.Count; $i++) {
        if ($listView_Registry.Columns[$i].Text -eq "Device") { $showDevice = $true; break }
    }
    update-RegistryListViewFromCache `
        -listViewparam $listView_Registry `
        -allListViewItems $allListViewItems_Registry `
        -ShowDevice $showDevice `
        -ShowInstallSource $checkbox_ShowInstallSource.Checked `
        -UseQuietUninstall $checkbox_QuietUninstallIfAvailable.Checked `
        -SourceData $filteredCache
    AdjustListViewColumns -listView $listView_Registry
}

$checkbox_HKCU64.Add_CheckedChanged({ Update-RegistryListViewFilter })
$checkbox_HKCU32.Add_CheckedChanged({ Update-RegistryListViewFilter })
$checkbox_HKLM64.Add_CheckedChanged({ Update-RegistryListViewFilter })
$checkbox_HKLM32.Add_CheckedChanged({ Update-RegistryListViewFilter })

$HKLM32_btn.Add_Click({ Open-RegeditHere $checkbox_HKLM32.Text })
$HKLM64_btn.Add_Click({ Open-RegeditHere $checkbox_HKLM64.Text })
$HKCU32_btn.Add_Click({ Open-RegeditHere $checkbox_HKCU32.Text })
$HKCU64_btn.Add_Click({ Open-RegeditHere $checkbox_HKCU64.Text })

$searchButton_Registry.Add_Click({
    Write-Log "Search button clicked in Tab3"
    $restrictToFilter   = $checkbox_RestrictToFilter.Checked
    $filterText         = $searchTextBox_Tab3.Text.Trim()
    if                    ($filterText -eq "Filter (empty = search everything)") { $filterText = "" }
    $selectedPaths = New-Object System.Collections.ArrayList
    if ($restrictToFilter) {
        if ($checkbox_HKCU64.Checked) { [void]$selectedPaths.Add("HKU:$($checkbox_HKCU64.Text.Substring(4))") }
        if ($checkbox_HKCU32.Checked) { [void]$selectedPaths.Add("HKU:$($checkbox_HKCU32.Text.Substring(4))") }
        if ($checkbox_HKLM64.Checked) { [void]$selectedPaths.Add("HKLM:$($checkbox_HKLM64.Text.Substring(5))") }
        if ($checkbox_HKLM32.Checked) { [void]$selectedPaths.Add("HKLM:$($checkbox_HKLM32.Text.Substring(5))") }
        if ($selectedPaths.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one registry hive to search.", "No Hive Selected", 
                                                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                                                    [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }
    else {
        [void]$selectedPaths.Add("HKU:$($checkbox_HKCU64.Text.Substring(4))")
        [void]$selectedPaths.Add("HKU:$($checkbox_HKCU32.Text.Substring(4))")
        [void]$selectedPaths.Add("HKLM:$($checkbox_HKLM64.Text.Substring(5))")
        [void]$selectedPaths.Add("HKLM:$($checkbox_HKLM32.Text.Substring(5))")
    }
    $useQuietUninstall  = $checkbox_QuietUninstallIfAvailable.Checked
    $showInstallSource  = $checkbox_ShowInstallSource.Checked
    $targetComputers    = Get-TargetComputersFromPanel
    $credential         = Get-CredentialFromPanel
    Update-DeviceColumnVisibility -ListView $listView_Registry -ShowDevice ($targetComputers.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($targetComputers[0]))
    Update-InstallSourceColumnVisibility -ListView $listView_Registry -ShowInstallSource $showInstallSource
    if ($targetComputers.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($targetComputers[0])) {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $searchButton_Registry.Enabled = $false
        $searchButton_Registry.Text = "Testing..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $connectionResults = Test-RemoteConnections -Computers $targetComputers -Credential $credential
            Update-ConnectionStatusDisplay -Results $connectionResults -Credential $credential
            $connectedComputers = @($script:RemoteConnectionResults | Where-Object { $_.Success } | ForEach-Object { $_.Computer })
            if ($connectedComputers.Count -eq 0) { return }
            $targetComputers = $connectedComputers
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $searchButton_Registry.Enabled = $true
            $searchButton_Registry.Text = "Search"
        }
    }
    PopulateRegistryListView -listViewparam $listView_Registry -registryPaths $selectedPaths -allListViewItems $allListViewItems_Registry -FilterText $filterText -RestrictToFilter $restrictToFilter -UseQuietUninstall $useQuietUninstall -ShowInstallSource $showInstallSource -TargetComputers $targetComputers -Credential $credential
})


$listView_Registry.Add_ColumnClick({
    param($s, $e)
    $colText = $listView_Registry.Columns[$e.Column].Text
    if ($colText -eq $([string][char]0x2198)) { 
        $script:RegistrySortColumn = -1
        $script:RegistrySortOrder  = [System.Windows.Forms.SortOrder]::None
        $listView_Registry.ListViewItemSorter = $null
    }
    else {
        if ($e.Column -eq $script:RegistrySortColumn) {
            $script:RegistrySortOrder = if ($script:RegistrySortOrder -eq [System.Windows.Forms.SortOrder]::Ascending) { 
                [System.Windows.Forms.SortOrder]::Descending 
            } else { 
                [System.Windows.Forms.SortOrder]::Ascending 
            }
        }
        else {
            $script:RegistrySortColumn = $e.Column
            $script:RegistrySortOrder = [System.Windows.Forms.SortOrder]::Ascending
        }
    }
    Update-RegistryListViewFilter
})


$launch_progressBar.Value = 85


#region Tab 4 : MSI CLEANUP

$script:CancelRequested              = $false
$script:ProductCache                 = @{}  # Structure: $ProductCache[$computerName][$guid]
$script:CacheVersion                 = 0
$script:UninstallTabCacheVersion     = -1
$script:FullCacheTabCacheVersion     = -1
$script:CompareTabCacheVersion       = -1
$script:UninstallPanelStates         = @{}
$script:SyncHash                     = $null
$script:BackgroundPowerShell         = $null
$script:BackgroundRunspace           = $null
$script:IsBackgroundOperationRunning = $false
$script:PendingTabIndex              = $null
$script:StopButtonBlinkTimer         = $null
$script:SelectedUninstallComputer    = $null
$script:LastSearchComputers          = @()
$script:PreviousTab4Index = 0

function Split-RegistryPath {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path) -or $Path.Length -lt 5) { return $null }
    $root    = $null
    $subPath = $null
    if ($Path[4] -eq ':' -and $Path.Length -gt 6 -and $Path[5] -eq '\') {
        $prefix = $Path.Substring(0, 4).ToUpperInvariant()
        switch ($prefix) {
            'HKLM' { $root = [Microsoft.Win32.Registry]::LocalMachine; $subPath = $Path.Substring(6) }
            'HKCU' { $root = [Microsoft.Win32.Registry]::CurrentUser;  $subPath = $Path.Substring(6) }
            'HKCR' { $root = [Microsoft.Win32.Registry]::ClassesRoot;  $subPath = $Path.Substring(6) }
        }
        if ($null -ne $root) { return @{ Root = $root; SubPath = $subPath } }
    }
    if ($Path.Length -gt 5 -and $Path[3] -eq ':' -and $Path[4] -eq '\') {
        if ($Path.Substring(0, 3).ToUpperInvariant() -eq 'HKU') {
            return @{ Root = [Microsoft.Win32.Registry]::Users; SubPath = $Path.Substring(5) }
        }
    }
    return $null
}

function Format-RegistryPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    $prefixToRemove = 'Microsoft.PowerShell.Core\Registry::'
    if ($Path.StartsWith($prefixToRemove, [StringComparison]::OrdinalIgnoreCase)) {
        $Path = $Path.Substring($prefixToRemove.Length)
    }
    switch -Regex ($Path) {
        '^HKLM:\\'              { return 'HKLM\' + $Path.Substring(6) }
        '^HKCU:\\'              { return 'HKCU\' + $Path.Substring(6) }
        '^HKCR:\\'              { return 'HKCR\' + $Path.Substring(6) }
        '^HKU:\\'               { return 'HKU\'  + $Path.Substring(5) }
        '^HKEY_LOCAL_MACHINE\\' { return 'HKLM\' + $Path.Substring(19) }
        '^HKEY_CURRENT_USER\\'  { return 'HKCU\' + $Path.Substring(18) }
        '^HKEY_CLASSES_ROOT\\'  { return 'HKCR\' + $Path.Substring(18) }
        '^HKEY_USERS\\'         { return 'HKU\'  + $Path.Substring(11) }
        default { return $Path }
    }
}

function ConvertTo-PowerShellPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $prefixToRemove = 'Microsoft.PowerShell.Core\Registry::'
    if ($Path.StartsWith($prefixToRemove, [StringComparison]::OrdinalIgnoreCase)) { $Path = $Path.Substring($prefixToRemove.Length) }
    if ($Path.Length -ge 5) {
        $prefix4 = $Path.Substring(0, [Math]::Min(5, $Path.Length))
        if ($prefix4 -eq 'HKLM:' -or $prefix4 -eq 'HKCU:' -or $prefix4 -eq 'HKCR:') { return $Path }
        if ($Path.Length -ge 4 -and $Path.Substring(0, 4) -eq 'HKU:') { return $Path }
    }
    switch -Regex ($Path) {
        '^HKLM\\'               { return 'HKLM:\' + $Path.Substring(5) }
        '^HKCU\\'               { return 'HKCU:\' + $Path.Substring(5) }
        '^HKCR\\'               { return 'HKCR:\' + $Path.Substring(5) }
        '^HKU\\'                { return 'HKU:\'  + $Path.Substring(4) }
        '^HKEY_LOCAL_MACHINE\\' { return 'HKLM:\' + $Path.Substring(19) }
        '^HKEY_CURRENT_USER\\'  { return 'HKCU:\' + $Path.Substring(18) }
        '^HKEY_CLASSES_ROOT\\'  { return 'HKCR:\' + $Path.Substring(18) }
        '^HKEY_USERS\\'         { return 'HKU:\'  + $Path.Substring(11) }
        default { return $Path }
    }
}

function Convert-GuidToCompressed {
    param([string]$Guid)
    $sb = [System.Text.StringBuilder]::new(32)
    $j  = 0
    for ($i = 0; $i -lt $Guid.Length -and $j -lt 32; $i++) {
        $c = $Guid[$i]
        if (($c -ge '0' -and $c -le '9') -or ($c -ge 'A' -and $c -le 'F') -or ($c -ge 'a' -and $c -le 'f')) {
            [void]$sb.Append([char]::ToUpperInvariant($c))
            $j++
        }
    }
    if ($sb.Length -ne 32) { Write-Log "Invalid GUID format : $Guid" -Level Warning; return $null }
    $clean  = $sb.ToString()
    $result = [System.Text.StringBuilder]::new(32)
    for ($i = 7; $i -ge 0; $i--)      { [void]$result.Append($clean[$i]) }
    for ($i = 11; $i -ge 8; $i--)     { [void]$result.Append($clean[$i]) }
    for ($i = 15; $i -ge 12; $i--)    { [void]$result.Append($clean[$i]) }
    for ($i = 16; $i -lt 32; $i += 2) { [void]$result.Append($clean[$i + 1]); [void]$result.Append($clean[$i]) }
    return $result.ToString()
}

function Convert-CompressedToGuid {
    param([string]$Compressed)
    if ($Compressed.Length -ne 32) { Write-Log "Invalid compressed GUID format : $Compressed" -Level Warning; return $null }
    $sb = [System.Text.StringBuilder]::new(38)
    [void]$sb.Append('{')
    for ($i = 7; $i -ge 0; $i--)      { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
    for ($i = 11; $i -ge 8; $i--)     { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
    for ($i = 15; $i -ge 12; $i--)    { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
    for ($i = 16; $i -lt 20; $i += 2) { [void]$sb.Append($Compressed[$i + 1]); [void]$sb.Append($Compressed[$i]) };   [void]$sb.Append('-')
    for ($i = 20; $i -lt 32; $i += 2) { [void]$sb.Append($Compressed[$i + 1]); [void]$sb.Append($Compressed[$i]) };   [void]$sb.Append('}')
    return $sb.ToString().ToUpperInvariant()
}

function ConvertTo-NormalizedGuid {
    param([string]$Guid)
    if ([string]::IsNullOrWhiteSpace($Guid)) { return $null }
    try {
        $parsedGuid = [System.Guid]::Parse($Guid.Trim())
        return $parsedGuid.ToString('B').ToUpperInvariant()
    }
    catch { return $null }
}

function Format-NodeDetailsRich {
    param($Node, [System.Windows.Forms.RichTextBox]$RichTextBox)
    $RichTextBox.Clear()
    if (!$Node.Tag) { return }
    $tag = $Node.Tag
    function Add-FormattedText {
        param([string]$Text, [bool]$Bold = $false)
        $start = $RichTextBox.TextLength
        $RichTextBox.AppendText($Text)
        $RichTextBox.Select($start, $Text.Length)
        if ($Bold) { $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font, [System.Drawing.FontStyle]::Bold) }
        else       { $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font, [System.Drawing.FontStyle]::Regular) }
        $RichTextBox.Select($RichTextBox.TextLength, 0)
    }
    if ($tag.Computer) {
        Add-FormattedText "Computer :" -Bold $true
        Add-FormattedText "`r`n$($tag.Computer)`r`n`r`n"
    }
    Add-FormattedText "Type :" -Bold $true
    Add-FormattedText "`r`n$($tag.Type)`r`n`r`n"
    if ($tag.Path) {
        $normalizedPath = Format-RegistryPath $tag.Path
        $pathType       = "RegistryPath"
        if     ($tag.Type -eq 'Folder')                                                     { $pathType = "FolderPath" }
        elseif ($tag.Type -in @('File', 'InstallerCache', 'LocalPackage', 'InstallSource')) { $pathType = "FilePath" }
        Add-FormattedText "${pathType} :" -Bold $true
        Add-FormattedText "`r`n$normalizedPath`r`n`r`n"
    }
    if ($tag.Data) {
        $data = $tag.Data
        if     ($tag.Type -eq 'Product') {
            foreach ($propName in @('DisplayName', 'ProductId', 'Version', 'Publisher', 'InstallLocation', 'InstallSource', 'LocalPackage', 'CompressedGuid')) {
                if ($data.$propName) { Add-FormattedText "${propName} :" -Bold $true; Add-FormattedText "`r`n$($data.$propName)`r`n`r`n" }
            }
        }
        elseif ($tag.Type -eq 'UninstallEntry') {
            foreach ($propName in @('DisplayName', 'DisplayVersion', 'UninstallString', 'QuietUninstallString', 'ModifyPath', 'InstallLocation', 'InstallSource', 'Publisher', 'ProductId', 'SystemComponent', 'Scope', 'Architecture')) {
                if ($data.$propName) { $value = $data.$propName; if ($propName -match 'Path' -and $value -match '^HKEY_') { $value = Format-RegistryPath $value }; Add-FormattedText "${propName} :" -Bold $true; Add-FormattedText "`r`n$value`r`n`r`n" }
            }
        }
        elseif ($tag.Type -in @('InstallerCache', 'LocalPackage', 'InstallSource')) {
            if ($data.FileSize)     { $sizeText = Format-FileSize -Bytes $data.FileSize; Add-FormattedText "FileSize :" -Bold $true; Add-FormattedText "`r`n$sizeText`r`n`r`n" }
            if ($data.LastModified) { 
                $formattedDate = try { ([datetime]$data.LastModified).ToString('yyyy-MM-dd HH:mm:ss') } catch { $data.LastModified }
                Add-FormattedText "LastModified :" -Bold $true
                Add-FormattedText "`r`n$formattedDate`r`n`r`n" 
            }
            if ($data.RegistryPath) { Add-FormattedText "RegistryPath :" -Bold $true;    Add-FormattedText "`r`n$(Format-RegistryPath $data.RegistryPath)`r`n`r`n" }
        }
        elseif ($data -is [PSCustomObject]) {
            $properties = $data.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' -and $_.Name -notin @('RegistryPath', 'ProductKeyPath', 'Type', 'Count', 'Keys', 'Values', 'SyncRoot', 'IsReadOnly', 'IsFixedSize', 'IsSynchronized') } | Sort-Object Name
            foreach ($prop in $properties) {
                if ($null -ne $prop.Value -and $prop.Value -ne '') {
                    $value = $prop.Value
                    if ($value -is [string] -and $value -match '^HKEY_|^HKLM|^HKCU') { $value = Format-RegistryPath $value }
                    elseif ($prop.Name -match 'Path' -and $value -is [string] -and $value -match '^HKEY_|^HKLM|^HKCU') { $value = Format-RegistryPath $value }
                    Add-FormattedText "$($prop.Name) :" -Bold $true
                    Add-FormattedText "`r`n$value`r`n`r`n"
                }
            }
        }
        elseif ($data -is [hashtable]) {
            foreach ($key in ($data.Keys | Sort-Object)) {
                if ($key -eq 'ClassPath') { continue }
                $value = $data[$key]
                if ($value -is [array]) { if ($value.Count -gt 0) { Add-FormattedText "${key} :" -Bold $true; Add-FormattedText "`r`n"; foreach ($item in $value) { Add-FormattedText "  - $item`r`n" }; Add-FormattedText "`r`n" } }
                else { if ($key -match 'Path' -and $value -is [string] -and $value -match '^HKEY_|^HKLM|^HKCU') { $value = Format-RegistryPath $value }; Add-FormattedText "${key} :" -Bold $true; Add-FormattedText "`r`n$value`r`n`r`n" }
            }
        }
    }
    $RichTextBox.Select(0, 0)
    $RichTextBox.ScrollToCaret()
}

#region Tab4 Helpers

function Get-ComputerNodeLabel {
    param([string]$ComputerName)
    if ([string]::IsNullOrWhiteSpace($ComputerName)) { return $env:COMPUTERNAME }
    $displayName = Get-ComputerDisplayName -Computer $ComputerName
    if ($displayName -ne $ComputerName) { return $displayName }
    return $ComputerName
}

function Get-LogPathFromCommandText {
    param([string]$Command)
    if ([string]::IsNullOrWhiteSpace($Command)) { return $null }
    if ($Command -match '/L\*v\s+"([^"]+)"')              { return $matches[1] }
    if ($Command -match '/L\*v\s+(\S+)')                   { return $matches[1] }
    if ($Command -match '/[Ll]og\s+"([^"]+)"')             { return $matches[1] }
    if ($Command -match '/[Ll]og\s+(\S+\.(log|txt))')     { return $matches[1] }
    return $null
}

function Test-UninstallPanelLogExists {
    param([string]$PanelKey)
    if (-not $script:UninstallPanelControls.ContainsKey($PanelKey)) { return }
    $panelInfo = $script:UninstallPanelControls[$PanelKey]
    $panel     = $panelInfo.Panel
    if (-not $panel -or -not $panel.Tag -or $panel.Tag.LogButtonShown) { return }
    $logPaths = $panel.Tag.CommandLogPaths
    if (-not $logPaths -or $logPaths.Count -eq 0) { return }
    $computerKey = $panel.Tag.Computer
    $isRemote    = ($panel.Tag.IsRemote -eq $true)
    foreach ($logPath in $logPaths) {
        if ([string]::IsNullOrWhiteSpace($logPath)) { continue }
        $exists = $false
        if ($isRemote) {
            $uncPath = $logPath -replace '^([A-Za-z]):\\', "\\$computerKey\`$1`$\"
            try { $exists = [System.IO.File]::Exists($uncPath) } catch { }
        }
        else {
            $exists = [System.IO.File]::Exists($logPath)
        }
        if ($exists) {
            $panel.Tag.LogButtonShown = $true
            $panel.Tag.LogPath        = $logPath
            $stateLabelPanel          = $panel.Tag.StateLabelPanel
            if (-not $stateLabelPanel) { return }
            $existingBtn = $null
            foreach ($ctrl in $stateLabelPanel.Controls) {
                if ($ctrl -is [System.Windows.Forms.Button] -and $ctrl.Text -eq "Show Log") { $existingBtn = $ctrl; break }
            }
            if (-not $existingBtn) {
                $effectiveWidth  = $panel.Tag.EffectiveWidth
                $logComputer     = if ($isRemote) { $computerKey } else { "" }
                $showLogBtn      = gen $stateLabelPanel "Button" "Show Log" ($effectiveWidth - 120) 0 80 20
                $showLogBtn.Tag  = @{ LogPath = $logPath; Computer = $logComputer }
                $showLogBtn.Add_Click({
                    $t = $this.Tag
                    Open-LogFile -LogPath $t.LogPath -ComputerName $t.Computer
                })
            }
            Write-Log "Log file found for panel : $logPath"
            return
        }
    }
}

function New-ConfiguredRunspace {
    param([string[]]$FunctionNames)
    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    foreach ($funcName in $FunctionNames) {
        $funcDef = Get-Command -Name $funcName -CommandType Function -ErrorAction SilentlyContinue
        if ($funcDef) {
            $funcEntry = [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($funcDef.Name, $funcDef.Definition)
            $initialSessionState.Commands.Add($funcEntry)
        }
        else { Write-Log "Function not found for runspace injection : $funcName" -Level Warning }
    }
    $runspace                = [runspacefactory]::CreateRunspace($initialSessionState)
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions  = "ReuseThread"
    $runspace.Open()
    return $runspace
}

function New-BackgroundOperationTimer {
    param(
        [Parameter(Mandatory = $true)][hashtable]$SyncHash,
        [Parameter(Mandatory = $true)][scriptblock]$OnUpdate,
        [Parameter(Mandatory = $true)][scriptblock]$OnComplete,
        [hashtable]$AdditionalState = @{},
        [int]$Interval = 50,
        [int]$MaxUpdatesPerTick = 500
    )
    $timer          = New-Object System.Windows.Forms.Timer
    $timer.Interval = $Interval
    $state = @{ SyncHash = $SyncHash; OnUpdate = $OnUpdate; OnComplete = $OnComplete; MaxUpdatesPerTick = $MaxUpdatesPerTick }
    foreach ($key in $AdditionalState.Keys) { $state[$key] = $AdditionalState[$key] }
    $timer.Tag = $state
    $timer.Add_Tick({
        $st            = $this.Tag
        $timerSyncHash = $st.SyncHash
        Update-ProgressUI_MSICleanupTab -Percent $timerSyncHash.ProgressPercent -Status $timerSyncHash.ProgressStatus
        $updatesProcessed = 0
        if ($timerSyncHash.TreeViewUpdates -and $timerSyncHash.TreeViewUpdates.Count -gt 0) {
            $script:SyncInProgress_CrossTab = $true
            try {
                while ($timerSyncHash.TreeViewUpdates.Count -gt 0 -and $updatesProcessed -lt $st.MaxUpdatesPerTick) {
                    $update = $timerSyncHash.TreeViewUpdates[0]
                    $timerSyncHash.TreeViewUpdates.RemoveAt(0)
                    $updatesProcessed++
                    & $st.OnUpdate $update $st
                }
            }
            finally { $script:SyncInProgress_CrossTab = $false }
        }
        $shouldStop = ($timerSyncHash.IsComplete -or $script:CancelRequested)
        if ($timerSyncHash.TreeViewUpdates) { $shouldStop = $shouldStop -and ($timerSyncHash.TreeViewUpdates.Count -eq 0) }
        if ($shouldStop) {
            $this.Stop()
            $this.Dispose()
            if ($script:BackgroundPowerShell) {
                try { $script:BackgroundPowerShell.Stop(); $script:BackgroundPowerShell.Dispose() } catch { }
                $script:BackgroundPowerShell = $null
            }
            if ($script:BackgroundRunspace) {
                try { $script:BackgroundRunspace.Close(); $script:BackgroundRunspace.Dispose() } catch { }
                $script:BackgroundRunspace = $null
            }
            & $st.OnComplete $timerSyncHash $st
            Complete-PendingTabSwitch
        }
    })
    return $timer
}

function Start-ProgressUI {
    param([string]$InitialStatus = "Processing...")
    $Tab4_statusLabel.Text          = $InitialStatus
    $Tab4_statusProgressBar.Value   = 0
    $Tab4_statusProgressBar.Visible = $true
    $Tab4_stopButton.Visible        = $true
    $Tab4_stopButton.Enabled        = $true
    $Tab4_stopButton.BackColor      = [System.Drawing.Color]::IndianRed
    $Tab4_stopButton.ForeColor      = [System.Drawing.Color]::White
    $searchButton_MSICleanupTab.Enabled = $false
    $resetButton_MSICleanupTab.Enabled  = $false
}

function Stop-ProgressUI {
    param([string]$FinalStatus = "Ready")
    $Tab4_statusLabel.Text          = $FinalStatus
    $Tab4_statusProgressBar.Value   = 0
    $Tab4_statusProgressBar.Visible = $false
    $Tab4_stopButton.Visible        = $false
    $Tab4_stopButton.Enabled        = $true
    $Tab4_stopButton.BackColor      = [System.Drawing.Color]::IndianRed
    $Tab4_stopButton.ForeColor      = [System.Drawing.Color]::White
    $searchButton_MSICleanupTab.Enabled = $true
    $resetButton_MSICleanupTab.Enabled  = $true
    $script:IsBackgroundOperationRunning = $script:CancelRequested = $false
    if ($script:SyncHash) { $script:SyncHash.CancelRequested = $false }
}

function Update-ProgressUI_MSICleanupTab {
    param([int]$Percent, [string]$Status)
    $Tab4_statusProgressBar.Value = [Math]::Min([Math]::Max($Percent, 0), 100)
    $Tab4_statusLabel.Text        = $Status
}

function Update-TextBoxScrollBars {
    param([System.Windows.Forms.TextBox]$TextBox)
    if (-not $TextBox -or $TextBox.IsDisposed -or -not $TextBox.Multiline) { return }
    $textSize    = [System.Windows.Forms.TextRenderer]::MeasureText($TextBox.Text, $TextBox.Font, [System.Drawing.Size]::new($TextBox.ClientSize.Width, [int]::MaxValue), [System.Windows.Forms.TextFormatFlags]::WordBreak -bor [System.Windows.Forms.TextFormatFlags]::TextBoxControl)
    $needsScroll = ($textSize.Height -gt $TextBox.ClientSize.Height)
    $desired     = if ($needsScroll) { [System.Windows.Forms.ScrollBars]::Vertical } else { [System.Windows.Forms.ScrollBars]::None }
    if ($TextBox.ScrollBars -ne $desired) {
        $TextBox.ScrollBars = $desired
        if ($desired -ne [System.Windows.Forms.ScrollBars]::None -and $TextBox.IsHandleCreated) {
            [NativeMethods]::SetWindowTheme($TextBox.Handle, "", "") | Out-Null
        }
    }
}

function New-RoundedPanel {
    param(
        [int]$Width = 400, [int]$Height = 280, [int]$Radius = 10,
        [System.Drawing.Color]$BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180),
        [System.Drawing.Color]$BackgroundColor = [System.Drawing.Color]::White,
        [bool]$IsChild = $false, [bool]$IsRecommended = $false, [int]$ChildIndent = 35
    )
    $normalMarginLeft = 5; $normalMarginRight = 5
    $outerPanel          = New-Object System.Windows.Forms.Panel
    $outerPanel.Height   = $Height
    $outerPanel.AutoSize = $false
    $outerPanel.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $panelWidth = $Width
    if ($IsChild) {
        $panelWidth        = $Width - $ChildIndent + $normalMarginLeft
        $outerPanel.Margin = New-Object System.Windows.Forms.Padding($ChildIndent, 5, $normalMarginRight, 5)
        $BorderColor       = [System.Drawing.Color]::FromArgb(200, 200, 200)
        $BackgroundColor   = [System.Drawing.Color]::FromArgb(250, 250, 250)
    }
    else { $outerPanel.Margin = New-Object System.Windows.Forms.Padding($normalMarginLeft, 5, $normalMarginRight, 5) }
    $outerPanel.Width       = $panelWidth
    $outerPanel.MinimumSize = New-Object System.Drawing.Size($panelWidth, $Height)
    $outerPanel.MaximumSize = New-Object System.Drawing.Size($panelWidth, $Height)
    if ($IsRecommended) { $BorderColor = [System.Drawing.Color]::FromArgb(76, 175, 80) }
    $type           = $outerPanel.GetType()
    $setStyleMethod = $type.GetMethod('SetStyle', [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
    $setStyleMethod.Invoke($outerPanel, @([System.Windows.Forms.ControlStyles]::SupportsTransparentBackColor, $true))
    $setStyleMethod.Invoke($outerPanel, @([System.Windows.Forms.ControlStyles]::OptimizedDoubleBuffer, $true))
    $setStyleMethod.Invoke($outerPanel, @([System.Windows.Forms.ControlStyles]::AllPaintingInWmPaint, $true))
    $setStyleMethod.Invoke($outerPanel, @([System.Windows.Forms.ControlStyles]::UserPaint, $true))
    $setStyleMethod.Invoke($outerPanel, @([System.Windows.Forms.ControlStyles]::ResizeRedraw, $true))
    $outerPanel.BackColor = [System.Drawing.Color]::Transparent
    $outerPanel.Tag = @{ Radius = $Radius; BorderColor = $BorderColor; BackgroundColor = $BackgroundColor; IsChild = $IsChild }
    $outerPanel.Add_Paint({
        param($s, $e)
        $r        = $s.Tag.Radius
        $diameter = $r * 2
        if ($s.Width -lt ($diameter + 4) -or $s.Height -lt ($diameter + 4)) { return }
        if ($null -eq $s.Tag.BackgroundColor -or $null -eq $s.Tag.BorderColor) { return }
        $g                 = $e.Graphics
        $g.SmoothingMode   = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $penWidth          = 1.5
        $offset            = [int][Math]::Ceiling($penWidth)
        $rect              = New-Object System.Drawing.Rectangle($offset, $offset, ($s.Width - $offset * 2 - 1), ($s.Height - $offset * 2 - 1))
        if ($rect.Width -lt $diameter -or $rect.Height -lt $diameter) { return }
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $path.AddArc($rect.X,                 $rect.Y,                  $diameter, $diameter, 180, 90)
        $path.AddArc($rect.Right - $diameter, $rect.Y,                  $diameter, $diameter, 270, 90)
        $path.AddArc($rect.Right - $diameter, $rect.Bottom - $diameter, $diameter, $diameter, 0,   90)
        $path.AddArc($rect.X,                 $rect.Bottom - $diameter, $diameter, $diameter, 90,  90)
        $path.CloseFigure()
        $bgBrush = New-Object System.Drawing.SolidBrush($s.Tag.BackgroundColor)
        $g.FillPath($bgBrush, $path)
        $bgBrush.Dispose()
        $borderPen           = New-Object System.Drawing.Pen($s.Tag.BorderColor, $penWidth)
        $borderPen.Alignment = [System.Drawing.Drawing2D.PenAlignment]::Center
        $g.DrawPath($borderPen, $path)
        $borderPen.Dispose()
        $path.Dispose()
    })
    $innerMargin          = $Radius / 2 + 3
    $innerPanel           = New-Object System.Windows.Forms.Panel
    $innerPanel.Location  = New-Object System.Drawing.Point($innerMargin, $innerMargin)
    $innerPanel.Size      = New-Object System.Drawing.Size(($panelWidth - $innerMargin * 2), ($Height - $innerMargin * 2))
    $innerPanel.BackColor = $BackgroundColor
    $innerPanel.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $outerPanel.Controls.Add($innerPanel)
    return @{ OuterPanel = $outerPanel; InnerPanel = $innerPanel; EffectiveWidth = $panelWidth }
}

function Get-DependencyRootKeyName {
    param([string]$RegistryPath)
    if ([string]::IsNullOrWhiteSpace($RegistryPath)) { return $null }
    $normalizedPath = Format-RegistryPath $RegistryPath
    $parts          = $normalizedPath -split '\\'
    $depIndex       = -1
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($parts[$i] -eq 'Dependencies') { $depIndex = $i; break }
    }
    if ($depIndex -ge 0 -and ($depIndex + 1) -lt $parts.Count) { return $parts[$depIndex + 1] }
    return $null
}

function Test-RegistryPathExists {
    param([string]$Path, [string]$ComputerName = "")
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $psPath = ConvertTo-PowerShellPath $Path
    if ([string]::IsNullOrWhiteSpace($psPath)) { return $false }
    # Remote check via Invoke-RemoteDataCollection
    if (-not [string]::IsNullOrWhiteSpace($ComputerName) -and $ComputerName -ne $env:COMPUTERNAME) {
        $result = Invoke-RemoteDataCollection -ComputerName $ComputerName -ScriptBlock {
            param($params)
            $testPath = $params.Path
            try {
                $parsed = $null
                if ($testPath.Length -ge 5) {
                    $prefix = $testPath.Substring(0, 4).ToUpperInvariant()
                    switch ($prefix) {
                        'HKLM' { $root = [Microsoft.Win32.Registry]::LocalMachine; $subPath = $testPath.Substring(6) }
                        'HKCU' { $root = [Microsoft.Win32.Registry]::CurrentUser;  $subPath = $testPath.Substring(6) }
                        'HKCR' { $root = [Microsoft.Win32.Registry]::ClassesRoot;  $subPath = $testPath.Substring(6) }
                    }
                    if ($root) {
                        $key = $root.OpenSubKey($subPath, $false)
                        if ($null -ne $key) { $key.Close(); return $true }
                    }
                }
                return $false
            } catch { return $false }
        } -ArgumentList @{ Path = $psPath } -Credential (Get-CredentialFromPanel) -TimeoutSeconds 30
        return ($result -eq $true)
    }
    # Local check
    $parsed = Split-RegistryPath $psPath
    if ($null -eq $parsed) { return $false }
    try {
        $key = $parsed.Root.OpenSubKey($parsed.SubPath, $false)
        if ($null -ne $key) { $key.Close(); return $true }
    } catch { }
    return $false
}

function Update-TreeViewContextMenu_MSICleanupTab {
    param($Node, [System.Windows.Forms.ContextMenuStrip]$ContextMenu, $RichTextBox)
    $ContextMenu.Items.Clear()
    if (!$Node.Tag) { return }
    $nodeTag   = $Node.Tag
    $nodeData  = $nodeTag.Data
    $menuItems = $ContextMenu.Items
    $hasCopyActions = $false
    if ($nodeData) {
        $propertyNames = if ($nodeData -is [System.Collections.IDictionary]) { [array]$nodeData.Keys } else { [array]$nodeData.PSObject.Properties.Name }
        [Array]::Sort($propertyNames)
        foreach ($propertyName in $propertyNames) {
            if ($propertyName -match 'type') { continue }
            $propertyValue = if ($nodeData -is [System.Collections.IDictionary]) { $nodeData[$propertyName] } else { $nodeData.$propertyName }
            if (![string]::IsNullOrWhiteSpace($propertyValue) -and ($propertyValue -is [string] -or $propertyValue -is [ValueType])) {
                $copyMenuItem     = [System.Windows.Forms.ToolStripMenuItem]::new("Copy $propertyName")
                $copyMenuItem.Tag = "$propertyValue"
                $copyMenuItem.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($this.Tag) })
                [void]$menuItems.Add($copyMenuItem)
                $hasCopyActions = $true
            }
        }
    }
    $isLocalContext  = [string]::IsNullOrWhiteSpace($nodeTag.Computer) -or $nodeTag.Computer -eq $env:COMPUTERNAME
    $computerName    = if ($isLocalContext) { "" } else { $nodeTag.Computer }
    $fileSystemPath  = if ($nodeData.FilePath) { $nodeData.FilePath } elseif ($nodeData.FolderPath) { $nodeData.FolderPath } else { $null }
    $isRegistryType  = $nodeTag.Type -match 'Registry|UninstallEntry|RegistryValue'
    if ($isRegistryType -or $fileSystemPath) {
        if ($hasCopyActions) { [void]$menuItems.Add([System.Windows.Forms.ToolStripSeparator]::new()) }
        if ($isRegistryType) {
            $regeditMenuItem     = [System.Windows.Forms.ToolStripMenuItem]::new("Open Regedit here")
            $regeditMenuItem.Tag = @{ Path = $nodeTag.Path; Computer = $computerName }
            $regeditMenuItem.Add_Click({
                $t = $this.Tag
                if ($t.Path) { Open-RegeditHere -Path $t.Path -ComputerName $t.Computer }
            })
            [void]$menuItems.Add($regeditMenuItem)
        }
        if ($fileSystemPath) {
            if ($isRegistryType) { [void]$menuItems.Add([System.Windows.Forms.ToolStripSeparator]::new()) }
            $isFile        = [bool]$nodeData.FilePath
            $explorerLabel = if ($isFile) { "Show in Explorer" } else { "Open Folder" }
            $explorerAction = if ($isFile) { 'Select' } else { 'Explore' }
            $explorerMenuItem     = [System.Windows.Forms.ToolStripMenuItem]::new($explorerLabel)
            $explorerMenuItem.Tag = @{ Path = $fileSystemPath; Computer = $computerName; Action = $explorerAction }
            $explorerMenuItem.Add_Click({
                $t = $this.Tag
                Open-FileSystemPath -Path $t.Path -ComputerName $t.Computer -Action $t.Action
            })
            [void]$menuItems.Add($explorerMenuItem)
            $propertiesMenuItem     = [System.Windows.Forms.ToolStripMenuItem]::new("Show Properties")
            $propertiesMenuItem.Tag = @{ Path = $fileSystemPath; Computer = $computerName }
            $propertiesMenuItem.Add_Click({
                $t = $this.Tag
                Open-FileSystemPath -Path $t.Path -ComputerName $t.Computer -Action 'Properties'
            })
            [void]$menuItems.Add($propertiesMenuItem)
            if ($isFile -and $fileSystemPath -match '\.(msi|msp|mst)$') {
                $msiPropsMenuItem     = [System.Windows.Forms.ToolStripMenuItem]::new("Show .MSI Properties")
                $msiPropsMenuItem.Tag = @{ Path = $fileSystemPath; Computer = $computerName }
                $msiPropsMenuItem.Add_Click({
                    $t  = $this.Tag
                    $fp = $t.Path
                    if ($t.Computer) { $fp = $fp -replace '^([A-Za-z]):\\', "\\$($t.Computer)\`$1`$\" }
                    $script:fromBrowseButton = $true
                    $textBoxPath_Tab1.Text = $fp
                    $tabControl.SelectedTab = $tabPage1
                    $findGuidButton.PerformClick()
                })
                [void]$menuItems.Add($msiPropsMenuItem)
            }
            if (-not $isFile) {
                $exploreMsiMenuItem     = [System.Windows.Forms.ToolStripMenuItem]::new("Explore MSI files")
                $exploreMsiMenuItem.Tag = @{ Path = $fileSystemPath; Computer = $computerName }
                $exploreMsiMenuItem.Add_Click({
                    $t  = $this.Tag
                    $fp = $t.Path
                    if ($t.Computer) { $fp = $fp -replace '^([A-Za-z]):\\', "\\$($t.Computer)\`$1`$\" }
                    $textBoxPath_Tab2.Text = $fp
                    $tabControl.SelectedTab = $tabPage2
                    $gotoButton.PerformClick()
                })
                [void]$menuItems.Add($exploreMsiMenuItem)
            }
        }
    }
}

function Add-TreeViewCategoryNode_MSICleanupTab {
    param([System.Windows.Forms.TreeNode]$ParentNode, [string]$Label, [int]$Count)
    $node          = $ParentNode.Nodes.Add("$Label ($Count)")
    $node.NodeFont = $script:CachedFonts.CategoryBold
    $node.Tag      = @{ Type = 'Category'; Label = $Label }
    return $node
}

function Add-TreeViewItemNode_MSICleanupTab {
    param([System.Windows.Forms.TreeNode]$ParentNode, [string]$Text, [string]$Type, [string]$Path, $Data, [string]$RegistryValue, [string]$Computer = "")
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $node = $ParentNode.Nodes.Add($Text)
    $tag  = @{ Type = $Type; Path = $Path; Data = $Data; Computer = $Computer }
    if ($RegistryValue) { $tag.RegistryValue = $RegistryValue }
    $node.Tag     = $tag
    $node.Checked = $false
    return $node
}

function Get-OrCreateComputerNode {
    param([System.Windows.Forms.TreeView]$TreeView, [string]$ComputerName, [hashtable]$ComputerNodes)
    $nodeKey = if ([string]::IsNullOrWhiteSpace($ComputerName)) { "" } else { $ComputerName }
    if ($ComputerNodes.ContainsKey($nodeKey)) { return $ComputerNodes[$nodeKey] }
    $nodeLabel = Get-ComputerNodeLabel -ComputerName $ComputerName
    $compNode  = $TreeView.Nodes.Add("$nodeLabel ")
    $compNode.NodeFont  = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $compNode.ForeColor = [System.Drawing.Color]::DarkGreen
    $compNode.Tag       = @{ Type = 'Computer'; Computer = $ComputerName; Data = @{ ComputerName = $ComputerName; DisplayName = $nodeLabel } }
    $compNode.Expand()
    $ComputerNodes[$nodeKey] = $compNode
    Write-Log "Created computer node : $nodeLabel"
    return $compNode
}

function Add-CategoryToTreeView_MSICleanupTab {
    param(
        [System.Windows.Forms.TreeView]$TreeView,
        [string]$ProductGuid,
        [string]$CategoryName,
        $CategoryData,
        [hashtable]$CurrentProductNodes,
        [bool]$SuppressUpdate = $false,
        [string]$ComputerName = "",
        [hashtable]$ComputerNodes = $null,
        [bool]$IsMultiComputer = $false
    )
    # Initialize cached fonts once
    if (-not $script:CachedFonts) {
        $script:CachedFonts = @{
            Bold         = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)
            CategoryBold = $null
            Italic       = $null
        }
    }
    if ($null -eq $script:CachedFonts.CategoryBold) {
        $script:CachedFonts.CategoryBold = [System.Drawing.Font]::new($TreeView.Font, [System.Drawing.FontStyle]::Bold)
        $script:CachedFonts.Italic       = [System.Drawing.Font]::new($TreeView.Font, [System.Drawing.FontStyle]::Italic)
    }
    if (-not $SuppressUpdate) { $TreeView.BeginUpdate() }
    try {
        $countValid = {
            param($items, $prop)
            $count = 0
            foreach ($item in $items) { if ($item -and (-not $prop -or $item.$prop)) { $count++ } }
            $count
        }
        $addSimpleCategory = {
            param($rootNode, $label, $entries, $filterProp, $textBuilder, $type, $pathProp, $regValueProp, $compName)
            $count = & $countValid $entries $filterProp
            if ($count -eq 0) { return }
            $catNode = Add-TreeViewCategoryNode_MSICleanupTab -ParentNode $rootNode -Label $label -Count $count
            foreach ($item in $entries) {
                if (-not $item -or ($filterProp -and -not $item.$filterProp)) { continue }
                $text   = & $textBuilder $item
                $path   = if ($pathProp)     { $item.$pathProp }     else { $null }
                $regVal = if ($regValueProp) { $item.$regValueProp } else { $null }
                [void](Add-TreeViewItemNode_MSICleanupTab -ParentNode $catNode -Text $text -Type $type -Path $path -Data $item -RegistryValue $regVal -Computer $compName)
            }
            if ($catNode.Nodes.Count -eq 0) { $catNode.Remove() }
        }
        $addCompareCategory = {
            param($rootNode, $label, $elements, $type, $textBuilder, $compName)
            if ($elements.Count -eq 0) { return $null }
            $catNode = Add-TreeViewCategoryNode_MSICleanupTab -ParentNode $rootNode -Label $label -Count $elements.Count
            foreach ($elem in $elements) {
                $node         = $catNode.Nodes.Add((& $textBuilder $elem))
                $node.Tag     = @{ Type = $type; Path = $elem.Path; Data = $elem.Data; RegistryValue = $elem.RegistryValue; Computer = $compName }
                $node.Checked = $false
            }
            $catNode
        }
        # Handle MoveProductUnderParent
        if ($CategoryName -eq 'MoveProductUnderParent') {
            Move-ProductNodeUnderParent -TreeView $TreeView -ChildGuid $ProductGuid -ParentGuid $CategoryData.ParentGuid -CurrentProductNodes $CurrentProductNodes -ComputerName $ComputerName
            return
        }
        # Composite key for multi-computer support
        $productKey = if ($IsMultiComputer -and $ComputerName) { "${ComputerName}|${ProductGuid}" } else { $ProductGuid }
        # Handle ProductRoot / CompareProductRoot creation
        if ($CategoryName -in @('ProductRoot', 'CompareProductRoot') -and -not $CurrentProductNodes.ContainsKey($productKey)) {
            $productName     = if ($CategoryData.DisplayName) { $CategoryData.DisplayName } else { "Unknown Product" }
            $productVersion  = if ($CategoryData.Version)     { $CategoryData.Version }     else { "N/A" }
            $productGuidText = if ($CategoryData.ProductId)   { $CategoryData.ProductId }   else { "N/A" }
            $rootNodeText    = "$productName - $productVersion - $productGuidText"
            $parentGuid      = $CategoryData.ParentGuid
            $parentKey       = if ($IsMultiComputer -and $ComputerName -and $parentGuid) { "${ComputerName}|${parentGuid}" } else { $parentGuid }
            # Determine parent node (computer node for multi-PC, or existing product for parent-child)
            $targetParentNode = $null
            if ($CategoryName -eq 'ProductRoot' -and $parentKey -and $CurrentProductNodes.ContainsKey($parentKey)) {
                $targetParentNode = $CurrentProductNodes[$parentKey].RootNode
            }
            elseif ($IsMultiComputer -and $ComputerNodes) {
                $targetParentNode = Get-OrCreateComputerNode -TreeView $TreeView -ComputerName $ComputerName -ComputerNodes $ComputerNodes
            }
            $isChildProduct = ($null -ne $targetParentNode -and $targetParentNode.Tag.Type -eq 'Product')
            $rootNode = if ($targetParentNode) {
                if ($isChildProduct) { $n = $targetParentNode.Nodes.Insert(0, $rootNodeText); $n.ForeColor = [System.Drawing.Color]::DarkBlue; $n }
                else                 { $targetParentNode.Nodes.Add($rootNodeText) }
            } else { $TreeView.Nodes.Add($rootNodeText) }
            if (-not $rootNodeText.EndsWith(' ')) { $rootNode.Text = "$rootNodeText " }
            $rootNode.Tag      = @{ Type = 'Product'; Data = $CategoryData; Path = $null; Computer = $ComputerName }
            $rootNode.NodeFont = $script:CachedFonts.Bold
            if     ($CategoryData.NoUninstallResidues) { $rootNode.ForeColor = [System.Drawing.Color]::DarkOrange }
            elseif ($isChildProduct)                   { $rootNode.ForeColor = [System.Drawing.Color]::DarkBlue }
            $CurrentProductNodes[$productKey] = @{
                RootNode       = $rootNode
                CategoryNodes  = @{}
                IsChildProduct = $isChildProduct
                ParentGuid     = $parentGuid
                Computer       = $ComputerName
            }
            # Add UninstallEntries for ProductRoot
            if ($CategoryName -eq 'ProductRoot' -and $CategoryData.UninstallEntries) {
                $validCount = & $countValid $CategoryData.UninstallEntries $null
                if ($validCount -gt 0) {
                    $uninstallNode = Add-TreeViewCategoryNode_MSICleanupTab -ParentNode $rootNode -Label "[Registry] Windows Uninstall" -Count $validCount
                    foreach ($entry in $CategoryData.UninstallEntries) {
                        if (-not $entry) { continue }
                        $displayText = if ($entry.DisplayName) { "$($entry.DisplayName) [$($entry.KeyName)]" } else { $entry.KeyName }
                        [void](Add-TreeViewItemNode_MSICleanupTab -ParentNode $uninstallNode -Text $displayText -Type 'UninstallEntry' -Path $entry.RegistryPath -Data $entry -Computer $ComputerName)
                    }
                    if ($uninstallNode.Nodes.Count -eq 0) { $uninstallNode.Remove() }
                }
            }
            Expand-ProductNode_MSICleanupTab -Node $rootNode
            if ($targetParentNode -and $targetParentNode.Tag.Type -eq 'Computer') { $targetParentNode.Expand() }
            if ($targetParentNode -and $targetParentNode.Tag.Type -eq 'Product') {
                Expand-ProductNode_MSICleanupTab -Node $targetParentNode -ExpandChildren $true
            }
            return
        }
        if (-not $CurrentProductNodes.ContainsKey($productKey)) { return }
        $rootNode = $CurrentProductNodes[$productKey].RootNode
        if (-not $CategoryData) { return }
        # Category definitions table
        $categoryDefs = @{
            'UserDataEntries' = {
                $entryCount = & $countValid $CategoryData $null
                if ($entryCount -eq 0) { return }
                $catNode   = Add-TreeViewCategoryNode_MSICleanupTab -ParentNode $rootNode -Label "[Registry] Installer UserData" -Count $entryCount
                $sidGroups = @{}
                foreach ($entry in $CategoryData) {
                    if ($entry) {
                        $sid = $entry.UserSID
                        if (-not $sidGroups[$sid]) { $sidGroups[$sid] = [System.Collections.Generic.List[object]]::new() }
                        $sidGroups[$sid].Add($entry)
                    }
                }
                foreach ($sid in $sidGroups.Keys) {
                    $group   = $sidGroups[$sid]
                    $sidNode = $catNode.Nodes.Add("$sid ($($group.Count))"); $sidNode.NodeFont = $script:CachedFonts.Italic
                    foreach ($entry in $group) {
                        $text = if ($entry.DisplayName) { "$($entry.DisplayName) [$($entry.CompressedGuid)]" } else { $entry.CompressedGuid }
                        [void](Add-TreeViewItemNode_MSICleanupTab -ParentNode $sidNode -Text $text -Type 'Registry' -Path $entry.ProductKeyPath -Data $entry -Computer $ComputerName)
                    }
                }
                if ($catNode.Nodes.Count -eq 0) { $catNode.Remove() }
            }
            'Dependencies' = {
                $entryCount = & $countValid $CategoryData $null
                if ($entryCount -eq 0) { return }
                $catNode   = Add-TreeViewCategoryNode_MSICleanupTab -ParentNode $rootNode -Label "[Registry] Installer Dependencies" -Count $entryCount
                $typeGroups = @{}
                foreach ($d in $CategoryData) {
                    if (-not $d) { continue }
                    # Skip entries with no meaningful identification data
                    $hasIdentifier = $d.DependencyType -or $d.DisplayName -or $d.DependencyRootKey -or $d.DependentGuid -or $d.RegistryPath
                    if (-not $hasIdentifier) { continue }
                    $depType = if ($d.DependencyType) { $d.DependencyType } else { 'Other' }
                    if (-not $typeGroups[$depType]) { $typeGroups[$depType] = [System.Collections.Generic.List[object]]::new() }
                    $typeGroups[$depType].Add($d)
                }
                foreach ($depType in ($typeGroups.Keys | Sort-Object)) {
                    $group    = $typeGroups[$depType]
                    if ($group.Count -eq 0) { continue }
                    $typeNode = $catNode.Nodes.Add("$depType ($($group.Count))")
                    $typeNode.NodeFont = $script:CachedFonts.Italic
                    foreach ($d in $group) {
                        $text = if     ($d.DisplayName -and $d.DisplayVersion) { "$($d.DisplayName) - $($d.DisplayVersion) - $($d.DependencyRootKey)" }
                                elseif ($d.DisplayName)                        { "$($d.DisplayName) - $($d.DependencyRootKey)" }
                                elseif ($d.DependencyRootKey)                  { $d.DependencyRootKey }
                                elseif ($d.DependentGuid)                      { $d.DependentGuid }
                                else                                           { continue }
                        [void](Add-TreeViewItemNode_MSICleanupTab -ParentNode $typeNode -Text $text -Type 'Registry' -Path $d.RegistryPath -Data $d -Computer $ComputerName)
                    }
                    if ($typeNode.Nodes.Count -eq 0) { $typeNode.Remove() }
                }
                if ($catNode.Nodes.Count -eq 0) { $catNode.Remove() }
            }
            'UpgradeCodes'       = { & $addSimpleCategory $rootNode "[Registry] Installer Upgrade Codes" $CategoryData 'RegistryPath' { param($u) Format-RegistryPath $u.RegistryPath }                                               'Registry'       'RegistryPath' $null $ComputerName }
            'Features'           = { & $addSimpleCategory $rootNode "[Registry] Installer Features"      $CategoryData 'RegistryPath' { param($f) Format-RegistryPath $f.RegistryPath }                                               'Registry'       'RegistryPath' $null $ComputerName }
            'Components'         = { & $addSimpleCategory $rootNode "[Registry] Installer Components"    $CategoryData 'ComponentId'  { param($c) "$($c.ComponentId) : $($c.ComponentPath)" }                                        'Registry'       'RegistryPath' $null $ComputerName }
            'InstallerProducts'  = { & $addSimpleCategory $rootNode "[Registry] Installer Products"      $CategoryData $null          { param($p) if ($p.ProductName) { "$($p.ProductName) [$($p.PackageCode)]" } else { $p.PackageCode } } 'Registry' 'RegistryPath' $null $ComputerName }
            'InstallerFolders'   = { & $addSimpleCategory $rootNode "[Registry] Installer Folders"       $CategoryData 'FolderPath'   { param($f) $t = $f.FolderPath; if ($f.PSObject.Properties['Exists'] -and -not $f.Exists) { $t += " [Not Found]" }; $t } 'RegistryValue' 'RegistryPath' 'RegistryValue' $ComputerName }
            'InstallerFiles'     = { & $addSimpleCategory $rootNode "[Registry] Installer Files"         $CategoryData 'FilePath'     { param($f) "$($f.FilePath) ($(Format-FileSize -Bytes $f.FileSize))" }                         'RegistryValue'  'RegistryPath' $null $ComputerName }
            'DiskFolders'        = { & $addSimpleCategory $rootNode "[Disk] Disk Folders"                $CategoryData 'FolderPath'   { param($f) $f.FolderPath }                                                                     'Folder'         'FolderPath'   $null $ComputerName }
            'DiskFiles' = {
                $count = & $countValid $CategoryData 'FilePath'
                if ($count -eq 0) { return }
                $catNode = Add-TreeViewCategoryNode_MSICleanupTab -ParentNode $rootNode -Label "[Disk] Disk Files" -Count $count
                foreach ($file in $CategoryData) {
                    if (-not $file -or -not $file.FilePath) { continue }
                    $node         = $catNode.Nodes.Add("$($file.FilePath) ($(Format-FileSize -Bytes $file.FileSize))")
                    $node.Tag     = @{ Type = $file.Type; Path = $file.FilePath; Data = $file; Computer = $ComputerName }
                    $node.Checked = $false
                }
            }
            'CompareUninstallEntries' = {
                $valid = @($CategoryData | Where-Object { $_ -and $_.Data })
                if ($valid.Count -eq 0) { return }
                $catNode = Add-TreeViewCategoryNode_MSICleanupTab -ParentNode $rootNode -Label "[Registry] Windows Uninstall" -Count $valid.Count
                foreach ($elem in $valid) {
                    $entry = $elem.Data
                    $text  = if ($entry.DisplayName) { "$($entry.DisplayName) [$($entry.KeyName)]" } else { $entry.KeyName }
                    if (![string]::IsNullOrWhiteSpace($text)) {
                        $node         = $catNode.Nodes.Add($text)
                        $node.Tag     = @{ Type = 'UninstallEntry'; Path = $elem.Path; Data = $entry; Computer = $ComputerName }
                        $node.Checked = $false
                    }
                }
                if ($catNode.Nodes.Count -eq 0) { $catNode.Remove() }
            }
            'CompareRegistryEntries' = {
                $patterns = @(
                    @{ P = 'UserData.*Products';  L = 'Installer UserData' },
                    @{ P = 'Dependencies';        L = 'Installer Dependencies' },
                    @{ P = 'UpgradeCodes';        L = 'Installer Upgrade Codes' },
                    @{ P = '\\Features\\';        L = 'Installer Features' },
                    @{ P = '\\Components\\';      L = 'Installer Components' },
                    @{ P = 'Installer\\Products'; L = 'Installer Products' }
                )
                $buckets = @{}; foreach ($p in $patterns) { $buckets[$p.L] = [System.Collections.Generic.List[object]]::new() }
                foreach ($elem in $CategoryData) {
                    foreach ($p in $patterns) {
                        if ($elem.Path -match $p.P) { $buckets[$p.L].Add($elem); break }
                    }
                }
                foreach ($p in $patterns) {
                    $elements = $buckets[$p.L]
                    if ($elements.Count -eq 0) { continue }
                    $catNode          = $rootNode.Nodes.Add("[Registry] $($p.L) ($($elements.Count))")
                    $catNode.NodeFont = $script:CachedFonts.CategoryBold
                    foreach ($elem in $elements) {
                        $node         = $catNode.Nodes.Add((Format-RegistryPath $elem.Path))
                        $node.Tag     = @{ Type = 'Registry'; Path = $elem.Path; Data = $elem.Data; Computer = $ComputerName }
                        $node.Checked = $false
                    }
                }
            }
            'CompareRegistryValues' = {
                $folders = @($CategoryData | Where-Object { $_.Data -and $_.Data.FolderPath })
                $files   = @($CategoryData | Where-Object { $_.Data -and $_.Data.FilePath })
                & $addCompareCategory $rootNode "[Registry] Installer Folders" $folders 'RegistryValue' { param($e) $e.Data.FolderPath } $ComputerName
                & $addCompareCategory $rootNode "[Registry] Installer Files"   $files   'RegistryValue' { param($e) "$($e.Data.FilePath) ($(Format-FileSize -Bytes $e.Data.FileSize))" } $ComputerName
            }
            'CompareDiskFolders' = { $valid = @($CategoryData | Where-Object { $_.Path }); & $addCompareCategory $rootNode "[Disk] Disk Folders" $valid 'Folder' { param($e) $e.Path } $ComputerName }
            'CompareDiskFiles' = {
                $valid = @($CategoryData | Where-Object { $_.Path })
                if ($valid.Count -eq 0) { return }
                $catNode = Add-TreeViewCategoryNode_MSICleanupTab -ParentNode $rootNode -Label "[Disk] Disk Files" -Count $valid.Count
                foreach ($elem in $valid) {
                    $sizeText     = if ($elem.Data -and $elem.Data.FileSize) { Format-FileSize -Bytes $elem.Data.FileSize } else { "0 KB" }
                    $fileType     = if ($elem.Data -and $elem.Data.Type)     { $elem.Data.Type }                            else { 'File' }
                    $node         = $catNode.Nodes.Add("$($elem.Path) ($sizeText)")
                    $node.Tag     = @{ Type = $fileType; Path = $elem.Path; Data = $elem.Data; Computer = $ComputerName }
                    $node.Checked = $false
                }
            }
        }
        if ($categoryDefs.ContainsKey($CategoryName)) { & $categoryDefs[$CategoryName] }
    }
    finally { if (-not $SuppressUpdate) { $TreeView.EndUpdate() } }
}

function Get-CheckedItems_MSICleanupTab {
    param([System.Windows.Forms.TreeView]$TreeView)
    $checkedItems = [System.Collections.ArrayList]::new()
    function Get-CheckedNodesRecursive {
        param($Nodes)
        foreach ($node in $Nodes) {
            if ($node.Checked -and $node.Tag -and $node.Tag.Type -notin @('Product', 'Computer', 'Category')) {
                $checkedItems.Add($node.Tag) | Out-Null
            }
            if ($node.Nodes.Count -gt 0) { Get-CheckedNodesRecursive -Nodes $node.Nodes }
        }
    }
    Get-CheckedNodesRecursive -Nodes $TreeView.Nodes
    return @($checkedItems)
}

function Get-TreeNodeCount_MSICleanupTab {
    param([System.Windows.Forms.TreeView]$TreeView)
    $script:nodeCount = 0
    function Get-NodesCountRecursive {
        param($Nodes)
        foreach ($node in $Nodes) {
            if ($node.Tag -and $node.Tag.Type -notin @('Product', 'Computer', 'Category')) { $script:nodeCount++ }
            if ($node.Nodes.Count -gt 0) { Get-NodesCountRecursive -Nodes $node.Nodes }
        }
    }
    Get-NodesCountRecursive -Nodes $TreeView.Nodes
    return $script:nodeCount
}

function Set-TreeViewCheckState_MSICleanupTab {
    param([System.Windows.Forms.TreeView]$TreeView, [bool]$Checked)
    function Set-NodesCheckStateRecursive {
        param($Nodes, [bool]$State)
        foreach ($node in $Nodes) {
            $node.Checked = $State
            if ($node.Nodes.Count -gt 0) { Set-NodesCheckStateRecursive -Nodes $node.Nodes -State $State }
        }
    }
    Set-NodesCheckStateRecursive -Nodes $TreeView.Nodes -State $Checked
}

function Expand-ProductNode_MSICleanupTab {
    param([System.Windows.Forms.TreeNode]$Node, [bool]$ExpandChildren = $false)
    if (-not $Node) { return }
    $Node.Expand()
    if ($ExpandChildren) {
        foreach ($child in $Node.Nodes) {
            if ($child.Tag -and $child.Tag.Type -eq 'Product') { $child.Expand() }
        }
    }
}

function Set-ChildNodesCheckState_MSICleanupTab {
    param([System.Windows.Forms.TreeNode]$Node, [bool]$Checked)
    foreach ($child in $Node.Nodes) {
        $child.Checked = $Checked
        if ($child.Nodes.Count -gt 0) { Set-ChildNodesCheckState_MSICleanupTab -Node $child -Checked $Checked }
    }
}

function Move-ProductNodeUnderParent {
    param(
        [System.Windows.Forms.TreeView]$TreeView,
        [string]$ChildGuid,
        [string]$ParentGuid,
        [hashtable]$CurrentProductNodes,
        [string]$ComputerName = ""
    )
    $isMultiComputer = $ComputerName -ne ""
    $childKey  = if ($isMultiComputer) { "${ComputerName}|${ChildGuid}" }  else { $ChildGuid }
    $parentKey = if ($isMultiComputer) { "${ComputerName}|${ParentGuid}" } else { $ParentGuid }
    if (-not $CurrentProductNodes.ContainsKey($childKey))  { return $false }
    if (-not $CurrentProductNodes.ContainsKey($parentKey)) { return $false }
    $childNode  = $CurrentProductNodes[$childKey].RootNode
    $parentNode = $CurrentProductNodes[$parentKey].RootNode
    if ($childNode.Parent -eq $parentNode) { return $false }
    $TreeView.BeginUpdate()
    try {
        $originalColor = $childNode.ForeColor
        $clonedNode    = $childNode.Clone()
        $childNode.Remove()
        $parentNode.Nodes.Insert(0, $clonedNode)
        if ($originalColor -eq [System.Drawing.Color]::DarkOrange) { $clonedNode.ForeColor = [System.Drawing.Color]::DarkOrange }
        else { $clonedNode.ForeColor = [System.Drawing.Color]::DarkBlue }
        $CurrentProductNodes[$childKey].RootNode       = $clonedNode
        $CurrentProductNodes[$childKey].IsChildProduct = $true
        Expand-ProductNode_MSICleanupTab -Node $clonedNode
        Expand-ProductNode_MSICleanupTab -Node $parentNode -ExpandChildren $true
    }
    finally { $TreeView.EndUpdate() }
    return $true
}

function Open-LogFile {
    param([string]$LogPath, [string]$ComputerName = "")
    if ([string]::IsNullOrWhiteSpace($LogPath)) { return }
    $isRemote = (-not [string]::IsNullOrWhiteSpace($ComputerName) -and $ComputerName -ne $env:COMPUTERNAME)
    if (-not $isRemote) {
        if ([System.IO.File]::Exists($LogPath)) { Start-Process "notepad.exe" -ArgumentList "`"$LogPath`"" }
        return
    }
    # Try UNC path first
    $uncPath = $LogPath -replace '^([A-Za-z]):\\', "\\$ComputerName\`$1`$\"
    if ([System.IO.File]::Exists($uncPath)) {
        Start-Process "notepad.exe" -ArgumentList "`"$uncPath`""
        return
    }
    # Fallback : copy log content via Invoke-Command
    try {
        $credential   = Get-CredentialFromPanel
        $invokeParams = @{
            ComputerName = $ComputerName
            ScriptBlock  = { param($p) if ([System.IO.File]::Exists($p)) { Get-Content $p -Raw -ErrorAction Stop } else { throw "File not found : $p" } }
            ArgumentList = @($LogPath)
            ErrorAction  = 'Stop'
        }
        if ($credential) { $invokeParams.Credential = $credential }
        $content  = Invoke-Command @invokeParams
        $tempFile = [System.IO.Path]::Combine($env:TEMP, "Remote_${ComputerName}_$([System.IO.Path]::GetFileName($LogPath))")
        [System.IO.File]::WriteAllText($tempFile, $content)
        Start-Process "notepad.exe" -ArgumentList "`"$tempFile`""
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Unable to read remote log file.`n`nPath : $LogPath`nComputer : $ComputerName`n`nError : $($_.Exception.Message)",
            "Log Access Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
}

function Add-ProductToCache {
    param([PSCustomObject]$Product, [string]$ComputerName = "")
    if (-not $Product -or -not $Product.ProductId) { return }
    $compKey = if ([string]::IsNullOrWhiteSpace($ComputerName)) { "" } else { $ComputerName }
    $guid    = $Product.ProductId
    if (-not $script:ProductCache.ContainsKey($compKey)) { $script:ProductCache[$compKey] = @{} }
    $logLabel = if ($compKey -eq "") { "LOCAL" } else { $compKey }
    Write-Log "Adding product to cache : [$logLabel] $guid - $($Product.DisplayName)"
    $elements = [System.Collections.Generic.List[hashtable]]::new()
    $standardMappings = @(
        @{ Collection = $Product.UninstallEntries;  PathProp = 'RegistryPath';   Type = 'UninstallEntry' }
        @{ Collection = $Product.UserDataEntries;   PathProp = 'ProductKeyPath'; Type = 'Registry' }
        @{ Collection = $Product.Dependencies;      PathProp = 'RegistryPath';   Type = 'Registry' }
        @{ Collection = $Product.UpgradeCodes;      PathProp = 'RegistryPath';   Type = 'Registry' }
        @{ Collection = $Product.Features;          PathProp = 'RegistryPath';   Type = 'Registry' }
        @{ Collection = $Product.Components;        PathProp = 'RegistryPath';   Type = 'Registry' }
        @{ Collection = $Product.InstallerProducts; PathProp = 'RegistryPath';   Type = 'Registry' }
    )
    foreach ($mapping in $standardMappings) {
        foreach ($item in $mapping.Collection) {
            $path = $item.($mapping.PathProp)
            if ($path) { $elements.Add(@{ Type = $mapping.Type; Path = $path; Data = $item }) }
        }
    }
    foreach ($folder in $Product.InstallerFolders) {
        if ($folder.RegistryPath -and $folder.RegistryValue) { $elements.Add(@{ Type = 'RegistryValue'; Path = $folder.RegistryPath; Data = $folder; RegistryValue = $folder.RegistryValue }) }
        if ($folder.FolderPath -and $folder.Exists)          { $elements.Add(@{ Type = 'Folder';        Path = $folder.FolderPath;   Data = $folder }) }
    }
    foreach ($file in $Product.InstallerFiles) {
        if ($file.RegistryPath) { $elements.Add(@{ Type = 'RegistryValue'; Path = $file.RegistryPath; Data = $file; RegistryValue = $null }) }
        if ($file.FilePath)     { $elements.Add(@{ Type = 'File';          Path = $file.FilePath;     Data = $file }) }
    }
    $script:ProductCache[$compKey][$guid] = @{
        ProductId       = $guid
        DisplayName     = $Product.DisplayName
        Version         = $Product.Version
        Publisher       = $Product.Publisher
        InstallLocation = $Product.InstallLocation
        InstallSource   = $Product.InstallSource
        LocalPackage    = $Product.LocalPackage
        CompressedGuid  = $Product.CompressedGuid
        ParentGuid      = $Product.ParentGuid
        Elements        = $elements
        FullProduct     = $Product
        Computer        = $compKey
    }
    $script:CacheVersion++
    Write-Log "Product cached with $($elements.Count) elements"
}

function Clear-ProductCache {
    $totalProducts = 0
    foreach ($compKey in $script:ProductCache.Keys) { $totalProducts += $script:ProductCache[$compKey].Count }
    $script:ProductCache = @{}
    $script:CacheVersion++
    Write-Log "Product cache cleared ($totalProducts products removed)"
}

function Get-CachedComputers {
    return @($script:ProductCache.Keys | Where-Object { $_ -ne "" })
}

#region Tab4 TreeView Synchro

$script:SyncInProgress_CrossTab   = $false
$script:SharedExpandedSyncPaths   = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:SharedSelectedSyncPath    = $null
$script:SharedSplitterRatio       = 0.6
$script:SharedCheckedSyncPaths    = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

function Get-NodeSyncPath {
    param([System.Windows.Forms.TreeNode]$Node)
    if (-not $Node) { return $null }
    $tag = $Node.Tag
    if (-not $tag) {
        # Grouping nodes without tags (SID groups, dependency type groups)
        $parentPath = Get-NodeSyncPath $Node.Parent
        if ($parentPath) {
            $label = $Node.Text -replace '\s*\(\d+\)\s*$', ''
            return "Group|$parentPath|$label"
        }
        return $null
    }
    if ($tag -isnot [hashtable]) { return $null }
    $computer = if ($tag.ContainsKey('Computer')) { $tag['Computer'] } else { "" }
    $type     = $tag['Type']
    switch ($type) {
        'Computer' {
            return "Computer|$computer"
        }
        'Product' {
            $productId = $null
            $data = $tag['Data']
            if ($data -is [hashtable])    { $productId = $data['ProductId'] }
            elseif ($data)                { try { $productId = $data.ProductId } catch {} }
            if ($productId)               { return "Product|$computer|$productId" }
        }
        'Category' {
            $parentPath = Get-NodeSyncPath $Node.Parent
            $label      = if ($tag.ContainsKey('Label')) { $tag['Label'] } else { $Node.Text -replace '\s*\(\d+\)\s*$', '' }
            if ($parentPath) { return "Category|$parentPath|$label" }
        }
        default {
            $path = $tag['Path']
            if ($path) {
                $rv = if ($tag.ContainsKey('RegistryValue')) { $tag['RegistryValue'] } else { $null }
                if ($rv) { return "Item|$computer|$type|$path|$rv" }
                return "Item|$computer|$type|$path"
            }
        }
    }
    return $null
}

function Find-NodeBySyncPath {
    param($Nodes, [string]$SyncPath)
    foreach ($node in $Nodes) {
        $np = Get-NodeSyncPath $node
        if ($np -eq $SyncPath) { return $node }
        $found = Find-NodeBySyncPath $node.Nodes $SyncPath
        if ($found) { return $found }
    }
    return $null
}

function Save-TreeViewSyncState {
    param([System.Windows.Forms.TreeView]$TreeView)
    if (-not $TreeView -or $TreeView.Nodes.Count -eq 0) { return }
    # Only update sync state for nodes present in this TreeView, leave others untouched
    $collectState = {
        param([System.Windows.Forms.TreeNode]$node)
        $sp = Get-NodeSyncPath $node
        if ($sp) {
            if ($node.IsExpanded) { [void]$script:SharedExpandedSyncPaths.Add($sp) }
            else                  { [void]$script:SharedExpandedSyncPaths.Remove($sp) }
            if ($node.Checked)    { [void]$script:SharedCheckedSyncPaths.Add($sp) }
            else                  { [void]$script:SharedCheckedSyncPaths.Remove($sp) }
        }
        foreach ($child in $node.Nodes) { & $collectState $child }
    }
    foreach ($root in $TreeView.Nodes) { & $collectState $root }
    # Save selection
    if ($TreeView.SelectedNode) {
        $script:SharedSelectedSyncPath = Get-NodeSyncPath $TreeView.SelectedNode
    }
}

function Restore-TreeViewSyncState {
    param([System.Windows.Forms.TreeView]$TreeView)
    if (-not $TreeView -or $TreeView.Nodes.Count -eq 0) { return }
    $script:SyncInProgress_CrossTab = $true
    try {
        $TreeView.BeginUpdate()
        try {
            # Restore expand/collapse state
            $restoreExpand = {
                param([System.Windows.Forms.TreeNode]$node)
                $sp = Get-NodeSyncPath $node
                if ($sp) {
                    if ($script:SharedExpandedSyncPaths.Contains($sp)) { $node.Expand() }
                    else { $node.Collapse() }
                }
                foreach ($child in $node.Nodes) { & $restoreExpand $child }
            }
            foreach ($root in $TreeView.Nodes) { & $restoreExpand $root }
            # Restore checked state
            $restoreChecked = {
                param([System.Windows.Forms.TreeNode]$node)
                $sp = Get-NodeSyncPath $node
                if ($sp) {
                    $shouldCheck = $script:SharedCheckedSyncPaths.Contains($sp)
                    if ($node.Checked -ne $shouldCheck) { $node.Checked = $shouldCheck }
                }
                foreach ($child in $node.Nodes) { & $restoreChecked $child }
            }
            foreach ($root in $TreeView.Nodes) { & $restoreChecked $root }
        }
        finally { $TreeView.EndUpdate() }
        # Restore selection
        $selectionRestored = $false
        if ($script:SharedSelectedSyncPath) {
            $target = Find-NodeBySyncPath $TreeView.Nodes $script:SharedSelectedSyncPath
            if ($target) {
                $TreeView.SelectedNode = $target
                $target.EnsureVisible()
                $selectionRestored = $true
            }
        }
        # Force layout recalculation before scroll commands
        if ($TreeView.IsHandleCreated) { $TreeView.Update() }
        # Scroll to top when no selection dictates the scroll position
        if (-not $selectionRestored -and $TreeView.IsHandleCreated) {
            [NativeMethods]::SendMessage($TreeView.Handle, [NativeMethods]::WM_VSCROLL, [NativeMethods]::SB_TOP, 0) | Out-Null
        }
        # Reset horizontal scroll to leftmost position
        if ($TreeView.IsHandleCreated) {
            [NativeMethods]::SendMessage($TreeView.Handle, 0x0114, 6, 0) | Out-Null
            [NativeMethods]::SendMessage($TreeView.Handle, 0x0114, 8, 0) | Out-Null
        }
        # Update clean button based on restored checked state
        $checked = Get-CheckedItems_MSICleanupTab -TreeView $TreeView
        $cleanButton_MSICleanupTab.Enabled = ($checked.Count -gt 0)
    }
    finally { $script:SyncInProgress_CrossTab = $false }
}

function Sync-SplitterRatio {
    param([System.Windows.Forms.SplitContainer]$Source)
    if ($script:SyncInProgress_CrossTab) { return }
    if (-not $Source -or $Source.Width -le 0) { return }
    $script:SharedSplitterRatio = $Source.SplitterDistance / [Math]::Max(1, $Source.Width)
}

function Restore-SplitterRatio {
    param([System.Windows.Forms.SplitContainer]$Target)
    if (-not $Target -or $Target.Width -le 0) { return }
    $newDistance = [int]($Target.Width * $script:SharedSplitterRatio)
    $newDistance = [Math]::Max($Target.Panel1MinSize, [Math]::Min($newDistance, $Target.Width - $Target.Panel2MinSize - $Target.SplitterWidth))
    try { $Target.SplitterDistance = $newDistance } catch {}
}

# Sync event handler scriptblock (shared by all 3 TreeViews)
$script:SyncAfterExpandHandler = {
    if ($script:SyncInProgress_CrossTab) { return }
    $sp = Get-NodeSyncPath $_.Node
    if ($sp) { [void]$script:SharedExpandedSyncPaths.Add($sp) }
}

$script:SyncAfterCollapseHandler = {
    if ($script:SyncInProgress_CrossTab) { return }
    $sp = Get-NodeSyncPath $_.Node
    if ($sp) { [void]$script:SharedExpandedSyncPaths.Remove($sp) }
}

$script:SyncAfterSelectHandler = {
    if ($script:SyncInProgress_CrossTab) { return }
    if ($this.SelectedNode) {
        $sp = Get-NodeSyncPath $this.SelectedNode
        if ($sp) { $script:SharedSelectedSyncPath = $sp }
    }
}

$script:SyncAfterCheckHandler = {
    param($s, $e)
    if ($script:SyncInProgress_CrossTab) { return }
    if ($e.Action -eq [System.Windows.Forms.TreeViewAction]::Unknown) { return }
    $collectCheckedSync = {
        param([System.Windows.Forms.TreeNode]$n)
        $sp = Get-NodeSyncPath $n
        if ($sp) {
            if ($n.Checked) { [void]$script:SharedCheckedSyncPaths.Add($sp) }
            else            { [void]$script:SharedCheckedSyncPaths.Remove($sp) }
        }
        foreach ($c in $n.Nodes) { & $collectCheckedSync $c }
    }
    & $collectCheckedSync $e.Node
}

$script:SyncSplitterMovedHandler = {
    Sync-SplitterRatio -Source $this
}

#region Tab4 UI

# Search Panel
$borderTop                            = gen $tabPage4                  "Panel"             0 0 0 1              'Dock=Top' 'BackColor=Gray'
$searchPanel_MSICleanupTab            = gen $tabPage4                  "Panel"                                  'Height=55' 'Dock=Top' 'Padding=10 10 10 10'
$titleLabel_MSICleanupTab             = gen $searchPanel_MSICleanupTab "Label"             "Product Title :"    10  10 100 19
$titleTextBox_MSICleanupTab           = gen $searchPanel_MSICleanupTab "TextBox"           ""                   120 10 300 19
$exactMatchCheckBox_MSICleanupTab     = gen $searchPanel_MSICleanupTab "CheckBox"          "Exact Match"        430 10 100 19
$guidLabel_MSICleanupTab              = gen $searchPanel_MSICleanupTab "Label"             "Product GUID :"     10  29 100 19
$guidTextBox_MSICleanupTab            = gen $searchPanel_MSICleanupTab "TextBox"           ""                   120 29 300 19
$exactMatchGuidCheckBox_MSICleanupTab = gen $searchPanel_MSICleanupTab "CheckBox"          "Exact Match"        430 29 100 19
$searchButton_MSICleanupTab           = gen $searchPanel_MSICleanupTab "Button"            "Search"             550 10 150 40
$fastModeCheckBox_MSICleanupTab       = gen $form "CheckBox" "$([string][char]0x21AA) Progressbar (slower)" 554 ($titleBarHeight+75) 150 17 'Checked=$false'
$resetButton_MSICleanupTab            = gen $searchPanel_MSICleanupTab "Button"            "Reset"              710 10 70  40

$titleTextBox_MSICleanupTab.TabIndex           = 0
$exactMatchCheckBox_MSICleanupTab.TabIndex     = 1
$guidTextBox_MSICleanupTab.TabIndex            = 2
$exactMatchGuidCheckBox_MSICleanupTab.TabIndex = 3
$searchButton_MSICleanupTab.TabIndex           = 4
$resetButton_MSICleanupTab.TabIndex            = 5

$rightButtonsPanel_MSICleanupTab          = gen $searchPanel_MSICleanupTab "FlowLayoutPanel" 0 0 110 40 'FlowDirection=RightToLeft' 'WrapContents=$false' 'AutoSize=$false' 'Anchor=Top, Right' 'Padding=0 0 0 0'
$rightButtonsPanel_MSICleanupTab.Location = [System.Drawing.Point]::new(($searchPanel_MSICleanupTab.ClientSize.Width - 112), 10)
$cleanButton_MSICleanupTab                = gen $rightButtonsPanel_MSICleanupTab "Button" "Clean Selection" 0 0 100 40 'Margin=5 0 0 0' 'Enabled=$false'

# Tab Control
$tabControl_MSICleanupTab    = gen $tabPage4                 "TabControl" 'Dock=Fill'
$tabSearch_MSICleanupTab     = gen $tabControl_MSICleanupTab "TabPage" "Search"                        'UseVisualStyleBackColor=$true'
$tabFullCache_MSICleanupTab  = gen $tabControl_MSICleanupTab "TabPage" "Full Cache"                    'UseVisualStyleBackColor=$true' 'Enabled=$false'
$tabUninstall_MSICleanupTab  = gen $tabControl_MSICleanupTab "TabPage" "Uninstall Commands from Cache" 'UseVisualStyleBackColor=$true' 'Enabled=$false'
$tabCompare_MSICleanupTab    = gen $tabControl_MSICleanupTab "TabPage" "Compare Residues from Cache"   'UseVisualStyleBackColor=$true' 'Enabled=$false'
# $tabControl_MSICleanupTab.TabStop = $false

# Search Tab Content
$splitContainerSearch_MSICleanupTab     = gen $tabSearch_MSICleanupTab                   "SplitContainer" 'Dock=Fill' 'Orientation=Vertical'
$treeViewSearch_MSICleanupTab           = gen $splitContainerSearch_MSICleanupTab.Panel1 "TreeView"       'Dock=Fill' 'CheckBoxes=$true' 'Font=Consolas, 9' 'HideSelection=$false'
$detailsRichTextBoxSearch_MSICleanupTab = gen $splitContainerSearch_MSICleanupTab.Panel2 "RichTextBox"    'Dock=Fill' 'Font=Consolas, 9' 'ReadOnly=$true' 'WordWrap=$true'
$detailsContextMenuSearch_MSICleanupTab = New-Object System.Windows.Forms.ContextMenuStrip

$treeViewSearch_MSICleanupTab.Add_AfterCheck({
    param($s, $e)
    if ($e.Action -ne [System.Windows.Forms.TreeViewAction]::Unknown) {
        $s.BeginUpdate()
        Set-ChildNodesCheckState_MSICleanupTab -Node $e.Node -Checked $e.Node.Checked
        $s.EndUpdate()
        $checked = Get-CheckedItems_MSICleanupTab -TreeView $s
        $cleanButton_MSICleanupTab.Enabled = ($checked.Count -gt 0)
    }
})

$titleTextBox_MSICleanupTab.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $searchButton_MSICleanupTab.PerformClick(); $e.Handled = $true; $e.SuppressKeyPress = $true }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $s.SelectAll(); $e.Handled = $true; $e.SuppressKeyPress = $true }
})
$titleTextBox_MSICleanupTab.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) { $_.Handled = $true }
})
$guidTextBox_MSICleanupTab.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $searchButton_MSICleanupTab.PerformClick(); $e.Handled = $true; $e.SuppressKeyPress = $true }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $s.SelectAll(); $e.Handled = $true; $e.SuppressKeyPress = $true }
})
$guidTextBox_MSICleanupTab.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) { $_.Handled = $true }
})

$tabControl_MSICleanupTab.Add_Selecting({
    param($s, $e)
    if ($script:IsBackgroundOperationRunning) {
        $script:PendingTabIndex = $e.TabPageIndex
        $e.Cancel = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        if ($Tab4_stopButton.Visible -and $Tab4_stopButton.Enabled -and ($null -eq $script:StopButtonBlinkTimer -or -not $script:StopButtonBlinkTimer.Enabled)) {
            if ($script:StopButtonBlinkTimer) { $script:StopButtonBlinkTimer.Dispose(); $script:StopButtonBlinkTimer = $null }
            $script:StopButtonBlinkTimer          = New-Object System.Windows.Forms.Timer
            $script:StopButtonBlinkTimer.Interval = 200
            $script:StopButtonBlinkTimer.Tag      = @{ IsHighlighted = $false }
            $script:StopButtonBlinkTimer.Add_Tick({
                if ($this.Tag.IsHighlighted) { $Tab4_stopButton.BackColor = [System.Drawing.Color]::IndianRed }
                else                         { $Tab4_stopButton.BackColor = [System.Drawing.Color]::Black }
                $this.Tag.IsHighlighted = -not $this.Tag.IsHighlighted
            })
            $script:StopButtonBlinkTimer.Start()
        }
    }
})

# Uninstall Tab Content
$uninstallMainPanel_MSICleanupTab = gen $tabUninstall_MSICleanupTab       "Panel"           'Dock=Fill'
$uninstallFlowPanel_MSICleanupTab = gen $uninstallMainPanel_MSICleanupTab "FlowLayoutPanel" 'Dock=Fill' 'FlowDirection=TopDown' 'WrapContents=$false' 'AutoScroll=$true' 'Padding=10 10 10 10'

# Full Cache Tab Content
$splitContainerFullCache     = gen $tabFullCache_MSICleanupTab     "SplitContainer" 'Dock=Fill' 'Orientation=Vertical'
$treeViewFullCache           = gen $splitContainerFullCache.Panel1 "TreeView"       'Dock=Fill' 'CheckBoxes=$true' 'Font=Consolas, 9' 'HideSelection=$false'
$detailsRichTextBoxFullCache = gen $splitContainerFullCache.Panel2 "RichTextBox"    'Dock=Fill' 'Font=Consolas, 9' 'ReadOnly=$true' 'WordWrap=$true'
$detailsContextMenuFullCache = New-Object System.Windows.Forms.ContextMenuStrip

$treeViewFullCache.Add_AfterCheck({
    param($s, $e)
    if ($e.Action -ne [System.Windows.Forms.TreeViewAction]::Unknown) {
        $s.BeginUpdate()
        Set-ChildNodesCheckState_MSICleanupTab -Node $e.Node -Checked $e.Node.Checked
        $s.EndUpdate()
        $checked = Get-CheckedItems_MSICleanupTab -TreeView $s
        $cleanButton_MSICleanupTab.Enabled = ($checked.Count -gt 0)
    }
})

# Compare Tab Content
$splitContainerCompare     = gen $tabCompare_MSICleanupTab       "SplitContainer" 'Dock=Fill' 'Orientation=Vertical'
$treeViewCompare           = gen $splitContainerCompare.Panel1   "TreeView"       'Dock=Fill' 'CheckBoxes=$true' 'Font=Consolas, 9' 'HideSelection=$false'
$detailsRichTextBoxCompare = gen $splitContainerCompare.Panel2   "RichTextBox"    'Dock=Fill' 'Font=Consolas, 9' 'ReadOnly=$true' 'WordWrap=$true'
$detailsContextMenuCompare = New-Object System.Windows.Forms.ContextMenuStrip

$treeViewCompare.Add_AfterCheck({
    param($s, $e)
    if ($e.Action -ne [System.Windows.Forms.TreeViewAction]::Unknown) {
        $s.BeginUpdate()
        Set-ChildNodesCheckState_MSICleanupTab -Node $e.Node -Checked $e.Node.Checked
        $s.EndUpdate()
        $checked = Get-CheckedItems_MSICleanupTab -TreeView $s
        $cleanButton_MSICleanupTab.Enabled = ($checked.Count -gt 0)
    }
})

# Tab Button Panel
$tabButtonPanel_MSICleanupTab          = gen $form                          "Panel"                 896 ($titleBarHeight+78) 336 21 'Anchor=Top, Right' 'BackColor=[System.Drawing.SystemColors]::Control' 'Visible=$false'
$buttonFlowPanel_MSICleanupTab         = gen $tabButtonPanel_MSICleanupTab  "FlowLayoutPanel" 'Dock=Right' 'FlowDirection=RightToLeft' 'WrapContents=$false' 'AutoSize=$true' 'Padding=0 2 5 0'
$collapseAllButton_MSICleanupTab       = gen $buttonFlowPanel_MSICleanupTab "Button" "Collapse All" 0 0 80 19 'Margin=3 0 3 0'
$expandAllButton_MSICleanupTab         = gen $buttonFlowPanel_MSICleanupTab "Button" "Expand All"   0 0 80 19 'Margin=3 0 3 0'
$deselectAllButton_MSICleanupTab       = gen $buttonFlowPanel_MSICleanupTab "Button" "Deselect All" 0 0 80 19 'Margin=3 0 3 0'
$selectAllButton_MSICleanupTab         = gen $buttonFlowPanel_MSICleanupTab "Button" "Select All"   0 0 70 19 'Margin=3 0 3 0'

# Status Strip
$Tab4_statusStrip            = gen $tabPage4         "StatusStrip"                        'Dock=Bottom' 'SizingGrip=$false'
$Tab4_statusLabel            = gen $Tab4_statusStrip "ToolStripStatusLabel" "Ready"       'Spring=$true' 'TextAlign=MiddleLeft'
$Tab4_statusProgressBar      = gen $Tab4_statusStrip "ToolStripProgressBar"               'Minimum=0' 'Maximum=100' 'Value=0' 'Visible=$false'
$Tab4_statusProgressBar.Size = [System.Drawing.Size]::new(400, 16)
$Tab4_stopButton             = gen $Tab4_statusStrip "ToolStripButton" "STOP"             'Visible=$false'

$searchPanel_MSICleanupTab.BringToFront()
$tabControl_MSICleanupTab.BringToFront()
$fastModeCheckBox_MSICleanupTab.BringToFront()
$tabButtonPanel_MSICleanupTab.BringToFront()
$panel_RemoteTarget.BringToFront()

$script:CurrentProductNodes          = @{}
$script:CurrentProductNodesFullCache = @{}
$script:CurrentProductNodesCompare   = @{}
$script:ComputerNodesSearch          = @{}
$script:ComputerNodesFullCache       = @{}
$script:ComputerNodesCompare         = @{}
$script:UninstallJobs                = @{}
$script:UninstallPanelControls       = @{}

function Update-TabStates {
    $hasCache = $false
    foreach ($compKey in $script:ProductCache.Keys) { if ($script:ProductCache[$compKey].Count -gt 0) { $hasCache = $true; break } }
    $tabFullCache_MSICleanupTab.Enabled = $hasCache
    $tabCompare_MSICleanupTab.Enabled   = $hasCache
    $hasUninstallEntries = $false
    foreach ($compKey in $script:ProductCache.Keys) {
        foreach ($guid in $script:ProductCache[$compKey].Keys) {
            $cachedProduct = $script:ProductCache[$compKey][$guid]
            if ($cachedProduct.FullProduct -and $cachedProduct.FullProduct.UninstallEntries) {
                foreach ($entry in $cachedProduct.FullProduct.UninstallEntries) {
                    if (![string]::IsNullOrWhiteSpace($entry.UninstallString) -or ![string]::IsNullOrWhiteSpace($entry.QuietUninstallString) -or ![string]::IsNullOrWhiteSpace($entry.ModifyPath)) {
                        $hasUninstallEntries = $true
                        break
                    }
                }
            }
            if ($hasUninstallEntries) { break }
        }
        if ($hasUninstallEntries) { break }
    }
    $tabUninstall_MSICleanupTab.Enabled = $hasUninstallEntries
}

function Get-ActiveTreeView {
    switch ($tabControl_MSICleanupTab.SelectedIndex) {
        0       { return $treeViewSearch_MSICleanupTab }
        1       { return $treeViewFullCache }
        3       { return $treeViewCompare }
        default { return $null }
    }
}

function Complete-PendingTabSwitch {
    if ($script:StopButtonBlinkTimer) {
        $script:StopButtonBlinkTimer.Stop()
        $script:StopButtonBlinkTimer.Dispose()
        $script:StopButtonBlinkTimer = $null
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    if ($null -ne $script:PendingTabIndex) {
        $pendingIndex           = $script:PendingTabIndex
        $script:PendingTabIndex = $null
        $tabControl_MSICleanupTab.SelectedIndex = $pendingIndex
    }
}

function Update-UninstallPanelTextBoxWidths {
    param([System.Windows.Forms.Panel]$Panel, [int]$EffectiveWidth)
    $innerPanel = $null
    foreach ($ctrl in $Panel.Controls) { if ($ctrl -is [System.Windows.Forms.Panel]) { $innerPanel = $ctrl; break } }
    if (-not $innerPanel) { return }
    $innerMargin      = 8
    $innerPanel.Width = $EffectiveWidth - ($innerMargin * 2)
    $contentWidth     = $innerPanel.Width - 10
    foreach ($ctrl in $innerPanel.Controls) {
        if ($ctrl -is [System.Windows.Forms.FlowLayoutPanel]) {
            foreach ($rowPanel in $ctrl.Controls) {
                if ($rowPanel -is [System.Windows.Forms.Panel]) {
                    $rowPanel.Width = $contentWidth
                    if ($rowPanel.Height -eq 50) {
                        foreach ($innerCtrl in $rowPanel.Controls) {
                            if ($innerCtrl -is [System.Windows.Forms.TextBox] -and $innerCtrl.Multiline) {
                                $innerCtrl.Width = $contentWidth - 205
                                Update-TextBoxScrollBars -TextBox $innerCtrl
                            }
                        }
                    }
                }
            }
        }
    }
}

#region Tab4 uninstall

function Update-UninstallTab {
    Write-Log "Updating Uninstall Tab"
    # Get list of computers with cached products
    $computersWithProducts = @()
    foreach ($compKey in $script:ProductCache.Keys) {
        if ($script:ProductCache[$compKey].Count -gt 0) {
            $computersWithProducts += $compKey
        }
    }
    $isMultiComputer = ($computersWithProducts.Count -gt 1) -or ($computersWithProducts.Count -eq 1 -and $computersWithProducts[0] -ne "")
    Write-Log "Uninstall tab : $($computersWithProducts.Count) computer(s), isMultiComputer=$isMultiComputer"
    $uninstallFlowPanel_MSICleanupTab.Controls.Clear()
    $script:UninstallPanelControls = @{}
    # Create computer selector panel if multi-computer
    if ($isMultiComputer) {
        $selectorPanel = gen $uninstallFlowPanel_MSICleanupTab "Panel" 0 0 0 50 'Margin=5 5 5 10' 'BackColor=240 240 245' 'BorderStyle=FixedSingle'
        $selectorPanel.Tag = @{ IsSelector = $true; IsUpdating = $false }
        $setStyle = $selectorPanel.GetType().GetMethod('SetStyle', [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
        $setStyle.Invoke($selectorPanel, @([System.Windows.Forms.ControlStyles]::OptimizedDoubleBuffer, $true))
        $setStyle.Invoke($selectorPanel, @([System.Windows.Forms.ControlStyles]::AllPaintingInWmPaint, $true))
        $setStyle.Invoke($selectorPanel, @([System.Windows.Forms.ControlStyles]::UserPaint, $true))
        $selectorLabel       = gen $selectorPanel "Label" "Select Computer :" 10 15 110 20 'Font=Segoe UI, 9, Bold'
        $radioScrollPanel    = gen $selectorPanel "Panel" "" 125 5 0 0 'AutoScroll=$true'
        $radioFlowPanel      = gen $radioScrollPanel "FlowLayoutPanel" "" 0 0 0 0 'FlowDirection=LeftToRight' 'WrapContents=$true' 'AutoSize=$true' 'AutoSizeMode=GrowAndShrink'
        $radioFlowPanel.Tag  = @{ RadioFlow = $true }
        $firstRadio = $true
        foreach ($compKey in $computersWithProducts) {
            $displayName = if ($compKey -eq "") { $env:COMPUTERNAME } else { Get-ComputerNodeLabel -ComputerName $compKey }
            $radioLabel = $displayName
            if ($compKey -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                try {
                    $hostEntry    = [System.Net.Dns]::GetHostEntry($compKey)
                    $resolvedName = ($hostEntry.HostName -split '\.')[0]
                    if ($resolvedName -and $resolvedName -ne $compKey) { $radioLabel = "$compKey ($resolvedName)" }
                }
                catch { }
            }
            $radio          = New-Object System.Windows.Forms.RadioButton
            $radio.Text     = $radioLabel
            $radio.Tag      = $compKey
            $radio.AutoSize = $true
            $radio.Margin   = New-Object System.Windows.Forms.Padding(5, 5, 15, 5)
            $radio.Checked  = $firstRadio
            if ($firstRadio) {
                $script:SelectedUninstallComputer = $compKey
                $firstRadio = $false
            }
            $radio.Add_CheckedChanged({
                param($s, $e)
                if ($s.Checked) {
                    $script:SelectedUninstallComputer = $s.Tag
                    Write-Log "Uninstall computer selection changed to : $($s.Tag)"
                    Rebuild-UninstallPanelsForComputer -ComputerKey $s.Tag
                }
            })
            $radioFlowPanel.Controls.Add($radio)
        }
        $uninstallFlowPanel_MSICleanupTab.Controls.Add($selectorPanel)
        # Uniform radio width for grid alignment
        $maxRadioWidth = 0
        foreach ($radio in $radioFlowPanel.Controls) {
            $measured = [System.Windows.Forms.TextRenderer]::MeasureText($radio.Text, $radio.Font).Width + 30
            if ($measured -gt $maxRadioWidth) { $maxRadioWidth = $measured }
        }
        foreach ($radio in $radioFlowPanel.Controls) {
            $radio.AutoSize = $false
            $radio.Width    = $maxRadioWidth
            $radio.Height   = 22
        }
        $maxScrollRows = 3
        $radioRowH     = 22 + 10  # Height + margin
        $recalcSelectorLayout = {
            $selPanel = $null
            foreach ($ctrl in $uninstallFlowPanel_MSICleanupTab.Controls) {
                if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Tag -is [hashtable] -and $ctrl.Tag.IsSelector) {
                    $selPanel = $ctrl; break
                }
            }
            if (-not $selPanel -or $selPanel.Tag.IsUpdating) { return }
            $selPanel.Tag.IsUpdating = $true
            try {
                $flowWidth = $uninstallFlowPanel_MSICleanupTab.ClientSize.Width
                $newWidth  = $flowWidth - 45
                if ($newWidth -lt 200) { return }
                $selPanel.Width = $newWidth
                # Find scroll panel and radio flow panel
                $scrollPanel = $null; $radioFlow = $null
                foreach ($inner in $selPanel.Controls) {
                    if ($inner -is [System.Windows.Forms.Panel] -and $inner.AutoScroll) { $scrollPanel = $inner; break }
                }
                if ($scrollPanel) {
                    foreach ($inner in $scrollPanel.Controls) {
                        if ($inner -is [System.Windows.Forms.FlowLayoutPanel]) { $radioFlow = $inner; break }
                    }
                }
                if (-not $radioFlow -or $radioFlow.Controls.Count -eq 0) { return }
                $radioFlowWidth = $newWidth - 135
                if ($radioFlowWidth -lt 50) { return }
                $scrollPanel.Width = $radioFlowWidth
                # Recalculate uniform radio width
                $maxW = 0
                foreach ($radio in $radioFlow.Controls) {
                    $measured = [System.Windows.Forms.TextRenderer]::MeasureText($radio.Text, $radio.Font).Width + 30
                    if ($measured -gt $maxW) { $maxW = $measured }
                }
                foreach ($radio in $radioFlow.Controls) { $radio.Width = $maxW }
                # Simulate flow layout to count rows
                $currentX  = 0
                $currentY  = 0
                $rowHeight = 0
                foreach ($radio in $radioFlow.Controls) {
                    $itemWidth  = $radio.Width + $radio.Margin.Left + $radio.Margin.Right
                    $itemHeight = $radio.Height + $radio.Margin.Top + $radio.Margin.Bottom
                    if ($currentX -gt 0 -and ($currentX + $itemWidth) -gt $radioFlowWidth) {
                        $currentY += $rowHeight
                        $currentX  = 0
                        $rowHeight = 0
                    }
                    $currentX += $itemWidth
                    if ($itemHeight -gt $rowHeight) { $rowHeight = $itemHeight }
                }
                $totalRadioHeight = $currentY + $rowHeight
                if ($totalRadioHeight -lt 30) { $totalRadioHeight = 30 }
                # Limit scroll area to 3 rows maximum
                $maxScrollH     = $maxScrollRows * $radioRowH
                $scrollH        = [Math]::Min($totalRadioHeight, $maxScrollH)
                $scrollPanel.Height = $scrollH
                $radioFlow.Width    = $radioFlowWidth
                $panelHeight = $scrollH + 15
                if ($panelHeight -lt 50) { $panelHeight = 50 }
                $selPanel.Height = $panelHeight
                # Center scroll panel and label vertically within selector
                $scrollPanel.Location = [System.Drawing.Point]::new($scrollPanel.Location.X, [int](($panelHeight - $scrollH) / 2))
                foreach ($ctrl in $selPanel.Controls) {
                    if ($ctrl -is [System.Windows.Forms.Label]) {
                        $ctrl.Location = [System.Drawing.Point]::new($ctrl.Location.X, [int](($panelHeight - $ctrl.Height) / 2))
                    }
                }
            }
            finally { $selPanel.Tag.IsUpdating = $false }
        }
        & $recalcSelectorLayout
        $script:RecalcSelectorLayout = $recalcSelectorLayout
    }
    else {
        $script:SelectedUninstallComputer = if ($computersWithProducts.Count -eq 1) { $computersWithProducts[0] } else { "" }
    }
    # Build panels for selected computer
    if ($script:SelectedUninstallComputer -ne $null) {
        Rebuild-UninstallPanelsForComputer -ComputerKey $script:SelectedUninstallComputer
    }
    $script:UninstallTabCacheVersion = $script:CacheVersion
}

function Rebuild-UninstallPanelsForComputer {
    param([string]$ComputerKey)
    Write-Log "Rebuilding uninstall panels for computer : $ComputerKey"
    # Remove existing product panels (keep selector if exists)
    $controlsToRemove = @()
    foreach ($ctrl in $uninstallFlowPanel_MSICleanupTab.Controls) {
        if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Tag -and $ctrl.Tag.Entry) {
            $controlsToRemove += $ctrl
        }
        elseif ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Controls.Count -gt 0) {
            foreach ($innerCtrl in $ctrl.Controls) {
                if ($innerCtrl -is [System.Windows.Forms.Panel] -and $innerCtrl.Tag -and $innerCtrl.Tag.Entry) {
                    $controlsToRemove += $ctrl
                    break
                }
            }
        }
        elseif ($ctrl -is [System.Windows.Forms.Label] -and $ctrl.Height -eq 1) {
            $controlsToRemove += $ctrl
        }
    }
    foreach ($ctrl in $controlsToRemove) { $uninstallFlowPanel_MSICleanupTab.Controls.Remove($ctrl) }
    $script:UninstallPanelControls = @{}
    if (-not $script:ProductCache.ContainsKey($ComputerKey)) {
        Write-Log "No products cached for computer : $ComputerKey"
        return
    }
    $isRemote = ($ComputerKey -ne "" -and $ComputerKey -ne $env:COMPUTERNAME)
    function Get-CustomUninstallCommand {
        param([string]$UninstallString, [string]$QuietUninstallString, [string]$ModifyPath)
        $commands = @(
            @{ Type = 'UninstallString';      Value = $UninstallString },
            @{ Type = 'QuietUninstallString'; Value = $QuietUninstallString },
            @{ Type = 'ModifyPath';           Value = $ModifyPath }
        )
        foreach ($cmd in $commands) {
            if ([string]::IsNullOrWhiteSpace($cmd.Value)) { continue }
            $cmdValue = $cmd.Value
            # Autodesk AdODIS exception
            if ($cmdValue -like '*C:\Program Files\Autodesk\AdODIS\V1\Installer.exe*') {
                if ($cmdValue     -match '"C:\\Program Files\\Autodesk\\AdODIS\\V1\\Installer\.exe"') { $customCmd = $cmdValue -replace '("C:\\Program Files\\Autodesk\\AdODIS\\V1\\Installer\.exe")', '$1 -q' }
                elseif ($cmdValue -match  'C:\\Program Files\\Autodesk\\AdODIS\\V1\\Installer\.exe')  { $customCmd = $cmdValue -replace  '(C:\\Program Files\\Autodesk\\AdODIS\\V1\\Installer\.exe)',  '"$1" -q' }
                Write-Log "Applied Autodesk AdODIS exception : $customCmd"
                return @{ Command = $customCmd; IsRuleBased = $true }
            }
            # Power BI Desktop exception
            if ($cmdValue -match 'PBIDesktopSetup_x64\.exe"\s+/uninstall\s*$') {
                $customCmd = $cmdValue -replace '(PBIDesktopSetup_x64\.exe"\s+/uninstall)\s*$', '$1 /silent /norestart /log C:\Windows\Temp\PBIDesktopSetup_x64_Uninstall.log'
                Write-Log "Applied Power BI Desktop exception : $customCmd"
                return @{ Command = $customCmd; IsRuleBased = $true }
            }
        }
        return $null
    }
    function Get-DefaultCustomCommand {
        param([string]$UninstallString, [string]$QuietUninstallString, [string]$ModifyPath, [string]$ProductGuid, [string]$DisplayName)
        $sourceCommand = $null
        if (![string]::IsNullOrWhiteSpace($UninstallString))          { $sourceCommand = $UninstallString }
        elseif (![string]::IsNullOrWhiteSpace($QuietUninstallString)) { $sourceCommand = $QuietUninstallString }
        elseif (![string]::IsNullOrWhiteSpace($ModifyPath))           { $sourceCommand = $ModifyPath }
        if ([string]::IsNullOrWhiteSpace($sourceCommand))             { return "" }
        if ($sourceCommand -match '^msiexec') {
            $cleanName = $DisplayName -replace '[^\w\s-]', '' -replace '\s+', '_'
            return "msiexec.exe /X$ProductGuid /qb /norestart /L*v `"C:\Windows\Temp\${cleanName}_${ProductGuid}_Uninstall.log`""
        }
        else { return $sourceCommand }
    }
    function Repair-MsiExecCommand {
        param([string]$Command, [ValidateSet('Uninstall', 'Modify')][string]$CommandType)
        if ([string]::IsNullOrWhiteSpace($Command)) { return $Command }
        $expectedSwitch = switch ($CommandType) { 'Uninstall' { 'X' }; 'Modify' { 'I' } }
        if ($Command -match '(?i)msiexec(\.exe)?\s+/([IX])') {
            $currentSwitch = $matches[2].ToUpper()
            if ($currentSwitch -ne $expectedSwitch) { $Command = $Command -replace '(?i)(msiexec(?:\.exe)?\s+)/[IX]', "`$1/$expectedSwitch" }
        }
        return $Command
    }
    $allEntries           = @{}
    $parentChildRelations = @{}
    $childToParent        = @{}
    $guidToParent         = @{}
    foreach ($guid in $script:ProductCache[$ComputerKey].Keys) {
        $cachedProduct = $script:ProductCache[$ComputerKey][$guid]
        $fullProduct   = $cachedProduct.FullProduct
        if (-not $fullProduct -or -not $fullProduct.UninstallEntries) { continue }
        $parentGuid = $fullProduct.ParentGuid
        if ($parentGuid) { $guidToParent[$guid] = $parentGuid }
        foreach ($entry in $fullProduct.UninstallEntries) {
            if ([string]::IsNullOrWhiteSpace($entry.UninstallString) -and [string]::IsNullOrWhiteSpace($entry.QuietUninstallString) -and [string]::IsNullOrWhiteSpace($entry.ModifyPath)) { continue }
            $entryKey = "$guid|$($entry.RegistryPath)"
            $isMsi    = $entry.UninstallString -match 'msiexec' -or $entry.QuietUninstallString -match 'msiexec'
            $allEntries[$entryKey] = @{
                Entry             = $entry
                ProductGuid       = $guid
                ParentGuid        = $parentGuid
                IsSystemComponent = ($entry.SystemComponent -eq 1)
                IsMsi             = $isMsi
            }
            if ($parentGuid) {
                if (-not $parentChildRelations.ContainsKey($parentGuid)) { $parentChildRelations[$parentGuid] = [System.Collections.Generic.List[string]]::new() }
                $parentChildRelations[$parentGuid].Add($entryKey)
                $childToParent[$entryKey] = $parentGuid
            }
        }
    }
    # Build product groups
    $productGroups = @{}
    $entryToGroup  = @{}
    foreach ($entryKey in $allEntries.Keys) {
        $item     = $allEntries[$entryKey]
        $rootGuid = $item.ProductGuid
        $currentGuid = $item.ParentGuid
        while ($currentGuid) { $rootGuid = $currentGuid; $currentGuid = $guidToParent[$currentGuid] }
        if (-not $productGroups.ContainsKey($rootGuid)) { $productGroups[$rootGuid] = [System.Collections.Generic.List[string]]::new() }
        $productGroups[$rootGuid].Add($entryKey)
        $entryToGroup[$entryKey] = $rootGuid
    }
    # Determine priority entries per group
    $priorityEntries  = @{}
    $groupHasPriority = @{}
    foreach ($rootGuid in $productGroups.Keys) {
        $groupEntries = $productGroups[$rootGuid]
        if ($groupEntries.Count -le 1) { continue }
        $nonSystemComponents = [System.Collections.Generic.List[string]]::new()
        $exeWrappers         = [System.Collections.Generic.List[string]]::new()
        foreach ($entryKey in $groupEntries) {
            $item = $allEntries[$entryKey]
            if (-not $item.IsSystemComponent) { $nonSystemComponents.Add($entryKey) }
            if (-not $item.IsMsi)             { $exeWrappers.Add($entryKey) }
        }
        $priorityKey = $null
        if ($nonSystemComponents.Count -eq 1)     { $priorityKey = $nonSystemComponents[0] }
        elseif ($exeWrappers.Count -eq 1)         { $priorityKey = $exeWrappers[0] }
        if ($priorityKey) {
            $priorityEntries[$priorityKey] = $groupEntries.Count - 1
            $groupHasPriority[$rootGuid]   = $priorityKey
        }
    }
    # Build ordered entries list
    $processedEntries = @{}
    $orderedEntries   = [System.Collections.Generic.List[hashtable]]::new()
    $addEntry = {
        param($EntryKey, $Item, $IsChild, $IsPriority, $ChildCount, $LogPrefix)
        $orderedEntries.Add(@{ EntryKey = $EntryKey; Item = $Item; IsChild = $IsChild; IsPriority = $IsPriority; ChildCount = $ChildCount })
        $processedEntries[$EntryKey] = $true
    }
    foreach ($rootGuid in $productGroups.Keys) {
        $groupEntries = $productGroups[$rootGuid]
        if ($groupHasPriority.ContainsKey($rootGuid)) {
            $priorityKey  = $groupHasPriority[$rootGuid]
            $priorityItem = $allEntries[$priorityKey]
            & $addEntry $priorityKey $priorityItem $false $true $priorityEntries[$priorityKey] "Added priority entry"
            foreach ($entryKey in $groupEntries) {
                if ($processedEntries.ContainsKey($entryKey)) { continue }
                & $addEntry $entryKey $allEntries[$entryKey] $true $false 0 "Added child entry"
            }
        }
        else {
            foreach ($entryKey in $groupEntries) {
                if ($processedEntries.ContainsKey($entryKey)) { continue }
                $item = $allEntries[$entryKey]
                if ($item.IsSystemComponent) { continue }
                & $addEntry $entryKey $item $false $false 0 "Added normal entry"
            }
        }
    }
    foreach ($entryKey in $allEntries.Keys) {
        if ($processedEntries.ContainsKey($entryKey)) { continue }
        $item              = $allEntries[$entryKey]
        if ($item.IsSystemComponent) { continue }
        & $addEntry $entryKey $item $false $false 0 "Added remaining entry"
    }
    foreach ($entryKey in $allEntries.Keys) {
        if ($processedEntries.ContainsKey($entryKey)) { continue }
        $item              = $allEntries[$entryKey]
        $groupRoot         = $entryToGroup[$entryKey]
        $isChildOfPriority = $groupHasPriority.ContainsKey($groupRoot)
        & $addEntry $entryKey $item $isChildOfPriority $false 0 "Added system component entry"
    }
    $baseWidth   = $uninstallFlowPanel_MSICleanupTab.ClientSize.Width - 45
    $childIndent = 35
    # Progress tracking for panel creation
    $totalEntries = $orderedEntries.Count
    $currentEntry = 0
    $script:IsBackgroundOperationRunning = $true
    $script:CancelRequested = $false
    Start-ProgressUI -InitialStatus "Loading uninstall panels (0 / $totalEntries)..."
    foreach ($orderedItem in $orderedEntries) {
        $currentEntry++
        if ($script:CancelRequested) {
            Write-Log "Uninstall panel loading cancelled at $currentEntry / $totalEntries"
            break
        }
        $percent = [int](($currentEntry / [Math]::Max(1, $totalEntries)) * 100)
        Update-ProgressUI_MSICleanupTab -Percent $percent -Status "Loading uninstall panel $currentEntry / $totalEntries..."
        [System.Windows.Forms.Application]::DoEvents()
        $entry       = $orderedItem.Item.Entry
        $productGuid = $orderedItem.Item.ProductGuid
        $isChild     = $orderedItem.IsChild
        $isPriority  = $orderedItem.IsPriority
        $panelKey    = "${ComputerKey}|$($entry.RegistryPath)"
        $panelHeight = 290
        if ($isPriority -and $orderedItem.ChildCount -gt 0) { $panelHeight = 325 }
        $roundedPanels  = New-RoundedPanel -Width $baseWidth -Height $panelHeight -IsChild $isChild -IsRecommended $isPriority -ChildIndent $childIndent
        $productPanel   = $roundedPanels.OuterPanel
        $innerPanel     = $roundedPanels.InnerPanel
        $effectiveWidth = $roundedPanels.EffectiveWidth
        $existingTag = $productPanel.Tag
        $productPanel.Tag = @{
            Radius          = $existingTag.Radius
            BorderColor     = $existingTag.BorderColor
            BackgroundColor = $existingTag.BackgroundColor
            IsChild         = $existingTag.IsChild
            Entry           = $entry
            State           = 'Exists'
            JobId           = $null
            ProductGuid     = $productGuid
            Computer        = $ComputerKey
            IsRemote        = $isRemote
        }
        $productFlowPanel             = gen $innerPanel       "FlowLayoutPanel"                                'Dock=Fill' 'FlowDirection=TopDown' 'WrapContents=$false' 'AutoSize=$false' 'Padding=5 4 5 4'
        $productFlowPanel.BackColor   = if ($productPanel.Tag.BackgroundColor) { $productPanel.Tag.BackgroundColor } else { [System.Drawing.Color]::White }
        $headerPanel                  = gen $productFlowPanel "Panel"            0 0 ($effectiveWidth - 30) 26 'Margin=0 0 0 2' 'BackColor=Transparent'
        $contentLeft                  = 0
        if ($isPriority) {
            $recommendedBadge = gen $headerPanel "Label" "$([char]0x2605) RECOMMENDED" 0 4 115 18 'Font=Segoe UI, 8, Bold' 'ForeColor=White' 'BackColor=76 175 80' 'AutoSize=$false' 'TextAlign=MiddleCenter'
            $contentLeft      = 120
        }
        elseif ($isChild) {
            $componentBadge = gen $headerPanel "Label" "Component" 0 4 75 18 'Font=Segoe UI, 8' 'ForeColor=White' 'BackColor=158 158 158' 'AutoSize=$false' 'TextAlign=MiddleCenter'
            $contentLeft    = 80
        }
        $productName    = if ($entry.DisplayName)    { $entry.DisplayName }            else { "Unknown Product" }
        $productVersion = if ($entry.DisplayVersion) { " - $($entry.DisplayVersion)" } else { "" }
        $titleText      = "$productName$productVersion"
        $titleLabel     = gen $headerPanel "Label" $titleText $contentLeft 3 0 0 'Font=Segoe UI, 10, Bold' 'AutoSize=$true'
        if ($isChild) { $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80) }
        $guidLabel = gen $headerPanel "Label" $productGuid 0 6 0 0 'Font=Consolas, 8' 'ForeColor=Gray' 'AutoSize=$true' 'Anchor=Top, Right'
        $guidLabel.Location = [System.Drawing.Point]::new(($headerPanel.Width - [System.Windows.Forms.TextRenderer]::MeasureText($productGuid, $guidLabel.Font).Width - 10), 6)
        $guidSize  = [System.Windows.Forms.TextRenderer]::MeasureText($productGuid, $guidLabel.Font)
        $guidLeft  = $headerPanel.Width - $guidSize.Width - 10
        $guidLabel.Location = [System.Drawing.Point]::new($guidLeft, 6)
        $headerPanel.Controls.Add($guidLabel)
        # Check registry exists (remote-aware)
        $registryExists = Test-RegistryPathExists -Path $entry.RegistryPath -ComputerName $ComputerKey
        $initialState   = if ($registryExists) { 'Exists' } else { 'Uninstalled' }
        $productPanel.Tag.State = $initialState
        if ($script:UninstallPanelStates.ContainsKey($panelKey)) {
            $savedStateInfo = $script:UninstallPanelStates[$panelKey]
            $savedState     = if ($savedStateInfo -is [hashtable]) { $savedStateInfo.State } else { $savedStateInfo }
            $savedExitCode  = if ($savedStateInfo -is [hashtable]) { $savedStateInfo.ExitCode } else { -1 }
            $savedLogPath   = if ($savedStateInfo -is [hashtable]) { $savedStateInfo.LogPath } else { $null }
            if ($savedState -eq 'Failed' -or $savedState -eq 'RebootPending' -or ($savedState -eq 'Uninstalled' -and !$registryExists)) {
                $productPanel.Tag.State    = $savedState
                $productPanel.Tag.ExitCode = $savedExitCode
                $productPanel.Tag.LogPath  = $savedLogPath
            }
        }
        $stateLabelPanel = gen $productFlowPanel "Panel"            0 0 ($effectiveWidth - 30) 25 'Margin=0 0 0 5' 'BackColor=Transparent'
        $stateTextLabel  = gen $stateLabelPanel  "Label" "State : " 3 3 0 0                       'AutoSize=$true'
        $stateValueLabel = gen $stateLabelPanel  "Label" ""         50 3 0 0                      'AutoSize=$true' 'Tag=StateValueLabel'
        $productPanel.Tag.StateLabelPanel = $stateLabelPanel
        $productPanel.Tag.EffectiveWidth  = $effectiveWidth
        switch ($productPanel.Tag.State) {
            'Exists'        { $stateValueLabel.Text = "Exists";         $stateValueLabel.ForeColor = [System.Drawing.Color]::Blue }
            'Uninstalled'   { $stateValueLabel.Text = "Uninstalled";    $stateValueLabel.ForeColor = [System.Drawing.Color]::Green }
            'RebootPending' { $stateValueLabel.Text = "Reboot Pending"; $stateValueLabel.ForeColor = [System.Drawing.Color]::Orange }
            'Failed'        {
                $exitCodeText              = if ($productPanel.Tag.ExitCode) { $productPanel.Tag.ExitCode } else { "unknown" }
                $stateValueLabel.Text      = "Failed with exit code : $exitCodeText"
                $stateValueLabel.ForeColor = [System.Drawing.Color]::Red
            }
        }
        if ($productPanel.Tag.LogPath) {
            $logComputer = if ($isRemote) { $ComputerKey } else { "" }
            $logExists   = if ($isRemote) { $true } else { [System.IO.File]::Exists($productPanel.Tag.LogPath) }
            if ($logExists) {
                $showLogBtnState     = gen $stateLabelPanel "Button" "Show Log" ($effectiveWidth - 120) 0 80 20
                $showLogBtnState.Tag = @{ LogPath = $productPanel.Tag.LogPath; Computer = $logComputer }
                $showLogBtnState.Add_Click({
                    $tag = $this.Tag
                    Open-LogFile -LogPath $tag.LogPath -ComputerName $tag.Computer
                })
            }
        }
        $customCommandResult = Get-CustomUninstallCommand -UninstallString $entry.UninstallString -QuietUninstallString $entry.QuietUninstallString -ModifyPath $entry.ModifyPath
        $customCommand       = ""
        if ($customCommandResult) { $customCommand = $customCommandResult.Command }
        else { $customCommand = Get-DefaultCustomCommand -UninstallString $entry.UninstallString -QuietUninstallString $entry.QuietUninstallString -ModifyPath $entry.ModifyPath -ProductGuid $entry.ProductId -DisplayName $entry.DisplayName }
        $methods = @(
            @{ Name = 'UninstallString';      Value = (Repair-MsiExecCommand -Command $entry.UninstallString -CommandType 'Uninstall'); IsCustom = $false },
            @{ Name = 'QuietUninstallString'; Value = $entry.QuietUninstallString;                                                      IsCustom = $false },
            @{ Name = 'ModifyPath';           Value = (Repair-MsiExecCommand -Command $entry.ModifyPath      -CommandType 'Modify');    IsCustom = $false },
            @{ Name = 'Custom';               Value = $customCommand;                                                                   IsCustom = $true }
        )
        foreach ($method in $methods) {
            $rowPanel = gen $productFlowPanel "Panel"               0 0 ($effectiveWidth - 30) 50 'Margin=0 2 0 2' 'BackColor=Transparent'
            $execBtn  = gen $rowPanel         "Button" $method.Name 0 10 130 30
            $copyBtn  = gen $rowPanel         "Button" "Copy"       135 10 50 30
            if ([string]::IsNullOrWhiteSpace($method.Value) -and !$method.IsCustom) {
                $lbl = gen $rowPanel "Label" "(Not available)" 195 18 0 0 'AutoSize=$true' 'ForeColor=Gray'
                $execBtn.Enabled = $false
                $copyBtn.Enabled = $false
            }
            else {
                $txt      = gen $rowPanel "TextBox" "" 195 5 ($rowPanel.Width - 205) 40 'Multiline=$true' 'ScrollBars=None'
                $txt.Text = if ($method.Value) { $method.Value } else { "" }
                $txt.Tag  = "CommandTextBox_$($method.Name)"
                if (-not $method.IsCustom) { $txt.ReadOnly = $true }
                Update-TextBoxScrollBars -TextBox $txt
                if ([string]::IsNullOrWhiteSpace($method.Value) -and !$method.IsCustom) {
                    $execBtn.Enabled = $false
                    $copyBtn.Enabled = $false
                }
                else {
                    $execBtn.Tag = @{ ProductPanel = $productPanel; TextBox = $txt; Entry = $entry; PanelKey = $panelKey; Computer = $ComputerKey }
                    $execBtn.Add_Click({
                        $t = $this.Tag
                        if (![string]::IsNullOrWhiteSpace($t.TextBox.Text)) {
                            # Refresh log paths from current textbox content before launching
                            $freshLog = Get-LogPathFromCommandText -Command $t.TextBox.Text
                            if ($freshLog) {
                                $existingPaths = $t.ProductPanel.Tag.CommandLogPaths
                                if (-not $existingPaths.Contains($freshLog)) { $existingPaths.Add($freshLog) }
                            }
                            Test-UninstallPanelLogExists -PanelKey $t.PanelKey
                            Start-UninstallJob -ProductPanel $t.ProductPanel -Command $t.TextBox.Text -Entry $t.Entry -PanelKey $t.PanelKey -ComputerName $t.Computer
                        }
                    })
                    $copyBtn.Tag = $txt
                    $copyBtn.Add_Click({
                        $textBox = $this.Tag
                        if (![string]::IsNullOrWhiteSpace($textBox.Text)) { [System.Windows.Forms.Clipboard]::SetText($textBox.Text) }
                    })
                }
            }
        }
        # Extract log paths from all command textboxes
        $commandLogPaths = [System.Collections.Generic.List[string]]::new()
        foreach ($ctrl in $productFlowPanel.Controls) {
            if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Height -eq 50) {
                foreach ($innerCtrl in $ctrl.Controls) {
                    if ($innerCtrl -is [System.Windows.Forms.TextBox] -and $innerCtrl.Multiline -and ![string]::IsNullOrWhiteSpace($innerCtrl.Text)) {
                        $extractedLog = Get-LogPathFromCommandText -Command $innerCtrl.Text
                        if ($extractedLog -and -not $commandLogPaths.Contains($extractedLog)) { $commandLogPaths.Add($extractedLog) }
                    }
                }
            }
        }
        $productPanel.Tag.CommandLogPaths = $commandLogPaths
        $productPanel.Tag.LogButtonShown  = $false
        if ($isPriority -and $orderedItem.ChildCount -gt 0) {
            $infoPanel = gen $productFlowPanel "Panel" 0 0 ($effectiveWidth - 30) 22 'Margin=0 5 0 0' 'BackColor=232 245 233'
            $infoLabel = gen $infoPanel        "Label" "  $([char]0x25BC) This uninstaller will MAYBE remove $($orderedItem.ChildCount) component(s) listed below" 3 3 0 0 'Font=Segoe UI, 8.5' 'ForeColor=46 125 50' 'AutoSize=$true'
        }
        if ($isChild) {
            $wrapperPanel          = gen $uninstallFlowPanel_MSICleanupTab "Panel" 0 0 $baseWidth ($productPanel.Height + 10) 'Margin=5 5 5 5' 'BackColor=Transparent'
            $indicatorLabel        = gen $wrapperPanel "Label" "$([char]0x221F)" 5 10 0 0 'Font=Consolas, 18' 'AutoSize=$true'
            $productPanel.Margin   = [System.Windows.Forms.Padding]::new(0)
            $productPanel.Location = [System.Drawing.Point]::new(30, 5)
            $wrapperPanel.Controls.Add($productPanel)
        }
        else { $uninstallFlowPanel_MSICleanupTab.Controls.Add($productPanel) }
        $script:UninstallPanelControls[$panelKey] = @{ Panel = $productPanel; StateLabel = $stateValueLabel; InnerPanel = $innerPanel }
    }
    $bottomSpacer = gen $uninstallFlowPanel_MSICleanupTab "Label" 'Height=1' 'Margin=0 0 0 15'
    $uninstallFlowPanel_MSICleanupTab.GetType().GetMethod('OnResize', [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($uninstallFlowPanel_MSICleanupTab, @([System.EventArgs]::Empty))
    # Initial log existence check for all panels
    foreach ($pk in $script:UninstallPanelControls.Keys) {
        Test-UninstallPanelLogExists -PanelKey $pk
    }
    $finalStatus = if ($script:CancelRequested) { "Loading cancelled ($currentEntry / $totalEntries panels)" } else { "Ready - $totalEntries uninstall entries loaded" }
    Stop-ProgressUI -FinalStatus $finalStatus
    Complete-PendingTabSwitch
}

function Start-UninstallJob {
    param(
        [System.Windows.Forms.Panel]$ProductPanel,
        [string]$Command,
        $Entry,
        [string]$PanelKey,
        [string]$ComputerName = ""
    )
    $isRemote   = (-not [string]::IsNullOrWhiteSpace($ComputerName) -and $ComputerName -ne $env:COMPUTERNAME)
    $credential = if ($isRemote) { Get-CredentialFromPanel } else { $null }
    $displayTarget = if ($isRemote) { $ComputerName } else { "local" }
    Write-Log "Starting uninstall job ($displayTarget) for : $($Entry.DisplayName)"
    Write-Log "Uninstall command : $Command"
    function Set-RichTextBoxWithBoldPrefix {
        param([System.Windows.Forms.RichTextBox]$RichTextBox, [string]$Text)
        $colonIndex = $Text.IndexOf(':')
        if ($colonIndex -gt 0) {
            $RichTextBox.Text = $Text
            $RichTextBox.Select(0, $colonIndex + 1)
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font, [System.Drawing.FontStyle]::Bold)
            $RichTextBox.Select(0, 0)
        }
        else { $RichTextBox.Text = $Text }
    }
    $originalControls = @()
    foreach ($ctrl in $ProductPanel.Controls) { $originalControls += $ctrl }
    $ProductPanel.Tag.OriginalControls = $originalControls
    $ProductPanel.Controls.Clear()
    $panelWidth = $ProductPanel.ClientSize.Width
    if ($panelWidth -lt 100) { $panelWidth = 800 }
    $splitContainer = gen $ProductPanel "SplitContainer" 'Dock=Fill' 'Orientation=Vertical' 'SplitterWidth=4' 'Panel1MinSize=150'
    $splitterPos    = [int]($splitContainer.Width * 0.5)
    if ($splitterPos -lt $splitContainer.Panel1MinSize) { $splitterPos = $splitContainer.Panel1MinSize }
    if ($splitterPos -gt ($splitContainer.Width - $splitContainer.Panel2MinSize - $splitContainer.SplitterWidth)) {
        $splitterPos = [Math]::Max($splitContainer.Panel1MinSize, $splitContainer.Width - $splitContainer.Panel2MinSize - $splitContainer.SplitterWidth - 10)
    }
    try { $splitContainer.SplitterDistance = $splitterPos } catch { Write-Log "Could not set SplitterDistance" -Level Warning }
    # Left Panel
    $leftPanel         = gen $splitContainer.Panel1 "Panel" 'Dock=Fill' 'Padding=10 10 10 10'
    $targetText        = if ($isRemote) { "Uninstall processing on $ComputerName..." } else { "Uninstall processing..." }
    $statusLabelProc   = gen $leftPanel             "Label" $targetText 'Font=Segoe UI, 10, Bold' 'ForeColor=DarkOrange' 'AutoSize=$false' 'Dock=Top' 'Height=25' 'Tag=StatusLabel'
    $leftInfoFlowPanel = gen $leftPanel             "FlowLayoutPanel" 'Dock=Fill' 'FlowDirection=TopDown' 'WrapContents=$false' 'AutoSize=$false' 'Padding=0 10 0 0'
    $productName     = if ($Entry.DisplayName)            { $Entry.DisplayName }            else { "Unknown Product" }
    $productVersion  = if ($Entry.DisplayVersion)         { $Entry.DisplayVersion }         else { "N/A" }
    $productGuidText = if ($ProductPanel.Tag.ProductGuid) { $ProductPanel.Tag.ProductGuid } else { $null }
    $displayNameRichTextBox            = gen $leftInfoFlowPanel "RichTextBox" 0 0 500 20 'ReadOnly=$true' 'BorderStyle=None' 'BackColor=[System.Drawing.SystemColors]::Control' 'WordWrap=$true' 'Margin=0 0 0 3'
    $displayNameRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::None
    $displayNameRichTextBox.Add_ContentsResized({ param($s, $e); $s.Height = $e.NewRectangle.Height + 2 })
    $versionRichTextBox            = gen $leftInfoFlowPanel "RichTextBox" 0 0 500 20 'ReadOnly=$true' 'BorderStyle=None' 'BackColor=[System.Drawing.SystemColors]::Control' 'WordWrap=$true' 'Margin=0 0 0 3'
    $versionRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::None
    $versionRichTextBox.Add_ContentsResized({ param($s, $e); $s.Height = $e.NewRectangle.Height + 2 })
    $guidRichTextBox            = gen $leftInfoFlowPanel "RichTextBox" 0 0 500 20 'ReadOnly=$true' 'BorderStyle=None' 'BackColor=[System.Drawing.SystemColors]::Control' 'WordWrap=$true' 'Margin=0 0 0 3' 'Visible=$false'
    $guidRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::None
    $guidRichTextBox.Add_ContentsResized({ param($s, $e); $s.Height = $e.NewRectangle.Height + 2 })
    if ($isRemote) {
        $targetRichTextBox            = gen $leftInfoFlowPanel "RichTextBox" 0 0 500 20 'ReadOnly=$true' 'BorderStyle=None' 'BackColor=[System.Drawing.SystemColors]::Control' 'WordWrap=$true' 'Margin=0 0 0 3'
        $targetRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::None
        $targetRichTextBox.Add_ContentsResized({ param($s, $e); $s.Height = $e.NewRectangle.Height + 2 })
    }
    $showLogDuringUninstall         = gen $leftInfoFlowPanel "Button" "Show Log" 0 0 100 25 'Margin=0 5 0 0' 'Visible=$false' 'Tag=ShowLogDuringUninstall'
    $showLogDuringUninstall.Tag     = @{ LogPath = $null; Computer = if ($isRemote) { $ComputerName } else { "" } }
    $showLogDuringUninstall.Add_Click({
        $t = $this.Tag
        if ($t.LogPath) { Open-LogFile -LogPath $t.LogPath -ComputerName $t.Computer }
    })
    $ProductPanel.Tag.ShowLogDuringUninstall = $showLogDuringUninstall
    $buttonPanel  = gen $leftPanel   "Panel"  'Dock=Bottom' 'Height=40'
    $cancelButton = gen $buttonPanel "Button" "Cancel" 0 5 100 30 'BackColor=IndianRed' 'ForeColor=White' 'FlatStyle=Flat'
    $leftPanel.Controls.SetChildIndex($buttonPanel, 0)
    $leftPanel.Controls.SetChildIndex($leftInfoFlowPanel, 1)
    $leftPanel.Controls.SetChildIndex($statusLabelProc, 2)
    # Right Panel
    $rightPanel    = gen $splitContainer.Panel2 "Panel"           'Dock=Fill' 'Padding=5 5 5 5'
    $infoFlowPanel = gen $rightPanel            "FlowLayoutPanel" 'Dock=Top' 'FlowDirection=TopDown' 'WrapContents=$false' 'AutoSize=$true' 'Padding=0 0 0 10'
    $commandLineRichTextBox            = gen $infoFlowPanel "RichTextBox" "Command line : (parsing...)" 0 0 500 20 'ReadOnly=$true' 'BorderStyle=None' 'BackColor=[System.Drawing.SystemColors]::Control' 'WordWrap=$true' 'Margin=0 3 0 3' 'Tag=CommandLineRichTextBox'
    $commandLineRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::None
    $commandLineRichTextBox.Add_ContentsResized({ param($s, $e); $s.Height = $e.NewRectangle.Height + 2 })
    $parseMethodRichTextBox            = gen $infoFlowPanel "RichTextBox" 0 0 500 20 'ReadOnly=$true' 'BorderStyle=None' 'BackColor=[System.Drawing.SystemColors]::Control' 'WordWrap=$true' 'Margin=0 3 0 3' 'Tag=ParseMethodRichTextBox'
    $parseMethodRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::None
    $parseMethodRichTextBox.Add_ContentsResized({ param($s, $e); $s.Height = $e.NewRectangle.Height + 2 })
    $logPathRichTextBox            = gen $infoFlowPanel "RichTextBox" "" 0 0 500 20 'ReadOnly=$true' 'BorderStyle=None' 'BackColor=[System.Drawing.SystemColors]::Control' 'WordWrap=$true' 'Margin=0 3 0 3' 'Visible=$false' 'Tag=LogPathRichTextBox'
    $logPathRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::None
    $logPathRichTextBox.Add_ContentsResized({ param($s, $e); $s.Height = $e.NewRectangle.Height + 2 })
    $processListView = gen $rightPanel "ListView" 'Dock=Fill' 'View=Details' 'FullRowSelect=$true' 'GridLines=$true' 'Font=Consolas, 9'
    $processListView.Columns.Add("PID", 60)           | Out-Null
    $processListView.Columns.Add("Process Name", 180) | Out-Null
    $processListView.Columns.Add("CPU %", 70)         | Out-Null
    $processListView.Columns.Add("RAM (MB)", 90)      | Out-Null
    $processListView.Columns.Add("Relation", 80)      | Out-Null
    $rightPanel.Controls.SetChildIndex($processListView, 0)
    $rightPanel.Controls.SetChildIndex($infoFlowPanel, 1)
    $ProductPanel.Tag.State                  = 'Running'
    $ProductPanel.Tag.ProcessListView        = $processListView
    $ProductPanel.Tag.CommandLineRichTextBox = $commandLineRichTextBox
    $ProductPanel.Tag.ParseMethodRichTextBox = $parseMethodRichTextBox
    $ProductPanel.Tag.LogPathRichTextBox     = $logPathRichTextBox
    $exePath        = $null
    $arguments      = $null
    $parseMethod    = $null
    $trimmedCommand = $Command.Trim()
    switch -Regex ($trimmedCommand) {
        '^msiexec(\.exe)?\s*(.*)$'                          { $exePath = "msiexec.exe";      $arguments = $matches[2]; $parseMethod = "MsiExec";            break }
        '^"([^"]+)"\s*(.*)$'                                { $exePath = $matches[1];        $arguments = $matches[2]; $parseMethod = "DoubleQuoted";       break }
        "^'([^']+)'\s*(.*)$"                                { $exePath = $matches[1];        $arguments = $matches[2]; $parseMethod = "SingleQuoted";       break }
        '^(.+?\.(exe|msi|bat|cmd|com|vbs|ps1|wsf))\s*(.*)$' { $exePath = $matches[1].Trim(); $arguments = $matches[3]; $parseMethod = "ExtensionDetection"; break }
        '^([^\s]+)\s*(.*)$'                                 { $exePath = $matches[1];        $arguments = $matches[2]; $parseMethod = "FirstToken";         break }
        default                                             { $exePath = $trimmedCommand;    $arguments = "";          $parseMethod = "FullCommand" }
    }
    if ($arguments) { $arguments = $arguments.Trim() }
    $msiLogPath = $null
    if ($parseMethod -eq "MsiExec" -and $arguments -notmatch '/L') {
        $sanitizedName = ($productName -replace '\s+', '_') -replace '[^\w\-_]', ''
        $sanitizedGuid = if ($productGuidText) { $productGuidText.Replace('{', '').Replace('}', '') } else { "NoGuid" }
        $msiLogPath    = "C:\Windows\Temp\$($sanitizedName)_$($sanitizedGuid).log"
        $arguments     = "$arguments /L*v `"$msiLogPath`""
        Write-Log "Added MSI logging : $msiLogPath $(if ($isRemote) { "(on $ComputerName)" } else { '' })"
    }
    Set-RichTextBoxWithBoldPrefix -RichTextBox $displayNameRichTextBox -Text "DisplayName : $productName"
    Set-RichTextBoxWithBoldPrefix -RichTextBox $versionRichTextBox     -Text "Version : $productVersion"
    if ($productGuidText) {
        Set-RichTextBoxWithBoldPrefix -RichTextBox $guidRichTextBox -Text "Product ID : $productGuidText"
        $guidRichTextBox.Visible = $true
    }
    if ($isRemote) {
        Set-RichTextBoxWithBoldPrefix -RichTextBox $targetRichTextBox -Text "Target : $ComputerName (remote)"
    }
    $displayExePath = if ($exePath -match '\s')                     { "`"$exePath`"" }  else { $exePath }
    $displayCommand = if ([string]::IsNullOrWhiteSpace($arguments)) { $displayExePath } else { "$displayExePath $arguments" }
    Set-RichTextBoxWithBoldPrefix -RichTextBox $commandLineRichTextBox -Text "Command line : $displayCommand"
    $modeLabel = if ($isRemote) { "$parseMethod (Remote)" } else { $parseMethod }
    Set-RichTextBoxWithBoldPrefix -RichTextBox $parseMethodRichTextBox -Text "Parse method : $modeLabel"
    if ($msiLogPath) {
        $logLabel = if ($isRemote) { "Log path (on $ComputerName) : $msiLogPath" } else { "Log path : $msiLogPath" }
        Set-RichTextBoxWithBoldPrefix -RichTextBox $logPathRichTextBox -Text $logLabel
        $logPathRichTextBox.Visible = $true
    }
    # Collect all known log paths for monitoring
    $allLogPaths = [System.Collections.Generic.List[string]]::new()
    if ($msiLogPath) { $allLogPaths.Add($msiLogPath) }
    if ($ProductPanel.Tag.CommandLogPaths) {
        foreach ($clp in $ProductPanel.Tag.CommandLogPaths) {
            if ($clp -and -not $allLogPaths.Contains($clp)) { $allLogPaths.Add($clp) }
        }
    }
    $jobSyncHash = [hashtable]::Synchronized(@{
        IsComplete       = $false
        ExitCode         = -1
        Success          = $false
        Error            = $null
        CancelRequested  = $false
        ParseMethod      = $parseMethod
        ExecutablePath   = $exePath
        Arguments        = $arguments
        LogPath          = $msiLogPath
        LogPaths         = @($allLogPaths)
        FoundLogPath     = $null
        StartTime        = Get-Date
        EndTime          = $null
        MainProcessId    = $null
        ProcessTree      = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
        PreviousCpuTimes = [hashtable]::Synchronized(@{})
    })
    $cancelButton.Tag = @{ SyncHash = $jobSyncHash; Panel = $ProductPanel }
    $cancelButton.Add_Click({
        $tag                          = $this.Tag
        $tag.SyncHash.CancelRequested = $true
        $this.Enabled                 = $false
        $this.Text                    = "Cancelling..."
        Write-Log "Cancel requested for uninstall job"
    })
    $runspace = New-ConfiguredRunspace
    $runspace.SessionStateProxy.SetVariable("SyncHash", $jobSyncHash)
    $runspace.SessionStateProxy.SetVariable("ExePath", $exePath)
    $runspace.SessionStateProxy.SetVariable("Arguments", $arguments)
    $runspace.SessionStateProxy.SetVariable("IsRemote", $isRemote)
    $runspace.SessionStateProxy.SetVariable("RemoteComputerName", $ComputerName)
    $runspace.SessionStateProxy.SetVariable("RemoteCredential", $credential)

    $uninstallScriptBlock = {
        # Local process tree monitoring
        function Get-ProcessTreeInfo {
            param([int]$MainProcessId, [hashtable]$PreviousCpuTimes, [datetime]$ProcessStartTime)
            $result          = @()
            $currentCpuTimes = @{}
            try { $allSystemProcs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Select-Object ProcessId, ParentProcessId, Name, CreationDate, WorkingSetSize }
            catch { return @{ Processes = @(); CpuTimes = $PreviousCpuTimes } }
            if (-not $allSystemProcs) { return @{ Processes = @(); CpuTimes = $PreviousCpuTimes } }
            $relevantPids = [System.Collections.Generic.HashSet[int]]::new()
            [void]$relevantPids.Add($MainProcessId)
            function Find-ChildrenRecursively {
                param([int]$ParentId)
                $children = $allSystemProcs | Where-Object { $_.ParentProcessId -eq $ParentId }
                foreach ($child in $children) {
                    if ($relevantPids.Add($child.ProcessId)) { Find-ChildrenRecursively -ParentId $child.ProcessId }
                }
            }
            Find-ChildrenRecursively -ParentId $MainProcessId
            $msiProcs = $allSystemProcs | Where-Object { $_.Name -eq 'msiexec.exe' }
            foreach ($msi in $msiProcs) {
                if ($msi.CreationDate -ge $ProcessStartTime.AddSeconds(-2)) {
                    if ($relevantPids.Add($msi.ProcessId)) { Find-ChildrenRecursively -ParentId $msi.ProcessId }
                }
            }
            $cpuCount = [Environment]::ProcessorCount
            foreach ($procData in $allSystemProcs) {
                if ($relevantPids.Contains($procData.ProcessId)) {
                    try {
                        $sysProc                                 = [System.Diagnostics.Process]::GetProcessById($procData.ProcessId)
                        $currentCpuTime                          = $sysProc.TotalProcessorTime.TotalMilliseconds
                        $currentCpuTimes[$procData.ProcessId]    = $currentCpuTime
                        $cpuPercent                              = 0
                        if ($PreviousCpuTimes -and $PreviousCpuTimes.ContainsKey($procData.ProcessId)) {
                            $prevTime   = $PreviousCpuTimes[$procData.ProcessId]
                            $cpuDelta   = $currentCpuTime - $prevTime
                            $cpuPercent = [math]::Min(100, [math]::Round(($cpuDelta / 1000) * 100 / $cpuCount, 1))
                        }
                        $ramMB    = [math]::Round($procData.WorkingSetSize / 1MB, 2)
                        $relation = "Descendant"
                        if     ($procData.ProcessId -eq $MainProcessId)       { $relation = "Main" }
                        elseif ($procData.Name -eq "msiexec.exe")              { $relation = "MsiExec" }
                        elseif ($procData.ParentProcessId -eq $MainProcessId) { $relation = "Child" }
                        $result += @{ PID = $procData.ProcessId; Name = $procData.Name; CPU = $cpuPercent; RAM = $ramMB; Relation = $relation }
                    } catch { }
                }
            }
            return @{ Processes = $result; CpuTimes = $currentCpuTimes }
        }
        # Combined remote monitor scriptblock : process tree with CPU + log file check
        $remoteCombinedMonitorScript = {
            param([int]$MainPid, [string]$StartTimeIso, [string[]]$LogPathsToCheck, [hashtable]$PrevCpuTimes, [double]$ElapsedSecondsSinceLastCheck)
            $startTime = [datetime]::Parse($StartTimeIso)
            $processes     = @()
            $newCpuTimes   = @{}
            $cpuCount      = [Environment]::ProcessorCount
            if ($cpuCount -lt 1) { $cpuCount = 1 }
            if ($ElapsedSecondsSinceLastCheck -le 0) { $ElapsedSecondsSinceLastCheck = 10 }
            try {
                $allProcs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue
                if ($allProcs) {
                    $relevantPids = [System.Collections.Generic.HashSet[int]]::new()
                    [void]$relevantPids.Add($MainPid)
                    $changed = $true
                    while ($changed) {
                        $changed = $false
                        foreach ($p in $allProcs) {
                            if ($relevantPids.Contains($p.ParentProcessId) -and $relevantPids.Add($p.ProcessId)) { $changed = $true }
                        }
                    }
                    foreach ($p in $allProcs) {
                        if ($p.Name -eq 'msiexec.exe' -and $p.CreationDate -ge $startTime.AddSeconds(-2)) {
                            if ($relevantPids.Add($p.ProcessId)) {
                                $changed = $true
                                while ($changed) {
                                    $changed = $false
                                    foreach ($pp in $allProcs) {
                                        if ($relevantPids.Contains($pp.ParentProcessId) -and $relevantPids.Add($pp.ProcessId)) { $changed = $true }
                                    }
                                }
                            }
                        }
                    }
                    foreach ($p in $allProcs) {
                        if ($relevantPids.Contains($p.ProcessId)) {
                            $ramMB      = [math]::Round($p.WorkingSetSize / 1MB, 2)
                            $cpuPercent = 0
                            try {
                                $sysProc        = [System.Diagnostics.Process]::GetProcessById($p.ProcessId)
                                $currentCpuMs   = $sysProc.TotalProcessorTime.TotalMilliseconds
                                $newCpuTimes[$p.ProcessId] = $currentCpuMs
                                if ($PrevCpuTimes -and $PrevCpuTimes.ContainsKey($p.ProcessId)) {
                                    $cpuDelta   = $currentCpuMs - $PrevCpuTimes[$p.ProcessId]
                                    $cpuPercent = [math]::Min(100, [math]::Round(($cpuDelta / ($ElapsedSecondsSinceLastCheck * 1000)) * 100 / $cpuCount, 1))
                                    if ($cpuPercent -lt 0) { $cpuPercent = 0 }
                                }
                            } catch { }
                            $relation = "Descendant"
                            if     ($p.ProcessId -eq $MainPid)       { $relation = "Main" }
                            elseif ($p.Name -eq "msiexec.exe")        { $relation = "MsiExec" }
                            elseif ($p.ParentProcessId -eq $MainPid) { $relation = "Child" }
                            $processes += @{ PID = $p.ProcessId; Name = $p.Name; CPU = $cpuPercent; RAM = $ramMB; Relation = $relation }
                        }
                    }
                }
            } catch { }
            # Log file existence check
            $foundLog = $null
            if ($LogPathsToCheck) {
                foreach ($lp in $LogPathsToCheck) {
                    if ($lp -and [System.IO.File]::Exists($lp)) { $foundLog = $lp; break }
                }
            }
            return @{ Processes = $processes; FoundLogPath = $foundLog; CpuTimes = $newCpuTimes }
        }
        # Remote process start + wait scriptblock
        $remoteStartProcessScript = {
            param([string]$ExePathR, [string]$ArgumentsR)
            $executableExists = $false
            $resolvedPath     = $ExePathR
            if ([System.IO.File]::Exists($ExePathR)) { $executableExists = $true }
            else {
                $cmd = Get-Command $ExePathR -ErrorAction SilentlyContinue
                if ($cmd) { $executableExists = $true; $resolvedPath = $cmd.Source }
            }
            if (-not $executableExists) { return @{ Error = "Executable not found : $ExePathR"; ExitCode = -1; PID = 0 } }
            $psi                        = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName               = $resolvedPath
            $psi.Arguments              = $ArgumentsR
            $psi.UseShellExecute        = $false
            $psi.CreateNoWindow         = $true
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $workDir = [System.IO.Path]::GetDirectoryName($resolvedPath)
            if ([string]::IsNullOrWhiteSpace($workDir)) { $workDir = $env:TEMP }
            $psi.WorkingDirectory = $workDir
            $proc                     = New-Object System.Diagnostics.Process
            $proc.StartInfo           = $psi
            $proc.EnableRaisingEvents = $true
            $proc.Start() | Out-Null
            $pid = $proc.Id
            $proc.BeginOutputReadLine()
            $proc.BeginErrorReadLine()
            $proc.WaitForExit()
            return @{ ExitCode = $proc.ExitCode; PID = $pid; Error = $null }
        }
        try {
            if ($IsRemote) {
                # =============================================
                # REMOTE EXECUTION PATH
                # =============================================
                $invokeBase = @{ ComputerName = $RemoteComputerName; ErrorAction = 'Stop' }
                if ($RemoteCredential) { $invokeBase.Credential = $RemoteCredential }
                # Start process as remote job
                $jobParams = @{ ComputerName = $RemoteComputerName; ScriptBlock = $remoteStartProcessScript; ArgumentList = @($ExePath, $Arguments); AsJob = $true; ErrorAction = 'Stop' }
                if ($RemoteCredential) { $jobParams.Credential = $RemoteCredential }
                $remoteJob = Invoke-Command @jobParams
                # Wait briefly for process to start, then find PID
                Start-Sleep -Milliseconds 1500
                $exeFileName = [System.IO.Path]::GetFileName($ExePath)
                try {
                    $pidResult = Invoke-Command @invokeBase -ScriptBlock {
                        param([string]$ExeName, [string]$StartTimeIso)
                        $startTime = [datetime]::Parse($StartTimeIso)
                        $procs = Get-CimInstance Win32_Process -Filter "Name='$ExeName'" -ErrorAction SilentlyContinue
                        $recent = $procs | Where-Object { $_.CreationDate -ge $startTime.AddSeconds(-10) } | Sort-Object CreationDate -Descending | Select-Object -First 1
                        if ($recent) { return $recent.ProcessId }
                        return 0
                    } -ArgumentList $exeFileName, $SyncHash.StartTime.ToString('o')
                    if ($pidResult -and $pidResult -gt 0) { $SyncHash.MainProcessId = $pidResult }
                } catch { }
                # Monitor loop (10-second interval for Invoke-Command, 1-second loop for cancel responsiveness)
                $lastRemoteCheckTime = [datetime]::MinValue
                $remoteCheckInterval = [timespan]::FromSeconds(10)
                $logAlreadyFound     = $false
                $remotePrevCpuTimes  = @{}
                while ($remoteJob.State -eq 'Running') {
                    if ($SyncHash.CancelRequested) {
                        if ($SyncHash.MainProcessId -and $SyncHash.MainProcessId -gt 0) {
                            try {
                                Invoke-Command @invokeBase -ScriptBlock {
                                    param([int]$Pid)
                                    function Stop-TreeRecursive {
                                        param([int]$ParentId)
                                        $children = Get-CimInstance Win32_Process -Filter "ParentProcessId=$ParentId" -ErrorAction SilentlyContinue
                                        foreach ($child in $children) { Stop-TreeRecursive -ParentId $child.ProcessId }
                                        try { Stop-Process -Id $ParentId -Force -ErrorAction SilentlyContinue } catch { }
                                    }
                                    Stop-TreeRecursive -ParentId $Pid
                                } -ArgumentList $SyncHash.MainProcessId
                            } catch { }
                        }
                        try { Stop-Job $remoteJob -Force } catch { }
                        break
                    }
                    $now = [datetime]::Now
                    if (($now - $lastRemoteCheckTime) -ge $remoteCheckInterval) {
                        $elapsedSec = if ($lastRemoteCheckTime -eq [datetime]::MinValue) { 10 } else { ($now - $lastRemoteCheckTime).TotalSeconds }
                        $lastRemoteCheckTime = $now
                        if ($SyncHash.MainProcessId -and $SyncHash.MainProcessId -gt 0) {
                            try {
                                $logPathsArg   = if ($logAlreadyFound) { @() } else { @($SyncHash.LogPaths) }
                                $monitorResult = Invoke-Command @invokeBase -ScriptBlock $remoteCombinedMonitorScript -ArgumentList $SyncHash.MainProcessId, $SyncHash.StartTime.ToString('o'), $logPathsArg, $remotePrevCpuTimes, $elapsedSec
                                $SyncHash.ProcessTree.Clear()
                                if ($monitorResult.Processes) { foreach ($p in $monitorResult.Processes) { $SyncHash.ProcessTree.Add($p) | Out-Null } }
                                if ($monitorResult.CpuTimes)  { $remotePrevCpuTimes = $monitorResult.CpuTimes }
                                if (-not $logAlreadyFound -and $monitorResult.FoundLogPath) {
                                    $logAlreadyFound       = $true
                                    $SyncHash.FoundLogPath = $monitorResult.FoundLogPath
                                }
                            } catch { }
                        }
                    }
                    Start-Sleep -Milliseconds 1000
                }
                # Collect result
                if (-not $SyncHash.CancelRequested) {
                    try {
                        $jobResult = Receive-Job $remoteJob -Wait -ErrorAction Stop
                        if ($jobResult.Error) {
                            $SyncHash.Error = $jobResult.Error
                        }
                        else {
                            $SyncHash.ExitCode = $jobResult.ExitCode
                            $SyncHash.Success  = ($jobResult.ExitCode -eq 0 -or $jobResult.ExitCode -eq 3010)
                        }
                    }
                    catch {
                        $SyncHash.Error = "Failed to retrieve remote job result : $($_.Exception.Message)"
                    }
                }
                $SyncHash.EndTime = Get-Date
                try { Remove-Job $remoteJob -Force -ErrorAction SilentlyContinue } catch { }
            }
            else {
                # =============================================
                # LOCAL EXECUTION PATH
                # =============================================
                $executableExists = $false
                $resolvedExePath  = $ExePath
                if ([System.IO.File]::Exists($ExePath)) { $executableExists = $true }
                else {
                    $resolvedCommand = Get-Command $ExePath -ErrorAction SilentlyContinue
                    if ($resolvedCommand) {
                        $executableExists        = $true
                        $resolvedExePath         = $resolvedCommand.Source
                        $SyncHash.ExecutablePath = $resolvedExePath
                    }
                }
                if (-not $executableExists) {
                    $SyncHash.Error   = "Executable not found : $ExePath"
                    $SyncHash.EndTime = Get-Date
                    return
                }
                $processInfo                        = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName               = $resolvedExePath
                $processInfo.Arguments              = $Arguments
                $processInfo.UseShellExecute        = $false
                $processInfo.RedirectStandardOutput = $true
                $processInfo.RedirectStandardError  = $true
                $processInfo.CreateNoWindow         = $true
                $workDir = [System.IO.Path]::GetDirectoryName($resolvedExePath)
                if ([string]::IsNullOrWhiteSpace($workDir)) { $workDir = $env:TEMP }
                $processInfo.WorkingDirectory = $workDir
                $process                      = New-Object System.Diagnostics.Process
                $process.StartInfo            = $processInfo
                $process.EnableRaisingEvents  = $true
                $process.Start() | Out-Null
                $SyncHash.MainProcessId = $process.Id
                $process.BeginOutputReadLine()
                $process.BeginErrorReadLine()
                $localPreviousCpuTimes = @{}
                $lastLocalLogCheck     = [datetime]::MinValue
                $localLogCheckInterval = [timespan]::FromSeconds(10)
                $localLogFound         = $false
                while (-not $process.HasExited) {
                    if ($SyncHash.CancelRequested) {
                        try { $process.Kill(); $process.WaitForExit(5000) } catch { }
                        break
                    }
                    $treeResult            = Get-ProcessTreeInfo -MainProcessId $process.Id -PreviousCpuTimes $localPreviousCpuTimes -ProcessStartTime $SyncHash.StartTime
                    $localPreviousCpuTimes = $treeResult.CpuTimes
                    $SyncHash.ProcessTree.Clear()
                    foreach ($p in $treeResult.Processes) { $SyncHash.ProcessTree.Add($p) | Out-Null }
                    # Periodic log file check (every 10 seconds, stop once found)
                    $nowLocal = [datetime]::Now
                    if (-not $localLogFound -and ($nowLocal - $lastLocalLogCheck) -ge $localLogCheckInterval) {
                        $lastLocalLogCheck = $nowLocal
                        foreach ($lp in $SyncHash.LogPaths) {
                            if ($lp -and [System.IO.File]::Exists($lp)) {
                                $localLogFound           = $true
                                $SyncHash.FoundLogPath   = $lp
                                break
                            }
                        }
                    }
                    Start-Sleep -Milliseconds 1000
                }
                if (-not $SyncHash.CancelRequested) { $process.WaitForExit() }
                $SyncHash.ExitCode = $process.ExitCode
                $SyncHash.Success  = ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010)
                $SyncHash.EndTime  = Get-Date
            }
        }
        catch {
            $SyncHash.Error   = $_.Exception.Message
            $SyncHash.EndTime = Get-Date
        }
        finally { $SyncHash.IsComplete = $true }
    }
    $powerShell          = [powershell]::Create()
    $powerShell.Runspace = $runspace
    $powerShell.AddScript($uninstallScriptBlock) | Out-Null
    $asyncResult            = $powerShell.BeginInvoke()
    $jobId                  = [guid]::NewGuid().ToString()
    $ProductPanel.Tag.JobId = $jobId
    $detectedLogPath = $msiLogPath
    if (-not $detectedLogPath) {
        # Detect any file path ending in .log or .txt from the full command
        if     ($Command -match '"([^"]+\.(?:log|txt))"')  { $detectedLogPath = $matches[1] }
        elseif ($Command -match    '(\S+\.(?:log|txt))\b') { $detectedLogPath = $matches[1] }
    }
    $script:UninstallJobs[$jobId] = @{
        PowerShell      = $powerShell
        Runspace        = $runspace
        AsyncResult     = $asyncResult
        SyncHash        = $jobSyncHash
        Panel           = $ProductPanel
        Command         = $Command
        DisplayName     = $Entry.DisplayName
        Entry           = $Entry
        PanelKey        = $PanelKey
        StartTime       = Get-Date
        ProcessListView = $processListView
        LogPath         = $detectedLogPath
        Computer        = $ComputerName
    }
    Write-Log "Uninstall job created : JobId=$jobId, Product=$($Entry.DisplayName), Target=$displayTarget"
}

function Update-ProcessMonitorDisplay {
    param([string]$JobId)
    if (-not $script:UninstallJobs.ContainsKey($JobId)) { return }
    $job      = $script:UninstallJobs[$JobId]
    $syncHash = $job.SyncHash
    $listView = $job.ProcessListView
    if (-not $listView -or $listView.IsDisposed) { return }
    try {
        $listView.BeginUpdate()
        $listView.Items.Clear()
        $processTree = @($syncHash.ProcessTree)
        foreach ($proc in $processTree) {
            $item = New-Object System.Windows.Forms.ListViewItem($proc.PID.ToString())
            $item.SubItems.Add($proc.Name)
            $item.SubItems.Add($proc.CPU.ToString("F1"))
            $item.SubItems.Add($proc.RAM.ToString("F2"))
            $item.SubItems.Add($proc.Relation)
            switch ($proc.Relation) {
                "Main"       { $item.BackColor = [System.Drawing.Color]::LightBlue }
                "Child"      { $item.BackColor = [System.Drawing.Color]::LightGreen }
                "Descendant" { $item.BackColor = [System.Drawing.Color]::LightYellow }
                "MsiExec"    { $item.BackColor = [System.Drawing.Color]::LightCoral }
            }
            $listView.Items.Add($item)
        }
        $listView.EndUpdate()
    }
    catch { Write-Log "Error updating process monitor : $_" -Level Warning }
}

function Update-UninstallPanelComplete {
    param(
        [System.Windows.Forms.Panel]$ProductPanel,
        [bool]$Success,
        [int]$ExitCode,
        [string]$ErrorMessage,
        $Entry,
        [string]$PanelKey,
        [string]$LogPath,
        [string]$ComputerName = ""
    )
    $isRemote       = (-not [string]::IsNullOrWhiteSpace($ComputerName) -and $ComputerName -ne $env:COMPUTERNAME)
    $registryExists = Test-RegistryPathExists -Path $Entry.RegistryPath -ComputerName $ComputerName
    Write-Log "Uninstall job completed : Product=$($Entry.DisplayName), JobSuccess=$Success, ExitCode=$ExitCode, RegistryExists=$registryExists, Remote=$isRemote"
    if ($ErrorMessage) { Write-Log "Uninstall error details : $ErrorMessage" -Level Error }
    $newState = 'Exists'
    if (!$registryExists) {
        if ($ExitCode -eq 3010) { $newState = 'RebootPending'; Write-Log "Product uninstalled successfully but requires reboot : $($Entry.DisplayName)" }
        else                    { $newState = 'Uninstalled';   Write-Log "Product registry key removed successfully : $($Entry.DisplayName)" }
    }
    elseif ($ExitCode -eq 3010) { $newState = 'RebootPending'; Write-Log "Uninstall requires reboot to complete : $($Entry.DisplayName)" }
    elseif (!$Success)          { $newState = 'Failed';        Write-Log "Uninstall failed for $($Entry.DisplayName) : ExitCode=$ExitCode" -Level Warning }
    else                        { $newState = 'Failed';        Write-Log "Uninstall reported success but registry key still exists : $($Entry.DisplayName)" -Level Warning }
    # Validate log path existence (local or remote via UNC)
    $validLogPath  = $null
    $logCandidates = [System.Collections.Generic.List[string]]::new()
    if ($LogPath) { $logCandidates.Add($LogPath) }
    if ($ProductPanel.Tag.CommandLogPaths) {
        foreach ($clp in $ProductPanel.Tag.CommandLogPaths) {
            if ($clp -and -not $logCandidates.Contains($clp)) { $logCandidates.Add($clp) }
        }
    }
    foreach ($candidate in $logCandidates) {
        if ($isRemote) {
            $uncPath = $candidate -replace '^([A-Za-z]):\\', "\\$ComputerName\`$1`$\"
            try { if ([System.IO.File]::Exists($uncPath)) { $validLogPath = $candidate; break } } catch { }
            if (-not $validLogPath) { $validLogPath = $candidate; break }
        }
        else {
            if ([System.IO.File]::Exists($candidate)) { $validLogPath = $candidate; break }
        }
    }
    if ($validLogPath) { Write-Log "Log file available : $validLogPath $(if ($isRemote) { "(on $ComputerName)" } else { '' })" }
    $ProductPanel.Tag.State                 = $newState
    $ProductPanel.Tag.ExitCode              = $ExitCode
    $ProductPanel.Tag.LogPath               = $validLogPath
    $script:UninstallPanelStates[$PanelKey] = @{ State = $newState; ExitCode = $ExitCode; LogPath = $validLogPath; Computer = $ComputerName }
    $script:CompareTabCacheVersion          = -1
    $ProductPanel.Controls.Clear()
    if ($ProductPanel.Tag.OriginalControls) {
        foreach ($ctrl in $ProductPanel.Tag.OriginalControls) { $ProductPanel.Controls.Add($ctrl) }
    }
    if ($script:UninstallPanelControls.ContainsKey($PanelKey)) {
        $stateLabel = $script:UninstallPanelControls[$PanelKey].StateLabel
        if ($stateLabel) {
            switch ($newState) {
                'Exists'        { $stateLabel.Text = "Exists";                            $stateLabel.ForeColor = [System.Drawing.Color]::Blue }
                'Uninstalled'   { $stateLabel.Text = "Uninstalled";                       $stateLabel.ForeColor = [System.Drawing.Color]::Green }
                'RebootPending' { $stateLabel.Text = "Reboot Pending";                    $stateLabel.ForeColor = [System.Drawing.Color]::Orange }
                'Failed'        { $stateLabel.Text = "Failed with exit code : $ExitCode"; $stateLabel.ForeColor = [System.Drawing.Color]::Red }
            }
            # Add Show Log button dynamically if log is available and button doesn't already exist
            if ($validLogPath -and $stateLabel.Parent) {
                $stateLabelPanel = $stateLabel.Parent
                $existingLogBtn  = $null
                foreach ($ctrl in $stateLabelPanel.Controls) {
                    if ($ctrl -is [System.Windows.Forms.Button] -and $ctrl.Text -eq "Show Log") { $existingLogBtn = $ctrl; break }
                }
                if (-not $existingLogBtn) {
                    $panelWidth       = $stateLabelPanel.Width
                    $showLogBtn       = gen $stateLabelPanel "Button" "Show Log" ($panelWidth - 100) 0 80 20
                    $logComputer      = if ($isRemote) { $ComputerName } else { "" }
                    $showLogBtn.Tag   = @{ LogPath = $validLogPath; Computer = $logComputer }
                    $showLogBtn.Add_Click({
                        $tag = $this.Tag
                        Open-LogFile -LogPath $tag.LogPath -ComputerName $tag.Computer
                    })
                }
            }
        }
    }
    Write-Log "Uninstall state updated : Product=$($Entry.DisplayName), State=$newState"
}

#region Tab4 cache

function Update-FullCacheTab {
    if ($script:FullCacheTabCacheVersion -eq $script:CacheVersion -and $treeViewFullCache.Nodes.Count -gt 0) {
        return
    }
    Write-Log "Updating Full Cache Tab"
    $treeViewFullCache.Nodes.Clear()
    $detailsRichTextBoxFullCache.Clear()
    $script:CurrentProductNodesFullCache = @{}
    $script:ComputerNodesFullCache       = @{}
    $computerKeys    = @($script:ProductCache.Keys)
    $isMultiComputer = ($computerKeys.Count -gt 1) -or ($computerKeys.Count -eq 1 -and $computerKeys[0] -ne "")
    # Suppress cross-tab sync events during tree construction
    $script:SyncInProgress_CrossTab = $true
    $treeViewFullCache.BeginUpdate()
    try {
        foreach ($compKey in $computerKeys) {
            $sortedGuids = @($script:ProductCache[$compKey].Keys | Sort-Object { $script:ProductCache[$compKey][$_].DisplayName })
            foreach ($guid in $sortedGuids) {
                $cachedProduct = $script:ProductCache[$compKey][$guid]
                $fullProduct   = $cachedProduct.FullProduct
                if (-not $fullProduct) { continue }
                # ProductRoot
                Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'ProductRoot' -CategoryData @{
                    DisplayName      = $fullProduct.DisplayName
                    ProductId        = $fullProduct.ProductId
                    Version          = $fullProduct.Version
                    Publisher        = $fullProduct.Publisher
                    InstallLocation  = $fullProduct.InstallLocation
                    InstallSource    = $fullProduct.InstallSource
                    LocalPackage     = $fullProduct.LocalPackage
                    CompressedGuid   = $fullProduct.CompressedGuid
                    UninstallEntries = $fullProduct.UninstallEntries
                    ParentGuid       = $null
                } -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -ComputerNodes $script:ComputerNodesFullCache -IsMultiComputer $isMultiComputer
                # Registry categories (with safe count check to handle deserialized PSObject-wrapped empty arrays)
                if ($fullProduct.UserDataEntries   -and @($fullProduct.UserDataEntries).Count -gt 0)   { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'UserDataEntries'   -CategoryData $fullProduct.UserDataEntries   -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                if ($fullProduct.Dependencies      -and @($fullProduct.Dependencies).Count -gt 0)      { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'Dependencies'      -CategoryData $fullProduct.Dependencies      -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                if ($fullProduct.UpgradeCodes      -and @($fullProduct.UpgradeCodes).Count -gt 0)      { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'UpgradeCodes'      -CategoryData $fullProduct.UpgradeCodes      -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                if ($fullProduct.Features          -and @($fullProduct.Features).Count -gt 0)          { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'Features'          -CategoryData $fullProduct.Features          -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                if ($fullProduct.Components        -and @($fullProduct.Components).Count -gt 0)        { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'Components'        -CategoryData $fullProduct.Components        -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                if ($fullProduct.InstallerProducts -and @($fullProduct.InstallerProducts).Count -gt 0) { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'InstallerProducts' -CategoryData $fullProduct.InstallerProducts -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                # Build filesystem items from InstallerFolders and InstallerFiles with full property propagation
                $registryFolderRefs = @()
                $diskFolders        = @()
                $registryFileRefs   = @()
                $diskFiles          = @()
                if ($fullProduct.InstallerFolders -and @($fullProduct.InstallerFolders).Count -gt 0) {
                    foreach ($folder in $fullProduct.InstallerFolders) {
                        if (-not $folder) { continue }
                        if ($folder.RegistryPath -and $folder.RegistryValue) {
                            $registryFolderRefs += [PSCustomObject]@{
                                FolderPath    = $folder.FolderPath
                                RegistryPath  = $folder.RegistryPath
                                RegistryValue = $folder.RegistryValue
                                Exists        = $folder.Exists
                            }
                        }
                        if ($folder.Exists) {
                            $diskFolders += [PSCustomObject]@{
                                FolderPath    = $folder.FolderPath
                                RegistryPath  = $folder.RegistryPath
                                RegistryValue = $folder.RegistryValue
                                Exists        = $folder.Exists
                            }
                        }
                    }
                }
                if ($fullProduct.InstallerFiles -and @($fullProduct.InstallerFiles).Count -gt 0) {
                    foreach ($file in $fullProduct.InstallerFiles) {
                        if (-not $file) { continue }
                        if ($file.RegistryPath) {
                            $registryFileRefs += [PSCustomObject]@{
                                FilePath     = $file.FilePath
                                RegistryPath = $file.RegistryPath
                                FileSize     = $file.FileSize
                                LastModified = $file.LastModified
                                Type         = $file.Type
                            }
                        }
                        $diskFiles += [PSCustomObject]@{
                            FilePath     = $file.FilePath
                            FileSize     = $file.FileSize
                            LastModified = $file.LastModified
                            Type         = $file.Type
                            RegistryPath = $file.RegistryPath
                        }
                    }
                }
                if ($registryFolderRefs.Count -gt 0) { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'InstallerFolders' -CategoryData $registryFolderRefs -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                if ($registryFileRefs.Count -gt 0)   { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'InstallerFiles'   -CategoryData $registryFileRefs   -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                if ($diskFolders.Count -gt 0)        { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'DiskFolders'      -CategoryData $diskFolders        -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
                if ($diskFiles.Count -gt 0)          { Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'DiskFiles'        -CategoryData $diskFiles          -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer }
            }
        }
        # Handle parent-child relationships
        foreach ($compKey in $computerKeys) {
            $sortedGuids = @($script:ProductCache[$compKey].Keys | Sort-Object { $script:ProductCache[$compKey][$_].DisplayName })
            foreach ($guid in $sortedGuids) {
                $cachedProduct = $script:ProductCache[$compKey][$guid]
                $fullProduct   = $cachedProduct.FullProduct
                if ($fullProduct -and $fullProduct.ParentGuid) {
                    $parentKey = if ($isMultiComputer) { "${compKey}|$($fullProduct.ParentGuid)" } else { $fullProduct.ParentGuid }
                    if ($script:CurrentProductNodesFullCache.ContainsKey($parentKey)) {
                        Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewFullCache -ProductGuid $guid -CategoryName 'MoveProductUnderParent' -CategoryData @{ ParentGuid = $fullProduct.ParentGuid } -CurrentProductNodes $script:CurrentProductNodesFullCache -SuppressUpdate $true -ComputerName $compKey -IsMultiComputer $isMultiComputer
                    }
                }
            }
        }
    }
    finally {
        $treeViewFullCache.EndUpdate()
        $script:SyncInProgress_CrossTab = $false
    }
    # Restore cross-tab sync state after rebuild
    if ($treeViewFullCache.Nodes.Count -gt 0 -and $script:SharedExpandedSyncPaths.Count -gt 0) {
        Restore-TreeViewSyncState -TreeView $treeViewFullCache
    }
    elseif ($treeViewFullCache.Nodes.Count -gt 0) {
        # Scroll to top only when no sync state to restore
        [NativeMethods]::SendMessage($treeViewFullCache.Handle, [NativeMethods]::WM_VSCROLL, [NativeMethods]::SB_TOP, 0)
    }
    $script:FullCacheTabCacheVersion = $script:CacheVersion
    $totalProducts = 0
    foreach ($ck in $script:ProductCache.Keys) { $totalProducts += $script:ProductCache[$ck].Count }
    $itemCount             = Get-TreeNodeCount_MSICleanupTab -TreeView $treeViewFullCache
    $Tab4_statusLabel.Text = "Full Cache : $totalProducts products, $itemCount total items"
    Write-Log "Full Cache Tab updated : $totalProducts products, $itemCount items"
}

#region Tab4 compare

function Update-CompareTab {
    if ($script:CompareTabCacheVersion -eq $script:CacheVersion -and $treeViewCompare.Nodes.Count -gt 0) {
        $itemCount    = Get-TreeNodeCount_MSICleanupTab -TreeView $treeViewCompare
        $productCount = $script:CurrentProductNodesCompare.Count
        $Tab4_statusLabel.Text = if ($itemCount -eq 0) { "No residual items found" } else { "Found $productCount products with $itemCount residual items" }
        return
    }
    if ($script:IsBackgroundOperationRunning) {
        [System.Windows.Forms.MessageBox]::Show("A background operation is already running.", "Operation In Progress", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $treeViewCompare.Nodes.Clear()
    $detailsRichTextBoxCompare.Clear()
    $script:CurrentProductNodesCompare = @{}
    $script:ComputerNodesCompare       = @{}
    $script:CancelRequested            = $false
    $script:IsBackgroundOperationRunning = $true
    Start-ProgressUI -InitialStatus "Comparing cache with current state..."
    $computerKeys    = @($script:ProductCache.Keys)
    $isMultiComputer = ($computerKeys.Count -gt 1) -or ($computerKeys.Count -eq 1 -and $computerKeys[0] -ne "")
    $fastMode        = -not $fastModeCheckBox_MSICleanupTab.Checked
    $script:SyncHash = [hashtable]::Synchronized(@{
        CancelRequested    = $false
        ProgressPercent    = 0
        ProgressStatus     = "Initializing..."
        TreeViewUpdates    = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
        FastModeResults    = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
        IsComplete         = $false
        FinalProductCount  = 0
        TotalUpdatesQueued = 0
        Error              = $null
    })
    # Clone cache for background processing
    $productCacheClone = @{}
    foreach ($compKey in $script:ProductCache.Keys) {
        $productCacheClone[$compKey] = @{}
        foreach ($guid in $script:ProductCache[$compKey].Keys) {
            $productCacheClone[$compKey][$guid] = $script:ProductCache[$compKey][$guid]
        }
    }
    # Prepare computer display names
    $computerDisplayNames = @{}
    foreach ($compKey in $computerKeys) {
        $computerDisplayNames[$compKey] = Get-ComputerNodeLabel -ComputerName $compKey
    }
    $script:BackgroundRunspace = New-ConfiguredRunspace -FunctionNames @('Format-RegistryPath', 'Split-RegistryPath', 'ConvertTo-PowerShellPath', 'Format-FileSize')
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("SyncHash", $script:SyncHash)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("ProductCacheClone", $productCacheClone)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("ComputerKeys", $computerKeys)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("IsMultiComputer", $isMultiComputer)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("ComputerDisplayNames", $computerDisplayNames)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("RemoteCredential", (Get-CredentialFromPanel))
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("FastMode", $fastMode)
    $compareScriptBlock = {
        function Write-Log { param([string]$Message); [void]$SyncHash.TreeViewUpdates.Add(@{ Type = 'Log'; Message = $Message }) }
        # Scriptblock for testing paths (used both inline and remote)
        $testPathsScriptBlock = {
            param($paths)
            $results = @{}
            foreach ($item in $paths) {
                $path   = $item.Path
                $type   = $item.Type
                $regVal = $item.RegistryValue
                $exists = $false
                switch ($type) {
                    'Registry' {
                        try {
                            $psPath = $path
                            $root   = $null
                            if     ($psPath -match '^HKLM\\') { $root = [Microsoft.Win32.Registry]::LocalMachine; $subPath = $psPath.Substring(5) }
                            elseif ($psPath -match '^HKCU\\') { $root = [Microsoft.Win32.Registry]::CurrentUser;  $subPath = $psPath.Substring(5) }
                            elseif ($psPath -match '^HKCR\\') { $root = [Microsoft.Win32.Registry]::ClassesRoot;  $subPath = $psPath.Substring(5) }
                            elseif ($psPath -match '^HKU\\')  { $root = [Microsoft.Win32.Registry]::Users;        $subPath = $psPath.Substring(4) }
                            if ($root) {
                                $key = $root.OpenSubKey($subPath, $false)
                                if ($key) { $key.Close(); $exists = $true }
                            }
                        } catch { }
                    }
                    'RegistryValue' {
                        try {
                            $psPath = $path
                            $root   = $null
                            if     ($psPath -match '^HKLM\\') { $root = [Microsoft.Win32.Registry]::LocalMachine; $subPath = $psPath.Substring(5) }
                            elseif ($psPath -match '^HKCU\\') { $root = [Microsoft.Win32.Registry]::CurrentUser;  $subPath = $psPath.Substring(5) }
                            elseif ($psPath -match '^HKCR\\') { $root = [Microsoft.Win32.Registry]::ClassesRoot;  $subPath = $psPath.Substring(5) }
                            elseif ($psPath -match '^HKU\\')  { $root = [Microsoft.Win32.Registry]::Users;        $subPath = $psPath.Substring(4) }
                            if ($root) {
                                $key = $root.OpenSubKey($subPath, $false)
                                if ($key) {
                                    if ($regVal) {
                                        $values = $key.GetValueNames()
                                        $exists = $values -contains $regVal
                                    } else {
                                        $exists = $true
                                    }
                                    $key.Close()
                                }
                            }
                        } catch { }
                    }
                    'Folder' { $exists = [System.IO.Directory]::Exists($path) }
                    'File'   { $exists = [System.IO.File]::Exists($path) }
                }
                $results[$path] = $exists
            }
            return $results
        }
        function Test-ElementsBatch {
            param(
                [ValidateSet('Registry', 'RegistryValue', 'Folder', 'File')][string]$TestType,
                [System.Collections.ArrayList]$Elements,
                [string]$ProductName,
                [string]$CategoryName,
                [hashtable]$SyncHash,
                [ref]$ProgressPercent,
                [string]$ComputerName = "",
                [scriptblock]$TestScript
            )
            $existingElements = [System.Collections.ArrayList]::new()
            if ($Elements.Count -eq 0) { return $existingElements }
            $pathsToTest   = [System.Collections.ArrayList]::new()
            $pathToElement = @{}
            foreach ($el in $Elements) {
                $path = if ($el.Path) { $el.Path } else { $null }
                if ($path) {
                    [void]$pathsToTest.Add(@{ Path = $path; Type = $TestType; RegistryValue = $el.RegistryValue })
                    if (-not $pathToElement.ContainsKey($path)) { $pathToElement[$path] = [System.Collections.ArrayList]::new() }
                    [void]$pathToElement[$path].Add($el)
                }
            }
            if ($pathsToTest.Count -eq 0) { return $existingElements }
            $SyncHash.ProgressStatus = "$ProductName : $CategoryName (testing $($pathsToTest.Count) paths)"
            $isRemote      = -not [string]::IsNullOrWhiteSpace($ComputerName) -and $ComputerName -ne $env:COMPUTERNAME
            $existsResults = @{}
            try {
                if ($isRemote) {
                    $invokeParams = @{
                        ComputerName = $ComputerName
                        ScriptBlock  = $TestScript
                        ArgumentList = @(,$pathsToTest)
                        ErrorAction  = 'Stop'
                    }
                    if ($RemoteCredential) { $invokeParams.Credential = $RemoteCredential }
                    $existsResults = Invoke-Command @invokeParams
                } else {
                    $existsResults = & $TestScript $pathsToTest
                }
            } catch {
                [void]$SyncHash.TreeViewUpdates.Add(@{ Type = 'Log'; Message = "Error testing $($pathsToTest.Count) paths on ${ComputerName} : $($_.Exception.Message)" })
                return $existingElements
            }
            foreach ($path in $existsResults.Keys) {
                if ($existsResults[$path] -eq $true) {
                    foreach ($el in $pathToElement[$path]) {
                        [void]$existingElements.Add($el)
                    }
                }
            }
            return $existingElements
        }
        function Test-AllElementsBatch {
            param(
                [System.Collections.ArrayList]$AllPaths,
                [string]$ComputerName,
                [scriptblock]$TestScript
            )
            if ($AllPaths.Count -eq 0) { return @{} }
            $isRemote      = -not [string]::IsNullOrWhiteSpace($ComputerName) -and $ComputerName -ne $env:COMPUTERNAME
            $existsResults = @{}
            try {
                if ($isRemote) {
                    $invokeParams = @{
                        ComputerName = $ComputerName
                        ScriptBlock  = $TestScript
                        ArgumentList = @(,$AllPaths)
                        ErrorAction  = 'Stop'
                    }
                    if ($RemoteCredential) { $invokeParams.Credential = $RemoteCredential }
                    $existsResults = Invoke-Command @invokeParams
                } else {
                    $existsResults = & $TestScript $AllPaths
                }
            } catch {
                Write-Log "Error testing $($AllPaths.Count) paths on ${ComputerName} : $($_.Exception.Message)"
            }
            return $existsResults
        }
        try {
            $totalProducts = 0
            foreach ($ck in $ComputerKeys) { $totalProducts += $ProductCacheClone[$ck].Count }
            if ($totalProducts -eq 0) { $totalProducts = 1 }
            $productIndex              = 0
            $productsWithResiduals     = 0
            $processedProductsForParent = @{}
            $categoryDefs = @(
                @{ Name = 'Uninstall entries';  Key = 'Uninstall';     TestFunc = 'Registry';      CategoryType = 'CompareUninstallEntries'; IsUninstall = $true;  Weight = 5;  Unwrap = $false }
                @{ Name = 'UserData';           Key = 'UserData';      TestFunc = 'Registry';      CategoryType = 'UserDataEntries';         Pattern = 'UserData.*Products'; Weight = 10; Unwrap = $true }
                @{ Name = 'Dependencies';       Key = 'Dependency';    TestFunc = 'Registry';      CategoryType = 'Dependencies';            Pattern = 'Dependencies';       Weight = 5;  Unwrap = $true }
                @{ Name = 'Upgrade codes';      Key = 'UpgradeCode';   TestFunc = 'Registry';      CategoryType = 'UpgradeCodes';            Pattern = 'UpgradeCodes';       Weight = 5;  Unwrap = $true }
                @{ Name = 'Features';           Key = 'Feature';       TestFunc = 'Registry';      CategoryType = 'Features';                Pattern = '\\Features\\';       Weight = 5;  Unwrap = $true }
                @{ Name = 'Components';         Key = 'Component';     TestFunc = 'Registry';      CategoryType = 'Components';              Pattern = '\\Components\\';     Weight = 20; Unwrap = $true }
                @{ Name = 'Installer products'; Key = 'Product';       TestFunc = 'Registry';      CategoryType = 'InstallerProducts';       Pattern = 'Installer\\Products'; Weight = 5; Unwrap = $true }
                @{ Name = 'Registry values';    Key = 'RegistryValue'; TestFunc = 'RegistryValue'; CategoryType = 'CompareRegistryValues';   Weight = 15; Unwrap = $false }
                @{ Name = 'Disk folders';       Key = 'Folder';        TestFunc = 'Folder';        CategoryType = 'CompareDiskFolders';      Weight = 15; Unwrap = $false }
                @{ Name = 'Disk files';         Key = 'File';          TestFunc = 'File';          CategoryType = 'CompareDiskFiles';        Weight = 15; Unwrap = $false }
            )
            $totalWeight = 0
            foreach ($def in $categoryDefs) { $totalWeight += $def.Weight }
            if ($totalWeight -eq 0) { $totalWeight = 1 }
            Write-Log "Starting comparison of $totalProducts cached products across $($ComputerKeys.Count) computer(s) $(if ($FastMode) { '(fast mode)' } else { '' })"
            foreach ($compKey in $ComputerKeys) {
                if ($SyncHash.CancelRequested) { break }
                $compDisplayName = $ComputerDisplayNames[$compKey]
                $isRemote        = -not [string]::IsNullOrWhiteSpace($compKey) -and $compKey -ne $env:COMPUTERNAME
                $useFastMode     = $FastMode
                Write-Log "Processing computer : $compDisplayName $(if ($useFastMode) { '(fast mode)' } else { '' })"
                if ($useFastMode) {
                    # FAST MODE : Collect all paths from all products, test in single Invoke-Command
                    $SyncHash.ProgressStatus = "[$compDisplayName] Collecting all paths..."
                    # Build master path list with product/category tracking
                    $masterPathList  = [System.Collections.ArrayList]::new()
                    $pathToProducts  = @{}
                    $productDataList = @{}
                    $sortedGuids = @($ProductCacheClone[$compKey].Keys | Sort-Object { $ProductCacheClone[$compKey][$_].DisplayName })
                    foreach ($guid in $sortedGuids) {
                        $cachedProduct = $ProductCacheClone[$compKey][$guid]
                        $productName   = if ($cachedProduct.DisplayName) { $cachedProduct.DisplayName } else { $guid }
                        $fullProduct   = $cachedProduct.FullProduct
                        $parentGuid    = if ($fullProduct) { $fullProduct.ParentGuid } else { $null }
                        $elementBuckets = @{}
                        foreach ($def in $categoryDefs) { $elementBuckets[$def.Key] = [System.Collections.ArrayList]::new() }
                        foreach ($element in $cachedProduct.Elements) {
                            $path = if ($element.Path) { $element.Path } else { "" }
                            switch ($element.Type) {
                                'UninstallEntry' { [void]$elementBuckets['Uninstall'].Add($element) }
                                'RegistryValue'  { [void]$elementBuckets['RegistryValue'].Add($element) }
                                'Folder'         { [void]$elementBuckets['Folder'].Add($element) }
                                'File'           { [void]$elementBuckets['File'].Add($element) }
                                'Registry' {
                                    foreach ($def in $categoryDefs) {
                                        if ($def.Pattern -and $path -match $def.Pattern) { [void]$elementBuckets[$def.Key].Add($element); break }
                                    }
                                }
                            }
                        }
                        $productDataList[$guid] = @{
                            CachedProduct  = $cachedProduct
                            ElementBuckets = $elementBuckets
                            ParentGuid     = $parentGuid
                        }
                        foreach ($def in $categoryDefs) {
                            $elements = $elementBuckets[$def.Key]
                            foreach ($el in $elements) {
                                $path = if ($el.Path) { $el.Path } else { $null }
                                if ($path) {
                                    $pathKey = "$path|$($el.RegistryValue)"
                                    if (-not $pathToProducts.ContainsKey($pathKey)) {
                                        $pathToProducts[$pathKey] = [System.Collections.ArrayList]::new()
                                        [void]$masterPathList.Add(@{ Path = $path; Type = $def.TestFunc; RegistryValue = $el.RegistryValue })
                                    }
                                    [void]$pathToProducts[$pathKey].Add(@{ Guid = $guid; Element = $el; CategoryDef = $def })
                                }
                            }
                        }
                    }
                    Write-Log "[$compDisplayName] Testing $($masterPathList.Count) paths in single call..."
                    $SyncHash.ProgressStatus = "[$compDisplayName] Testing $($masterPathList.Count) paths..."
                    $existsResults = Test-AllElementsBatch -AllPaths $masterPathList -ComputerName $compKey -TestScript $testPathsScriptBlock
                    $SyncHash.ProgressStatus = "[$compDisplayName] Processing results..."
                    $productResiduals = @{}
                    foreach ($pathKey in $pathToProducts.Keys) {
                        $pathParts = $pathKey -split '\|', 2
                        $path      = $pathParts[0]
                        if ($existsResults.ContainsKey($path) -and $existsResults[$path] -eq $true) {
                            foreach ($ref in $pathToProducts[$pathKey]) {
                                $guid = $ref.Guid
                                $def  = $ref.CategoryDef
                                if (-not $productResiduals.ContainsKey($guid)) { $productResiduals[$guid] = @{} }
                                if (-not $productResiduals[$guid].ContainsKey($def.Key)) { $productResiduals[$guid][$def.Key] = [System.Collections.ArrayList]::new() }
                                [void]$productResiduals[$guid][$def.Key].Add($ref.Element)
                            }
                        }
                    }
                    # Store results for bulk processing in OnComplete
                    $sortedResidualGuids = @($productResiduals.Keys | Sort-Object { $productDataList[$_].CachedProduct.DisplayName })
                    foreach ($guid in $sortedResidualGuids) {
                        $productIndex++
                        $cachedProduct         = $productDataList[$guid].CachedProduct
                        $parentGuid            = $productDataList[$guid].ParentGuid
                        $residualCategories    = $productResiduals[$guid]
                        $hasUninstallResiduals = $residualCategories.ContainsKey('Uninstall') -and $residualCategories['Uninstall'].Count -gt 0
                        $productsWithResiduals++
                        [void]$SyncHash.FastModeResults.Add(@{
                            Type     = 'CompareProductRoot'
                            Guid     = $guid
                            Computer = $compKey
                            Data     = @{
                                DisplayName         = $cachedProduct.DisplayName
                                ProductId           = $cachedProduct.ProductId
                                Version             = $cachedProduct.FullProduct.Version
                                Publisher           = $cachedProduct.FullProduct.Publisher
                                InstallLocation     = $cachedProduct.FullProduct.InstallLocation
                                InstallSource       = $cachedProduct.FullProduct.InstallSource
                                LocalPackage        = $cachedProduct.FullProduct.LocalPackage
                                CompressedGuid      = $cachedProduct.FullProduct.CompressedGuid
                                ParentGuid          = $null
                                NoUninstallResidues = -not $hasUninstallResiduals
                            }
                        })
                        $productKey = if ($IsMultiComputer) { "${compKey}|${guid}" } else { $guid }
                        $processedProductsForParent[$productKey] = $true
                        foreach ($def in $categoryDefs) {
                            if ($residualCategories.ContainsKey($def.Key) -and $residualCategories[$def.Key].Count -gt 0) {
                                $dataToSend = if ($def.Unwrap) { @($residualCategories[$def.Key] | ForEach-Object { $_.Data }) } else { $residualCategories[$def.Key] }
                                [void]$SyncHash.FastModeResults.Add(@{ Type = $def.CategoryType; Guid = $guid; Computer = $compKey; Data = $dataToSend })
                            }
                        }
                        if (-not $hasUninstallResiduals) {
                            [void]$SyncHash.FastModeResults.Add(@{ Type = 'UpdateProductColor'; Guid = $guid; Computer = $compKey; NoUninstallResidues = $true })
                        }
                    }
                    # Second pass : parent-child nesting
                    foreach ($guid in $sortedResidualGuids) {
                        $parentGuid = $productDataList[$guid].ParentGuid
                        if ($parentGuid -and $productResiduals.ContainsKey($parentGuid)) {
                            [void]$SyncHash.FastModeResults.Add(@{ Type = 'MoveProductUnderParent'; Guid = $guid; Computer = $compKey; Data = @{ ParentGuid = $parentGuid } })
                        }
                    }
                    $SyncHash.ProgressPercent = [int][Math]::Min(99, (($ComputerKeys.IndexOf($compKey) + 1) / $ComputerKeys.Count) * 100)
                }
                else {
                    # NORMAL MODE : Per-category Invoke-Command
                    $sortedGuids = @($ProductCacheClone[$compKey].Keys | Sort-Object { $ProductCacheClone[$compKey][$_].DisplayName })
                    foreach ($guid in $sortedGuids) {
                        if ($SyncHash.CancelRequested) { break }
                        $cachedProduct = $ProductCacheClone[$compKey][$guid]
                        $productIndex++
                        $productName = if ($cachedProduct.DisplayName) { $cachedProduct.DisplayName } else { $guid }
                        $shortName   = if ($productName.Length -gt 40) { $productName.Substring(0, 37) + "..." } else { $productName }
                        if ($IsMultiComputer) { $shortName = "[$compDisplayName] $shortName" }
                        Write-Log "[$productIndex/$totalProducts] Processing : $productName"
                        $safeProductCount    = [Math]::Max(1, $totalProducts)
                        $productBasePercent  = 5 + (($productIndex - 1) / $safeProductCount) * 90
                        $productPercentRange = 90 / $safeProductCount
                        # Build element buckets
                        $elementBuckets = @{}
                        foreach ($def in $categoryDefs) { $elementBuckets[$def.Key] = [System.Collections.ArrayList]::new() }
                        foreach ($element in $cachedProduct.Elements) {
                            $path = if ($element.Path) { $element.Path } else { "" }
                            switch ($element.Type) {
                                'UninstallEntry' { [void]$elementBuckets['Uninstall'].Add($element) }
                                'RegistryValue'  { [void]$elementBuckets['RegistryValue'].Add($element) }
                                'Folder'         { [void]$elementBuckets['Folder'].Add($element) }
                                'File'           { [void]$elementBuckets['File'].Add($element) }
                                'Registry' {
                                    foreach ($def in $categoryDefs) {
                                        if ($def.Pattern -and $path -match $def.Pattern) { [void]$elementBuckets[$def.Key].Add($element); break }
                                    }
                                }
                            }
                        }
                        $productHasResiduals   = $false
                        $hasUninstallResiduals = $false
                        $fullProduct           = $cachedProduct.FullProduct
                        $parentGuid            = if ($fullProduct) { $fullProduct.ParentGuid } else { $null }
                        $cumulativeWeight      = 0
                        foreach ($def in $categoryDefs) {
                            if ($SyncHash.CancelRequested) { break }
                            $elements                 = $elementBuckets[$def.Key]
                            $safeWeight               = [Math]::Max(1, $totalWeight)
                            $categoryPercent          = $productBasePercent + ($cumulativeWeight / $safeWeight) * $productPercentRange
                            $SyncHash.ProgressPercent = [int][Math]::Min(99, $categoryPercent)
                            $SyncHash.ProgressStatus  = "[$productIndex/$totalProducts] $shortName : $($def.Name)..."
                            $testTypeMap = @{ 'Registry' = 'Registry'; 'RegistryValue' = 'RegistryValue'; 'Folder' = 'Folder'; 'File' = 'File' }
                            $residuals   = Test-ElementsBatch -TestType $testTypeMap[$def.TestFunc] -Elements $elements -ProductName $shortName -CategoryName $def.Name -SyncHash $SyncHash -ProgressPercent ([ref]$categoryPercent) -ComputerName $compKey -TestScript $testPathsScriptBlock
                            $cumulativeWeight += $def.Weight
                            if ($residuals.Count -gt 0) {
                                if (-not $productHasResiduals) {
                                    $productHasResiduals   = $true
                                    $productsWithResiduals++
                                    $noUninstallRes        = -not ($def.IsUninstall -eq $true)
                                    [void]$SyncHash.TreeViewUpdates.Add(@{
                                        Type     = 'CompareProductRoot'
                                        Guid     = $guid
                                        Computer = $compKey
                                        Data     = @{
                                            DisplayName         = $cachedProduct.DisplayName
                                            ProductId           = $cachedProduct.ProductId
                                            Version             = $cachedProduct.FullProduct.Version
                                            Publisher           = $cachedProduct.FullProduct.Publisher
                                            InstallLocation     = $cachedProduct.FullProduct.InstallLocation
                                            InstallSource       = $cachedProduct.FullProduct.InstallSource
                                            LocalPackage        = $cachedProduct.FullProduct.LocalPackage
                                            CompressedGuid      = $cachedProduct.FullProduct.CompressedGuid
                                            ParentGuid          = $null
                                            NoUninstallResidues = $noUninstallRes
                                        }
                                    })
                                    $productKey = if ($IsMultiComputer) { "${compKey}|${guid}" } else { $guid }
                                    if ($parentGuid) {
                                        $parentKey = if ($IsMultiComputer) { "${compKey}|${parentGuid}" } else { $parentGuid }
                                        if ($processedProductsForParent.ContainsKey($parentKey)) {
                                            [void]$SyncHash.TreeViewUpdates.Add(@{ Type = 'MoveProductUnderParent'; Guid = $guid; Computer = $compKey; Data = @{ ParentGuid = $parentGuid } })
                                        }
                                    }
                                }
                                if ($def.IsUninstall -eq $true) { $hasUninstallResiduals = $true }
                                $dataToSend = if ($def.Unwrap) { @($residuals | ForEach-Object { $_.Data }) } else { $residuals }
                                [void]$SyncHash.TreeViewUpdates.Add(@{ Type = $def.CategoryType; Guid = $guid; Computer = $compKey; Data = $dataToSend })
                            }
                        }
                        if ($productHasResiduals -and -not $hasUninstallResiduals) {
                            [void]$SyncHash.TreeViewUpdates.Add(@{ Type = 'UpdateProductColor'; Guid = $guid; Computer = $compKey; NoUninstallResidues = $true })
                        }
                        $productKey = if ($IsMultiComputer) { "${compKey}|${guid}" } else { $guid }
                        $processedProductsForParent[$productKey] = $productHasResiduals
                        if ($parentGuid -and $productHasResiduals) {
                            $parentKey = if ($IsMultiComputer) { "${compKey}|${parentGuid}" } else { $parentGuid }
                            if (-not $processedProductsForParent.ContainsKey($parentKey)) {
                                [void]$SyncHash.TreeViewUpdates.Add(@{ Type = 'PendingParentMove'; Guid = $guid; Computer = $compKey; ParentGuid = $parentGuid })
                            }
                        }
                        if (-not $productHasResiduals) { Write-Log "[$productIndex/$totalProducts] No residuals for : $shortName" }
                    }
                }
            }
            $SyncHash.FinalProductCount = $productsWithResiduals
            $SyncHash.ProgressPercent   = 100
            Write-Log "Background scan complete : $productsWithResiduals products with residuals"
        }
        catch {
            $SyncHash.Error = $_.Exception.Message
            Write-Log "Error during comparison : $($_.Exception.Message)"
        }
        finally { $SyncHash.IsComplete = $true }
    }
    $script:BackgroundPowerShell          = [powershell]::Create()
    $script:BackgroundPowerShell.Runspace = $script:BackgroundRunspace
    [void]$script:BackgroundPowerShell.AddScript($compareScriptBlock)
    [void]$script:BackgroundPowerShell.BeginInvoke()
    $script:PendingParentMoves = @{}
    $compareTimer = New-BackgroundOperationTimer -SyncHash $script:SyncHash -AdditionalState @{
        PendingParentMoves = @{}
        IsMultiComputer    = $isMultiComputer
    } -OnUpdate {
        param($update, $state)
        $compKey         = $update.Computer
        $isMultiComp     = $state.IsMultiComputer
        switch ($update.Type) {
            'Log' { Write-Log $update.Message }
            'PendingParentMove' {
                $childGuid  = $update.Guid
                $parentGuid = $update.ParentGuid
                $parentKey  = if ($isMultiComp) { "${compKey}|${parentGuid}" } else { $parentGuid }
                if ($script:CurrentProductNodesCompare.ContainsKey($parentKey)) {
                    Move-ProductNodeUnderParent -TreeView $treeViewCompare -ChildGuid $childGuid -ParentGuid $parentGuid -CurrentProductNodes $script:CurrentProductNodesCompare -ComputerName $compKey
                }
                else {
                    if (-not $state.PendingParentMoves.ContainsKey($parentKey)) { $state.PendingParentMoves[$parentKey] = [System.Collections.Generic.List[hashtable]]::new() }
                    $state.PendingParentMoves[$parentKey].Add(@{ ChildGuid = $childGuid; Computer = $compKey })
                }
            }
            'UpdateProductColor' {
                $productKey = if ($isMultiComp) { "${compKey}|$($update.Guid)" } else { $update.Guid }
                if ($script:CurrentProductNodesCompare.ContainsKey($productKey) -and $update.NoUninstallResidues) {
                    $script:CurrentProductNodesCompare[$productKey].RootNode.ForeColor = [System.Drawing.Color]::DarkOrange
                }
            }
            'MoveProductUnderParent' {
                $parentKey = if ($isMultiComp) { "${compKey}|$($update.Data.ParentGuid)" } else { $update.Data.ParentGuid }
                if ($script:CurrentProductNodesCompare.ContainsKey($parentKey)) {
                    Move-ProductNodeUnderParent -TreeView $treeViewCompare -ChildGuid $update.Guid -ParentGuid $update.Data.ParentGuid -CurrentProductNodes $script:CurrentProductNodesCompare -ComputerName $compKey
                }
            }
            default {
                Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewCompare -ProductGuid $update.Guid -CategoryName $update.Type -CategoryData $update.Data -CurrentProductNodes $script:CurrentProductNodesCompare -SuppressUpdate $true -ComputerName $compKey -ComputerNodes $script:ComputerNodesCompare -IsMultiComputer $isMultiComp
                $productKey = if ($isMultiComp) { "${compKey}|$($update.Guid)" } else { $update.Guid }
                if ($script:CurrentProductNodesCompare.ContainsKey($productKey)) {
                    $script:CurrentProductNodesCompare[$productKey].RootNode.Expand()
                }
                if ($update.Type -eq 'CompareProductRoot' -and $state.PendingParentMoves.ContainsKey($productKey)) {
                    foreach ($pending in $state.PendingParentMoves[$productKey]) {
                        $childKey = if ($isMultiComp) { "$($pending.Computer)|$($pending.ChildGuid)" } else { $pending.ChildGuid }
                        if ($script:CurrentProductNodesCompare.ContainsKey($childKey)) {
                            Move-ProductNodeUnderParent -TreeView $treeViewCompare -ChildGuid $pending.ChildGuid -ParentGuid $update.Guid -CurrentProductNodes $script:CurrentProductNodesCompare -ComputerName $pending.Computer
                        }
                    }
                    $state.PendingParentMoves.Remove($productKey)
                }
            }
        }
    } -OnComplete {
        param($timerSyncHash, $state)
        $state.PendingParentMoves = @{}
        # Suppress cross-tab sync events during tree construction
        $script:SyncInProgress_CrossTab = $true
        try {
            # Process fast mode results in bulk
            if ($timerSyncHash.FastModeResults -and $timerSyncHash.FastModeResults.Count -gt 0) {
                $treeViewCompare.BeginUpdate()
                try {
                    foreach ($update in $timerSyncHash.FastModeResults) {
                        $compKey     = $update.Computer
                        $isMultiComp = $state.IsMultiComputer
                        switch ($update.Type) {
                            'UpdateProductColor' {
                                $productKey = if ($isMultiComp) { "${compKey}|$($update.Guid)" } else { $update.Guid }
                                if ($script:CurrentProductNodesCompare.ContainsKey($productKey) -and $update.NoUninstallResidues) {
                                    $script:CurrentProductNodesCompare[$productKey].RootNode.ForeColor = [System.Drawing.Color]::DarkOrange
                                }
                            }
                            'MoveProductUnderParent' {
                                $parentKey = if ($isMultiComp) { "${compKey}|$($update.Data.ParentGuid)" } else { $update.Data.ParentGuid }
                                if ($script:CurrentProductNodesCompare.ContainsKey($parentKey)) {
                                    Move-ProductNodeUnderParent -TreeView $treeViewCompare -ChildGuid $update.Guid -ParentGuid $update.Data.ParentGuid -CurrentProductNodes $script:CurrentProductNodesCompare -ComputerName $compKey
                                }
                            }
                            default {
                                Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewCompare -ProductGuid $update.Guid -CategoryName $update.Type -CategoryData $update.Data -CurrentProductNodes $script:CurrentProductNodesCompare -SuppressUpdate $true -ComputerName $compKey -ComputerNodes $script:ComputerNodesCompare -IsMultiComputer $isMultiComp
                            }
                        }
                    }
                }
                finally {
                    $treeViewCompare.EndUpdate()
                }
            }
            # Final expansion
            foreach ($compNodeKey in $script:ComputerNodesCompare.Keys)      { $script:ComputerNodesCompare[$compNodeKey].Expand() }
            foreach ($productKey in $script:CurrentProductNodesCompare.Keys) { $script:CurrentProductNodesCompare[$productKey].RootNode.Expand() }
        }
        finally { $script:SyncInProgress_CrossTab = $false }
        # Restore cross-tab sync state after rebuild
        if ($treeViewCompare.Nodes.Count -gt 0 -and $script:SharedExpandedSyncPaths.Count -gt 0) {
            Restore-TreeViewSyncState -TreeView $treeViewCompare
        }
        elseif ($treeViewCompare.Nodes.Count -gt 0) {
            # Scroll to top only when no sync state to restore
            [NativeMethods]::SendMessage($treeViewCompare.Handle, [NativeMethods]::WM_VSCROLL, [NativeMethods]::SB_TOP, 0)
        }
        if ($script:CancelRequested) {
            Write-Log "Comparison cancelled"
            Stop-ProgressUI -FinalStatus "Comparison cancelled"
            $script:CompareTabCacheVersion = -1
        }
        elseif ($timerSyncHash.Error) {
            Write-Log "Comparison error : $($timerSyncHash.Error)" -Level Error
            Stop-ProgressUI -FinalStatus "Comparison error : $($timerSyncHash.Error)"
            $script:CompareTabCacheVersion = -1
        }
        else {
            $script:CompareTabCacheVersion = $script:CacheVersion
            $itemCount    = Get-TreeNodeCount_MSICleanupTab -TreeView $treeViewCompare
            $productCount = $script:CurrentProductNodesCompare.Count
            $hasItems     = $itemCount -gt 0
            $statusText   = if ($hasItems) { "Found $productCount products with $itemCount residual items" } else { "No residual items found" }
            Write-Log "Comparison complete : $statusText"
            Stop-ProgressUI -FinalStatus $statusText
            $cleanButton_MSICleanupTab.Enabled = $hasItems
        }
    }
    $compareTimer.Start()
}

#region Tab4 search

$searchButton_MSICleanupTab.Add_Click({
    # -------------------------------------------------------------------------
    # 1. VALIDATIONS & PRE-CHECKS
    # -------------------------------------------------------------------------
    if ($script:IsBackgroundOperationRunning) {
        [System.Windows.Forms.MessageBox]::Show("A background operation is already running. Use the STOP button to cancel it.", "Operation In Progress", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $tabControl_MSICleanupTab.SelectedIndex = 0
    $searchTitles = @($titleTextBox_MSICleanupTab.Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $searchGuids  = @($guidTextBox_MSICleanupTab.Text -split ';' | ForEach-Object {
        $raw = ($_.Trim()) -replace '[{}]', ''
        if ($raw -ne '' -and $raw -match '^[A-Fa-f0-9]{32}$' -and $raw -notmatch '-') {
            $expanded = Convert-CompressedToGuid $raw
            if ($expanded) { $expanded -replace '[{}]', '' } else { $raw }
        } else { $raw }
    } | Where-Object { $_ -ne '' })
    if ($searchTitles.Count -eq 0 -and $searchGuids.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a product title or GUID to search.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    # -------------------------------------------------------------------------
    # 2. TARGET RESOLUTION & CONNECTIVITY TEST
    # -------------------------------------------------------------------------
    $targetComputers = Get-TargetComputersFromPanel
    $credential      = Get-CredentialFromPanel
    $isRemoteSearch  = ($targetComputers.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($targetComputers[0]))
    if ($isRemoteSearch) {
        Write-Log "Remote search requested for $($targetComputers.Count) computer(s)"
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $searchButton_MSICleanupTab.Enabled = $false
        $searchButton_MSICleanupTab.Text    = "Testing..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $connectionResults  = Test-RemoteConnections -Computers $targetComputers -Credential $credential
            Update-ConnectionStatusDisplay -Results $connectionResults -Credential $credential
            $connectedComputers = @($script:RemoteConnectionResults | Where-Object { $_.Success } | ForEach-Object { $_.Computer })
            if ($connectedComputers.Count -eq 0) {
                Write-Log "No remote computers are accessible" -Level Warning
                [System.Windows.Forms.MessageBox]::Show("No remote computers are accessible.", "Connection Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $targetComputers = $connectedComputers
            Write-Log "Connected to $($targetComputers.Count) remote computer(s)"
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $searchButton_MSICleanupTab.Enabled = $true
            $searchButton_MSICleanupTab.Text    = "Search"
        }
    }
    else { 
        $targetComputers = @("")
    }
    $script:LastSearchComputers = $targetComputers
    $isMultiComputer = ($targetComputers.Count -gt 1) -or ($targetComputers.Count -eq 1 -and $targetComputers[0] -ne "")
    $computerDisplayNames = @{}
    foreach ($comp in $targetComputers) { $computerDisplayNames[$comp] = Get-ComputerNodeLabel -ComputerName $comp }
    # -------------------------------------------------------------------------
    # 3. UI RESET
    # -------------------------------------------------------------------------
    $treeViewSearch_MSICleanupTab.Nodes.Clear()
    $detailsRichTextBoxSearch_MSICleanupTab.Clear()
    $script:CurrentProductNodes  = @{}
    $script:ComputerNodesSearch  = @{}
    $script:CancelRequested      = $false
    $script:IsBackgroundOperationRunning = $true
    Start-ProgressUI -InitialStatus "Initializing unified search..."
    # -------------------------------------------------------------------------
    # 4. PREPARE SYNC HASH
    # -------------------------------------------------------------------------
    $script:SyncHash = [hashtable]::Synchronized(@{
        CancelRequested    = $false
        ProgressPercent    = 0
        ProgressStatus     = "Initializing..."
        TreeViewUpdates    = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
        FastModeResults    = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
        IsComplete         = $false
        TotalProductCount  = 0
        Error              = $null
    })
    $searchParams = @{
        SearchTitles         = $searchTitles
        SearchGuids          = $searchGuids
        ExactMatch           = $exactMatchCheckBox_MSICleanupTab.Checked
        ExactMatchGuid       = $exactMatchGuidCheckBox_MSICleanupTab.Checked
        TargetComputers      = $targetComputers
        IsMultiComputer      = $isMultiComputer
        ComputerDisplayNames = $computerDisplayNames
        Credential           = $credential
        FastMode             = -not $fastModeCheckBox_MSICleanupTab.Checked
    }
    # -------------------------------------------------------------------------
    # 5. REMOTE-EXECUTABLE SCRIPTBLOCKS (Phases)
    # -------------------------------------------------------------------------
    # Phase 1 : Scan Uninstall + UserData -> Returns product groups
    $phase1ScanScriptBlock = {
        param($params)
        $SearchTitles   = @($params.SearchTitles)
        $SearchGuids    = @($params.SearchGuids)
        $ExactMatch     = $params.ExactMatch
        $ExactMatchGuid = $params.ExactMatchGuid
        $hasTitles      = ($SearchTitles.Count -gt 0)
        $hasGuids       = ($SearchGuids.Count -gt 0)
        # Helper functions
        function Convert-GuidToCompressed {
            param([string]$Guid)
            $sb = [System.Text.StringBuilder]::new(32); $j = 0
            for ($i = 0; $i -lt $Guid.Length -and $j -lt 32; $i++) {
                $c = $Guid[$i]
                if (($c -ge '0' -and $c -le '9') -or ($c -ge 'A' -and $c -le 'F') -or ($c -ge 'a' -and $c -le 'f')) { [void]$sb.Append([char]::ToUpperInvariant($c)); $j++ }
            }
            if ($sb.Length -ne 32) { return $null }
            $clean = $sb.ToString(); $result = [System.Text.StringBuilder]::new(32)
            for ($i = 7; $i -ge 0; $i--)      { [void]$result.Append($clean[$i]) }
            for ($i = 11; $i -ge 8; $i--)     { [void]$result.Append($clean[$i]) }
            for ($i = 15; $i -ge 12; $i--)    { [void]$result.Append($clean[$i]) }
            for ($i = 16; $i -lt 32; $i += 2) { [void]$result.Append($clean[$i + 1]); [void]$result.Append($clean[$i]) }
            return $result.ToString()
        }
        function Convert-CompressedToGuid {
            param([string]$Compressed)
            if ($Compressed.Length -ne 32) { return $null }
            $sb = [System.Text.StringBuilder]::new(38); [void]$sb.Append('{')
            for ($i = 7; $i -ge 0; $i--)      { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
            for ($i = 11; $i -ge 8; $i--)     { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
            for ($i = 15; $i -ge 12; $i--)    { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
            for ($i = 16; $i -lt 20; $i += 2) { [void]$sb.Append($Compressed[$i + 1]); [void]$sb.Append($Compressed[$i]) };   [void]$sb.Append('-')
            for ($i = 20; $i -lt 32; $i += 2) { [void]$sb.Append($Compressed[$i + 1]); [void]$sb.Append($Compressed[$i]) };   [void]$sb.Append('}')
            return $sb.ToString().ToUpperInvariant()
        }
        function ConvertTo-NormalizedGuid {
            param([string]$Guid)
            if ([string]::IsNullOrWhiteSpace($Guid)) { return $null }
            try { $parsedGuid = [System.Guid]::Parse($Guid.Trim()); return $parsedGuid.ToString('B').ToUpperInvariant() }
            catch { return $null }
        }

        $LM = [Microsoft.Win32.Registry]::LocalMachine; $CU = [Microsoft.Win32.Registry]::CurrentUser
        $UninstallPaths = @(
            @{ Root = $LM; SubPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';              Prefix = 'HKLM'; Scope = 'Machine'; Arch = '64-bit' }
            @{ Root = $LM; SubPath = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall';  Prefix = 'HKLM'; Scope = 'Machine'; Arch = '32-bit' }
            @{ Root = $CU; SubPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';              Prefix = 'HKCU'; Scope = 'User';    Arch = '64-bit' }
            @{ Root = $CU; SubPath = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall';  Prefix = 'HKCU'; Scope = 'User';    Arch = '32-bit' }
        )

        # Pre-compute dashless and compressed forms for each GUID search term
        $guidDescriptors = [System.Collections.ArrayList]::new()
        foreach ($rawGuid in $SearchGuids) {
            $dashless   = $rawGuid -replace '[{}\-]', ''
            $compressed = Convert-GuidToCompressed "{$($rawGuid -replace '[{}]', '')}"
            [void]$guidDescriptors.Add(@{ Dashless = $dashless; Compressed = $compressed })
        }
        $diagnostics = [System.Collections.ArrayList]::new()
        [void]$diagnostics.Add("Search parameters : Titles=[$($SearchTitles -join '; ')], GUIDs=[$($SearchGuids -join '; ')], ExactMatch=$ExactMatch, ExactMatchGuid=$ExactMatchGuid")
        [void]$diagnostics.Add("GUID descriptors built : $($guidDescriptors.Count)")
        # Scan Uninstall
        $uninstallResults    = [System.Collections.ArrayList]::new()
        $uninstallKeysTotal  = 0
        $uninstallKeysOpened = 0
        foreach ($regPath in $UninstallPaths) {
            $parentKey = $regPath.Root.OpenSubKey($regPath.SubPath, $false)
            if ($null -eq $parentKey) { continue }
            try {
                foreach ($keyName in $parentKey.GetSubKeyNames()) {
                    $uninstallKeysTotal++
                    # Pre-filter by GUID when no title search (title requires reading every key)
                    if ($hasGuids -and -not $hasTitles) {
                        $dashlessKeyName   = ($keyName -replace '[{}\-]', '')
                        $guidPreFilterPass = $false
                        foreach ($desc in $guidDescriptors) {
                            if ($ExactMatchGuid) { if ($dashlessKeyName -eq $desc.Dashless) { $guidPreFilterPass = $true; break } }
                            else                 { if ($dashlessKeyName -like "*$($desc.Dashless)*") { $guidPreFilterPass = $true; break } }
                        }
                        if (-not $guidPreFilterPass) { continue }
                    }
                    $uninstallKeysOpened++
                    $subKey = $parentKey.OpenSubKey($keyName, $false)
                    if ($null -eq $subKey) { continue }
                    try {
                        $displayName = $subKey.GetValue('DisplayName'); $displayVersion = $subKey.GetValue('DisplayVersion')
                        $uninstallString = $subKey.GetValue('UninstallString'); $quietUninstallString = $subKey.GetValue('QuietUninstallString'); $modifyPath = $subKey.GetValue('ModifyPath')
                        $match = $false
                        if ($hasGuids) {
                            $dashlessKeyName = ($keyName -replace '[{}\-]', '')
                            foreach ($desc in $guidDescriptors) {
                                if ($ExactMatchGuid) { if ($dashlessKeyName -eq $desc.Dashless) { $match = $true; break } }
                                else                 { if ($dashlessKeyName -like "*$($desc.Dashless)*") { $match = $true; break } }
                                if (-not $match) {
                                    $escapedGuid = [regex]::Escape($desc.Dashless)
                                    if (($uninstallString -match $escapedGuid) -or ($quietUninstallString -match $escapedGuid) -or ($modifyPath -match $escapedGuid)) { $match = $true; break }
                                }
                            }
                        }
                        if (-not $match -and $hasTitles -and $displayName) {
                            foreach ($title in $SearchTitles) {
                                if ($ExactMatch) { if ($displayName -eq $title) { $match = $true; break } }
                                else             { if ($displayName -like "*$title*") { $match = $true; break } }
                            }
                        }
                        if (!$match) { continue }
                        $productId = $null
                        foreach ($src in @($keyName, $uninstallString, $quietUninstallString, $modifyPath)) { if ($src -match '\{[A-F0-9-]{36}\}') { $productId = $matches[0]; break } }
                        [void]$uninstallResults.Add(@{
                            ProductId = $productId; Scope = $regPath.Scope; Architecture = $regPath.Arch; SystemComponent = $subKey.GetValue('SystemComponent')
                            DisplayName = $displayName; DisplayVersion = $displayVersion; Publisher = $subKey.GetValue('Publisher')
                            UninstallString = $uninstallString; QuietUninstallString = $quietUninstallString; ModifyPath = $modifyPath
                            InstallLocation = $subKey.GetValue('InstallLocation'); InstallSource = $subKey.GetValue('InstallSource')
                            KeyName = $keyName; RegistryPath = "$($regPath.Prefix)\$($regPath.SubPath)\$keyName"
                        })
                    } finally { $subKey.Close() }
                }
            } finally { $parentKey.Close() }
        }
        [void]$diagnostics.Add("Uninstall scan : $uninstallKeysTotal keys scanned, $uninstallKeysOpened passed filter, $($uninstallResults.Count) matched")
        # Scan UserData
        $userDataResults     = [System.Collections.ArrayList]::new()
        $userDataKeysScanned = 0
        try {
            $userDataKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData', $false)
            if ($userDataKey) {
                try {
                    foreach ($sidName in $userDataKey.GetSubKeyNames()) {
                        $productsKey = $userDataKey.OpenSubKey("$sidName\Products", $false)
                        if ($null -eq $productsKey) { continue }
                        try {
                            foreach ($compressedGuid in $productsKey.GetSubKeyNames()) {
                                $installPropsKey = $productsKey.OpenSubKey("$compressedGuid\InstallProperties", $false)
                                if ($null -eq $installPropsKey) { continue }
                                try {
                                    $userDataKeysScanned++
                                    $displayName = $installPropsKey.GetValue('DisplayName'); $uninstallString = $installPropsKey.GetValue('UninstallString')
                                    $match = $false
                                    if ($hasGuids) {
                                        $expandedDashless = (Convert-CompressedToGuid $compressedGuid) -replace '[{}\-]', ''
                                        foreach ($desc in $guidDescriptors) {
                                            if ($ExactMatchGuid) { if ($compressedGuid -eq $desc.Compressed) { $match = $true; break } }
                                            else                 { if ($expandedDashless -like "*$($desc.Dashless)*") { $match = $true; break } }
                                        }
                                    }
                                    if (-not $match -and $hasTitles -and $displayName) {
                                        foreach ($title in $SearchTitles) {
                                            if ($ExactMatch) { if ($displayName -eq $title) { $match = $true; break } }
                                            else             { if ($displayName -like "*$title*") { $match = $true; break } }
                                        }
                                    }
                                    if ($match) {
                                        $productId = if ($uninstallString -match '\{[A-F0-9-]+\}') { $matches[0] } else { $null }
                                        [void]$userDataResults.Add(@{
                                            RegistryPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$sidName\Products\$compressedGuid\InstallProperties"
                                            ProductKeyPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$sidName\Products\$compressedGuid"
                                            CompressedGuid = $compressedGuid; ProductId = $productId; DisplayName = $displayName
                                            DisplayVersion = $installPropsKey.GetValue('DisplayVersion'); UninstallString = $uninstallString
                                            InstallLocation = $installPropsKey.GetValue('InstallLocation'); InstallSource = $installPropsKey.GetValue('InstallSource')
                                            LocalPackage = $installPropsKey.GetValue('LocalPackage'); Publisher = $installPropsKey.GetValue('Publisher'); UserSID = $sidName
                                        })
                                    }
                                } finally { $installPropsKey.Close() }
                            }
                        } finally { $productsKey.Close() }
                    }
                } finally { $userDataKey.Close() }
            }
        } catch { }
        # Group results
        $productGroups = @{}
        foreach ($entry in $uninstallResults) {
            $guid = $entry.ProductId; if (-not $guid) { foreach ($prop in @('UninstallString', 'QuietUninstallString', 'ModifyPath')) { if ($entry[$prop] -match '\{[A-F0-9-]{36}\}') { $guid = $matches[0]; break } } }
            $normalizedGuid = ConvertTo-NormalizedGuid $guid; if (-not $normalizedGuid) { continue }
            if (-not $productGroups.ContainsKey($normalizedGuid)) { $productGroups[$normalizedGuid] = @{ ProductId = $normalizedGuid; ParentGuid = $null; DisplayName = $null; Version = $null; Publisher = $null; InstallLocation = $null; InstallSource = $null; LocalPackage = $null; UninstallEntries = [System.Collections.ArrayList]::new(); UserDataEntries = [System.Collections.ArrayList]::new() } }
            $group = $productGroups[$normalizedGuid]; [void]$group.UninstallEntries.Add($entry)
            if (-not $group.DisplayName) { $group.DisplayName = $entry.DisplayName }; if (-not $group.Version) { $group.Version = $entry.DisplayVersion }; if (-not $group.Publisher) { $group.Publisher = $entry.Publisher }; if (-not $group.InstallLocation) { $group.InstallLocation = $entry.InstallLocation }; if (-not $group.InstallSource) { $group.InstallSource = $entry.InstallSource }
        }
        foreach ($entry in $userDataResults) {
            $guid = Convert-CompressedToGuid $entry.CompressedGuid; $normalizedGuid = ConvertTo-NormalizedGuid $guid; if (-not $normalizedGuid) { continue }
            if (-not $productGroups.ContainsKey($normalizedGuid)) { $productGroups[$normalizedGuid] = @{ ProductId = $normalizedGuid; ParentGuid = $null; DisplayName = $null; Version = $null; Publisher = $null; InstallLocation = $null; InstallSource = $null; LocalPackage = $null; UninstallEntries = [System.Collections.ArrayList]::new(); UserDataEntries = [System.Collections.ArrayList]::new() } }
            $group = $productGroups[$normalizedGuid]; [void]$group.UserDataEntries.Add($entry)
            if (-not $group.DisplayName) { $group.DisplayName = $entry.DisplayName }; if (-not $group.Version) { $group.Version = $entry.DisplayVersion }; if (-not $group.Publisher) { $group.Publisher = $entry.Publisher }; if (-not $group.InstallLocation) { $group.InstallLocation = $entry.InstallLocation }; if (-not $group.InstallSource) { $group.InstallSource = $entry.InstallSource }; if (-not $group.LocalPackage) { $group.LocalPackage = $entry.LocalPackage }
        }
        [void]$diagnostics.Add("UserData scan : $userDataKeysScanned entries scanned, $($userDataResults.Count) matched")
        [void]$diagnostics.Add("Product groups assembled : $($productGroups.Count)")
        return @{ Products = $productGroups; Diagnostics = $diagnostics }
    }

    # Phase 2 : Scan categories for a single product
    $phase2CategoryScriptBlock = {
        param($params)
        $Guid            = $params.Guid
        $CompressedGuid  = $params.CompressedGuid
        $DisplayName     = $params.DisplayName
        $InstallLocation = $params.InstallLocation
        $LocalPackage    = $params.LocalPackage
        $UserDataEntries = $params.UserDataEntries
        $CategoryName    = $params.CategoryName
        function Convert-CompressedToGuid {
            param([string]$Compressed)
            if ($Compressed.Length -ne 32) { return $null }
            $sb = [System.Text.StringBuilder]::new(38); [void]$sb.Append('{')
            for ($i = 7; $i -ge 0; $i--)      { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
            for ($i = 11; $i -ge 8; $i--)     { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
            for ($i = 15; $i -ge 12; $i--)    { [void]$sb.Append($Compressed[$i]) };                                          [void]$sb.Append('-')
            for ($i = 16; $i -lt 20; $i += 2) { [void]$sb.Append($Compressed[$i + 1]); [void]$sb.Append($Compressed[$i]) };   [void]$sb.Append('-')
            for ($i = 20; $i -lt 32; $i += 2) { [void]$sb.Append($Compressed[$i + 1]); [void]$sb.Append($Compressed[$i]) };   [void]$sb.Append('}')
            return $sb.ToString().ToUpperInvariant()
        }
        function ConvertTo-NormalizedGuid {
            param([string]$Guid)
            if ([string]::IsNullOrWhiteSpace($Guid)) { return $null }
            try { $parsedGuid = [System.Guid]::Parse($Guid.Trim()); return $parsedGuid.ToString('B').ToUpperInvariant() }
            catch { return $null }
        }
        function Get-DependencyRootKeyName {
            param([string]$RegistryPath)
            if ([string]::IsNullOrWhiteSpace($RegistryPath)) { return $null }
            $parts = $RegistryPath -split '\\'; $depIndex = -1
            for ($i = 0; $i -lt $parts.Count; $i++) { if ($parts[$i] -eq 'Dependencies') { $depIndex = $i; break } }
            if ($depIndex -ge 0 -and ($depIndex + 1) -lt $parts.Count) { return $parts[$depIndex + 1] }
            return $null
        }
        $LM = [Microsoft.Win32.Registry]::LocalMachine; $CU = [Microsoft.Win32.Registry]::CurrentUser
        $results = [System.Collections.ArrayList]::new()
        $parentGuidFound = $null
        switch ($CategoryName) {
            'Dependencies' {
                $depPaths = @(
                    @{ Root = $LM; SubPath = 'SOFTWARE\Classes\Installer\Dependencies';             Prefix = 'HKLM' }
                    @{ Root = $LM; SubPath = 'SOFTWARE\Wow6432Node\Classes\Installer\Dependencies'; Prefix = 'HKLM' }
                    @{ Root = $CU; SubPath = 'SOFTWARE\Classes\Installer\Dependencies';             Prefix = 'HKCU' }
                )
                $normalizedGuidClean = $Guid -replace '[{}]', ''
                foreach ($basePath in $depPaths) {
                    $parentKey = $basePath.Root.OpenSubKey($basePath.SubPath, $false); if ($null -eq $parentKey) { continue }
                    try {
                        foreach ($depKeyName in $parentKey.GetSubKeyNames()) {
                            $normalizedKeyName = $depKeyName -replace '[{}]', ''
                            if ($normalizedKeyName -eq $normalizedGuidClean) { 
                                $depRootKey = Get-DependencyRootKeyName -RegistryPath "$($basePath.Prefix)\$($basePath.SubPath)\$depKeyName"
                                [void]$results.Add(@{ ParentGuid = $Guid; DependentGuid = $null; DependencyType = 'ParentKey'; DependencyRootKey = $depRootKey; RegistryPath = "$($basePath.Prefix)\$($basePath.SubPath)\$depKeyName" }) 
                            }
                            $dependentsKey = $parentKey.OpenSubKey("$depKeyName\Dependents", $false)
                            if ($dependentsKey) { 
                                try { 
                                    foreach ($dependentKeyName in $dependentsKey.GetSubKeyNames()) { 
                                        $normalizedDepKeyName = $dependentKeyName -replace '[{}]', ''
                                        if ($normalizedDepKeyName -eq $normalizedGuidClean) { 
                                            $depRootKey = Get-DependencyRootKeyName -RegistryPath "$($basePath.Prefix)\$($basePath.SubPath)\$depKeyName"
                                            $parentGuidFromDep = ConvertTo-NormalizedGuid $depRootKey
                                            if ($parentGuidFromDep -and $parentGuidFromDep -ne $Guid) { $parentGuidFound = $parentGuidFromDep }
                                            [void]$results.Add(@{ ParentGuid = $Guid; DependentGuid = $dependentKeyName; DependencyType = 'DependentsChild'; DependencyRootKey = $depRootKey; RegistryPath = "$($basePath.Prefix)\$($basePath.SubPath)\$depKeyName\Dependents\$dependentKeyName" }) 
                                        } 
                                    } 
                                } finally { $dependentsKey.Close() } 
                            }
                        }
                    } finally { $parentKey.Close() }
                }
            }
            'UpgradeCodes' {
                $ucPaths = @(
                    @{ Root = $LM; SubPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes'; Prefix = 'HKLM' }
                    @{ Root = $LM; SubPath = 'SOFTWARE\Classes\Installer\UpgradeCodes';                          Prefix = 'HKLM' }
                )
                foreach ($path in $ucPaths) { 
                    $parentKey = $path.Root.OpenSubKey($path.SubPath, $false); if ($null -eq $parentKey) { continue }
                    try { 
                        foreach ($upgradeCode in $parentKey.GetSubKeyNames()) { 
                            $upgradeKey = $parentKey.OpenSubKey($upgradeCode, $false)
                            if ($upgradeKey) { 
                                try { 
                                    foreach ($valueName in $upgradeKey.GetValueNames()) { 
                                        if ($valueName -eq $CompressedGuid) { [void]$results.Add(@{ CompressedProductGuid = $valueName; UpgradeCode = $upgradeCode; RegistryPath = "$($path.Prefix)\$($path.SubPath)\$upgradeCode" }) } 
                                    } 
                                } finally { $upgradeKey.Close() } 
                            } 
                        } 
                    } finally { $parentKey.Close() } 
                }
            }
            'Features' {
                $fPaths = @( @{ Root = $LM; SubPath = 'SOFTWARE\Classes\Installer\Features'; Prefix = 'HKLM' }; @{ Root = $CU; SubPath = 'SOFTWARE\Classes\Installer\Features'; Prefix = 'HKCU' } )
                foreach ($path in $fPaths) { 
                    $featureKey = $path.Root.OpenSubKey("$($path.SubPath)\$CompressedGuid", $false)
                    if ($featureKey) { $featureKey.Close(); [void]$results.Add(@{ CompressedGuid = $CompressedGuid; RegistryPath = "$($path.Prefix)\$($path.SubPath)\$CompressedGuid" }) } 
                }
            }
            'Components' {
                $cPath = @{ Root = $LM; SubPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components'; Prefix = 'HKLM' }
                $parentKey = $cPath.Root.OpenSubKey($cPath.SubPath, $false); if ($null -eq $parentKey) { break }
                try { 
                    foreach ($componentId in $parentKey.GetSubKeyNames()) { 
                        $componentKey = $parentKey.OpenSubKey($componentId, $false)
                        if ($componentKey) { 
                            try { 
                                foreach ($valueName in $componentKey.GetValueNames()) { 
                                    if ($valueName -eq $CompressedGuid) { [void]$results.Add(@{ CompressedProductGuid = $valueName; ComponentId = $componentId; ComponentPath = $componentKey.GetValue($valueName); RegistryPath = "$($cPath.Prefix)\$($cPath.SubPath)\$componentId" }) } 
                                } 
                            } finally { $componentKey.Close() } 
                        } 
                    } 
                } finally { $parentKey.Close() }
            }
            'InstallerProducts' {
                $pPaths = @( @{ Root = $LM; SubPath = 'SOFTWARE\Classes\Installer\Products'; Prefix = 'HKLM' }; @{ Root = $CU; SubPath = 'SOFTWARE\Classes\Installer\Products'; Prefix = 'HKCU' } )
                foreach ($path in $pPaths) { 
                    $productKey = $path.Root.OpenSubKey("$($path.SubPath)\$CompressedGuid", $false)
                    if ($productKey) { 
                        try { [void]$results.Add(@{ CompressedGuid = $CompressedGuid; PackageCode = $productKey.GetValue('PackageCode'); RegistryPath = "$($path.Prefix)\$($path.SubPath)\$CompressedGuid"; ProductName = $productKey.GetValue('ProductName') }) } 
                        finally { $productKey.Close() } 
                    } 
                }
            }
            'InstallerFolders' {
                $foldersKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders', $false)
                if ($foldersKey) { 
                    try { 
                        foreach ($folderPath in $foldersKey.GetValueNames()) { 
                            $match = ($DisplayName -and $folderPath -like "*$DisplayName*") -or ($InstallLocation -and $folderPath -like "$InstallLocation*")
                            if (!$match -and $LocalPackage) { $parentPath = [System.IO.Path]::GetDirectoryName($LocalPackage); if ($folderPath -eq $parentPath -or $folderPath -eq "$parentPath\") { $match = $true } }
                            if ($match) { 
                                $exists = [System.IO.Directory]::Exists($folderPath)
                                [void]$results.Add(@{ FolderPath = $folderPath; Exists = $exists; RegistryPath = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders'; RegistryValue = $folderPath }) 
                            } 
                        } 
                    } finally { $foldersKey.Close() } 
                }
            }
            'InstallerFiles' {
                $installerPath = "C:\Windows\Installer\$Guid"
                if ([System.IO.Directory]::Exists($installerPath)) { 
                    try { 
                        foreach ($fileInfo in ([System.IO.DirectoryInfo]::new($installerPath)).GetFiles('*', [System.IO.SearchOption]::AllDirectories)) { 
                            [void]$results.Add(@{ FilePath = $fileInfo.FullName; FileSize = $fileInfo.Length; LastModified = $fileInfo.LastWriteTime.ToString('o'); Type = 'InstallerCache'; RegistryPath = $null }) 
                        } 
                    } catch { } 
                }
                if ($LocalPackage -and [System.IO.File]::Exists($LocalPackage)) { 
                    try { 
                        $fileInfo = [System.IO.FileInfo]::new($LocalPackage)
                        $regPath  = $null
                        foreach ($entry in $UserDataEntries) { 
                            if ($entry.LocalPackage -eq $LocalPackage) { $regPath = $entry.RegistryPath; break } 
                        }
                        [void]$results.Add(@{ FilePath = $LocalPackage; FileSize = $fileInfo.Length; LastModified = $fileInfo.LastWriteTime.ToString('o'); Type = 'LocalPackage'; RegistryPath = $regPath }) 
                    } catch { } 
                }
            }
        }
        return @{ Results = $results; ParentGuid = $parentGuidFound }
    }
    # -------------------------------------------------------------------------
    # 6. RUNSPACE CONTROLLER
    # -------------------------------------------------------------------------
    $script:BackgroundRunspace = New-ConfiguredRunspace -FunctionNames @()
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("SyncHash", $script:SyncHash)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("SearchParams", $searchParams)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("Phase1Script", $phase1ScanScriptBlock)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("Phase2Script", $phase2CategoryScriptBlock)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("Phase2ScriptString", $phase2CategoryScriptBlock.ToString())

    $controllerScriptBlock = {
        $Phase1Script = [scriptblock]::Create($Phase1Script.ToString())
        $Phase2Script = [scriptblock]::Create($Phase2Script.ToString())
        function Write-SearchLog { param([string]$Message) ; [void]$SyncHash.TreeViewUpdates.Add(@{ Type = 'Log'; Message = $Message }) }
        function Send-Update     { param($Update)          ; [void]$SyncHash.TreeViewUpdates.Add($Update) }
        function Convert-GuidToCompressed {
            param([string]$Guid)
            $sb = [System.Text.StringBuilder]::new(32); $j = 0
            for ($i = 0; $i -lt $Guid.Length -and $j -lt 32; $i++) {
                $c = $Guid[$i]
                if (($c -ge '0' -and $c -le '9') -or ($c -ge 'A' -and $c -le 'F') -or ($c -ge 'a' -and $c -le 'f')) { [void]$sb.Append([char]::ToUpperInvariant($c)); $j++ }
            }
            if ($sb.Length -ne 32) { return $null }
            $clean = $sb.ToString(); $result = [System.Text.StringBuilder]::new(32)
            for ($i = 7; $i -ge 0; $i--)      { [void]$result.Append($clean[$i]) }
            for ($i = 11; $i -ge 8; $i--)     { [void]$result.Append($clean[$i]) }
            for ($i = 15; $i -ge 12; $i--)    { [void]$result.Append($clean[$i]) }
            for ($i = 16; $i -lt 32; $i += 2) { [void]$result.Append($clean[$i + 1]); [void]$result.Append($clean[$i]) }
            return $result.ToString()
        }
        # Scriptblock for fast mode : processes all products and categories in one remote call
        $phase2AllScriptBlock = {
            param($params)
            $Phase2Code      = [scriptblock]::Create($params.Phase2ScriptString)
            $categories      = @('Dependencies', 'UpgradeCodes', 'Features', 'Components', 'InstallerProducts', 'InstallerFolders', 'InstallerFiles')
            $results         = [System.Collections.ArrayList]::new()
            $productIndex    = 0
            $totalProducts   = $params.ProductList.Count
            foreach ($prodData in $params.ProductList) {
                $productIndex++
                $guid                 = $prodData.Guid
                $compressed           = $prodData.CompressedGuid
                $group                = $prodData.Group
                $allCategoryData      = @{}
                $parentGuidDiscovered = $null
                # Add ProductRoot first
                [void]$results.Add(@{ 
                    Type = 'ProductRoot'; Guid = $guid
                    Data = @{ 
                        CompressedGuid   = $compressed; ProductId = $guid; ParentGuid = $null
                        DisplayName      = $group.DisplayName; Version = $group.Version; Publisher = $group.Publisher
                        InstallLocation  = $group.InstallLocation; InstallSource = $group.InstallSource; LocalPackage = $group.LocalPackage
                        UninstallEntries = @($group.UninstallEntries) 
                    } 
                })
                # Add UserDataEntries if present
                if ($group.UserDataEntries -and $group.UserDataEntries.Count -gt 0) {
                    [void]$results.Add(@{ Type = 'UserDataEntries'; Guid = $guid; Data = @($group.UserDataEntries) })
                }
                foreach ($catName in $categories) {
                    $phase2Params = @{
                        Guid            = $guid
                        CompressedGuid  = $compressed
                        DisplayName     = $group.DisplayName
                        InstallLocation = $group.InstallLocation
                        LocalPackage    = $group.LocalPackage
                        UserDataEntries = $group.UserDataEntries
                        CategoryName    = $catName
                    }
                    try {
                        $catResult = & $Phase2Code $phase2Params
                        if ($catResult -and $catResult.Results -and $catResult.Results.Count -gt 0) {
                            $allCategoryData[$catName] = @($catResult.Results)
                            if ($catName -ne 'InstallerFiles') {
                                [void]$results.Add(@{ Type = $catName; Guid = $guid; Data = @($catResult.Results) })
                            }
                        }
                        if ($catResult -and $catResult.ParentGuid) { $parentGuidDiscovered = $catResult.ParentGuid }
                    } catch { }
                }
                # Post-process files
                $diskFolders = @(); $diskFiles = @(); $registryFiles = @()
                if ($allCategoryData.InstallerFolders) {
                    foreach ($folder in $allCategoryData.InstallerFolders) {
                        if ($folder.Exists) { $diskFolders += @{ FolderPath = $folder.FolderPath } }
                    }
                }
                if ($allCategoryData.InstallerFiles) {
                    foreach ($file in $allCategoryData.InstallerFiles) {
                        if ($file.RegistryPath) { $registryFiles += $file }
                        $diskFiles += $file
                    }
                }
                if ($registryFiles.Count -gt 0) { [void]$results.Add(@{ Type = 'InstallerFiles'; Guid = $guid; Data = @($registryFiles) }) }
                if ($diskFolders.Count -gt 0)   { [void]$results.Add(@{ Type = 'DiskFolders';    Guid = $guid; Data = @($diskFolders) }) }
                if ($diskFiles.Count -gt 0)     { [void]$results.Add(@{ Type = 'DiskFiles';      Guid = $guid; Data = @($diskFiles) }) }
                # CacheProduct
                [void]$results.Add(@{ 
                    Type = 'CacheProduct'; Guid = $guid
                    Data = @{ 
                        CompressedGuid    = $compressed; ProductId = $guid; ParentGuid = $parentGuidDiscovered
                        DisplayName       = $group.DisplayName; Version = $group.Version; Publisher = $group.Publisher
                        InstallLocation   = $group.InstallLocation; InstallSource = $group.InstallSource; LocalPackage = $group.LocalPackage
                        InstallerProducts = if ($allCategoryData.InstallerProducts) { @($allCategoryData.InstallerProducts) } else { @() }
                        InstallerFolders  = if ($allCategoryData.InstallerFolders)  { @($allCategoryData.InstallerFolders) }  else { @() }
                        InstallerFiles    = if ($allCategoryData.InstallerFiles)    { @($allCategoryData.InstallerFiles) }    else { @() }
                        Dependencies      = if ($allCategoryData.Dependencies)      { @($allCategoryData.Dependencies) }      else { @() }
                        UpgradeCodes      = if ($allCategoryData.UpgradeCodes)      { @($allCategoryData.UpgradeCodes) }      else { @() }
                        Features          = if ($allCategoryData.Features)          { @($allCategoryData.Features) }          else { @() }
                        Components        = if ($allCategoryData.Components)        { @($allCategoryData.Components) }        else { @() }
                        UninstallEntries  = @($group.UninstallEntries)
                        UserDataEntries   = @($group.UserDataEntries)
                    }
                })
                if ($parentGuidDiscovered) { $group.ParentGuid = $parentGuidDiscovered }
            }
            # Parent moves
            $guidToParent = @{}
            foreach ($prodData in $params.ProductList) { if ($prodData.Group.ParentGuid) { $guidToParent[$prodData.Guid] = $prodData.Group.ParentGuid } }
            foreach ($guid in $guidToParent.Keys) {
                $parentGuid = $guidToParent[$guid]
                $parentExists = $false
                foreach ($pd in $params.ProductList) { if ($pd.Guid -eq $parentGuid) { $parentExists = $true; break } }
                if ($parentExists) { [void]$results.Add(@{ Type = 'MoveProductUnderParent'; Guid = $guid; Data = @{ ParentGuid = $parentGuid } }) }
            }
            return $results
        }

        try {
            $targetComputers      = $SearchParams.TargetComputers
            $isMultiComputer      = $SearchParams.IsMultiComputer
            $computerDisplayNames = $SearchParams.ComputerDisplayNames
            $fastMode             = $SearchParams.FastMode
            $totalComputers       = [Math]::Max(1, $targetComputers.Count)
            $currentCompIndex     = 0
            $grandTotalProducts   = 0
            $categories = @('Dependencies', 'UpgradeCodes', 'Features', 'Components', 'InstallerProducts', 'InstallerFolders', 'InstallerFiles')
            foreach ($computerName in $targetComputers) {
                if ($SyncHash.CancelRequested) { break }
                $currentCompIndex++
                $isLocal      = [string]::IsNullOrWhiteSpace($computerName)
                $displayLabel = if ($isLocal) { $env:COMPUTERNAME } else { $computerDisplayNames[$computerName] }
                $useFastMode  = $fastMode
                Write-SearchLog "[$currentCompIndex/$totalComputers] Starting search on $displayLabel $(if ($useFastMode) { '(fast mode)' } else { '' })"
                Write-SearchLog "Search criteria : Titles=[$($SearchParams.SearchTitles -join '; ')], GUIDs=[$($SearchParams.SearchGuids -join '; ')], ExactMatch=$($SearchParams.ExactMatch), ExactMatchGuid=$($SearchParams.ExactMatchGuid)"
                $SyncHash.ProgressStatus  = "[$currentCompIndex/$totalComputers] $displayLabel : Scanning registry..."
                $SyncHash.ProgressPercent = [int]((($currentCompIndex - 1) / $totalComputers) * 100)
                # Phase 1 : Initial scan
                $phase1Params = @{
                    SearchTitles   = $SearchParams.SearchTitles
                    SearchGuids    = $SearchParams.SearchGuids
                    ExactMatch     = $SearchParams.ExactMatch
                    ExactMatchGuid = $SearchParams.ExactMatchGuid
                }
                $phase1Result = $null
                try {
                    if ($isLocal) {
                        $phase1Result = & $Phase1Script $phase1Params
                    }
                    else {
                        $invokeParams = @{ ComputerName = $computerName; ScriptBlock = $Phase1Script; ArgumentList = @($phase1Params); ErrorAction = 'Stop' }
                        if ($SearchParams.Credential) { $invokeParams.Credential = $SearchParams.Credential }
                        $phase1Result = Invoke-Command @invokeParams
                    }
                }
                catch {
                    Write-SearchLog "Error during Phase 1 on $displayLabel : $($_.Exception.Message)"
                    continue
                }
                $productGroups = if ($phase1Result -is [hashtable] -and $phase1Result.ContainsKey('Products')) { $phase1Result.Products } else { $phase1Result }
                $phase1Diags   = if ($phase1Result -is [hashtable] -and $phase1Result.ContainsKey('Diagnostics')) { $phase1Result.Diagnostics } else { $null }
                if ($phase1Diags) { foreach ($diagMsg in $phase1Diags) { Write-SearchLog "[$displayLabel] $diagMsg" } }
                if ($null -eq $productGroups -or $productGroups.Count -eq 0) {
                    Write-SearchLog "[$currentCompIndex/$totalComputers] No products found on $displayLabel"
                    continue
                }
                $sortedProductGuids = @($productGroups.Keys | Sort-Object { $productGroups[$_].DisplayName })
                $productCount       = $sortedProductGuids.Count
                $grandTotalProducts += $productCount
                $SyncHash.TotalProductCount = $grandTotalProducts
                Write-SearchLog "[$currentCompIndex/$totalComputers] Found $productCount product(s) on $displayLabel"
                # Send ProductRoot nodes immediately if fast mode not achecked
                if (-not $useFastMode) {
                    foreach ($guid in $sortedProductGuids) {
                        $group      = $productGroups[$guid]
                        $compressed = Convert-GuidToCompressed $guid
                        Send-Update @{ 
                            Type = 'ProductRoot'; Guid = $guid; Computer = $computerName
                            Data = @{ 
                                CompressedGuid = $compressed; ProductId = $guid; ParentGuid = $null
                                DisplayName = $group.DisplayName; Version = $group.Version; Publisher = $group.Publisher
                                InstallLocation = $group.InstallLocation; InstallSource = $group.InstallSource; LocalPackage = $group.LocalPackage
                                UninstallEntries = @($group.UninstallEntries) 
                            } 
                        }
                        if ($group.UserDataEntries -and $group.UserDataEntries.Count -gt 0) {
                            Send-Update @{ Type = 'UserDataEntries'; Guid = $guid; Computer = $computerName; Data = @($group.UserDataEntries) }
                        }
                    }
                }
                # Phase 2 : Process categories
                if ($useFastMode) {
                    # FAST MODE : Single Invoke-Command for all products and categories
                    $SyncHash.ProgressStatus = "[$currentCompIndex/$totalComputers] $displayLabel : Scanning..."
                    $productList = [System.Collections.ArrayList]::new()
                    foreach ($guid in $sortedProductGuids) {
                        $group      = $productGroups[$guid]
                        $compressed = Convert-GuidToCompressed $guid
                        [void]$productList.Add(@{ Guid = $guid; CompressedGuid = $compressed; Group = $group })
                    }
                    $phase2AllParams = @{
                        Phase2ScriptString = $Phase2ScriptString
                        ProductList        = $productList
                    }
                    try {
                        if ($isLocal) {
                            $allResults = & $phase2AllScriptBlock $phase2AllParams
                        }
                        else {
                            $invokeParams = @{ ComputerName = $computerName; ScriptBlock = $phase2AllScriptBlock; ArgumentList = @($phase2AllParams); ErrorAction = 'Stop' }
                            if ($SearchParams.Credential) { $invokeParams.Credential = $SearchParams.Credential }
                            $allResults = Invoke-Command @invokeParams
                        }
                        # Store for bulk processing in OnComplete
                        if (-not $SyncHash.FastModeResults) { $SyncHash.FastModeResults = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new()) }
                        foreach ($result in $allResults) {
                            $result['Computer'] = $computerName
                            [void]$SyncHash.FastModeResults.Add($result)
                        }
                    }
                    catch {
                        Write-SearchLog "Error during fast Phase 2 on $displayLabel : $($_.Exception.Message)"
                    }
                    $SyncHash.ProgressPercent = [int][Math]::Min(99, ($currentCompIndex / $totalComputers) * 100)
                }
                else {
                    # NORMAL MODE : Separate Invoke-Command per category per product
                    $productIndex = 0
                    foreach ($guid in $sortedProductGuids) {
                        if ($SyncHash.CancelRequested) { break }
                        $productIndex++
                        $group      = $productGroups[$guid]
                        $compressed = Convert-GuidToCompressed $guid
                        $shortName  = if ($group.DisplayName -and $group.DisplayName.Length -gt 40) { $group.DisplayName.Substring(0, 37) + "..." } elseif ($group.DisplayName) { $group.DisplayName } else { $guid }
                        if ($isMultiComputer) { $shortName = "[$displayLabel] $shortName" }
                        Write-SearchLog "[$currentCompIndex/$totalComputers] [$productIndex/$productCount] Processing : $shortName"
                        $parentGuidDiscovered = $null
                        $allCategoryData = @{}
                        $categoryIndex   = 0
                        $totalCategories = $categories.Count
                        foreach ($catName in $categories) {
                            if ($SyncHash.CancelRequested) { break }
                            $categoryIndex++
                            $compBasePercent   = (($currentCompIndex - 1) / $totalComputers) * 100
                            $compRangePercent  = 100 / $totalComputers
                            $productBaseOffset = (($productIndex - 1) / [Math]::Max(1, $productCount)) * $compRangePercent
                            $productRange      = $compRangePercent / [Math]::Max(1, $productCount)
                            $categoryProgress  = $categoryIndex / [Math]::Max(1, $totalCategories)
                            $SyncHash.ProgressPercent = [int][Math]::Min(99, $compBasePercent + $productBaseOffset + ($categoryProgress * $productRange))
                            $SyncHash.ProgressStatus  = "[$currentCompIndex/$totalComputers] $displayLabel : [$productIndex/$productCount] $shortName - $catName"
                            $phase2Params = @{
                                Guid            = $guid
                                CompressedGuid  = $compressed
                                DisplayName     = $group.DisplayName
                                InstallLocation = $group.InstallLocation
                                LocalPackage    = $group.LocalPackage
                                UserDataEntries = @($group.UserDataEntries)
                                CategoryName    = $catName
                            }
                            $catResult = $null
                            try {
                                if ($isLocal) {
                                    $catResult = & $Phase2Script $phase2Params
                                }
                                else {
                                    $invokeParams = @{ ComputerName = $computerName; ScriptBlock = $Phase2Script; ArgumentList = @($phase2Params); ErrorAction = 'Stop' }
                                    if ($SearchParams.Credential) { $invokeParams.Credential = $SearchParams.Credential }
                                    $catResult = Invoke-Command @invokeParams
                                }
                            }
                            catch {
                                Write-SearchLog "Error scanning $catName for $shortName : $($_.Exception.Message)"
                                continue
                            }
                            if ($catResult -and $catResult.Results -and $catResult.Results.Count -gt 0) {
                                $allCategoryData[$catName] = @($catResult.Results)
                                if ($catName -ne 'InstallerFiles') {
                                    Send-Update @{ Type = $catName; Guid = $guid; Computer = $computerName; Data = @($catResult.Results) }
                                }
                            }
                            if ($catResult -and $catResult.ParentGuid) {
                                $parentGuidDiscovered = $catResult.ParentGuid
                            }
                        }
                        # Post-process files
                        $diskFolders = @(); $diskFiles = @(); $registryFiles = @()
                        if ($allCategoryData.InstallerFolders) {
                            foreach ($folder in $allCategoryData.InstallerFolders) {
                                if ($folder.Exists) { $diskFolders += @{ FolderPath = $folder.FolderPath } }
                            }
                        }
                        if ($allCategoryData.InstallerFiles) {
                            foreach ($file in $allCategoryData.InstallerFiles) {
                                if ($file.RegistryPath) { $registryFiles += $file }
                                $diskFiles += $file
                            }
                        }
                        if ($registryFiles.Count -gt 0) { Send-Update @{ Type = 'InstallerFiles'; Guid = $guid; Computer = $computerName; Data = @($registryFiles) } }
                        if ($diskFolders.Count -gt 0)   { Send-Update @{ Type = 'DiskFolders';    Guid = $guid; Computer = $computerName; Data = @($diskFolders) } }
                        if ($diskFiles.Count -gt 0)     { Send-Update @{ Type = 'DiskFiles';      Guid = $guid; Computer = $computerName; Data = @($diskFiles) } }
                        Send-Update @{ 
                            Type = 'CacheProduct'; Guid = $guid; Computer = $computerName
                            Data = @{ 
                                CompressedGuid    = $compressed; ProductId = $guid; ParentGuid = $parentGuidDiscovered
                                DisplayName       = $group.DisplayName; Version = $group.Version; Publisher = $group.Publisher
                                InstallLocation   = $group.InstallLocation; InstallSource = $group.InstallSource; LocalPackage = $group.LocalPackage
                                InstallerProducts = if ($allCategoryData.InstallerProducts) { @($allCategoryData.InstallerProducts) } else { @() }
                                InstallerFolders  = if ($allCategoryData.InstallerFolders)  { @($allCategoryData.InstallerFolders) }  else { @() }
                                InstallerFiles    = if ($allCategoryData.InstallerFiles)    { @($allCategoryData.InstallerFiles) }    else { @() }
                                Dependencies      = if ($allCategoryData.Dependencies)      { @($allCategoryData.Dependencies) }      else { @() }
                                UpgradeCodes      = if ($allCategoryData.UpgradeCodes)      { @($allCategoryData.UpgradeCodes) }      else { @() }
                                Features          = if ($allCategoryData.Features)          { @($allCategoryData.Features) }          else { @() }
                                Components        = if ($allCategoryData.Components)        { @($allCategoryData.Components) }        else { @() }
                                UninstallEntries  = @($group.UninstallEntries)
                                UserDataEntries   = @($group.UserDataEntries)
                            } 
                        }
                        if ($parentGuidDiscovered) { $group.ParentGuid = $parentGuidDiscovered }
                    }
                    # Parent moves
                    foreach ($guid in $sortedProductGuids) {
                        $group = $productGroups[$guid]
                        if ($group.ParentGuid -and $productGroups.ContainsKey($group.ParentGuid)) {
                            Send-Update @{ Type = 'MoveProductUnderParent'; Guid = $guid; Computer = $computerName; Data = @{ ParentGuid = $group.ParentGuid } }
                        }
                    }
                }
            }
            Write-SearchLog "Search complete : $grandTotalProducts product(s) found"
        }
        catch { 
            $SyncHash.Error = $_.Exception.Message
            Write-SearchLog "Critical Error : $($_.Exception.Message)"
        }
        finally { 
            $SyncHash.IsComplete = $true 
        }
    }
    $script:BackgroundPowerShell          = [powershell]::Create()
    $script:BackgroundPowerShell.Runspace = $script:BackgroundRunspace
    [void]$script:BackgroundPowerShell.AddScript($controllerScriptBlock)
    [void]$script:BackgroundPowerShell.BeginInvoke()
    # -------------------------------------------------------------------------
    # 7. UI TIMER
    # -------------------------------------------------------------------------
    $searchTimer = New-BackgroundOperationTimer -SyncHash $script:SyncHash -AdditionalState @{
        IsMultiComputer = $isMultiComputer
        FastMode        = -not $fastModeCheckBox_MSICleanupTab.Checked
    } -OnUpdate {
        param($update, $state)
        $isMultiComp = $state.IsMultiComputer
        $compKey     = $update.Computer
        switch ($update.Type) {
            'Log' {
                Write-Log $update.Message
            }
            'CacheProduct' {
                $compCacheKey = if ([string]::IsNullOrWhiteSpace($compKey)) { "" } else { $compKey }
                Add-ProductToCache -Product ([PSCustomObject]$update.Data) -ComputerName $compCacheKey
            }
            'MoveProductUnderParent' {
                $parentKey = if ($isMultiComp) { "${compKey}|$($update.Data.ParentGuid)" } else { $update.Data.ParentGuid }
                if ($script:CurrentProductNodes.ContainsKey($parentKey)) {
                    Move-ProductNodeUnderParent -TreeView $treeViewSearch_MSICleanupTab -ChildGuid $update.Guid -ParentGuid $update.Data.ParentGuid -CurrentProductNodes $script:CurrentProductNodes -ComputerName $compKey
                }
            }
            default {
                Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewSearch_MSICleanupTab -ProductGuid $update.Guid -CategoryName $update.Type -CategoryData $update.Data -CurrentProductNodes $script:CurrentProductNodes -SuppressUpdate $true -ComputerName $compKey -ComputerNodes $script:ComputerNodesSearch -IsMultiComputer $isMultiComp
                $productKey = if ($isMultiComp) { "${compKey}|$($update.Guid)" } else { $update.Guid }
                if ($script:CurrentProductNodes.ContainsKey($productKey) -and $update.Type -eq 'ProductRoot') {
                    $script:CurrentProductNodes[$productKey].RootNode.Expand()
                }
            }
        }
    } -OnComplete {
        param($timerSyncHash, $state)
        # Suppress cross-tab sync events during tree construction
        $script:SyncInProgress_CrossTab = $true
        try {
            # Process fast mode results in bulk
            if ($state.FastMode -and $timerSyncHash.FastModeResults -and $timerSyncHash.FastModeResults.Count -gt 0) {
                $treeViewSearch_MSICleanupTab.BeginUpdate()
                try {
                    foreach ($update in $timerSyncHash.FastModeResults) {
                        $compKey     = $update.Computer
                        $isMultiComp = $state.IsMultiComputer
                        switch ($update.Type) {
                            'CacheProduct' {
                                $compCacheKey = if ([string]::IsNullOrWhiteSpace($compKey)) { "" } else { $compKey }
                                Add-ProductToCache -Product ([PSCustomObject]$update.Data) -ComputerName $compCacheKey
                            }
                            'MoveProductUnderParent' {
                                $parentKey = if ($isMultiComp) { "${compKey}|$($update.Data.ParentGuid)" } else { $update.Data.ParentGuid }
                                if ($script:CurrentProductNodes.ContainsKey($parentKey)) {
                                    Move-ProductNodeUnderParent -TreeView $treeViewSearch_MSICleanupTab -ChildGuid $update.Guid -ParentGuid $update.Data.ParentGuid -CurrentProductNodes $script:CurrentProductNodes -ComputerName $compKey
                                }
                            }
                            default {
                                Add-CategoryToTreeView_MSICleanupTab -TreeView $treeViewSearch_MSICleanupTab -ProductGuid $update.Guid -CategoryName $update.Type -CategoryData $update.Data -CurrentProductNodes $script:CurrentProductNodes -SuppressUpdate $true -ComputerName $compKey -ComputerNodes $script:ComputerNodesSearch -IsMultiComputer $isMultiComp
                            }
                        }
                    }
                    # Expand all product nodes at once
                    foreach ($productKey in $script:CurrentProductNodes.Keys) {
                        $script:CurrentProductNodes[$productKey].RootNode.Expand()
                    }
                }
                finally {
                    $treeViewSearch_MSICleanupTab.EndUpdate()
                }
            }
        }
        finally { $script:SyncInProgress_CrossTab = $false }
        # Restore cross-tab sync state after rebuild
        if ($treeViewSearch_MSICleanupTab.Nodes.Count -gt 0 -and $script:SharedExpandedSyncPaths.Count -gt 0) {
            Restore-TreeViewSyncState -TreeView $treeViewSearch_MSICleanupTab
        }
        elseif ($treeViewSearch_MSICleanupTab.Nodes.Count -gt 0) {
            # Scroll to top only when no sync state to restore
            [NativeMethods]::SendMessage($treeViewSearch_MSICleanupTab.Handle, [NativeMethods]::WM_VSCROLL, [NativeMethods]::SB_TOP, 0)
        }
        if ($script:CancelRequested) {
            Write-Log "Search cancelled"
            Stop-ProgressUI -FinalStatus "Search cancelled"
        }
        elseif ($timerSyncHash.Error) {
            Write-Log "Search error : $($timerSyncHash.Error)" -Level Error
            Stop-ProgressUI -FinalStatus "Error : $($timerSyncHash.Error)"
        }
        else {
            $itemCount    = Get-TreeNodeCount_MSICleanupTab -TreeView $treeViewSearch_MSICleanupTab
            $productCount = $timerSyncHash.TotalProductCount
            Write-Log "Search complete : $productCount products, $itemCount items"
            Stop-ProgressUI -FinalStatus "Found $productCount product(s) with $itemCount total items"
            $cleanButton_MSICleanupTab.Enabled = ($itemCount -gt 0)
        }
        Update-TabStates
    }
    $searchTimer.Start()
})
#region Tab4 events

foreach ($tv in @($treeViewSearch_MSICleanupTab, $treeViewFullCache, $treeViewCompare)) {
    $tv.Add_AfterExpand($script:SyncAfterExpandHandler)
    $tv.Add_AfterCollapse($script:SyncAfterCollapseHandler)
    $tv.Add_AfterSelect($script:SyncAfterSelectHandler)
    $tv.Add_AfterCheck($script:SyncAfterCheckHandler)
}

foreach ($sc in @($splitContainerSearch_MSICleanupTab, $splitContainerFullCache, $splitContainerCompare)) {
    $sc.Add_SplitterMoved($script:SyncSplitterMovedHandler)
    $sc.Add_VisibleChanged({
        param($s, $e)
        if ($s.Visible -and $s.Width -gt 0 -and -not $s.Tag.InitialRatioApplied) {
            $s.SplitterDistance = [int]($s.Width * $script:SharedSplitterRatio)
            $s.Tag = @{ InitialRatioApplied = $true }
        }
    })
}

$tabControl_MSICleanupTab.Add_SelectedIndexChanged({
    if (-not $script:IsBackgroundOperationRunning) { $Tab4_statusLabel.Text = "Ready" }
    $isUninstallTab = ($tabControl_MSICleanupTab.SelectedIndex -eq 2)
    $buttonFlowPanel_MSICleanupTab.Visible  = -not $isUninstallTab
    $activeTreeView = Get-ActiveTreeView
    if ($activeTreeView) {
        $itemCount = Get-TreeNodeCount_MSICleanupTab -TreeView $activeTreeView
        $cleanButton_MSICleanupTab.Enabled = ($itemCount -gt 0)
    }
    else { $cleanButton_MSICleanupTab.Enabled = $false }
    # Save sync state from the previously active TreeView
    $previousTreeView = switch ($script:PreviousTab4Index) {
        0       { $treeViewSearch_MSICleanupTab }
        1       { $treeViewFullCache }
        3       { $treeViewCompare }
        default { $null }
    }
    if ($previousTreeView -and $previousTreeView.Nodes.Count -gt 0) {
        Save-TreeViewSyncState -TreeView $previousTreeView
    }
    switch ($tabControl_MSICleanupTab.SelectedIndex) {
        0 {
            $splitContainerSearch_MSICleanupTab.SplitterDistance = [int]($splitContainerSearch_MSICleanupTab.Width * 0.6)
            Restore-SplitterRatio -Target $splitContainerSearch_MSICleanupTab
            $titleTextBox_MSICleanupTab.Focus()
            if ($treeViewSearch_MSICleanupTab.Nodes.Count -gt 0) {
                Restore-TreeViewSyncState -TreeView $treeViewSearch_MSICleanupTab
            }
        }
        1 {
            Update-FullCacheTab
            $splitContainerSearch_MSICleanupTab.SplitterDistance = [int]($splitContainerSearch_MSICleanupTab.Width * 0.6)
            Restore-SplitterRatio -Target $splitContainerFullCache
            if ($treeViewFullCache.Nodes.Count -gt 0) {
                Restore-TreeViewSyncState -TreeView $treeViewFullCache
            }
        }
        2 {
            if ($script:UninstallTabCacheVersion -ne $script:CacheVersion) { Update-UninstallTab }
        }
        3 {
            $splitContainerSearch_MSICleanupTab.SplitterDistance = [int]($splitContainerSearch_MSICleanupTab.Width * 0.6)
            Update-CompareTab
            Restore-SplitterRatio -Target $splitContainerCompare
            if ($treeViewCompare.Nodes.Count -gt 0) {
                Restore-TreeViewSyncState -TreeView $treeViewCompare
            }
        }
    }
    $script:PreviousTab4Index = $tabControl_MSICleanupTab.SelectedIndex
})


$resetButton_MSICleanupTab.Add_Click({
    if ($script:IsBackgroundOperationRunning) {
        [System.Windows.Forms.MessageBox]::Show("Cannot clear while a background operation is running. Use the STOP button first.", "Operation In Progress", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $hasContentToClear = ($treeViewSearch_MSICleanupTab.Nodes.Count -gt 0) -or ($treeViewFullCache.Nodes.Count -gt 0) -or ($treeViewCompare.Nodes.Count -gt 0) -or ($uninstallFlowPanel_MSICleanupTab.Controls.Count -gt 0) -or ($script:ProductCache.Count -gt 0)
    $exactMatchCheckBox_MSICleanupTab.Checked = $false
    $treeViewSearch_MSICleanupTab.Nodes.Clear()
    $detailsRichTextBoxSearch_MSICleanupTab.Clear()
    $treeViewFullCache.Nodes.Clear()
    $detailsRichTextBoxFullCache.Clear()
    $treeViewCompare.Nodes.Clear()
    $detailsRichTextBoxCompare.Clear()
    $uninstallFlowPanel_MSICleanupTab.Controls.Clear()
    $script:CurrentProductNodes          = @{}
    $script:CurrentProductNodesFullCache = @{}
    $script:CurrentProductNodesCompare   = @{}
    $script:ComputerNodesSearch          = @{}
    $script:ComputerNodesFullCache       = @{}
    $script:ComputerNodesCompare         = @{}
    $script:UninstallPanelControls       = @{}
    $script:UninstallPanelStates         = @{}
    $script:UninstallTabCacheVersion     = -1
    $script:FullCacheTabCacheVersion     = -1
    $script:CompareTabCacheVersion       = -1
    $script:SelectedUninstallComputer    = $null
    $script:LastSearchComputers          = @()
    $script:SharedExpandedSyncPaths.Clear()
    $script:SharedCheckedSyncPaths.Clear()
    $script:SharedSelectedSyncPath = $null
    $script:SharedSplitterRatio    = 0.6
    $script:PreviousTab4Index      = 0
    Clear-ProductCache
    $cleanButton_MSICleanupTab.Enabled  = $false
    $tabUninstall_MSICleanupTab.Enabled = $false
    $tabFullCache_MSICleanupTab.Enabled = $false
    $tabCompare_MSICleanupTab.Enabled   = $false
    $tabControl_MSICleanupTab.SelectedIndex = 0
    Update-TabStates
    $Tab4_statusLabel.Text = "Ready"
    if (-not $hasContentToClear) { $titleTextBox_MSICleanupTab.Clear(); $guidTextBox_MSICleanupTab.Clear() }
})

$Tab4_stopButton.Add_Click({
    Write-Log "Stop button clicked - requesting cancellation"
    $script:CancelRequested = $true
    if ($script:SyncHash) { $script:SyncHash.CancelRequested = $true }
    $Tab4_stopButton.Enabled   = $false
    $Tab4_stopButton.BackColor = [System.Drawing.Color]::LightGray
    $Tab4_stopButton.ForeColor = [System.Drawing.Color]::Black
    $Tab4_statusLabel.Text     = "Cancelling..."
})

$selectAllButton_MSICleanupTab.Add_Click(   { $activeTreeView = Get-ActiveTreeView; if ($activeTreeView) { Set-TreeViewCheckState_MSICleanupTab -TreeView $activeTreeView -Checked $true } })
$deselectAllButton_MSICleanupTab.Add_Click( { $activeTreeView = Get-ActiveTreeView; if ($activeTreeView) { Set-TreeViewCheckState_MSICleanupTab -TreeView $activeTreeView -Checked $false } })
$expandAllButton_MSICleanupTab.Add_Click(   { $activeTreeView = Get-ActiveTreeView; if ($activeTreeView) { $activeTreeView.ExpandAll() } })
$collapseAllButton_MSICleanupTab.Add_Click( { $activeTreeView = Get-ActiveTreeView; if ($activeTreeView) { $activeTreeView.CollapseAll() } })

$treeViewSearch_MSICleanupTab.Add_AfterSelect({      param($s, $e); if ($e.Node.Tag) { Format-NodeDetailsRich -Node $e.Node -RichTextBox $detailsRichTextBoxSearch_MSICleanupTab; Update-TreeViewContextMenu_MSICleanupTab -Node $e.Node -ContextMenu $detailsContextMenuSearch_MSICleanupTab -RichTextBox $detailsRichTextBoxSearch_MSICleanupTab } else { $detailsRichTextBoxSearch_MSICleanupTab.Clear(); $detailsContextMenuSearch_MSICleanupTab.Items.Clear() } })
$treeViewSearch_MSICleanupTab.Add_NodeMouseClick({   param($s, $e); if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) { $treeViewSearch_MSICleanupTab.SelectedNode = $e.Node; if ($e.Node.Tag) { Format-NodeDetailsRich -Node $e.Node -RichTextBox $detailsRichTextBoxSearch_MSICleanupTab; Update-TreeViewContextMenu_MSICleanupTab -Node $e.Node -ContextMenu $detailsContextMenuSearch_MSICleanupTab -RichTextBox $detailsRichTextBoxSearch_MSICleanupTab; if ($detailsContextMenuSearch_MSICleanupTab.Items.Count -gt 0) { $detailsContextMenuSearch_MSICleanupTab.Show($treeViewSearch_MSICleanupTab, $e.Location) } } } })
$treeViewFullCache.Add_AfterSelect({   param($s, $e); if ($e.Node.Tag) { Format-NodeDetailsRich -Node $e.Node -RichTextBox $detailsRichTextBoxFullCache; Update-TreeViewContextMenu_MSICleanupTab -Node $e.Node -ContextMenu $detailsContextMenuFullCache -RichTextBox $detailsRichTextBoxFullCache } else { $detailsRichTextBoxFullCache.Clear(); $detailsContextMenuFullCache.Items.Clear() } })
$treeViewFullCache.Add_NodeMouseClick({ param($s, $e); if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) { $treeViewFullCache.SelectedNode = $e.Node; if ($e.Node.Tag) { Format-NodeDetailsRich -Node $e.Node -RichTextBox $detailsRichTextBoxFullCache; Update-TreeViewContextMenu_MSICleanupTab -Node $e.Node -ContextMenu $detailsContextMenuFullCache -RichTextBox $detailsRichTextBoxFullCache; if ($detailsContextMenuFullCache.Items.Count -gt 0) { $detailsContextMenuFullCache.Show($treeViewFullCache, $e.Location) } } } })
$treeViewCompare.Add_AfterSelect({     param($s, $e); if ($e.Node.Tag) { Format-NodeDetailsRich -Node $e.Node -RichTextBox $detailsRichTextBoxCompare; Update-TreeViewContextMenu_MSICleanupTab -Node $e.Node -ContextMenu $detailsContextMenuCompare -RichTextBox $detailsRichTextBoxCompare } else { $detailsRichTextBoxCompare.Clear(); $detailsContextMenuCompare.Items.Clear() } })
$treeViewCompare.Add_NodeMouseClick({  param($s, $e); if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) { $treeViewCompare.SelectedNode = $e.Node; if ($e.Node.Tag) { Format-NodeDetailsRich -Node $e.Node -RichTextBox $detailsRichTextBoxCompare; Update-TreeViewContextMenu_MSICleanupTab -Node $e.Node -ContextMenu $detailsContextMenuCompare -RichTextBox $detailsRichTextBoxCompare; if ($detailsContextMenuCompare.Items.Count -gt 0) { $detailsContextMenuCompare.Show($treeViewCompare, $e.Location) } } } })

$cleanButton_MSICleanupTab.Add_Click({
    # ═══════════════════════════════════════════════════════════════
    # 1. VALIDATION
    # ═══════════════════════════════════════════════════════════════
    $activeTreeView = Get-ActiveTreeView
    if (-not $activeTreeView) { [System.Windows.Forms.MessageBox]::Show("Please select a tab with items to clean.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }
    if ($script:IsBackgroundOperationRunning) { [System.Windows.Forms.MessageBox]::Show("A background operation is already running.", "Operation In Progress", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }
    $itemsToClean = Get-CheckedItems_MSICleanupTab -TreeView $activeTreeView
    if ($itemsToClean.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select items to clean.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }

    # ═══════════════════════════════════════════════════════════════
    # 2. GROUP BY COMPUTER
    # ═══════════════════════════════════════════════════════════════
    $itemsByComputer = [ordered]@{}
    foreach ($item in $itemsToClean) {
        $compKey = if ($item.Computer -and $item.Computer -ne "" -and $item.Computer -ne $env:COMPUTERNAME) { $item.Computer } else { "" }
        if (-not $itemsByComputer.Contains($compKey)) { $itemsByComputer[$compKey] = [System.Collections.ArrayList]::new() }
        [void]$itemsByComputer[$compKey].Add($item)
    }
    $hasRemoteItems = @($itemsByComputer.Keys | Where-Object { $_ -ne "" }).Count -gt 0
    $credential     = if ($hasRemoteItems) { Get-CredentialFromPanel } else { $null }

    # ═══════════════════════════════════════════════════════════════
    # 3. CONFIRMATION
    # ═══════════════════════════════════════════════════════════════
    $computerCount = $itemsByComputer.Count
    $confirmMsg    = "You have selected $($itemsToClean.Count) item(s) for removal"
    if ($computerCount -gt 1 -or $hasRemoteItems) {
        $names      = @($itemsByComputer.Keys | ForEach-Object { if ($_ -eq "") { $env:COMPUTERNAME } else { Get-ComputerNodeLabel -ComputerName $_ } })
        $confirmMsg += " across $computerCount computer(s) :`n$($names -join ', ')"
    }
    $confirmMsg += ".`n`nThis operation cannot be undone.`n`nContinue?"
    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "Confirm Cleanup", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    Write-Log "Starting cleanup of $($itemsToClean.Count) items across $computerCount computer(s)"

    # ═══════════════════════════════════════════════════════════════
    # 4. PREPARE SERIALIZABLE DATA
    # ═══════════════════════════════════════════════════════════════
    $categoryMap = @{ 'Folder'='Folder'; 'File'='File'; 'InstallerCache'='File'; 'LocalPackage'='File'; 'InstallSource'='File'; 'Registry'='Registry'; 'RegistryValue'='RegistryValue'; 'UninstallEntry'='UninstallEntry' }
    $allItemsGrouped = [ordered]@{}
    foreach ($compKey in $itemsByComputer.Keys) {
        $items = [System.Collections.ArrayList]::new()
        $itemsByComputer[$compKey] | Where-Object { $_.Type -eq 'Folder' } | Sort-Object { ($_.Path -split '\\').Count } -Descending | ForEach-Object { [void]$items.Add(@{ Path = $_.Path; Category = 'Folder'; RegistryValue = $null }) }
        $itemsByComputer[$compKey] | Where-Object { $_.Type -ne 'Folder' } | ForEach-Object { [void]$items.Add(@{ Path = $_.Path; Category = $categoryMap[$_.Type]; RegistryValue = $_.RegistryValue }) }
        $allItemsGrouped[$compKey] = $items
    }

    # ═══════════════════════════════════════════════════════════════
    # 5. SYNCHASH + RUNSPACE
    # ═══════════════════════════════════════════════════════════════
    $script:IsBackgroundOperationRunning = $true
    $script:CancelRequested = $false
    Start-ProgressUI -InitialStatus "Preparing cleanup..."
    $script:SyncHash = [hashtable]::Synchronized(@{
        CancelRequested = $false; ProgressPercent = 0; ProgressStatus = "Initializing..."; IsComplete = $false; Error = $null
        Results       = @{}                            # compKey -> @{ SuccessCount; FailCount; FailedItems }
        ComputerOrder = @($allItemsGrouped.Keys)
    })
    $script:BackgroundRunspace = New-ConfiguredRunspace -FunctionNames @('ConvertTo-PowerShellPath', 'Split-RegistryPath')
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("SyncHash", $script:SyncHash)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("AllItemsGrouped", $allItemsGrouped)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("RemoteCredential", $credential)
    $script:BackgroundRunspace.SessionStateProxy.SetVariable("LocalComputerName", $env:COMPUTERNAME)

    # ═══════════════════════════════════════════════════════════════
    # 6. CLEANUP SCRIPTBLOCK
    # ═══════════════════════════════════════════════════════════════
    $cleanupScriptBlock = {
        # ── Self-contained scriptblock for Invoke-Command (remote) ──
        $remoteCleanupScript = {
            param([array]$Items)
            function Parse-Reg { param([string]$p)
                $p = $p -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''; $r = $null; $s = $null
                if     ($p -match '^HKLM[:\\]+(.+)$')           { $r = [Microsoft.Win32.Registry]::LocalMachine; $s = $matches[1] }
                elseif ($p -match '^HKCU[:\\]+(.+)$')           { $r = [Microsoft.Win32.Registry]::CurrentUser;  $s = $matches[1] }
                elseif ($p -match '^HKU[:\\]+(.+)$')            { $r = [Microsoft.Win32.Registry]::Users;        $s = $matches[1] }
                elseif ($p -match '^HKCR[:\\]+(.+)$')           { $r = [Microsoft.Win32.Registry]::ClassesRoot;  $s = $matches[1] }
                elseif ($p -match '^HKEY_LOCAL_MACHINE\\(.+)$') { $r = [Microsoft.Win32.Registry]::LocalMachine; $s = $matches[1] }
                elseif ($p -match '^HKEY_CURRENT_USER\\(.+)$')  { $r = [Microsoft.Win32.Registry]::CurrentUser;  $s = $matches[1] }
                elseif ($p -match '^HKEY_USERS\\(.+)$')         { $r = [Microsoft.Win32.Registry]::Users;        $s = $matches[1] }
                elseif ($p -match '^HKEY_CLASSES_ROOT\\(.+)$')  { $r = [Microsoft.Win32.Registry]::ClassesRoot;  $s = $matches[1] }
                if ($s) { $s = $s.TrimStart('\') }
                return @{ Root = $r; SubPath = $s }
            }
            $sc = 0; $fc = 0; $fl = [System.Collections.ArrayList]::new(); $dp = [System.Collections.ArrayList]::new()
            foreach ($e in $Items) {
                try {
                    switch ($e.Category) {
                        'Folder' {
                            $skip = $false; foreach ($d in $dp) { if ($e.Path.StartsWith($d)) { $skip = $true; break } }; if ($skip) { $sc++; continue }
                            if ([System.IO.Directory]::Exists($e.Path)) { [System.IO.Directory]::Delete($e.Path, $true); [void]$dp.Add($e.Path + '\') }; $sc++
                        }
                        'File' { if ([System.IO.File]::Exists($e.Path)) { [System.IO.File]::Delete($e.Path) }; $sc++ }
                        { $_ -in @('Registry', 'UninstallEntry') } {
                            $pr = Parse-Reg $e.Path; if ($null -eq $pr.Root) { throw "Invalid registry path" }
                            $k = $pr.Root.OpenSubKey($pr.SubPath, $false); if ($k) { $k.Close(); $pr.Root.DeleteSubKeyTree($pr.SubPath, $false) }; $sc++
                        }
                        'RegistryValue' {
                            $pr = Parse-Reg $e.Path; if ($null -eq $pr.Root) { throw "Invalid registry path" }
                            if ($e.RegistryValue) { $k = $pr.Root.OpenSubKey($pr.SubPath, $true); if ($k) { $k.DeleteValue($e.RegistryValue, $false); $k.Close() } }
                            else                  { $k = $pr.Root.OpenSubKey($pr.SubPath, $false); if ($k) { $k.Close(); $pr.Root.DeleteSubKeyTree($pr.SubPath, $false) } }; $sc++
                        }
                        default { $sc++ }
                    }
                }
                catch { $fc++; [void]$fl.Add(@{ Type = $e.Category; Path = $e.Path; Error = $_.Exception.Message; RegistryValue = $e.RegistryValue }) }
            }
            return @{ SuccessCount = $sc; FailCount = $fc; FailedItems = @($fl) }
        }

        # ── Local deletion helper (uses injected ConvertTo-PowerShellPath / Split-RegistryPath) ──
        function Remove-LocalItem {
            param([string]$Path, [string]$Category, [string]$RegistryValue)
            try {
                switch ($Category) {
                    'Folder' { if ([System.IO.Directory]::Exists($Path)) { [System.IO.Directory]::Delete($Path, $true) }; return @{ Success = $true; Error = $null; Deleted = $true } }
                    'File'   { if ([System.IO.File]::Exists($Path))      { [System.IO.File]::Delete($Path) };             return @{ Success = $true; Error = $null } }
                    { $_ -in @('Registry', 'UninstallEntry') } {
                        $ps = ConvertTo-PowerShellPath $Path; if ([string]::IsNullOrWhiteSpace($ps)) { return @{ Success = $false; Error = "Invalid registry path" } }
                        $pr = Split-RegistryPath $ps;         if ($null -eq $pr) { return @{ Success = $false; Error = "Unsupported registry root" } }
                        $k = $pr.Root.OpenSubKey($pr.SubPath, $false); if ($k) { $k.Close(); $pr.Root.DeleteSubKeyTree($pr.SubPath, $false) }
                        return @{ Success = $true; Error = $null }
                    }
                    'RegistryValue' {
                        $ps = ConvertTo-PowerShellPath $Path; if ([string]::IsNullOrWhiteSpace($ps)) { return @{ Success = $false; Error = "Invalid registry path" } }
                        $pr = Split-RegistryPath $ps;         if ($null -eq $pr) { return @{ Success = $false; Error = "Unsupported registry root" } }
                        if ($RegistryValue) { $k = $pr.Root.OpenSubKey($pr.SubPath, $true); if ($k) { $k.DeleteValue($RegistryValue, $false); $k.Close() } }
                        else                { $k = $pr.Root.OpenSubKey($pr.SubPath, $false); if ($k) { $k.Close(); $pr.Root.DeleteSubKeyTree($pr.SubPath, $false) } }
                        return @{ Success = $true; Error = $null }
                    }
                }
                return @{ Success = $true; Error = $null }
            }
            catch { return @{ Success = $false; Error = $_.Exception.Message } }
        }

        # ── Main controller ──
        try {
            $totalItems = 0; foreach ($k in $AllItemsGrouped.Keys) { $totalItems += $AllItemsGrouped[$k].Count }
            $processedItems = 0

            foreach ($compKey in $AllItemsGrouped.Keys) {
                if ($SyncHash.CancelRequested) { break }
                $items    = @($AllItemsGrouped[$compKey])
                $isRemote = ($compKey -ne "")
                $label    = if ($isRemote) { $compKey } else { $LocalComputerName }

                if ($isRemote) {
                    # ── REMOTE : single Invoke-Command per computer ──
                    $SyncHash.ProgressStatus = "Cleaning $($items.Count) items on $label..."
                    try {
                        $ip = @{ ComputerName = $compKey; ScriptBlock = $remoteCleanupScript; ArgumentList = @(, $items); ErrorAction = 'Stop' }
                        if ($RemoteCredential) { $ip.Credential = $RemoteCredential }
                        $res = Invoke-Command @ip
                        $SyncHash.Results[$compKey] = @{ SuccessCount = $res.SuccessCount; FailCount = $res.FailCount; FailedItems = @($res.FailedItems) }
                    }
                    catch {
                        $errMsg = $_.Exception.Message
                        $SyncHash.Results[$compKey] = @{ SuccessCount = 0; FailCount = $items.Count; FailedItems = @($items | ForEach-Object { @{ Type = $_.Category; Path = $_.Path; Error = $errMsg; RegistryValue = $_.RegistryValue } }) }
                    }
                    $processedItems += $items.Count
                    $SyncHash.ProgressPercent = [int](($processedItems / $totalItems) * 100)
                }
                else {
                    # ── LOCAL : item-by-item with per-item progress ──
                    $sc = 0; $fc = 0; $fl = [System.Collections.ArrayList]::new(); $dp = [System.Collections.ArrayList]::new()
                    foreach ($entry in $items) {
                        if ($SyncHash.CancelRequested) { break }
                        $processedItems++
                        $short = if ($entry.Path.Length -gt 60) { "..." + $entry.Path.Substring($entry.Path.Length - 57) } else { $entry.Path }
                        $SyncHash.ProgressPercent = [int](($processedItems / $totalItems) * 100)
                        $SyncHash.ProgressStatus  = "[$processedItems/$totalItems] Removing : $short"
                        if ($entry.Category -eq 'Folder') {
                            $skip = $false; foreach ($d in $dp) { if ($entry.Path.StartsWith($d)) { $skip = $true; break } }
                            if ($skip) { $sc++; continue }
                        }
                        $r = Remove-LocalItem -Path $entry.Path -Category $entry.Category -RegistryValue $entry.RegistryValue
                        if ($r.Success) { $sc++; if ($r.Deleted) { [void]$dp.Add($entry.Path + '\') } }
                        else            { $fc++; [void]$fl.Add(@{ Type = $entry.Category; Path = $entry.Path; Error = $r.Error; RegistryValue = $entry.RegistryValue }) }
                    }
                    $SyncHash.Results[""] = @{ SuccessCount = $sc; FailCount = $fc; FailedItems = @($fl) }
                }
            }
            $SyncHash.ProgressPercent = 100; $SyncHash.ProgressStatus = "Cleanup completed"
        }
        catch   { $SyncHash.Error = $_.Exception.Message }
        finally { $SyncHash.IsComplete = $true }
    }

    $script:BackgroundPowerShell          = [powershell]::Create()
    $script:BackgroundPowerShell.Runspace = $script:BackgroundRunspace
    $script:BackgroundPowerShell.AddScript($cleanupScriptBlock) | Out-Null
    $script:BackgroundPowerShell.BeginInvoke() | Out-Null

    # ═══════════════════════════════════════════════════════════════
    # 7. TIMER + RESULT FORM
    # ═══════════════════════════════════════════════════════════════
    $cleanupTimer = New-BackgroundOperationTimer -SyncHash $script:SyncHash -AdditionalState @{ ActiveTabIndex = $tabControl_MSICleanupTab.SelectedIndex; HasRemoteItems = $hasRemoteItems } -OnUpdate { } -OnComplete {
        param($timerSyncHash, $state)

        # ── Aggregate results ──
        $results       = $timerSyncHash.Results
        $computerOrder = @($timerSyncHash.ComputerOrder)
        $totalSuccess  = 0; $totalFail = 0
        $computersWithFailures = [System.Collections.ArrayList]::new()
        foreach ($ck in $computerOrder) {
            if ($results.ContainsKey($ck)) {
                $totalSuccess += $results[$ck].SuccessCount
                $totalFail    += $results[$ck].FailCount
                if ($results[$ck].FailCount -gt 0) { [void]$computersWithFailures.Add($ck) }
            }
        }
        $activeTabIndex = $state.ActiveTabIndex
        $statusMsg      = if ($script:CancelRequested) { "Cleanup cancelled" } elseif ($timerSyncHash.Error) { "Error" } else { "Cleanup completed" }
        Write-Log "$statusMsg : $totalSuccess successful, $totalFail failed" -Level $(if ($timerSyncHash.Error) { 'Error' } else { 'Info' })
        Stop-ProgressUI -FinalStatus "$statusMsg : $totalSuccess successful, $totalFail failed"

        # ── RichTextBox helpers ──
        $writeRtf = {
            param([System.Windows.Forms.RichTextBox]$rtb, [string]$text, [System.Drawing.Color]$color, [bool]$bold = $false)
            $rtb.SelectionFont  = if ($bold) { [System.Drawing.Font]::new($rtb.Font, [System.Drawing.FontStyle]::Bold) } else { $rtb.Font }
            $rtb.SelectionColor = $color
            $rtb.AppendText($text)
        }
        $writeFailedItem = {
            param([System.Windows.Forms.RichTextBox]$rtb, $item, $rtfWriter)
            & $rtfWriter $rtb "Type : "                ([System.Drawing.Color]::Black)    $true
            & $rtfWriter $rtb "$($item.Type)`r`n"      ([System.Drawing.Color]::DarkBlue) $false
            & $rtfWriter $rtb "Path : "                ([System.Drawing.Color]::Black)    $true
            & $rtfWriter $rtb "$($item.Path)`r`n"      ([System.Drawing.Color]::DarkRed)  $false
            if ($item.Error) {
                & $rtfWriter $rtb "Error : "           ([System.Drawing.Color]::Black)    $true
                & $rtfWriter $rtb "$($item.Error)`r`n" ([System.Drawing.Color]::Red)      $false
            }
        }
        $populateRtbForComputer = {
            param([System.Windows.Forms.RichTextBox]$rtb, [string]$compKey, $res, $rtfW, $fiW)
            $rtb.Clear()
            if (-not $res.ContainsKey($compKey) -or $res[$compKey].FailCount -eq 0) {
                & $rtfW $rtb "All items were successfully removed.`r`n" ([System.Drawing.Color]::DarkGreen) $true; $rtb.Select(0, 0); $rtb.ScrollToCaret(); return
            }
            $fc = $res[$compKey].FailCount; $fi = @($res[$compKey].FailedItems)
            & $rtfW $rtb "FAILED ITEMS ($fc)`r`n" ([System.Drawing.Color]::Red) $true
            & $rtfW $rtb (("-" * 60) + "`r`n") ([System.Drawing.Color]::Gray) $false
            foreach ($f in $fi) { & $rtfW $rtb "`r`n" ([System.Drawing.Color]::Black) $false; & $fiW $rtb $f $rtfW }
            $rtb.Select(0, 0); $rtb.ScrollToCaret()
        }

        # ── Build result form ──
        $resultForm = gen $null "Form" "Cleanup Results" 0 0 700 500 'StartPosition=CenterParent' 'Font=Segoe UI, 9'
        $resultForm.MinimumSize = [System.Drawing.Size]::new(500, 350)

        # Summary
        $summaryPanel      = gen $resultForm "Panel" 'Dock=Top' 'Height=80' 'Padding=15 10 15 10'
        $summaryLabel      = gen $summaryPanel "Label" 'Dock=Fill' 'Font=Segoe UI, 11'
        $summaryText       = "Cleanup operation completed.`n`nSuccessful : $totalSuccess`nFailed : $totalFail"
        if ($computerOrder.Count -gt 1 -or $state.HasRemoteItems) { $summaryText += " (across $($computerOrder.Count) computer(s))" }
        $summaryLabel.Text      = $summaryText
        $summaryLabel.ForeColor = if ($totalFail -eq 0) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkOrange }

        # Computer selector (only if multiple computers have failures)
        $selectedComputer = [ref]$(if ($computersWithFailures.Count -gt 0) { $computersWithFailures[0] } elseif ($computerOrder.Count -gt 0) { $computerOrder[0] } else { "" })
        $selectorPanel    = $null
        if ($computersWithFailures.Count -gt 1) {
            $selectorPanel   = gen $resultForm "Panel" 'Dock=Top' 'Height=40' 'BackColor=240 240 245' 'BorderStyle=FixedSingle' 'Padding=10 5 10 5'
            $selectorLabel   = gen $selectorPanel "Label" "Computer :" 10 10 70 20 'Font=Segoe UI, 9, Bold'
            $radioScrollPanel = gen $selectorPanel "Panel" "" 85 3 0 34 'Dock=Fill' 'AutoScroll=$true'
            $radioFlow       = gen $radioScrollPanel "FlowLayoutPanel" "" 0 0 0 0 'FlowDirection=LeftToRight' 'WrapContents=$true' 'AutoSize=$true' 'AutoSizeMode=GrowAndShrink'
            $firstRadio = $true
            foreach ($ck in $computersWithFailures) {
                $dn = if ($ck -eq "") { $env:COMPUTERNAME } else { Get-ComputerNodeLabel -ComputerName $ck }
                $fc = $results[$ck].FailCount
                $radio          = New-Object System.Windows.Forms.RadioButton
                $radio.Text     = "$dn ($fc errors)"
                $radio.Tag      = $ck
                $radio.AutoSize = $true
                $radio.Margin   = [System.Windows.Forms.Padding]::new(5, 5, 15, 5)
                $radio.Checked  = $firstRadio
                if ($firstRadio) { $selectedComputer.Value = $ck; $firstRadio = $false }
                $radio.Add_CheckedChanged({
                    param($s, $e)
                    if ($s.Checked) { $selectedComputer.Value = $s.Tag; & $populateRtbForComputer $detailsRtb $s.Tag $results $writeRtf $writeFailedItem }
                }.GetNewClosure())
                $radioFlow.Controls.Add($radio)
            }
        }
        elseif ($computersWithFailures.Count -eq 1) { $selectedComputer.Value = $computersWithFailures[0] }

        # Details
        $detailsRtb = gen $resultForm "RichTextBox" 'Dock=Fill' 'Font=Consolas, 9' 'ReadOnly=$true' 'WordWrap=$true' 'BackColor=White'

        # Button panel
        $buttonPanel           = gen $resultForm "Panel" 'Dock=Bottom' 'Height=50'
        $buttonFlowPanelResult = gen $buttonPanel "FlowLayoutPanel" 'Dock=Fill' 'Padding=10 10 10 10' 'FlowDirection=RightToLeft' 'WrapContents=$false'
        $closeButton           = gen $buttonFlowPanelResult "Button" "Close" 0 0 100 30 'Margin=5 0 5 0'
        $closeButton.Add_Click({ $resultForm.Close() })
        $forceRetryButton         = gen $buttonFlowPanelResult "Button" "Force ACL and retry" 0 0 140 30 'FlatStyle=Flat' 'Margin=5 0 5 0' 'BackColor=220 53 69' 'ForeColor=White'
        $forceRetryButton.Visible = ($totalFail -gt 0)
        $forceRetryButton.Tag     = @{
            Results = $results; ComputerOrder = $computerOrder; SelectedComputer = $selectedComputer
            SummaryLabel = $summaryLabel; DetailsRtb = $detailsRtb; SelectorPanel = $selectorPanel
            WriteRtf = $writeRtf; WriteFailedItem = $writeFailedItem; PopulateRtb = $populateRtbForComputer
            ActiveTabIndex = $activeTabIndex
        }

        # ═══════════════════════════════════════════════════════════
        # FORCE ACL AND RETRY HANDLER
        # ═══════════════════════════════════════════════════════════
        $forceRetryButton.Add_Click({
            $tag      = $this.Tag
            $compKey  = $tag.SelectedComputer.Value
            $results  = $tag.Results
            $rtb      = $tag.DetailsRtb
            $writeRtf = $tag.WriteRtf; $writeFailedItem = $tag.WriteFailedItem
            if (-not $results.ContainsKey($compKey) -or $results[$compKey].FailCount -eq 0) { return }
            $this.Enabled = $false; $this.Text = "Processing..."
            $failedItems  = @($results[$compKey].FailedItems)
            $isRemote     = ($compKey -ne "" -and $compKey -ne $env:COMPUTERNAME)
            $credential   = if ($isRemote) { Get-CredentialFromPanel } else { $null }

            # ── Self-contained Force ACL scriptblock (local + Invoke-Command) ──
            $forceAclScript = {
                param([array]$FailedItems)
                $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
                function Parse-Reg { param([string]$p)
                    $p = $p -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''; $r = $null; $s = $null
                    if     ($p -match '^HKLM[:\\]+(.+)$')           { $r = [Microsoft.Win32.Registry]::LocalMachine; $s = $matches[1] }
                    elseif ($p -match '^HKCU[:\\]+(.+)$')           { $r = [Microsoft.Win32.Registry]::CurrentUser;  $s = $matches[1] }
                    elseif ($p -match '^HKU[:\\]+(.+)$')            { $r = [Microsoft.Win32.Registry]::Users;        $s = $matches[1] }
                    elseif ($p -match '^HKCR[:\\]+(.+)$')           { $r = [Microsoft.Win32.Registry]::ClassesRoot;  $s = $matches[1] }
                    elseif ($p -match '^HKEY_LOCAL_MACHINE\\(.+)$') { $r = [Microsoft.Win32.Registry]::LocalMachine; $s = $matches[1] }
                    elseif ($p -match '^HKEY_CURRENT_USER\\(.+)$')  { $r = [Microsoft.Win32.Registry]::CurrentUser;  $s = $matches[1] }
                    elseif ($p -match '^HKEY_USERS\\(.+)$')         { $r = [Microsoft.Win32.Registry]::Users;        $s = $matches[1] }
                    elseif ($p -match '^HKEY_CLASSES_ROOT\\(.+)$')  { $r = [Microsoft.Win32.Registry]::ClassesRoot;  $s = $matches[1] }
                    if ($s) { $s = $s.TrimStart('\') }; return @{ Root = $r; SubPath = $s }
                }
                function Set-FsAcl { param($path, [bool]$isDir, $usr)
                    $rts = [System.Security.AccessControl.FileSystemRights]::FullControl
                    $inh = if ($isDir) { [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit } else { [System.Security.AccessControl.InheritanceFlags]::None }
                    $inf = if ($isDir) { [System.IO.DirectoryInfo]::new($path) } else { [System.IO.FileInfo]::new($path) }
                    if (-not $isDir) { $inf.Attributes = [System.IO.FileAttributes]::Normal }
                    $a = $inf.GetAccessControl(); $a.SetOwner($usr)
                    $a.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new($usr, $rts, $inh, [System.Security.AccessControl.PropagationFlags]::None, [System.Security.AccessControl.AccessControlType]::Allow))
                    $inf.SetAccessControl($a)
                }
                function Set-RegAcl { param($root, $sub, $usr)
                    $ar = [System.Security.AccessControl.RegistryAccessRule]::new($usr, [System.Security.AccessControl.RegistryRights]::FullControl, [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit, [System.Security.AccessControl.PropagationFlags]::None, [System.Security.AccessControl.AccessControlType]::Allow)
                    $k = $root.OpenSubKey($sub, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, [System.Security.AccessControl.RegistryRights]::TakeOwnership)
                    if ($k) { $a = $k.GetAccessControl(); $a.SetOwner($usr); $k.SetAccessControl($a); $k.Close() }
                    $k = $root.OpenSubKey($sub, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, [System.Security.AccessControl.RegistryRights]::ChangePermissions)
                    if ($k) { $a = $k.GetAccessControl(); $a.AddAccessRule($ar); $k.SetAccessControl($a); $k.Close() }
                }
                $retryResults = [System.Collections.ArrayList]::new()
                foreach ($item in $FailedItems) {
                    $ok = $false; $err = $null
                    try {
                        switch ($item.Type) {
                            'Folder' {
                                if ([System.IO.Directory]::Exists($item.Path)) {
                                    Set-FsAcl $item.Path $true $user
                                    foreach ($f in [System.IO.Directory]::GetFiles($item.Path, '*', [System.IO.SearchOption]::AllDirectories))       { try { Set-FsAcl $f $false $user } catch { } }
                                    foreach ($d in [System.IO.Directory]::GetDirectories($item.Path, '*', [System.IO.SearchOption]::AllDirectories)) { try { Set-FsAcl $d $true  $user } catch { } }
                                    [System.IO.Directory]::Delete($item.Path, $true)
                                }; $ok = $true
                            }
                            'File' {
                                if ([System.IO.File]::Exists($item.Path)) { Set-FsAcl $item.Path $false $user; [System.IO.File]::Delete($item.Path) }; $ok = $true
                            }
                            { $_ -in @('Registry', 'RegistryValue', 'UninstallEntry') } {
                                $pr = Parse-Reg $item.Path; if ($null -eq $pr.Root) { throw "Invalid registry path" }
                                $k = $pr.Root.OpenSubKey($pr.SubPath, $false)
                                if ($k) {
                                    $k.Close(); Set-RegAcl $pr.Root $pr.SubPath $user
                                    $ke = $pr.Root.OpenSubKey($pr.SubPath, $false)
                                    if ($ke) { foreach ($sn in $ke.GetSubKeyNames()) { try { Set-RegAcl $pr.Root "$($pr.SubPath)\$sn" $user } catch { } }; $ke.Close() }
                                    if ($item.Type -eq 'RegistryValue' -and $item.RegistryValue) {
                                        $k = $pr.Root.OpenSubKey($pr.SubPath, $true); if ($k) { $k.DeleteValue($item.RegistryValue, $false); $k.Close() }
                                    } else { $pr.Root.DeleteSubKeyTree($pr.SubPath, $false) }
                                }; $ok = $true
                            }
                        }
                    } catch { $err = $_.Exception.Message }
                    [void]$retryResults.Add(@{ Type = $item.Type; Path = $item.Path; Success = $ok; Error = $err; RegistryValue = $item.RegistryValue })
                }
                return @($retryResults)
            }

            $rtb.Clear()
            $targetLabel = if ($isRemote) { " ON $compKey" } else { "" }
            & $writeRtf $rtb "FORCE ACL RETRY$targetLabel IN PROGRESS...`r`n" ([System.Drawing.Color]::DarkOrange) $true
            & $writeRtf $rtb (("-" * 60) + "`r`n`r`n") ([System.Drawing.Color]::Gray) $false
            $rtb.Refresh()

            $retryResults = $null
            if ($isRemote) {
                & $writeRtf $rtb "Sending $($failedItems.Count) item(s) to $compKey...`r`n`r`n" ([System.Drawing.Color]::Black) $false; $rtb.Refresh()
                try {
                    $ip = @{ ComputerName = $compKey; ScriptBlock = $forceAclScript; ArgumentList = @(, $failedItems); ErrorAction = 'Stop' }
                    if ($credential) { $ip.Credential = $credential }
                    $retryResults = Invoke-Command @ip
                }
                catch {
                    & $writeRtf $rtb "Remote execution error : $($_.Exception.Message)`r`n" ([System.Drawing.Color]::Red) $true
                    Write-Log "Force ACL remote error on $compKey : $($_.Exception.Message)" -Level Error
                }
            }
            else { $retryResults = & $forceAclScript $failedItems }

            $retrySc = 0; $retryFc = 0; $stillFailed = [System.Collections.ArrayList]::new()
            if ($retryResults) {
                foreach ($rr in $retryResults) {
                    & $writeRtf $rtb "Processing : $($rr.Path)`r`n" ([System.Drawing.Color]::Black) $false
                    if ($rr.Success) {
                        $retrySc++
                        & $writeRtf $rtb "  -> SUCCESS`r`n" ([System.Drawing.Color]::DarkGreen) $false
                        Write-Log "Force ACL retry succeeded : $($rr.Path)"
                    }
                    else {
                        $retryFc++
                        [void]$stillFailed.Add(@{ Type = $rr.Type; Path = $rr.Path; Error = $rr.Error; RegistryValue = $rr.RegistryValue })
                        & $writeRtf $rtb "  -> FAILED : $($rr.Error)`r`n" ([System.Drawing.Color]::Red) $false
                        Write-Log "Force ACL retry failed : $($rr.Path) - $($rr.Error)" -Level Error
                    }
                    $rtb.ScrollToCaret(); $rtb.Refresh()
                }
            }

            # Update per-computer results
            $results[$compKey] = @{ SuccessCount = $results[$compKey].SuccessCount + $retrySc; FailCount = $retryFc; FailedItems = @($stillFailed) }

            # Retry summary
            & $writeRtf $rtb ("`r`n" + ("-" * 60) + "`r`n") ([System.Drawing.Color]::Gray) $false
            & $writeRtf $rtb "RETRY SUMMARY : $retrySc succeeded, $retryFc failed`r`n" ([System.Drawing.Color]::Black) $true
            if ($stillFailed.Count -gt 0) {
                & $writeRtf $rtb ("`r`n" + ("-" * 60) + "`r`n") ([System.Drawing.Color]::Gray) $false
                & $writeRtf $rtb "STILL FAILED ITEMS ($($stillFailed.Count))`r`n`r`n" ([System.Drawing.Color]::Red) $true
                foreach ($f in $stillFailed) { & $writeFailedItem $rtb $f $writeRtf; & $writeRtf $rtb "`r`n" ([System.Drawing.Color]::Black) $false }
            }
            $rtb.Select(0, 0); $rtb.ScrollToCaret()

            # Recalculate global totals
            $newSuccess = 0; $newFail = 0
            foreach ($ck in $tag.ComputerOrder) { if ($results.ContainsKey($ck)) { $newSuccess += $results[$ck].SuccessCount; $newFail += $results[$ck].FailCount } }
            $tag.SummaryLabel.Text      = "Force ACL retry completed.`n`nSuccessful : $newSuccess`nStill failed : $newFail"
            $tag.SummaryLabel.ForeColor = if ($newFail -eq 0) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkOrange }

            # Update selector radios
            if ($tag.SelectorPanel) {
                foreach ($ctrl in $tag.SelectorPanel.Controls) {
                    if ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.AutoScroll) { $ctrl = $ctrl.Controls[0] }
                    if ($ctrl -is [System.Windows.Forms.FlowLayoutPanel]) {
                        foreach ($radio in $ctrl.Controls) {
                            if ($radio -is [System.Windows.Forms.RadioButton] -and $results.ContainsKey($radio.Tag)) {
                                $rfc = $results[$radio.Tag].FailCount
                                $rdn = if ($radio.Tag -eq "") { $env:COMPUTERNAME } else { Get-ComputerNodeLabel -ComputerName $radio.Tag }
                                if ($rfc -gt 0) { $radio.Text = "$rdn ($rfc errors)"; $radio.ForeColor = [System.Drawing.Color]::Black }
                                else            { $radio.Text = "$rdn (OK)";           $radio.ForeColor = [System.Drawing.Color]::Green }
                            }
                        }
                    }
                }
            }

            if ($newFail -eq 0) { $this.Visible = $false } else { $this.Enabled = $true; $this.Text = "Force ACL and retry" }
            Write-Log "Force ACL retry completed : $retrySc succeeded, $retryFc still failed"
            switch ($tag.ActiveTabIndex) { 0 { $script:CompareTabCacheVersion = -1 }; 1 { $script:FullCacheTabCacheVersion = -1 }; 3 { $script:CompareTabCacheVersion = -1 } }
        })

        # Populate initial RTB content
        & $populateRtbForComputer $detailsRtb $selectedComputer.Value $results $writeRtf $writeFailedItem

        # Z-order
        $resultForm.Controls.SetChildIndex($detailsRtb, 0)
        if ($selectorPanel) { $resultForm.Controls.SetChildIndex($selectorPanel, 1); $resultForm.Controls.SetChildIndex($summaryPanel, 2) }
        else                { $resultForm.Controls.SetChildIndex($summaryPanel, 1) }
        $resultForm.Controls.SetChildIndex($buttonPanel, $resultForm.Controls.Count - 1)
        $resultForm.AcceptButton = $closeButton
        $resultForm.ShowDialog($form) | Out-Null

        # Refresh active tab
        switch ($activeTabIndex) {
            0 { $searchButton_MSICleanupTab.PerformClick() }
            1 { $script:FullCacheTabCacheVersion = -1; Update-FullCacheTab }
            3 { $script:CompareTabCacheVersion = -1; Update-CompareTab }
        }
    }
    $cleanupTimer.Start()
})

# Uninstall job monitor timer
$uninstallTimer          = New-Object System.Windows.Forms.Timer
$uninstallTimer.Interval = 500
$uninstallTimer.Add_Tick({
    $completedJobs = @()
    foreach ($jobId in $script:UninstallJobs.Keys) {
        $jobInfo  = $script:UninstallJobs[$jobId]
        $syncHash = $jobInfo.SyncHash
        if (-not $syncHash.IsComplete) {
            Update-ProcessMonitorDisplay -JobId $jobId
            # Show log button during uninstall when log file is detected
            if ($syncHash.FoundLogPath -and $jobInfo.Panel -and $jobInfo.Panel.Tag.ShowLogDuringUninstall) {
                $logBtn = $jobInfo.Panel.Tag.ShowLogDuringUninstall
                if (-not $logBtn.Visible) {
                    $logBtn.Tag     = @{ LogPath = $syncHash.FoundLogPath; Computer = if ($jobInfo.Panel.Tag.IsRemote) { $jobInfo.Computer } else { "" } }
                    $logBtn.Visible = $true
                    Write-Log "Log file detected during uninstall : $($syncHash.FoundLogPath)"
                }
            }
        }
        if ($syncHash.IsComplete) {
            $completedJobs += $jobId
            $success      = $syncHash.Success
            $exitCode     = $syncHash.ExitCode
            $errorMessage = $syncHash.Error
            $effectiveLogPath = $jobInfo.LogPath
            if (-not $effectiveLogPath -and $syncHash.FoundLogPath) { $effectiveLogPath = $syncHash.FoundLogPath }
            Write-Log "Uninstall job completed : JobId=$jobId, ExitCode=$exitCode, Success=$success"
            Update-UninstallPanelComplete -ProductPanel $jobInfo.Panel -Success $success -ExitCode $exitCode -ErrorMessage $errorMessage -Entry $jobInfo.Entry -PanelKey $jobInfo.PanelKey -LogPath $effectiveLogPath -ComputerName $jobInfo.Computer
            Test-UninstallPanelLogExists -PanelKey $jobInfo.PanelKey
            try {
                $jobInfo.PowerShell.EndInvoke($jobInfo.AsyncResult)
                $jobInfo.PowerShell.Dispose()
                $jobInfo.Runspace.Close()
                $jobInfo.Runspace.Dispose()
            }
            catch { Write-Log "Error disposing runspace : $_" -Level Warning }
        }
    }
    foreach ($jobId in $completedJobs) { $script:UninstallJobs.Remove($jobId) }
})
$uninstallTimer.Start()

#endregion Tab 4 : MSI CLEANUP

#region EXPORT

function Show-NonBlockingMessage {
    param([string]$message, [string]$title = "Information", [int]$timeout = 0)
    Add-Type -AssemblyName System.Windows.Forms ; Add-Type -AssemblyName System.Drawing
    $successform = New-Object System.Windows.Forms.Form ; $successform.Text = $title ; $successform.Size = [Drawing.Size]::new(300, 150) ; $successform.StartPosition="CenterScreen" ; $successform.TopMost=$true
    $label       = New-Object System.Windows.Forms.Label ; $label.Text=$message ; $label.AutoSize=$true ; $label.Location=[Drawing.Point]::new(20,20) ; $null=$successform.Controls.Add($label)
    $button      = New-Object System.Windows.Forms.Button ; $button.Text="OK" ; $button.Size=[Drawing.Size]::new(75,30) ; $button.Location=[Drawing.Point]::new(110,70) ; $button.Add_Click({ param($csender,$cargs) ($csender.FindForm()).Close() }) ; $null=$successform.Controls.Add($button)
    if ($timeout -gt 0) {
        $timer=New-Object System.Windows.Forms.Timer ; $timer.Interval=$timeout*1000 ; $timer.Tag=$successform
        $timer.Add_Tick({ param($tsender,$targs) $tsender.Stop() ; ($tsender.Tag).Close() })
        $timer.Start()
    }
    $successform.Show()
}

function ListExport {
    param(
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.ListView]$listView,
        [Parameter(Mandatory=$true)]
        [bool]$all
    )
    $itemsToExport = if ($all) { $listView.Items } else { $listView.SelectedItems }
    if ($itemsToExport.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Nothing to export.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    $formatForm                 = New-Object System.Windows.Forms.Form
    $formatForm.Text            = "Export" ; $formatForm.Width = 300 ; $formatForm.Height = 150
    $formatForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formatForm.StartPosition   = "CenterScreen"
    $formatForm.MinimizeBox = $false ; $formatForm.MaximizeBox = $false ; $formatForm.ShowIcon = $false ; 
    $label = New-Object System.Windows.Forms.Label ; $label.Text = "Choose an export format :" ; $label.AutoSize = $true ; $label.Top = 10 ; $label.Left = 10
    $formatForm.Controls.Add($label)
    $comboBox = New-Object System.Windows.Forms.ComboBox ; $comboBox.Top = 30 ; $comboBox.Left = 10 ; $comboBox.Width = 100
    $comboBox.Items.Add("OGV Hybrid") ; $comboBox.Items.Add("XLSX") ; $comboBox.Items.Add("CSV") ; 
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList ; $comboBox.SelectedIndex = 0
    $formatForm.Controls.Add($comboBox)
    $openafterbtn = New-Object System.Windows.Forms.Checkbox ; $openafterbtn.Text = "Open after export" ; $openafterbtn.AutoSize = $true ; $openafterbtn.Top = 30 ; $openafterbtn.Left = 120
    $formatForm.Controls.Add($openafterbtn)
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Export" ; $okButton.Top = 55 ; $okButton.Left = 10 ; $okButton.Enabled = $true
    $formatForm.Controls.Add($okButton)
    $cancelButton = New-Object System.Windows.Forms.Button ; $cancelButton.Text = "Cancel" ; $cancelButton.Top = 55 ; $cancelButton.Left = 85 ; $cancelButton.Add_Click({ $formatForm.Close() })
    $formatForm.Controls.Add($cancelButton)
    $ExProgress = New-Object System.Windows.Forms.ProgressBar
    $ExProgress.Top = 80 ; $ExProgress.Left = 10 ; $ExProgress.Width = 260 ; $ExProgress.Style = 'Continuous' ; $ExProgress.Minimum = 0 ; $ExProgress.Maximum = 100 ; $ExProgress.Value = 0 ; $ExProgress.Visible = $false
    $formatForm.Controls.Add($ExProgress)
    $okButton.Add_Click({
        $okButton.Enabled = $false
        $comboBox.Enabled = $false
        $columns = $listView.Columns
        $data = New-Object System.Collections.Generic.List[System.Object[]]
        foreach ($item in $itemsToExport) {
            $row = @($null) * $columns.Count
            foreach ($col in $columns) { $index = $col.Index ; if ($item.SubItems.Count -gt $index) { $row[$index] = $item.SubItems[$index].Text } else { $row[$index] = "" } }
            $data.Add($row)
        }
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $dateTimeString = (Get-Date -Format "yyyy-MM-dd_HH-mm")
        switch ($comboBox.SelectedItem) {
            "OGV Hybrid"  { $saveFileDialog.Filter = "Batch (*.bat)|*.bat"  ; $saveFileDialog.FileName = "Export_MSI_$dateTimeString.bat" }
            "XLSX"        { $saveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx"; $saveFileDialog.FileName = "Export_MSI_$dateTimeString.xlsx" }
            "CSV"         { $saveFileDialog.Filter = "CSV (*.csv)|*.csv"    ; $saveFileDialog.FileName = "Export_MSI_$dateTimeString.csv" }
        }
        $saveFileDialog.Title = "Select a path for export"
        $result = $saveFileDialog.ShowDialog()        
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { $okButton.Enabled=$true ; $comboBox.Enabled=$true ; return }
        $filePath = $saveFileDialog.FileName
        $ExProgress.Visible = $true
        [System.Windows.Forms.Application]::DoEvents()
        $totalSteps = $data.Count
        $progressStep = [int](100 / $totalSteps)
        $currentProgress = 0
        switch ($comboBox.SelectedItem) {
            "OGV Hybrid" {
                $headers = $columns | ForEach-Object { $_.Text }
                $objects = New-Object System.Collections.ArrayList
                foreach ($row in $data) {
                    $obj = [ordered]@{}
                    for ($i = 0; $i -lt $headers.Count; $i++) { $obj[$headers[$i]] = $row[$i] }
                    $objects.Add([PSCustomObject]$obj) | Out-Null
                    $currentProgress += $progressStep
                    $ExProgress.Value = [int][math]::Min($currentProgress, 100)
                    [System.Windows.Forms.Application]::DoEvents()
                }
                $scriptContent = @"
<# ::
    cls & @echo off & title Export_MSI_$dateTimeString
    copy /y "%~f0" "%TEMP%\%~n0.ps1" >NUL && powershell -Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TEMP%\%~n0.ps1"
    exit /b
#>

`$jsonData = @'
$(($objects | ConvertTo-Json))
'@

`$data = ConvertFrom-Json -InputObject `$jsonData
`$data | Select-Object '$($headers -join "','")' | Out-GridView -Title 'Export_MSI_$dateTimeString' -Wait
"@
                [System.IO.File]::WriteAllLines($filePath, $scriptContent, [System.Text.Encoding]::UTF8)
            }
            "XLSX" {
                $headers   = $columns | ForEach-Object { $_.Text }
                $dataArray = $data.ToArray()
                Export-ToXlsx -Path $filePath -Data $dataArray -Columns $headers -ProgressBar $ExProgress
            }
            "CSV" {
                $headers  = ($columns | ForEach-Object { $_.Text }) -join ";"
                $csvLines = New-Object System.Collections.Generic.List[string]
                $csvLines.Add($headers)
                foreach ($row in $data) { $csvLines.Add(($row -join ";")) ; $currentProgress+=$progressStep ; $ExProgress.Value=[int][math]::Min($currentProgress, 100) ; [System.Windows.Forms.Application]::DoEvents() }
                [System.IO.File]::WriteAllLines($filePath, $csvLines)
            }
        }
        $ExProgress.Value = 100
        [System.Windows.Forms.Application]::DoEvents()
        if ($openafterbtn.Checked) { Start-Process -FilePath $filePath }
        $formatForm.Close()
        Show-NonBlockingMessage -message "Export Complete" -title "Success" -timeout 2
    })
    $formatForm.ShowDialog()
}

function Export-ToXlsx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][Object[]]$Data,
        [Parameter(Mandatory=$false)][string[]]$Columns,
        [Parameter(Mandatory=$false)]$ProgressBar
    )
    [void][Reflection.Assembly]::LoadWithPartialName("WindowsBase")
    if ([System.IO.File]::Exists($Path)) {Remove-Item $Path}
    if (-not $Columns -and $Data.Count -gt 0) { 
        $Columns = if ($Data[0] -is [System.Management.Automation.PSObject]) {($Data[0]|Get-Member -MemberType NoteProperty,Property|Select-Object -ExpandProperty Name)} else {"Value"} 
    }
    $Package=[System.IO.Packaging.Package]::Open($Path,[System.IO.FileMode]::Create,[System.IO.FileAccess]::ReadWrite)
    $Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $RelNamespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    $AddCell = {
        param($Doc,$ParentNode,$Namespace,$ColIndex,$RowIndex,$Value)
        if ($null -eq $Value) {$Value = ""}
        $Cell=$Doc.CreateElement("c",$Namespace)
        $Cell.SetAttribute("r",([char](65+$ColIndex-1))+[string]$RowIndex)
        $Cell.SetAttribute("t","inlineStr")
        $InlineStr=$Doc.CreateElement("is",$Namespace)
        $TextNode=$Doc.CreateElement("t",$Namespace)
        $TextNode.InnerText=$Value
        $InlineStr.AppendChild($TextNode)|Out-Null
        $Cell.AppendChild($InlineStr)|Out-Null
        $ParentNode.AppendChild($Cell)|Out-Null
    }
    $WorkbookUri=[Uri]"/xl/workbook.xml"
    $WorkbookPart=$Package.CreatePart($WorkbookUri,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    $WorkbookXml=[xml]"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><workbook xmlns='$Namespace' xmlns:r='$RelNamespace'><sheets><sheet name='Sheet1' sheetId='1' r:id='rId1'/></sheets></workbook>"
    $WorkbookXml.Save($WorkbookPart.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write))
    $Package.CreateRelationship($WorkbookUri,[System.IO.Packaging.TargetMode]::Internal,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument","rId1")|Out-Null
    $SheetUri=[Uri]"/xl/worksheets/sheet1.xml"
    $SheetPart=$Package.CreatePart($SheetUri,"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    $SheetXml=[xml]"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><worksheet xmlns='$Namespace' xmlns:r='$RelNamespace'><sheetData/></worksheet>"
    $NamespaceManager=New-Object System.Xml.XmlNamespaceManager($SheetXml.NameTable)
    $NamespaceManager.AddNamespace("x",$Namespace)
    $SheetData=$SheetXml.SelectSingleNode("//x:sheetData",$NamespaceManager)
    $RowIndex=1
    $HeaderRow=$SheetXml.CreateElement("row",$Namespace)
    $HeaderRow.SetAttribute("r",[string]$RowIndex)
    $SheetData.AppendChild($HeaderRow)|Out-Null
    $ColIndex=1
    $Columns|ForEach-Object{&$AddCell $SheetXml $HeaderRow $Namespace $ColIndex $RowIndex $_;$ColIndex++}
    $RowIndex++
    $progressIncrement=100/$Data.Count
    $currentProgress=0
    foreach($RowData in $Data){
        $Row=$SheetXml.CreateElement("row",$Namespace)
        $Row.SetAttribute("r",[string]$RowIndex)
        $SheetData.AppendChild($Row)|Out-Null
        for($ColIndex=1;$ColIndex -le $Columns.Count;$ColIndex++){
            $Header=$Columns[$ColIndex-1]
            $CellValue=if($RowData -is [System.Management.Automation.PSObject]){($RowData|Select-Object -ExpandProperty $Header -ErrorAction SilentlyContinue)}else{$RowData[$ColIndex-1]}
            &$AddCell $SheetXml $Row $Namespace $ColIndex $RowIndex $CellValue
        }
        $RowIndex++
        if ($ProgressBar) {
            $currentProgress+=$progressIncrement
            $ProgressBar.Value=[int][math]::Min($currentProgress,100)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    $SheetXml.Save($SheetPart.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write))
    ($Package.GetPart($WorkbookUri)).CreateRelationship($SheetUri,[System.IO.Packaging.TargetMode]::Internal,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet","rId1")|Out-Null
    $Package.Close()
}


#region GOBAL EVENTS

$launch_progressBar.Value = 90

ConfigureListViewContextMenu -listView $listView_Explore
ConfigureListViewContextMenu -listView $listView_Registry

$form.add_OnWindowMessage({
    param($s, $m)
    $updateColumns = {
        param($listView)
        $currentWidth = $listView.ClientSize.Width
        if (-not $script:lastWidth[$listView] -or $script:lastWidth[$listView] -ne $currentWidth) {
            $script:lastWidth[$listView] = $currentWidth
            AdjustListViewColumns -listView $listView
        }
    }
    $updateUninstallPanels = {
        if (-not $uninstallFlowPanel_MSICleanupTab.IsHandleCreated -or $uninstallFlowPanel_MSICleanupTab.Controls.Count -eq 0) { return }
        if ($script:RecalcSelectorLayout) { & $script:RecalcSelectorLayout }
        $newWidth = $uninstallFlowPanel_MSICleanupTab.ClientSize.Width - 45
        foreach ($ctrl in $uninstallFlowPanel_MSICleanupTab.Controls) {
            if ($ctrl -is [System.Windows.Forms.Panel]) {
                $isWrapper  = $false
                $childPanel = $null
                foreach ($innerCtrl in $ctrl.Controls) {
                    if ($innerCtrl -is [System.Windows.Forms.Panel] -and $innerCtrl.Tag -and $innerCtrl.Tag.Entry) {
                        $isWrapper  = $true
                        $childPanel = $innerCtrl
                        break
                    }
                }
                if ($isWrapper -and $childPanel) {
                    $ctrl.Width             = $newWidth
                    $childPanelWidth        = $newWidth - 30
                    $childPanel.Width       = $childPanelWidth
                    $childPanel.MinimumSize = [System.Drawing.Size]::new($childPanelWidth, $childPanel.Height)
                    $childPanel.MaximumSize = [System.Drawing.Size]::new($childPanelWidth, $childPanel.Height)
                    Update-UninstallPanelTextBoxWidths -Panel $childPanel -EffectiveWidth $childPanelWidth
                }
                elseif ($ctrl.Tag -and $ctrl.Tag.Entry) {
                    $ctrl.Width       = $newWidth
                    $ctrl.MinimumSize = [System.Drawing.Size]::new($newWidth, $ctrl.Height)
                    $ctrl.MaximumSize = [System.Drawing.Size]::new($newWidth, $ctrl.Height)
                    Update-UninstallPanelTextBoxWidths -Panel $ctrl -EffectiveWidth $newWidth
                }
            }
        }
    }
    $processAction = {
        foreach ($listView in @($listView_Explore, $listView_Registry, $listView_Tab1Props, $listView_Tab1Features)) {
            if ($listView.IsHandleCreated -and $listView.Visible) { & $updateColumns $listView }
        }
        & $updateUninstallPanels
    }
    switch ($m.Msg) {
        0x0233 {
            $files = [DragDropFix]::GetDroppedFiles($m.WParam)
            $validFiles = @($files | Where-Object { $_ -match "\.(msi|msix|mst|msp)$" })
            if ($validFiles.Count -gt 0) {
                $script:fromBrowseButton = $true
                $textBoxPath_Tab1.Text = $validFiles -join ";"
                foreach ($file in $validFiles) { Invoke-MsiLoad -MsiPath $file }
            }
            $m.Result = [IntPtr]::Zero
            return
        }
        0x231  {
            $script:resizePending = $true
            if ($rightPanel_Tab1.IsHandleCreated) { [NativeMethods]::SendMessage($rightPanel_Tab1.Handle, 0x000B, 0, 0) }
        }
        0x232  {
            if ($script:resizePending) {
                $script:resizePending = $false
                if ($rightPanel_Tab1.IsHandleCreated) {
                    [NativeMethods]::SendMessage($rightPanel_Tab1.Handle, 0x000B, 1, 0)
                    $rightPanel_Tab1.Invalidate($true)
                    $rightPanel_Tab1.Update()
                }
                Update-ButtonPositions
                & $processAction
            }
        }
        0x0005 { if (-not $script:resizePending) { & $processAction } }
    }
})

$form.Add_Resize({
    $btnAbout.Location = [System.Drawing.Point]::new(
        [int](($form.ClientSize.Width - $btnAbout.Width) / 2),
        [int](($titleBarHeight - $btnAbout.Height) / 2)
    )
})

$tabControl.Add_SelectedIndexChanged({
    $isTab4 = ($tabControl.SelectedTab -eq $tabPage4)
    $tabButtonPanel_MSICleanupTab.Visible   = $isTab4
    $fastModeCheckBox_MSICleanupTab.Visible = $isTab4
    $splitContainerSearch_MSICleanupTab.SplitterDistance = [int]($splitContainerSearch_MSICleanupTab.Width * 0.6)
    if ($isTab4) { $tabButtonPanel_MSICleanupTab.BringToFront(); $panel_RemoteTarget.BringToFront() }
    if ($tabControl.SelectedTab -eq $tabPage3) {
        $script:currentTabIndex = 2
        Update-RemoteTargetPanelVisibility
        AdjustListViewColumns -listView $listView_Registry
        return
    }
    if ($tabControl.SelectedTab -eq $tabPage1) { Update-ButtonPositions                             ; $script:currentTabIndex = 0 ; Update-RemoteTargetPanelVisibility }
    if ($tabControl.SelectedTab -eq $tabPage2) { AdjustListViewColumns -listView $listView_Explore  ; $script:currentTabIndex = 1 ; Update-RemoteTargetPanelVisibility }
    if ($tabControl.SelectedTab -eq $tabPage3) { AdjustListViewColumns -listView $listView_Registry ; $script:currentTabIndex = 2 ; Update-RemoteTargetPanelVisibility }
    if ($tabControl.SelectedTab -eq $tabPage4) { $tabButtonPanel_MSICleanupTab.BringToFront(); $fastModeCheckBox_MSICleanupTab.BringToFront() ; $script:currentTabIndex = 3 ; Update-RemoteTargetPanelVisibility }
})

$form.Add_Load({
    $launch_progressBar.Value = 95
    $loadingLabel.Text        = "Finalizing..."
    $form.MinimumSize         = New-Object System.Drawing.Size(1122, 600)
})

$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

$form.Add_Shown({
    $tabControl.Location = [System.Drawing.Point]::new(0, $titleBarHeight)
    $tabControl.Size     = [System.Drawing.Size]::new($form.ClientSize.Width, ($form.ClientSize.Height - $titleBarHeight))
    Update-ButtonPositions
    Update-RemoteTargetPanelVisibility
    $isTab4 = ($tabControl.SelectedTab -eq $tabPage4)
    $tabButtonPanel_MSICleanupTab.Visible   = $isTab4
    $fastModeCheckBox_MSICleanupTab.Visible = $isTab4
    $launch_progressBar.Value = 100
    $loadingLabel.Text        = "Complete"
    $loadingForm.Close()
    $form.Activate()
    $form.BringToFront()
    [DwmHelper]::SetRoundedCorners($form)
    [DragDropFix]::Enable($form.Handle)
    Set-CustomPlaceholder -TextBox $textBox_TargetDevice   -PlaceholderText "Target Devices (empty = local)"
    Set-CustomPlaceholder -TextBox $textBox_CredentialID   -PlaceholderText "Credential ID"
    Set-PasswordPlaceholderState -ShowPlaceholder $true
    Set-CustomPlaceholder -TextBox $searchTextBox_Tab3 -PlaceholderText "Filter (empty = search everything)"
    $uninstallFlowPanel_MSICleanupTab.GetType().GetMethod('OnResize', [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($uninstallFlowPanel_MSICleanupTab, @([System.EventArgs]::Empty))
})

$form.Add_Click(    { $form.ActiveControl = $null })
$tabPage1.Add_Click({ $form.ActiveControl = $null })
$tabPage2.Add_Click({ $form.ActiveControl = $null })
$tabPage3.Add_Click({ $form.ActiveControl = $null })
$tabPage4.Add_Click({ $form.ActiveControl = $null })

function Invoke-ApplicationCleanup {
    Write-Log "Performing application cleanup"
    try   { Remove-ScheduledTrustedHosts }
    catch { Write-Log "Error cleaning TrustedHosts : $_"    -Level Error }
    try {
        foreach ($target in $script:SmbSessionsToCleanup) {
            $smbResult  = [NetSession.Native]::DisconnectSmb($target)
            $credResult = [NetSession.Native]::DeleteDomainCredential($target)
            Write-Log "Session cleanup for $target : smb=$smbResult, cred=$credResult"
        }
        $script:SmbSessionsToCleanup.Clear()
    }
    catch { Write-Log "Error cleaning sessions : $_" -Level Error }
    try   { if ($script:SecurePassword) { $script:SecurePassword.Dispose() } }
    catch { Write-Log "Error disposing SecurePassword : $_" -Level Error }
    try {
        if ($script:WindowsInstallerCOM) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:WindowsInstallerCOM)
            $script:WindowsInstallerCOM = $null
            Write-Log "WindowsInstaller COM object released"
        }
    }
    catch { Write-Log "Error releasing WindowsInstaller COM : $_" -Level Error }
    # ── Remove TaskBar shortcut if user did not actually pin the app ──
    $taskbarLnk = [System.IO.Path]::Combine($script:TaskbarPinDir, $script:LnkName)
    if ([System.IO.File]::Exists($taskbarLnk)) {
        $actuallyPinned = Test-AppPinned -Target Taskbar
        if (-not $actuallyPinned) {
            try {
                [System.IO.File]::Delete($taskbarLnk)
                Write-Log "Cleaned up TaskBar shortcut (user did not pin)"
            }
            catch { Write-Log "Failed to clean up TaskBar shortcut : $_" -Level Warning }
        }
        else {
            Write-Log "TaskBar shortcut kept (user has pinned the app)"
        }
    }
    # ── Clean up temporary Start Menu shortcut ──
    $startMenuLnk = [System.IO.Path]::Combine($script:StartMenuDir, $script:LnkName)
    if ([System.IO.File]::Exists($startMenuLnk)) {
        $isUserPinned = $false
        try {
            $shell = New-Object -ComObject WScript.Shell
            $existing = $shell.CreateShortcut($startMenuLnk)
            $isUserPinned = $existing.Description.Contains('[UserPinned]')
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($existing)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
        }
        catch { }
        if (-not $isUserPinned) {
            try {
                [System.IO.File]::Delete($startMenuLnk)
                Write-Log "Cleaned up temporary Start Menu shortcut"
            }
            catch { Write-Log "Failed to clean up Start Menu shortcut : $_" -Level Warning }
        }
        else {
            Write-Log "Start Menu shortcut kept (user pinned)"
        }
    }
}

[System.Windows.Forms.Application]::add_ThreadException({
    param($s, $e)
    Write-Log "Unhandled thread exception : $($e.Exception.Message)" -Level Error
    Invoke-ApplicationCleanup
})

[AppDomain]::CurrentDomain.add_UnhandledException({
    param($s, $e)
    Write-Log "Unhandled domain exception : $($e.ExceptionObject.Message)" -Level Error
    Invoke-ApplicationCleanup
})

$form.Add_FormClosing({
    param($s, $e)
    if ($script:UninstallJobs.Count -gt 0) {
        $result = [System.Windows.Forms.MessageBox]::Show("Uninstall operations are still in progress. Are you sure you want to close?", "Operations In Progress", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq [System.Windows.Forms.DialogResult]::No) { $e.Cancel = $true }
    }
})

$form.Add_FormClosed({ 
    $uninstallTimer.Stop()
    $uninstallTimer.Dispose()
    Invoke-ApplicationCleanup
})


[System.Windows.Forms.Application]::Run($form)
Write-Log "MSI Tools ended"