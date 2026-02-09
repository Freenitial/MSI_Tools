# MSI Tools — Technical Documentation

> **Version** : 1.0  
> **Author** : Léo Gillet — Freenitial  
> **Requirements** : PowerShell 5.1+ / .NET Framework 4.5+  
> **Privbiliges** : Works without admin rights — but recommended for network and cleanup

---

## Table of Contents

- [1. Overview](#1-overview)
- [2. Technical Architecture](#2-technical-architecture)
- [3. Launch and Application Identity](#3-launch-and-application-identity)
- [4. Tab 1 — .MSI Properties](#4-tab-1--msi-properties)
- [5. Tab 2 — Explore .MSI Files](#5-tab-2--explore-msi-files)
- [6. Tab 3 — Explore Uninstall Keys](#6-tab-3--explore-uninstall-keys)
- [7. Tab 4 — MSI Residues Cleanup](#7-tab-4--msi-residues-cleanup)
- [8. Network Connectivity and Remote Access](#8-network-connectivity-and-remote-access)
- [9. Security and Impact Surface](#9-security-and-impact-surface)
- [10. Files Created on the System](#10-files-created-on-the-system)
- [11. Cleanup at Shutdown](#11-cleanup-at-shutdown)
- [12. Known Limitations](#12-known-limitations)

---

## 1. Overview

MSI Tools is a diagnostic and maintenance tool for Windows Installer (MSI) installations. It provides:

| Feature | Description |
|---|---|
| **MSI property reading** | Extraction of metadata (GUID, name, version, icon, features, detected properties) from `.msi` files |
| **MSI file exploration** | Recursive directory scan to inventory `.msi`, `.mst`, `.msp`, `.msix` files |
| **Uninstall registry exploration** | Reading of the 4 `Uninstall` hives (HKLM x64/x32, HKCU x64/x32) with filtering |
| **MSI residue cleanup** | Search and deletion of orphaned registry traces left by uninstalled or corrupted MSIs |
| **Remote execution** | All of the above features can target remote machines via WinRM/PSRemoting/SMB |

---

## 2. Technical Architecture

### 2.1 File Format

The script uses the **polyglot batch/PowerShell** pattern:

```
<# :
    @echo off
    powershell -NoLogo -NoProfile -Ex Bypass -Window Hidden -Command ...
    exit /b
#>
# PowerShell code here
```

- The `.bat` file self-executes in PowerShell
- `-ExecutionPolicy Bypass` is required as the script is not signed
- `-Window Hidden` hides the console (can be re-shown via the "Show console" checkbox in the title bar)

### 2.2 Technology Stack

| Component | Usage |
|---|---|
| **WinForms** | Graphical interface (no WPF) |
| **Inline C# (`Add-Type`)** | Win32 interop, shortcut management, elevated drag & drop, ListView sorting, visual theme |
| **COM `WindowsInstaller.Installer`** | Reading `.msi` files (Property, Feature, Icon, ComboBox, CheckBox tables, etc.) |
| **`Microsoft.Win32.Registry`** | Native registry access (not `Get-ItemProperty`) for performance |
| **Runspaces** | Long-running background operations (scan, comparison, uninstall) without blocking the UI |
| **WinRM / `Invoke-Command`** | Remote execution of all scans and cleanup operations |

### 2.3 Custom Title Bar

The application uses `FormBorderStyle = 'None'` with a custom WinForms title bar (`Win11TitleBar`) that:

- Supports drag-to-move, double-click maximize/restore
- Provides minimize/maximize/close buttons
- Handles resizing via `WM_NCHITTEST` (6px grip zones)
- Applies Windows 11 rounded corners via `DwmSetWindowAttribute`

### 2.4 Taskbar Identity (AppUserModelID)

To make Windows display a distinct icon in the taskbar (instead of the PowerShell icon):

1. `SetCurrentProcessExplicitAppUserModelID` is called at launch
2. A `.lnk` shortcut is created in the Start Menu with the same AppUserModelID via the `IPropertyStore` COM interface
3. This shortcut is **temporary**: it is deleted at shutdown, unless the user clicks "Keep in Start Menu" in the About dialog

---

## 3. Launch and Application Identity

### Startup Sequence

```
1. Batch wrapper
2. Add-Type of C# classes (TaskBarHelper, ShortcutHelper, CustomForm, DragDropFix, etc.)
3. SetAppId for taskbar identity
4. Creation of the loading form
5. Logging initialization (%TEMP%\MSI_Tools\)
6. UI construction (tabs, controls)
7. Creation of the WindowsInstaller.Installer COM object (lazy, on first use)
8. Display of the main form
9. DwmSetWindowAttribute for rounded corners
10. DragDropFix::Enable for drag & drop in elevated mode
```

### Admin Detection

Without admin rights, the network panel is hidden and replaced by a "Restart as Admin" button.

---

## 4. Tab 1 — .MSI Properties

### Operation

1. The user drops a `.msi` file (drag & drop) or uses Browse/the text field
2. The `WindowsInstaller.Installer` COM opens the MSI database in **read-only** mode (`OpenDatabase(..., 0)`)
3. The following tables are read:

| MSI Table | Extracted Data |
|---|---|
| `Property` | All properties (ProductCode, ProductName, etc.) |
| `Feature` | List of features with their default install level |
| `ComboBox`, `ListBox`, `RadioButton`, `CheckBox` | Detected UI options to build the "Detected Properties" panel |
| `Icon` | Extraction of the product icon (embedded ICO or EXE) |

### "Detected Properties" Panel

Properties are automatically categorized:

- **UI Tables**: Properties with 2+ possible values found in the ComboBox/ListBox/RadioButton/CheckBox tables
- **Boolean**: Properties whose default value is `0`, `1`, `True`, `False`, `Yes`, `No` — the tool infers the opposite value

Each option is displayed as a clickable button (copies `PROPERTY=VALUE` to the clipboard).
Checkboxes allow selecting multiple properties and copying them in bulk as command-line arguments.

### Multi-MSI Cache

Multiple `.msi` files can be loaded simultaneously.
They are displayed in a selector at the top of the tab with radio buttons. The cache (`$script:MsiFileCache`) stores per file:

```
Key = full file path
Value = @{
    FileName, ProductName, Results (all extracted data),
    SelectedListView (active tab), Icon, IconBytes
}
```

### Drag & Drop in Elevated Mode

Windows blocks `WM_DROPFILES` to an elevated process. The tool works around this via:

```csharp
ChangeWindowMessageFilterEx(hwnd, 0x0233, 1, IntPtr.Zero); // WM_DROPFILES
ChangeWindowMessageFilterEx(hwnd, 0x004A, 1, IntPtr.Zero); // WM_COPYDATA
ChangeWindowMessageFilterEx(hwnd, 0x0049, 1, IntPtr.Zero); // WM_COPYGLOBALDATA
DragAcceptFiles(hwnd, true);
```

The `WM_DROPFILES` (`0x0233`) handler in `OnWindowMessage` calls `DragQueryFile` to retrieve the files.

---

## 5. Tab 2 — Explore .MSI Files

### TreeView (Left Panel)

- **This Device**: Local drives with sub-folders expandable on demand
- **Fast Access**: Windows Quick Access folders (via `Shell.Application` COM)
- **Network**: Dynamically added via the "Goto" field (shares via `net view`)

The TreeView uses **owner-draw** (`OwnerDrawText`) for visual selection and **lazy loading**: sub-folders are only loaded when the parent node is expanded.

#### Node Coloring

| Color | Meaning |
|---|---|
| Black | Read + write access |
| Orange | Read-only access (no Write permission) |
| Red | No read access |

### File Scan

The `Get-FilesRecursive` function:
- Uses `[System.IO.Directory]::GetFiles()` and `GetDirectories()` (no PowerShell cmdlets for performance)
- Filters via a compiled regex: `\.(msi|msix|mst|msp)$`
- Supports configurable depth (No / 1-5 / All)
- Is interruptible via the STOP button (`$script:stopRequested`)
- Updates the progress counter every 50 iterations or 100ms

### ListView (Right Panel)

For each file found, `Get-MsiInfo` is called in simple mode (without `-Full`) to extract only `ProductCode`, `ProductName`, `ProductVersion`.

MSP/MST columns can be hidden via checkboxes.

### Enhanced Context Menu

The ListView context menu is dynamically built by `ConfigureListViewContextMenu` with conditional items:

| Condition | Displayed Items |
|---|---|
| Local path, single selection | CMD here, CMD Admin here, PowerShell Admin here |
| UNC admin share path (`\\server\c$\...`), single selection | PSSession here |
| Always | Copy, Export, Open Parent Folder |
| `.msi` file | Show .MSI Properties (switches to Tab 1) |
| GUID available | Search in Explore Uninstall Keys, Search in MSI Residues Cleanup |

---

## 6. Tab 3 — Explore Uninstall Keys

### Scanned Registry Paths

```
HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall                    (64-bit)
HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall        (32-bit)
HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall                    (64-bit)
HKCU\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall        (32-bit)
```

For remote machines, HKCU keys are remapped to `HKU:<SID>\...` using the console user's SID detected via `WTSQuerySessionInformation` or `Win32_ComputerSystem`.

### Cache Architecture

The scan produces a cache (`$script:RegistryCache`) which is an `ArrayList` of `PSCustomObject` containing all entries.
Filtering changes (hive checkboxes, text field, display options) **do not re-trigger the scan** but filter/reformat from the cache.

### Filtering Options

| Option | Behavior |
|---|---|
| **Restrict Search to** (checkbox) | If checked, only the checked hives are scanned. If unchecked, all 4 hives are scanned and the checkboxes serve only as a post-scan filter |
| **QuietUninstallString if available** | Displays `QuietUninstallString` instead of `UninstallString` when available |
| **Show InstallSource** | Dynamically adds/removes the `InstallSource` column |
| **Text field** | Real-time filtering across all columns, supports multi-term with `;` as separator |

### "Show" Buttons (Regedit)

The "Show" buttons next to each registry path open Regedit directly at the corresponding key via `Open-RegeditHere`.

---

## 7. Tab 4 — MSI Residues Cleanup

This is the most complex feature. It is organized into 4 sub-tabs.

### 7.1 Search (Sub-tab "Search")

#### Search Process

The search runs in 2 phases, within a dedicated runspace to avoid blocking the UI:

**Phase 1 — Initial Scan** (scriptblock `$phase1ScanScriptBlock`):
1. Scan of the 4 Uninstall hives (same logic as Tab 3)
2. Scan of `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\<SID>\Products\<CompressedGUID>\InstallProperties`
3. Matching by title (exact or partial) and/or by GUID (exact or partial)
4. Compressed GUIDs (Windows Installer 32-character format) are automatically converted to standard GUID format `{XXXXXXXX-XXXX-...}`
5. Grouping of results by normalized product GUID

**Phase 2 — Category Scan** (scriptblock `$phase2CategoryScriptBlock`):
For each product found, the following locations are scanned:

| Category | Location |
|---|---|
| **Dependencies** | `HKLM\SOFTWARE\Classes\Installer\Dependencies` + Wow6432Node + HKCU |
| **UpgradeCodes** | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes` + `Classes\Installer\UpgradeCodes` |
| **Features** | `HKLM\SOFTWARE\Classes\Installer\Features\<CompressedGUID>` + HKCU |
| **Components** | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components` |
| **InstallerProducts** | `HKLM\SOFTWARE\Classes\Installer\Products\<CompressedGUID>` + HKCU |
| **InstallerFolders** | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders` (values matching the product name or InstallLocation) |
| **InstallerFiles** | Folder `C:\Windows\Installer\<GUID>` + `LocalPackage` file referenced in UserData |

#### Fast Mode vs Progressbar Mode

- **Fast mode** (checkbox unchecked, default): A single `Invoke-Command` per remote machine for all categories of all products. Faster but no granular progress.
- **Progressbar mode** (checkbox checked): One `Invoke-Command` per category per product. Slower but with detailed progress.

#### Parent-Child Detection

Dependencies are analyzed to detect parent/child relationships:
- If a product A appears in the `Dependents` of a product B, A is considered a child of B
- The TreeView displays children indented under their parent
- Another method is used in the Uninstall sub-tab to determine which uninstaller is "recommended" based on whether the SystemComponent=1 key is present (= visibility in the Control Panel)

### 7.2 Product Cache

Each product found is stored in `$script:ProductCache`:

```
$script:ProductCache[computerName][guid] = @{
    ProductId, DisplayName, Version, Publisher,
    InstallLocation, InstallSource, LocalPackage, CompressedGuid,
    ParentGuid, Elements (list of all found elements),
    FullProduct (complete object with all categories),
    Computer
}
```

This cache feeds the 3 other sub-tabs.

### 7.3 Full Cache (Sub-tab)

Displays the entire cache in a TreeView, organized by product then by category. **Read-only**, serves as a complete reference.

### 7.4 Uninstall Commands (Sub-tab)

For each Uninstall entry found in the cache, a rounded panel is generated with:

| Field | Description |
|---|---|
| **UninstallString** | Raw uninstall command from the registry |
| **QuietUninstallString** | Silent command if available |
| **ModifyPath** | Modification/repair command |
| **Custom** | Automatically generated command |

#### Custom Command Generation

1. **Specific rules**: Certain detected products may offer custom uninstall commands
2. **Default command**: For MSIs, generates `msiexec.exe /X<GUID> /qb /norestart /L*v "C:\Windows\Temp\<Name>_<GUID>_Uninstall.log"`

#### Recommendation

When a product has a parent, the tool identifies the "recommended" uninstaller based on these criteria:
- If only one entry is not a `SystemComponent`, that one is selected
- If only one entry is not an MSI wrapper (it's an `.exe`), that one is selected

The recommended entry is displayed with a green "★ RECOMMENDED" badge and components are indented below it.

#### Uninstall Execution

Execution runs in a dedicated runspace:

1. **Command parsing**: Detection of the type (msiexec, path in quotes, extension detection, first token)
2. **Process launch**: `System.Diagnostics.Process` with `RedirectStandardOutput/Error`, `CreateNoWindow = true`
3. **Real-time monitoring**: ListView displaying child processes with PID, name, CPU%, RAM, relationship (Main/Child/Descendant/MsiExec)
4. **Cancellable**: The Cancel button kills the process and its descendants
5. **Post-execution verification**: Checks whether the Uninstall registry key still exists to determine the outcome

**For remote execution**:
- The process is launched via `Invoke-Command -AsJob`
- Monitoring uses `Get-CimInstance Win32_Process` periodically (every 10 seconds)
- Log files are accessible via UNC (`\\server\c$\...`) or `Invoke-Command`

### 7.5 Compare Residues (Sub-tab)

Compares the cache with the current system state to identify **orphaned residues**:

1. For each element of each product in the cache, tests whether the path (registry or file) still exists
2. Only displays elements that still exist → these are the residues
3. Products shown in orange in the TreeView no longer have an Uninstall entry but still have residues in other locations

The existence test is performed in batch:
- **Local**: Direct test via `[Microsoft.Win32.Registry]` and `[System.IO.File/Directory]`
- **Remote**: A single `Invoke-Command` with all paths to test (fast mode) or per category (progressbar mode)

### 7.6 Cleanup ("Clean Selection" Button)

1. The user checks the items to delete in the active TreeView
2. Confirmation with count and list of targeted machines
3. Items are grouped by machine and serialized
4. **Local**: Item-by-item deletion with progress
5. **Remote**: A single `Invoke-Command` per machine with all items

Deletion types:

| Type | Action |
|---|---|
| `Folder` | `[System.IO.Directory]::Delete($path, $true)` (recursive) |
| `File` | `[System.IO.File]::Delete($path)` |
| `Registry` / `UninstallEntry` | `Registry.DeleteSubKeyTree($subPath, $false)` |
| `RegistryValue` | `RegistryKey.DeleteValue($valueName, $false)` |

Folders are sorted by descending depth to delete children before parents. A tracking of already-deleted folders prevents errors on sub-folders.

#### Force ACL and Retry

On failure (access denied), the "Force ACL and retry" button:
1. Takes ownership of the file/key via `SetOwner`
2. Adds a `FullControl` ACL for the current user
3. For folders, applies recursively to files and sub-folders
4. For registry, applies to sub-keys
5. Retries the deletion

### 7.7 Cross-Tab Synchronization

The 3 TreeViews (Search, Full Cache, Compare) share a synchronization state:

- **Expand/Collapse**: Memorized in `$script:SharedExpandedSyncPaths` (HashSet)
- **Selection**: Memorized in `$script:SharedSelectedSyncPath`
- **Checkboxes**: Memorized in `$script:SharedCheckedSyncPaths` (HashSet)
- **Splitter ratio**: Shared across the 3 SplitContainers

Each node is identified by a unique "SyncPath" (e.g., `Product|computerName|{GUID}`, `Category|...|Label`, `Item|...|type|path`).

When switching sub-tabs, the state is saved from the outgoing TreeView and restored in the incoming TreeView.

---

## 8. Network Connectivity and Remote Access

### 8.1 Connection Panel

The panel at the top right (visible on tabs 3 and 4) contains:

| Field | Function |
|---|---|
| **Target Devices** | Machine names/IPs, separated by `,`, `;` or spaces. Supports Ctrl+Shift+V to paste multi-line (auto-fills all 3 fields) |
| **Credential ID** | Username (format `DOMAIN\user` or `user@domain`) |
| **Password** | Secure password (stored in `SecureString`, never in cleartext in memory after input) |
| **Test** | Launches `Test-RemoteConnections` |

### 8.2 Connection Process (`Test-RemoteConnections`)

For each target machine, in parallel (runspace pool, max 10 threads):

```
1. TCP port 445 test (150ms timeout)
2. Test-WSMan (WinRM)
3. RemoteRegistry service verification
   - If stopped and StartType=Manual → OK (auto-starts)
   - If stopped and StartType=Disabled → Proposes repair
   - If stopped and other StartType → Attempts to start via Invoke-Command
4. Registry access test via Invoke-Command (opens HKLM\SOFTWARE)
5. Reverse DNS resolution for IPs
6. Console user detection (WTS API or Win32_ComputerSystem)
```

### 8.3 TrustedHosts Management

For machines accessed by IP (outside the domain), WinRM requires the IP to be in TrustedHosts.

- **Automatic addition**: Via the connection failure dialog, "Repair" button
- **Automatic cleanup**: At application shutdown, added IPs are removed (`$script:TrustedHostsToRemove`)
- **Cache**: `$script:TrustedHostsCache` avoids repeated reads of `WSMan:\localhost\Client\TrustedHosts`

### 8.4 SMB Sessions and Credentials

To allow `explorer.exe` (non-elevated process) to access admin shares:

```csharp
// SMB session for the PowerShell process (elevated)
WNetAddConnection2(\\target\IPC$, password, username, CONNECT_TEMPORARY)

// Windows credential for explorer.exe (non-elevated)
CredWrite(target, CRED_TYPE_DOMAIN_PASSWORD, username, password, CRED_PERSIST_LOCAL_MACHINE)
```

These sessions and credentials are **cleaned up at shutdown** via `WNetCancelConnection2` and `CredDelete`.

### 8.5 Open-RegeditHere (Remote)

Opening Regedit on a remote machine uses native UI automation:

```
1. Launch regedit.exe /m (multi-instance mode)
2. Find the "Connect Network Registry" menu via GetMenu/GetSubMenu
3. Send WM_COMMAND to open the dialog
4. Inject the machine name via WM_SETTEXT into the RICHEDIT50W
5. Detect credential prompts (CredentialUIBroker.exe) and wait
6. Navigate via the address bar: Tab→Tab→Delete→SetText→Enter
```

Configurable timeout: 10s for automated operations, 60s when a credential prompt is detected.

---

## 9. Security and Impact Surface

### 9.1 What the Tool Reads

| Resource | Method | Usage |
|---|---|---|
| `.msi` files | COM `WindowsInstaller.Installer` in read-only mode | Tab 1 |
| File system | `[System.IO.Directory/File]` | Tab 2, Tab 4 |
| Registry (4 Uninstall hives) | `[Microsoft.Win32.Registry]` in read-only mode | Tab 3 |
| Registry (Installer UserData, Dependencies, UpgradeCodes, Features, Components, Products, Folders) | `[Microsoft.Win32.Registry]` in read-only mode | Tab 4 |
| File/folder ACLs | `GetAccessControl()` | Tab 2 (node coloring) |
| Windows services | `Get-Service WinRM`, `sc.exe query RemoteRegistry` | Network connectivity |
| WTS Sessions | `WTSEnumerateSessions`, `WTSQuerySessionInformation` | Console user detection |
| Active Directory | `[adsisearcher]` (LDAP) | Console user full name resolution |
| Quick Access | `Shell.Application` COM | Tab 2, favorite folders |

### 9.2 What the Tool Writes

| Resource | Condition | Detail |
|---|---|---|
| `%TEMP%\MSI_Tools\*.log` | Always | Log files (rotation at 10 files) |
| `%APPDATA%\MSI_Tools\MSI_Tools.ico` | On first launch | Persistent icon for taskbar identity |
| `%APPDATA%\...\Start Menu\Programs\MSI Tools.lnk` | On launch | Temporary shortcut (deleted at shutdown unless "Keep in Start Menu") |
| `WSMan:\localhost\Client\TrustedHosts` | User action (Repair) | IPs added, cleaned up at shutdown |
| SMB sessions (`\\target\IPC$`) | Remote connection with credentials | Cleaned up at shutdown |
| Windows credentials (`CredWrite`) | Remote connection with credentials | Deleted at shutdown |
| RemoteRegistry service | Remote connection (if disabled) | StartType change + start |
| `HKCU\...\Applets\Regedit\LastKey` | Open-RegeditHere (local, no existing instance) | Registry value for navigation |
| **Registry keys** | Tab 4, Clean Selection + confirmation | Deletion via `DeleteSubKeyTree` / `DeleteValue` |
| **Files/folders** | Tab 4, Clean Selection + confirmation | Deletion via `File.Delete` / `Directory.Delete` |
| **ACLs** | Tab 4, Force ACL + confirmation | `SetOwner` + `AddAccessRule` |
| Temporary files | EXE icon extraction | Written to `%TEMP%`, deleted immediately after |

### 9.3 What the Tool Executes

| Process | Condition | Detail |
|---|---|---|
| `regedit.exe /m` | Open-RegeditHere | New multi-mode instance |
| `explorer.exe` | Open Parent Folder / Open File | Explorer navigation |
| `cmd.exe` | Context menu "CMD here" | Console in the selected folder |
| `powershell.exe` | Context menu "PowerShell Admin here" / "PSSession here" | With `-Verb runas` or `Enter-PSSession` |
| `notepad.exe` | Show Log | Opening log files |
| `sc.exe` | RemoteRegistry verification | `sc.exe \\target query RemoteRegistry` |
| `net view` | Network navigation (Tab 2) | Share enumeration |
| `winrm quickconfig -force` | WinRM configuration (user action) | Initial WinRM setup |
| Uninstall command | Tab 4, Uninstall Commands (user action) | Executed with `RedirectStandardOutput`, `CreateNoWindow` |

### 9.4 Password Management

- The password is stored in a `System.Security.SecureString`
- The text field only displays `●` characters (`U+25CF`)
- Keystrokes are intercepted via `KeyPress` → `SecurePassword.AppendChar()`
- The TextBox text never contains the actual password
- The `SecureString` is `Dispose()`d at shutdown and upon reset
- For SMB sessions, the password is temporarily extracted via `GetNetworkCredential().Password` (required for `WNetAddConnection2`)

---

## 10. Files Created on the System

### Persistent (if "Keep in Start Menu")

| File | Path |
|---|---|
| Icon | `%APPDATA%\MSI_Tools\MSI_Tools.ico` |
| Start Menu shortcut | `%APPDATA%\Microsoft\Windows\Start Menu\Programs\MSI Tools.lnk` |

### Temporary (automatically cleaned up)

| File | Path | Cleanup |
|---|---|---|
| Logs | `%TEMP%\MSI_Tools\MSI_Tools_YYYYMMDD.log` | Rotation at 10 files |
| Extracted icons | `%TEMP%\msi_icon_*.exe` | Immediate after extraction |
| Copied icons | `%TEMP%\<ProductName>.png` | No automatic cleanup |
| Copied remote logs | `%TEMP%\Remote_<Computer>_*.log` | No automatic cleanup |
| Start Menu shortcut (if not pinned) | `%APPDATA%\...\Start Menu\Programs\MSI Tools.lnk` | At shutdown |

---

## 11. Cleanup at Shutdown

The `Invoke-ApplicationCleanup` function is called:
- Via `Form.FormClosed`
- Via `Application.ThreadException` (unhandled exception on the UI thread)
- Via `AppDomain.UnhandledException` (fatal exception)

A confirmation is prompted if uninstall jobs are still in progress.

---

## 12. Known Limitations

| Limitation | Detail |
|---|---|
| **PowerShell 5.1** | Not tested with PowerShell 7+ (COM classes and WinForms may behave differently) |
| **x64 only** | The script redirects to `Sysnative` if launched in x86, but the UI is designed for x64 |
| **No code signing** | Requires `-ExecutionPolicy Bypass` |
| **Regedit automation** | Depends on the regedit UI structure (fragile across Windows versions). Supports EN/FR via menu scanning |
| **RemoteRegistry** | Required for certain Tab 4 Cleanup operations. The tool can start it but cannot configure it if the service is disabled by GPO |
| **Console user detection** | Uses WTS API then WMI as fallback. Does not work if no user is logged in to the console |
| **Parent-child detection** | Based on the Dependencies registry. May miss relationships if the installer does not use this standard mechanism (e.g., C++ redistributables) |
| **Components scan** | Only scans `S-1-5-18` (LocalSystem). Components installed per-user under other SIDs are not detected |
| **Force ACL** | May fail on keys protected by TrustedInstaller (would require a specific token) |
