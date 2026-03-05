#  <img width="128" height="128" alt="Untitled 1ss" src="https://github.com/user-attachments/assets/8ba8a4c0-966d-4427-96b6-8e25783aabde" />  **MSI Tools**

**Graphical PowerShell tool for MSI diagnostics, registry exploration, remote analysis, and residue cleanup**

--------------------

### Features ✨

- 📦 MSI file properties analysis + Auto-detect possible arguments
- 🔍 Fast file detection and GUID finder
- 📂 File system explorer
- ⚙️ Customizable recursion level for MSI scan
- 🔑 Registry explorer for installed software
- 📊 Sorting and filtering, Smart columns resize
- 💾 Export data to CSV, XLS, or Self-launching OGV (hybrid batch)
- 🎯 Multi-selection support
- 🧹 MSI residue cleanup — find and remove corrupted msi registry traces
- 🖥️ Full remote computer support via WinRM/PSRemoting/SMB
  - Built-in connection panel with credential management and connectivity testing
  - All tabs support remote targets — scans, exploration, uninstall, and cleanup

--------------------

### Tab 1: .MSI Properties
- Drag & drop or browse for MSI files
- Extract all properties, features, and product icon
- **Detected Properties** panel: automatically identifies configurable install options from UI tables (ComboBox, ListBox, RadioButton, CheckBox) and boolean properties
- Multi-MSI cache: load and switch between several MSI files

![ezgif-6a2088b7915f7a0b](https://github.com/user-attachments/assets/b173a131-34b0-4986-a340-8dcdf249ee31)

--------------------

### Tab 2: Explore Folders
- Browse file system with a tree view
- Recursive search for MSI, MST, MSP files
- Filtering and sorting options
- Right-click menu with multiple actions
- Scan with progressbar tracking
- Quick Access folders and network share support

![Tab2](https://github.com/user-attachments/assets/ce8f4ec0-f242-4d89-b584-45956151ab1a)

--------------------

### Tab 3: Explore Registry
- View installed software from registry
- Remote registry capable
- Filtering and sorting options
- Right-click menu with multiple actions

![Tab3](https://github.com/user-attachments/assets/1961f0f9-755d-471a-b650-3bc0e2046029)

--------------------

### Tab 4: MSI Residues Cleanup
- **Search**: find MSI traces by product name and/or GUID across 8+ registry locations and installer folders
- **Full Cache**: read-only TreeView of all discovered products and their registry/file footprint
- **Uninstall Commands**: auto-generated uninstall commands with recommended entry detection, real-time process monitoring, and remote execution support
- **Compare Residues**: compares cache against live system state to identify orphaned entries
- **Clean Selection**: delete checked residues (registry keys, values, files, folders) with confirmation — supports Force ACL and retry for access-denied scenarios
- **Cross-tab synchronization**: expand/collapse, selection, and checkbox states are preserved across sub-tabs
- Fast mode (single batch call) or Progressbar mode (granular progress per category)

![VideoProject1-ezgif com-video-to-gif-converter](https://github.com/user-attachments/assets/a2793ce4-7a78-4ebc-9367-9470602a57fa)

--------------------

### Remote Capabilities 🌐

All tabs support remote targets when running as administrator:

| Feature | Method |
|---|---|
| **Connection panel** | Target devices, credentials, one-click connectivity test |
| **Connectivity test** | TCP 445, WinRM, RemoteRegistry service check, registry access validation |
| **TrustedHosts** | Automatic management for IP-based targets (added on connect, removed at shutdown) |
| **SMB sessions** | Credential injection for admin shares (`\\server\c$`) |
| **Remote Regedit** | Automated UI-driven connection to remote registry |

--------------------

### Requirements 📋

- Windows 10
-PowerShell 5.1+

--------------------

### Getting Started 🚀

1. Download `MSI_Tools.bat`
2. Double-click to run (or right-click → Run as administrator for full features)
3. The batch wrapper auto-launches PowerShell with the appropriate parameters

--------------------
