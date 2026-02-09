# MSI Tools — Documentation Technique

> **Version** : 1.0  
> **Auteur** : Léo Gillet — Freenitial  
> **Runtime** : PowerShell 5.1+ / .NET Framework 4.5+  
> **Privilèges** : Fonctionne sans droits admin — Droits admin recommandés pour les fonctionnalités réseau et nettoyage

---

## Table des matières

- [1. Vue d'ensemble](#1-vue-densemble)
- [2. Architecture technique](#2-architecture-technique)
- [3. Lancement et identité de l'application](#3-lancement-et-identité-de-lapplication)
- [4. Tab 1 — .MSI Properties](#4-tab-1--msi-properties)
- [5. Tab 2 — Explore .MSI Files](#5-tab-2--explore-msi-files)
- [6. Tab 3 — Explore Uninstall Keys](#6-tab-3--explore-uninstall-keys)
- [7. Tab 4 — MSI Residues Cleanup](#7-tab-4--msi-residues-cleanup)
- [8. Connectivité réseau et accès distant](#8-connectivité-réseau-et-accès-distant)
- [9. Sécurité et surface d'impact](#9-sécurité-et-surface-dimpact)
- [10. Fichiers créés sur le système](#10-fichiers-créés-sur-le-système)
- [11. Nettoyage à la fermeture](#11-nettoyage-à-la-fermeture)
- [12. Limitations connues](#12-limitations-connues)

---

## 1. Vue d'ensemble

MSI Tools est un outil de diagnostic et de maintenance des installations Windows Installer (MSI). Il permet de :

| Fonctionnalité | Description |
|---|---|
| **Lecture de propriétés MSI** | Extraction des métadonnées (GUID, nom, version, icône, features, propriétés détectées) depuis des fichiers `.msi` |
| **Exploration de fichiers MSI** | Scan récursif de répertoires pour inventorier les fichiers `.msi`, `.mst`, `.msp`, `.msix` |
| **Exploration du registre Uninstall** | Lecture des 4 hives `Uninstall` (HKLM x64/x32, HKCU x64/x32) avec filtrage |
| **Nettoyage de résidus MSI** | Recherche et suppression des traces orphelines issues du registre laissées par des MSI désinstallés ou corrumpus |
| **Exécution distante** | Toutes les fonctionnalités ci-dessus peuvent cibler des machines distantes via WinRM/PSRemoting/SMB |

---

## 2. Architecture technique

### 2.1 Format du fichier

Le script utilise le pattern **polyglot batch/PowerShell** :

```
<# :
    @echo off
    powershell -NoLogo -NoProfile -Ex Bypass -Window Hidden -Command ...
    exit /b
#>
# Code PowerShell ici
```

- Le fichier `.bat` s'auto-exécute en PowerShell
- `-ExecutionPolicy Bypass` est nécessaire car le script n'est pas signé
- `-Window Hidden` masque la console (réaffichable via la checkbox "Show console" dans la barre de titre)

### 2.2 Stack technologique

| Composant | Usage |
|---|---|
| **WinForms** | Interface graphique (aucun WPF) |
| **C# inline (`Add-Type`)** | Interop Win32, gestion des raccourcis, drag & drop élevé, tri des ListView, thème visuel |
| **COM `WindowsInstaller.Installer`** | Lecture des fichiers `.msi` (tables Property, Feature, Icon, ComboBox, CheckBox, etc.) |
| **`Microsoft.Win32.Registry`** | Accès registre natif (pas de `Get-ItemProperty`) pour les performances |
| **Runspaces** | Opérations longues en arrière-plan (scan, comparaison, désinstallation) sans bloquer l'UI |
| **WinRM / `Invoke-Command`** | Exécution distante de tous les scans et opérations de nettoyage |

### 2.3 Barre de titre personnalisée

L'application utilise `FormBorderStyle = 'None'` avec une barre de titre WinForms custom (`Win11TitleBar`) qui :

- Supporte le drag-to-move, le double-clic maximize/restore
- Fournit les boutons minimize/maximize/close
- Gère le redimensionnement via `WM_NCHITTEST` (grip zones de 6px)
- Applique les coins arrondis Windows 11 via `DwmSetWindowAttribute`

### 2.4 Identité taskbar (AppUserModelID)

Pour que Windows affiche une icône distincte dans la barre des tâches (au lieu de l'icône PowerShell) :

1. `SetCurrentProcessExplicitAppUserModelID` est appelé au lancement
2. Un raccourci `.lnk` est créé dans le Start Menu avec le même AppUserModelID via l'interface COM `IPropertyStore`
3. Ce raccourci est **temporaire** : il est supprimé à la fermeture, sauf si l'utilisateur clique "Keep in Start Menu" dans la boîte About

---

## 3. Lancement et identité de l'application

### Séquence de démarrage

```
1. Wrapper batch
2. Add-Type des classes C# (TaskBarHelper, ShortcutHelper, CustomForm, DragDropFix, etc.)
3. SetAppId pour l'identité taskbar
4. Création du formulaire de loading
5. Initialisation du logging (%TEMP%\MSI_Tools\)
6. Construction de l'UI (tabs, contrôles)
7. Création de l'objet COM WindowsInstaller.Installer (lazy, au premier usage)
8. Affichage du formulaire principal
9. DwmSetWindowAttribute pour les coins arrondis
10. DragDropFix::Enable pour le drag & drop en mode élevé
```

### Détection admin

Sans droits admin, le panneau réseau est masqué, remplacé par un bouton "Restart as Admin"

---

## 4. Tab 1 — .MSI Properties

### Fonctionnement

1. L'utilisateur dépose un fichier `.msi` (drag & drop) ou utilise Browse/le champ texte
2. Le COM `WindowsInstaller.Installer` ouvre la base de données MSI en **lecture seule** (`OpenDatabase(..., 0)`)
3. Les tables suivantes sont lues :

| Table MSI | Données extraites |
|---|---|
| `Property` | Toutes les propriétés (ProductCode, ProductName, etc.) |
| `Feature` | Liste des features avec leur niveau d'installation par défaut |
| `ComboBox`, `ListBox`, `RadioButton`, `CheckBox` | Options UI détectées pour construire le panneau "Detected Properties" |
| `Icon` | Extraction de l'icône du produit (ICO ou EXE embarqué) |

### Panneau "Detected Properties"

Les propriétés sont catégorisées automatiquement :

- **UI Tables** : Propriétés ayant 2+ valeurs possibles retrouvées dans les tables ComboBox/ListBox/RadioButton/CheckBox
- **Boolean** : Propriétés dont la valeur par défaut est `0`, `1`, `True`, `False`, `Yes`, `No` — l'outil infère la valeur opposée

Chaque option est affichée sous forme de bouton cliquable (copie `PROPERTY=VALUE` dans le presse-papiers). 
Les checkboxes permettent de sélectionner plusieurs propriétés et de les copier en bloc comme arguments de ligne de commande.

### Cache multi-MSI

Plusieurs fichiers `.msi` peuvent être chargés simultanément. 
Ils sont affichés dans un sélecteur en haut du tab avec des radio buttons. Le cache (`$script:MsiFileCache`) stocke par fichier :

```
Key = chemin complet du fichier
Value = @{
    FileName, ProductName, Results (toutes les données extraites),
    SelectedListView (onglet actif), Icon, IconBytes
}
```

### Drag & Drop en mode élevé

Windows bloque le `WM_DROPFILES` vers un processus élevé. L'outil contourne cela via :

```csharp
ChangeWindowMessageFilterEx(hwnd, 0x0233, 1, IntPtr.Zero); // WM_DROPFILES
ChangeWindowMessageFilterEx(hwnd, 0x004A, 1, IntPtr.Zero); // WM_COPYDATA
ChangeWindowMessageFilterEx(hwnd, 0x0049, 1, IntPtr.Zero); // WM_COPYGLOBALDATA
DragAcceptFiles(hwnd, true);
```

Le handler `WM_DROPFILES` (`0x0233`) dans `OnWindowMessage` appelle `DragQueryFile` pour récupérer les fichiers.

---

## 5. Tab 2 — Explore .MSI Files

### TreeView (panneau gauche)

- **This Device** : Drives locaux avec sous-dossiers expandables à la demande
- **Fast Access** : Dossiers Quick Access Windows (via `Shell.Application` COM)
- **Network** : Ajouté dynamiquement via le champ "Goto" (partages via `net view`)

Le TreeView utilise un **owner-draw** (`OwnerDrawText`) pour la sélection visuelle et un **lazy loading** : les sous-dossiers ne sont chargés qu'à l'expansion du nœud parent.

#### Coloration des nœuds

| Couleur | Signification |
|---|---|
| Noir | Accès lecture + écriture |
| Orange | Accès lecture seul (pas de droit Write) |
| Rouge | Pas d'accès lecture |

### Scan de fichiers

La fonction `Get-FilesRecursive` :
- Utilise `[System.IO.Directory]::GetFiles()` et `GetDirectories()` (pas de cmdlet PowerShell pour la performance)
- Filtre via une regex compilée : `\.(msi|msix|mst|msp)$`
- Supporte une profondeur configurable (No / 1-5 / All)
- Est interruptible via le bouton STOP (`$script:stopRequested`)
- Met à jour le compteur de progression toutes les 50 itérations ou 100ms

### ListView (panneau droit)

Pour chaque fichier trouvé, `Get-MsiInfo` est appelé en mode simple (sans `-Full`) pour extraire uniquement `ProductCode`, `ProductName`, `ProductVersion`.

Les colonnes MSP/MST sont masquables via les checkboxes.

### Menu contextuel enrichi

Le menu contextuel des ListViews est construit dynamiquement par `ConfigureListViewContextMenu` avec des items conditionnels :

| Condition | Items affichés |
|---|---|
| Chemin local, sélection unique | CMD here, CMD Admin here, PowerShell Admin here |
| Chemin UNC admin share (`\\server\c$\...`), sélection unique | PSSession here |
| Toujours | Copy, Export, Open Parent Folder |
| Fichier `.msi` | Show .MSI Properties (bascule vers Tab 1) |
| GUID disponible | Search in Explore Uninstall Keys, Search in MSI Residues Cleanup |

---

## 6. Tab 3 — Explore Uninstall Keys

### Chemins de registre scannés

```
HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall                    (64-bit)
HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall        (32-bit)
HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall                    (64-bit)
HKCU\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall        (32-bit)
```

Pour les machines distantes, les clés HKCU sont remappées vers `HKU:<SID>\...` en utilisant le SID de l'utilisateur console détecté via `WTSQuerySessionInformation` ou `Win32_ComputerSystem`.

### Architecture de cache

Le scan produit un cache (`$script:RegistryCache`) qui est un `ArrayList` de `PSCustomObject` contenant toutes les entrées. 
Les changements de filtrage (checkboxes hive, champ texte, options d'affichage) **ne relancent pas le scan** mais filtrent/reformatent depuis le cache.

### Options de filtrage

| Option | Comportement |
|---|---|
| **Restrict Search to** (checkbox) | Si coché, seules les hives cochées sont scannées. Si décoché, toutes les 4 hives sont scannées et les checkboxes servent uniquement de filtre post-scan |
| **QuietUninstallString if available** | Affiche `QuietUninstallString` à la place de `UninstallString` quand elle existe |
| **Show InstallSource** | Ajoute/retire dynamiquement la colonne `InstallSource` |
| **Champ texte** | Filtre en temps réel sur toutes les colonnes, supporte le multi-terme avec `;` comme séparateur |

### Boutons "Show" (Regedit)

Les boutons "Show" à côté de chaque chemin de registre ouvrent Regedit directement à la clé correspondante via `Open-RegeditHere`.

---

## 7. Tab 4 — MSI Residues Cleanup

C'est la fonctionnalité la plus complexe. Elle est organisée en 4 sous-onglets.

### 7.1 Recherche (sous-onglet "Search")

#### Processus de recherche

La recherche s'exécute en 2 phases, dans un runspace dédié pour ne pas bloquer l'UI :

**Phase 1 — Scan initial** (scriptblock `$phase1ScanScriptBlock`) :
1. Scan des 4 hives Uninstall (même logique que Tab 3)
2. Scan de `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\<SID>\Products\<CompressedGUID>\InstallProperties`
3. Correspondance par titre (exact ou partiel) et/ou par GUID (exact ou partiel)
4. Les GUIDs compressés (format Windows Installer 32 caractères) sont automatiquement convertis en GUID standard `{XXXXXXXX-XXXX-...}`
5. Groupement des résultats par GUID produit normalisé

**Phase 2 — Scan des catégories** (scriptblock `$phase2CategoryScriptBlock`) :
Pour chaque produit trouvé, les emplacements suivants sont scannés :

| Catégorie | Emplacement |
|---|---|
| **Dependencies** | `HKLM\SOFTWARE\Classes\Installer\Dependencies` + Wow6432Node + HKCU |
| **UpgradeCodes** | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes` + `Classes\Installer\UpgradeCodes` |
| **Features** | `HKLM\SOFTWARE\Classes\Installer\Features\<CompressedGUID>` + HKCU |
| **Components** | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components` |
| **InstallerProducts** | `HKLM\SOFTWARE\Classes\Installer\Products\<CompressedGUID>` + HKCU |
| **InstallerFolders** | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders` (valeurs correspondant au nom du produit ou InstallLocation) |
| **InstallerFiles** | Dossier `C:\Windows\Installer\<GUID>` + fichier `LocalPackage` référencé dans UserData |

#### Mode Fast vs Progressbar

- **Fast mode** (checkbox non cochée, par défaut) : Un seul `Invoke-Command` par machine distante pour toutes les catégories de tous les produits. Plus rapide mais pas de progression granulaire.
- **Progressbar mode** (checkbox cochée) : Un `Invoke-Command` par catégorie par produit. Plus lent mais progression détaillée.

#### Détection parent-enfant

Les dépendances sont analysées pour détecter les relations parent/enfant :
- Si un produit A apparaît dans les `Dependents` d'un produit B, A est considéré comme enfant de B
- Le TreeView affiche les enfants indentés sous leur parent
- Une autre méthode est utilisée dans le sous-onglet Uninstall pour déterminer quel désinstalleur est "recommandé" selon si la clé SystemComponent=1 est présente (= visibilité dans le panneau de configuration)

### 7.2 Cache produit

Chaque produit trouvé est stocké dans `$script:ProductCache` :

```
$script:ProductCache[computerName][guid] = @{
    ProductId, DisplayName, Version, Publisher,
    InstallLocation, InstallSource, LocalPackage, CompressedGuid,
    ParentGuid, Elements (liste de tous les éléments trouvés),
    FullProduct (objet complet avec toutes les catégories),
    Computer
}
```

Ce cache alimente les 3 autres sous-onglets.

### 7.3 Full Cache (sous-onglet)

Affiche l'intégralité du cache dans un TreeView, organisé par produit puis par catégorie. **Lecture seule**, sert de référence complète.

### 7.4 Uninstall Commands (sous-onglet)

Pour chaque entrée Uninstall trouvée dans le cache, un panneau arrondi est généré avec :

| Champ | Description |
|---|---|
| **UninstallString** | Commande de désinstallation brute du registre |
| **QuietUninstallString** | Commande silencieuse si disponible |
| **ModifyPath** | Commande de modification/réparation |
| **Custom** | Commande générée automatiquement |

#### Génération de la commande Custom

1. **Règles spécifiques** : Certains produits détectés peuvent proposer des commandes de désinstallation personnalisées
2. **Commande par défaut** : Pour les MSI, génère `msiexec.exe /X<GUID> /qb /norestart /L*v "C:\Windows\Temp\<Name>_<GUID>_Uninstall.log"`

#### Recommandation

Quand un produit a un parent, l'outil identifie le désinstalleur "recommandé" selon ces critères :
- Si un seul entry n'est pas un `SystemComponent`, c'est lui
- Si un seul entry n'est pas un MSI wrapper (c'est un `.exe`), c'est lui

L'entry recommandé est affiché avec un badge vert "★ RECOMMENDED" et les composants sont indentés en dessous.

#### Exécution de désinstallation

L'exécution se fait dans un runspace dédié :

1. **Parsing de la commande** : Détection du type (msiexec, chemin entre guillemets, détection d'extension, premier token)
2. **Lancement du processus** : `System.Diagnostics.Process` avec `RedirectStandardOutput/Error`, `CreateNoWindow = true`
3. **Monitoring en temps réel** : ListView affichant les processus enfants avec PID, nom, CPU%, RAM, relation (Main/Child/Descendant/MsiExec)
4. **Annulable** : Le bouton Cancel tue le processus et ses descendants
5. **Vérification post-exécution** : Vérifie si la clé de registre Uninstall existe encore pour déterminer l'état

**Pour l'exécution distante** :
- Le processus est lancé via `Invoke-Command -AsJob`
- Le monitoring utilise `Get-CimInstance Win32_Process` périodiquement (toutes les 10 secondes)
- Les fichiers log sont accessibles via UNC (`\\server\c$\...`) ou `Invoke-Command`

### 7.5 Compare Residues (sous-onglet)

Compare le cache avec l'état actuel du système pour identifier les **résidus orphelins** :

1. Pour chaque élément de chaque produit en cache, teste si le chemin (registre ou fichier) existe encore
2. N'affiche que les éléments qui existent encore → ce sont les résidus
3. Les produits en orange dans le TreeView n'ont plus d'entrée Uninstall mais ont encore des résidus dans d'autres emplacements

Le test d'existence se fait en batch :
- **Local** : Test direct via `[Microsoft.Win32.Registry]` et `[System.IO.File/Directory]`
- **Remote** : Un seul `Invoke-Command` avec tous les chemins à tester (fast mode) ou par catégorie (progressbar mode)

### 7.6 Nettoyage (bouton "Clean Selection")

1. L'utilisateur coche les éléments à supprimer dans le TreeView actif
2. Confirmation avec comptage et liste des machines ciblées
3. Les éléments sont groupés par machine et sérialisés
4. **Local** : Suppression item par item avec progression
5. **Remote** : Un seul `Invoke-Command` par machine avec tous les items

Types de suppression :

| Type | Action |
|---|---|
| `Folder` | `[System.IO.Directory]::Delete($path, $true)` (récursif) |
| `File` | `[System.IO.File]::Delete($path)` |
| `Registry` / `UninstallEntry` | `Registry.DeleteSubKeyTree($subPath, $false)` |
| `RegistryValue` | `RegistryKey.DeleteValue($valueName, $false)` |

Les dossiers sont triés par profondeur décroissante pour supprimer les enfants avant les parents. Un tracking des dossiers déjà supprimés évite les erreurs sur les sous-dossiers.

#### Force ACL and Retry

En cas d'échec (accès refusé), le bouton "Force ACL and retry" :
1. Prend possession du fichier/clé via `SetOwner`
2. Ajoute une ACL `FullControl` pour l'utilisateur courant
3. Pour les dossiers, applique récursivement aux fichiers et sous-dossiers
4. Pour le registre, applique aux sous-clés
5. Retente la suppression

### 7.7 Synchronisation cross-tab

Les 3 TreeViews (Search, Full Cache, Compare) partagent un état de synchronisation :

- **Expand/Collapse** : Mémorisé dans `$script:SharedExpandedSyncPaths` (HashSet)
- **Selection** : Mémorisé dans `$script:SharedSelectedSyncPath`
- **Checkboxes** : Mémorisées dans `$script:SharedCheckedSyncPaths` (HashSet)
- **Splitter ratio** : Partagé entre les 3 SplitContainers

Chaque nœud est identifié par un "SyncPath" unique (ex: `Product|computerName|{GUID}`, `Category|...|Label`, `Item|...|type|path`).

Quand on change de sous-onglet, l'état est sauvegardé depuis le TreeView sortant et restauré dans le TreeView entrant.

---

## 8. Connectivité réseau et accès distant

### 8.1 Panneau de connexion

Le panneau en haut à droite (visible sur les onglets 3 et 4) contient :

| Champ | Fonction |
|---|---|
| **Target Devices** | Noms/IPs des machines, séparés par `,`, `;` ou espaces. Supporte le Ctrl+Shift+V pour coller multi-lignes (auto-rempli les 3 champs) |
| **Credential ID** | Nom d'utilisateur (format `DOMAIN\user` ou `user@domain`) |
| **Password** | Mot de passe sécurisé (stocké dans `SecureString`, jamais en clair en mémoire après saisie) |
| **Test** | Lance `Test-RemoteConnections` |

### 8.2 Processus de connexion (`Test-RemoteConnections`)

Pour chaque machine cible, en parallèle (runspace pool, max 10 threads) :

```
1. Test TCP port 445 (timeout 150ms)
2. Test-WSMan (WinRM)
3. Vérification du service RemoteRegistry
   - Si arrêté et StartType=Manual → OK (s'auto-démarre)
   - Si arrêté et StartType=Disabled → Propose la réparation
   - Si arrêté et autre StartType → Tente de le démarrer via Invoke-Command
4. Test d'accès au registre via Invoke-Command (ouvre HKLM\SOFTWARE)
5. Résolution DNS inverse pour les IPs
6. Détection de l'utilisateur console (WTS API ou Win32_ComputerSystem)
```

### 8.3 Gestion des TrustedHosts

Pour les machines accédées par IP (hors domaine), WinRM exige que l'IP soit dans TrustedHosts.

- **Ajout automatique** : Via la boîte de dialogue des échecs de connexion, bouton "Repair"
- **Nettoyage automatique** : À la fermeture de l'application, les IPs ajoutées sont retirées (`$script:TrustedHostsToRemove`)
- **Cache** : `$script:TrustedHostsCache` évite les lectures répétées de `WSMan:\localhost\Client\TrustedHosts`

### 8.4 Sessions SMB et credentials

Pour permettre à `explorer.exe` (processus non-élevé) d'accéder aux partages admin :

```csharp
// Session SMB pour le processus PowerShell (élevé)
WNetAddConnection2(\\target\IPC$, password, username, CONNECT_TEMPORARY)

// Credential Windows pour explorer.exe (non-élevé)
CredWrite(target, CRED_TYPE_DOMAIN_PASSWORD, username, password, CRED_PERSIST_LOCAL_MACHINE)
```

Ces sessions et credentials sont **nettoyés à la fermeture** via `WNetCancelConnection2` et `CredDelete`.

### 8.5 Open-RegeditHere (distant)

L'ouverture de Regedit sur une machine distante utilise l'automatisation de l'UI native :

```
1. Lance regedit.exe /m (mode multi-instance)
2. Trouve le menu "Connect Network Registry" via GetMenu/GetSubMenu
3. Envoie WM_COMMAND pour ouvrir la boîte de dialogue
4. Injecte le nom de machine via WM_SETTEXT dans le RICHEDIT50W
5. Détecte les prompts de credentials (CredentialUIBroker.exe) et attend
6. Navigue via la barre d'adresse : Tab→Tab→Delete→SetText→Enter
```

Timeout configurable : 10s pour les opérations automatisées, 60s quand un prompt de credentials est détecté.

---

## 9. Sécurité et surface d'impact

### 9.1 Ce que l'outil lit

| Ressource | Méthode | Usage |
|---|---|---|
| Fichiers `.msi` | COM `WindowsInstaller.Installer` en lecture seule | Tab 1 |
| Système de fichiers | `[System.IO.Directory/File]` | Tab 2, Tab 4 |
| Registre (4 hives Uninstall) | `[Microsoft.Win32.Registry]` en lecture seule | Tab 3 |
| Registre (Installer UserData, Dependencies, UpgradeCodes, Features, Components, Products, Folders) | `[Microsoft.Win32.Registry]` en lecture seule | Tab 4 |
| ACLs fichiers/dossiers | `GetAccessControl()` | Tab 2 (coloration des nœuds) |
| Services Windows | `Get-Service WinRM`, `sc.exe query RemoteRegistry` | Connectivité réseau |
| WTS Sessions | `WTSEnumerateSessions`, `WTSQuerySessionInformation` | Détection utilisateur console |
| Active Directory | `[adsisearcher]` (LDAP) | Résolution du nom complet de l'utilisateur console |
| Quick Access | `Shell.Application` COM | Tab 2, dossiers favoris |

### 9.2 Ce que l'outil écrit

| Ressource | Condition | Détail |
|---|---|---|
| `%TEMP%\MSI_Tools\*.log` | Toujours | Fichiers de log (rotation à 10 fichiers) |
| `%APPDATA%\MSI_Tools\MSI_Tools.ico` | Au premier lancement | Icône persistante pour l'identité taskbar |
| `%APPDATA%\...\Start Menu\Programs\MSI Tools.lnk` | Au lancement | Raccourci temporaire (supprimé à la fermeture sauf si "Keep in Start Menu") |
| `WSMan:\localhost\Client\TrustedHosts` | Action utilisateur (Repair) | IPs ajoutées, nettoyées à la fermeture |
| Sessions SMB (`\\target\IPC$`) | Connexion distante avec credentials | Nettoyées à la fermeture |
| Credentials Windows (`CredWrite`) | Connexion distante avec credentials | Supprimés à la fermeture |
| Service RemoteRegistry | Connexion distante (si désactivé) | Changement de StartType + démarrage |
| `HKCU\...\Applets\Regedit\LastKey` | Open-RegeditHere (local, pas d'instance existante) | Valeur de registre pour la navigation |
| **Clés de registre** | Tab 4, Clean Selection + confirmation | Suppression via `DeleteSubKeyTree` / `DeleteValue` |
| **Fichiers/dossiers** | Tab 4, Clean Selection + confirmation | Suppression via `File.Delete` / `Directory.Delete` |
| **ACLs** | Tab 4, Force ACL + confirmation | `SetOwner` + `AddAccessRule` |
| Fichiers temporaires | Extraction d'icône EXE | Écrits dans `%TEMP%`, supprimés immédiatement après |

### 9.3 Ce que l'outil exécute

| Processus | Condition | Détail |
|---|---|---|
| `regedit.exe /m` | Open-RegeditHere | Nouvelle instance multi-mode |
| `explorer.exe` | Open Parent Folder / Open File | Navigation dans l'explorateur |
| `cmd.exe` | Menu contextuel "CMD here" | Console dans le dossier sélectionné |
| `powershell.exe` | Menu contextuel "PowerShell Admin here" / "PSSession here" | Avec `-Verb runas` ou `Enter-PSSession` |
| `notepad.exe` | Show Log | Ouverture de fichiers log |
| `sc.exe` | Vérification RemoteRegistry | `sc.exe \\target query RemoteRegistry` |
| `net view` | Navigation réseau (Tab 2) | Enumération des partages |
| `winrm quickconfig -force` | Config WinRM (action utilisateur) | Configuration initiale de WinRM |
| Commande de désinstallation | Tab 4, Uninstall Commands (action utilisateur) | Exécutée avec `RedirectStandardOutput`, `CreateNoWindow` |

### 9.4 Gestion du mot de passe

- Le mot de passe est stocké dans un `System.Security.SecureString`
- Le champ texte n'affiche que des caractères `●` (`U+25CF`)
- Les frappes sont interceptées via `KeyPress` → `SecurePassword.AppendChar()`
- Le texte du TextBox ne contient jamais le mot de passe réel
- Le `SecureString` est `Dispose()` à la fermeture et lors de la réinitialisation
- Pour les sessions SMB, le mot de passe est temporairement extrait via `GetNetworkCredential().Password` (nécessaire pour `WNetAddConnection2`)

---

## 10. Fichiers créés sur le système

### Persistants (si "Keep in Start Menu")

| Fichier | Chemin |
|---|---|
| Icône | `%APPDATA%\MSI_Tools\MSI_Tools.ico` |
| Raccourci Start Menu | `%APPDATA%\Microsoft\Windows\Start Menu\Programs\MSI Tools.lnk` |

### Temporaires (nettoyés automatiquement)

| Fichier | Chemin | Nettoyage |
|---|---|---|
| Logs | `%TEMP%\MSI_Tools\MSI_Tools_YYYYMMDD.log` | Rotation à 10 fichiers |
| Icônes extraites | `%TEMP%\msi_icon_*.exe` | Immédiat après extraction |
| Icons copiées | `%TEMP%\<ProductName>.png` | Pas de nettoyage automatique |
| Logs distants copiés | `%TEMP%\Remote_<Computer>_*.log` | Pas de nettoyage automatique |
| Raccourci Start Menu (si non pinned) | `%APPDATA%\...\Start Menu\Programs\MSI Tools.lnk` | À la fermeture |

---

## 11. Nettoyage à la fermeture

La fonction `Invoke-ApplicationCleanup` est appelée :
- Via `Form.FormClosed`
- Via `Application.ThreadException` (exception non gérée sur le thread UI)
- Via `AppDomain.UnhandledException` (exception fatale)

Une confirmation est demandée si des jobs de désinstallation sont encore en cours.

---

## 12. Limitations connues

| Limitation | Détail |
|---|---|
| **PowerShell 5.1** | Pas testé avec PowerShell 7+ (les classes COM et WinForms peuvent se comporter différemment) |
| **x64 uniquement** | Le script redirige vers `Sysnative` si lancé en x86, mais l'UI est conçue pour x64 |
| **Pas de signature de code** | Nécessite `-ExecutionPolicy Bypass` |
| **Regedit automation** | Dépend de la structure UI de regedit (fragile entre versions de Windows). Supporte EN/FR via scan des menus |
| **RemoteRegistry** | Nécessaire pour certaines opérations Tab 4 Cleanup. L'outil peut le démarrer mais ne peut pas le configurer si le service est désactivé par GPO |
| **Détection utilisateur console** | Utilise WTS API puis WMI en fallback. Ne fonctionne pas si aucun utilisateur n'est connecté en console |
| **Détection parent-enfant** | Basée sur les Dependencies registry. Peut manquer des relations si l'installeur n'utilise pas ce mécanisme standard (exemple : redist c++)|
| **Scan de Components** | Ne scanne que `S-1-5-18` (LocalSystem). Les composants installés per-user sous d'autres SIDs ne sont pas détectés |
| **Force ACL** | Peut échouer sur des clés protégées par TrustedInstaller (nécessiterait un token spécifique) |
