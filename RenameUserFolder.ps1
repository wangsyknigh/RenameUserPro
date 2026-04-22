<#
.SYNOPSIS
    将 Windows 用户文件夹从中文名改为英文名
.DESCRIPTION
    适用于 Windows 10/11，解决因中文用户文件夹名导致部分专业软件运行异常的问题。
    通过修改注册表 ProfileImagePath 并创建符号链接，将用户文件夹重命名为英文，
    同时保持账户登录名不变。支持单用户/批量模式，内置回滚机制。
.NOTES
    版本：1.0
    要求：管理员权限，PowerShell 5.1+
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$NewUserName,
    [switch]$BatchMode,
    [switch]$AutoRollback,
    [switch]$SkipConfirm,
    [switch]$NoWait,
    [switch]$ForceLocalAccount,
    [switch]$StopOnError,
    [switch]$DryRun,
    [switch]$RebuildSearchIndex,
    [switch]$SkipShortcutRepair,
    [string]$ReportPath
)

# =====================================================
# 基础配置
# =====================================================
$ErrorActionPreference = "Stop"
$ScriptVersion = "1.0"
$LogFile = Join-Path $env:TEMP "RenameUserPro_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:RollbackActions = @()
$script:RollbackKeys = @{}
$script:UsedNames = @()
$script:NameCache = @{}
$script:IsMicrosoftAccount = $false
$script:NoWaitFlag = $false
$script:LoadedHives = @()
$script:Results = @()
$script:AllSkipped = $false
$script:RollbackInvoked = $false

$UsersBasePath = Join-Path $env:SystemDrive "Users"

# 注册 Ctrl+C 事件处理
$consoleHandler = [ConsoleCancelEventHandler]{
    param($sender, $e)
    $e.Cancel = $true
    if (-not $script:RollbackInvoked) {
        $script:RollbackInvoked = $true
        Write-Host "`n检测到 Ctrl+C，正在执行回滚..." -ForegroundColor Red
        Invoke-Rollback
    }
}
[Console]::CancelKeyPress += $consoleHandler

# 全局 Mutex 防止并发
$script:Mutex = New-Object System.Threading.Mutex($false, "Global\RenameUserPro_Unique_Mutex")
if (-not $script:Mutex.WaitOne(0)) {
    Write-Host "另一个脚本实例正在运行，请稍后再试。" -ForegroundColor Red
    exit 1
}

# 启用长路径支持
function Enable-LongPaths {
    $key = "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem"
    $value = Get-ItemProperty -Path $key -Name "LongPathsEnabled" -ErrorAction SilentlyContinue
    if (-not $value -or $value.LongPathsEnabled -ne 1) {
        if ($DryRun) {
            Write-Log "[DRYRUN] 将启用长路径支持" "DRYRUN"
        } else {
            try {
                Set-ItemProperty -Path $key -Name "LongPathsEnabled" -Value 1 -Force
                Write-Log "已启用长路径支持" "SUCCESS"
            } catch { Write-Log "启用长路径失败: $_" "WARN" }
        }
    }
}

# =====================================================
# 日志系统
# =====================================================
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "ERROR"   { "Red" }
        "WARN"    { "Yellow" }
        "SUCCESS" { "Green" }
        "STEP"    { "Cyan" }
        "HINT"    { "Magenta" }
        "SKIP"    { "DarkGray" }
        "MSA"     { "Blue" }
        "PERF"    { "DarkCyan" }
        "DRYRUN"  { "DarkYellow" }
        default   { "White" }
    }
    $line = "$time [$Level] $Message"
    Write-Host $line -ForegroundColor $color
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch {}
}

function Wait-ForExit {
    param([string]$Message = "按 Enter 退出")
    if ($BatchMode -or $NoWait -or $script:NoWaitFlag) {
        Write-Log "操作完成，自动退出" "HINT"
        return
    }
    Read-Host "`n$Message"
}

function Invoke-IfNotDryRun {
    param([scriptblock]$Action, [string]$Description = "执行操作")
    if ($DryRun) {
        Write-Log "[DRYRUN] 跳过: $Description" "DRYRUN"
        return $false
    }
    & $Action
    return $true
}

# =====================================================
# 回滚机制
# =====================================================
function Add-RollbackAction {
    param(
        [string]$Type,
        [string]$Description,
        [scriptblock]$Action,
        [int]$Priority = 50
    )
    $uniqueId = "$Type|$Description|$(Get-Date -Format 'yyyyMMddHHmmssfff')|$(Get-Random)"
    if ($script:RollbackKeys.ContainsKey($uniqueId)) { return }
    $script:RollbackKeys[$uniqueId] = $true
    $script:RollbackActions += [PSCustomObject]@{
        Id          = $uniqueId
        Type        = $Type
        Description = $Description
        Action      = $Action
        Priority    = $Priority
        Timestamp   = Get-Date
    }
    Write-Log "已记录回滚点 [$Priority]: $Description" "HINT"
}

function Invoke-Rollback {
    if ($script:RollbackInvoked) { return }
    $script:RollbackInvoked = $true
    Write-Log "===================================================" "ERROR"
    Write-Log "发生错误或用户中断，开始自动回滚..." "ERROR"
    Write-Log "===================================================" "ERROR"
    $sortedActions = $script:RollbackActions | Sort-Object Priority
    foreach ($action in $sortedActions) {
        Write-Log "回滚: $($action.Description)" "WARN"
        if ($DryRun) {
            Write-Log "[DRYRUN] 将执行回滚: $($action.Description)" "DRYRUN"
        } else {
            try {
                & $action.Action
                Write-Log "  ✓ 成功" "SUCCESS"
            } catch {
                Write-Log "  ✗ 失败: $_" "ERROR"
            }
        }
    }
    Write-Log "回滚完成，系统已恢复原状" "SUCCESS"
    exit 1
}

# =====================================================
# 系统兼容性检查
# =====================================================
function Test-SystemCompatibility {
    Write-Log "检查系统兼容性..." "STEP"
    $osInfo = Get-CimInstance Win32_OperatingSystem
    $buildNumber = [int]$osInfo.BuildNumber
    $caption = $osInfo.Caption
    $isWin10 = $caption -like "*Windows 10*"
    $isWin11 = $caption -like "*Windows 11*"
    if (-not ($isWin10 -or $isWin11)) {
        throw "此脚本仅支持 Windows 10 和 Windows 11。检测到: $caption"
    }
    if ($isWin10 -and $buildNumber -lt 10240) {
        throw "Windows 10 版本过旧（Build $buildNumber）。需要 Build 10240 或更高版本。"
    }
    if ($isWin11 -and $buildNumber -lt 22000) {
        throw "Windows 11 版本过旧（Build $buildNumber）。需要 Build 22000 (21H2) 或更高版本。"
    }
    Write-Log "系统兼容: $caption (Build $buildNumber)" "SUCCESS"
    if ($PSVersionTable.PSVersion.Major -lt 5 -or 
        ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
        throw "需要 PowerShell 5.1 或更高版本。当前版本: $($PSVersionTable.PSVersion)"
    }
    Write-Log "PowerShell 版本: $($PSVersionTable.PSVersion)" "SUCCESS"
}

# =====================================================
# 预检（BitLocker、WSL、漫游、压缩/加密）
# =====================================================
function Test-PreCheck {
    param([string]$UserName, [string]$ProfilePath, [string]$SID)
    
    try {
        $driveRoot = (Get-Item $ProfilePath -ErrorAction SilentlyContinue).PSDrive.Root
        if (-not $driveRoot) { $driveRoot = $ProfilePath.Substring(0, 3) }
        $blv = Get-BitLockerVolume -MountPoint $driveRoot -ErrorAction SilentlyContinue
        if ($blv -and $blv.ProtectionStatus -eq 1) {
            Write-Log "⚠️  驱动器 $driveRoot 已启用 BitLocker 加密" "WARN"
            Write-Log "   重命名后可能导致解锁失败，建议先暂停保护或解密" "ERROR"
            if (-not $SkipConfirm) {
                $ans = Read-Host "是否继续? (输入 YES 继续)"
                if ($ans -ne "YES") { throw "用户取消操作" }
            }
        }
    } catch { Write-Log "BitLocker 检查失败: $_" "WARN" }
    
    try {
        $folder = Get-Item $ProfilePath -Force -ErrorAction SilentlyContinue
        if ($folder.Attributes -band [System.IO.FileAttributes]::Compressed) {
            Write-Log "⚠️  用户文件夹启用了 NTFS 压缩，重命名后可能影响性能，但数据安全。" "WARN"
        }
        if ($folder.Attributes -band [System.IO.FileAttributes]::Encrypted) {
            Write-Log "⚠️  用户文件夹启用了 EFS 加密！重命名后可能导致无法访问。" "ERROR"
            if (-not $SkipConfirm) {
                $ans = Read-Host "是否继续? (输入 YES 继续，否则请先解密文件夹)"
                if ($ans -ne "YES") { throw "用户取消操作" }
            }
        }
    } catch { }
    
    $wslConfig = Join-Path $ProfilePath ".wslconfig"
    if (Test-Path $wslConfig) {
        $content = Get-Content $wslConfig -Raw
        if ($content -match $UserName) {
            Write-Log "⚠️  WSL 配置文件包含旧用户名，重命名后需手动修改 $wslConfig" "WARN"
        }
    }
    $profileKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
    if (Test-Path $profileKey) {
        $centralProfile = (Get-ItemProperty -Path $profileKey -Name "CentralProfile" -ErrorAction SilentlyContinue).CentralProfile
        $roamingProfile = (Get-ItemProperty -Path $profileKey -Name "RoamingProfile" -ErrorAction SilentlyContinue).RoamingProfile
        if ($centralProfile -or $roamingProfile) {
            Write-Log "❌ 检测到用户 $UserName 启用了漫游配置文件或文件夹重定向！" "ERROR"
            Write-Log "   本脚本不支持此类环境，请先联系管理员禁用漫游配置文件。" "HINT"
            throw "漫游配置文件不支持重命名"
        }
    }
}

function Test-FolderRedirection {
    param([string]$UserName, [string]$ProfilePath, [string]$SID)
    $targetShellFolders = "Registry::HKEY_USERS\$SID\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    if (Test-Path $targetShellFolders) {
        $props = Get-ItemProperty $targetShellFolders -ErrorAction SilentlyContinue
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Value -is [string] -and $prop.Value -like "*$ProfilePath*" -and $prop.Value -like "\\*") {
                Write-Log "⚠️  检测到文件夹重定向: $($prop.Name) -> $($prop.Value)" "WARN"
                Write-Log "   重命名后可能导致重定向失效，请先取消重定向" "ERROR"
                if (-not $SkipConfirm) {
                    $ans = Read-Host "是否继续? (输入 YES 继续)"
                    if ($ans -ne "YES") { throw "用户取消操作" }
                }
                break
            }
        }
    }
}

# =====================================================
# 用户登录检测（纯WMI）
# =====================================================
function Test-UserLoggedIn {
    param([string]$UserName)
    try {
        $sessions = Get-CimInstance Win32_LogonSession -ErrorAction Stop | Where-Object { $_.LogonType -in 2,10 }
        foreach ($session in $sessions) {
            $user = Get-CimAssociatedInstance -InputObject $session -ResultClassName Win32_UserAccount -ErrorAction SilentlyContinue
            if ($user -and $user.Name -eq $UserName) {
                Write-Log "用户 $UserName 当前有活动会话！" "ERROR"
                Write-Log "请先完全注销该用户，然后以其他管理员账户（如 Administrator）运行此脚本。" "HINT"
                return $true
            }
        }
    } catch {
        Write-Log "WMI 登录检测失败: $_" "WARN"
        $processes = Get-Process -IncludeUserName -ErrorAction SilentlyContinue | Where-Object { $_.UserName -like "*\$UserName" }
        if ($processes) {
            Write-Log "用户 $UserName 仍有进程在运行，请注销后重试。" "ERROR"
            return $true
        }
    }
    return $false
}

# =====================================================
# 注册表蜂巢加载/卸载（重试+挂起清理）
# =====================================================
function Mount-UserRegistryHive {
    param([string]$SID, [string]$ProfilePath)
    $regPath = "Registry::HKEY_USERS\$SID"
    if (Test-Path $regPath) {
        Write-Log "用户注册表蜂巢已加载: $SID" "SUCCESS"
        return $true
    }
    $ntuserPath = Join-Path $ProfilePath "NTUSER.DAT"
    if (-not (Test-Path $ntuserPath)) {
        Write-Log "无法找到 NTUSER.DAT: $ntuserPath" "ERROR"
        return $false
    }
    if ($DryRun) {
        Write-Log "[DRYRUN] 将加载注册表蜂巢: $SID from $ntuserPath" "DRYRUN"
        return $true
    }
    try {
        $process = Start-Process -FilePath "reg.exe" -ArgumentList "load `"HKU\$SID`" `"$ntuserPath`"" -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            Write-Log "reg load 失败，退出代码: $($process.ExitCode)" "ERROR"
            return $false
        }
        Start-Sleep 1
        if (Test-Path $regPath) {
            Write-Log "已加载用户注册表蜂巢: $SID" "SUCCESS"
            $script:LoadedHives += $SID
            $localSID = $SID
            Add-RollbackAction -Type "Registry" -Description "卸载注册表蜂巢 $localSID" -Action {
                for ($i=0; $i -lt 10; $i++) {
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                    try {
                        $unloadProc = Start-Process -FilePath "reg.exe" -ArgumentList "unload `"HKU\$localSID`"" -NoNewWindow -Wait -PassThru
                        if ($unloadProc.ExitCode -eq 0) { break }
                    } catch { }
                    Start-Sleep 1
                }
                Write-Log "蜂巢卸载完成（或将在下次重启后清理）" "HINT"
            } -Priority 100
            return $true
        }
    } catch {
        Write-Log "加载注册表蜂巢失败: $_" "ERROR"
    }
    return $false
}

# =====================================================
# 辅助检测函数
# =====================================================
function Test-ProfilePathMismatch {
    param([string]$UserName, [string]$SID)
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
    if (!(Test-Path $regPath)) { return $false }
    $actualProfilePath = (Get-ItemProperty -Path $regPath -Name "ProfileImagePath" -ErrorAction SilentlyContinue).ProfileImagePath
    $expectedPath = Join-Path $UsersBasePath $UserName
    
    $containsNonAscii = $actualProfilePath -match '[^\x00-\x7F]'
    if ($containsNonAscii) {
        Write-Log "检测到路径仍包含非 ASCII 字符: $actualProfilePath" "SKIP"
        return $false
    }
    
    if ($actualProfilePath -and ($actualProfilePath -ne $expectedPath)) {
        $pathDir = Split-Path $actualProfilePath -Leaf
        if ($pathDir -ne $UserName -and (Test-Path $actualProfilePath)) {
            Write-Log "检测到路径已修改: $actualProfilePath" "SKIP"
            return $true
        }
    }
    return $false
}

function Test-MicrosoftAccount {
    param([string]$UserName)
    Write-Log "检测账户类型..." "STEP"
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine)
        $userPrincipal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($context, $UserName)
        if ($userPrincipal) {
            $extendedName = $userPrincipal.GetUnderlyingObject().Properties("msDS-UserPrincipalName").Value
            if ($extendedName -like "*@outlook.com" -or $extendedName -like "*@live.com" -or $extendedName -like "*@hotmail.com") {
                Write-Log "检测到 Microsoft 账户关联！" "MSA"
                $script:IsMicrosoftAccount = $true
                return $true
            }
        }
    } catch {
        Write-Log "高级检测失败，回退注册表方式" "WARN"
        $identityPath = "HKLM:\SOFTWARE\Microsoft\IdentityStore\LogonCache"
        if (Test-Path $identityPath) {
            $cacheItems = Get-ChildItem $identityPath -ErrorAction SilentlyContinue
            foreach ($item in $cacheItems) {
                $userProperties = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
                if ($userProperties.UserName -eq $UserName -or $item.PSChildName -like "*$UserName*") {
                    if ($userProperties.ProviderID -eq "MicrosoftAccount" -or 
                        $userProperties.ProviderID -eq "{D6886603-9D2F-4EB2-B667-1971041FA96B}") {
                        Write-Log "检测到 Microsoft 账户关联！" "MSA"
                        $script:IsMicrosoftAccount = $true
                        return $true
                    }
                }
            }
        }
    }
    Write-Log "本地账户，风险较低" "SUCCESS"
    $script:IsMicrosoftAccount = $false
    return $false
}

function Check-OneDriveSyncStatus {
    param([string]$UserName)
    Write-Log "检查 OneDrive 同步状态..." "STEP"
    $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if (-not $oneDriveProcess) {
        Write-Log "OneDrive 未运行，跳过同步检查" "SKIP"
        return $true
    }
    Write-Log "⚠️  OneDrive 正在运行，建议先完成所有文件同步。" "WARN"
    Write-Log "   注意：即使选择继续，未同步的文件在断开链接后仍会保留在本地，但云端可能不会更新。" "HINT"
    if (-not $SkipConfirm) {
        $ans = Read-Host "是否继续？输入 UPLOAD 等待同步，输入 CONTINUE 继续，输入 EXIT 退出"
        if ($ans -eq "UPLOAD") {
            Write-Log "请等待 OneDrive 同步完成，按任意键继续..." "HINT"
            Read-Host
        } elseif ($ans -ne "CONTINUE") {
            throw "用户取消操作"
        }
    }
    return $true
}

function Suspend-OneDrive {
    param([string]$UserName, [string]$NewUserName, [string]$NewProfilePath, [string]$SID)
    Write-Log "处理 OneDrive（安全断开）..." "STEP"
    Check-OneDriveSyncStatus -UserName $UserName
    Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep 2
    $oneDriveExe = Get-Command "OneDrive.exe" -ErrorAction SilentlyContinue
    if (-not $oneDriveExe) {
        $oneDriveExe = "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe"
        if (-not (Test-Path $oneDriveExe)) { $oneDriveExe = "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe" }
    }
    if (Test-Path $oneDriveExe) {
        if ($DryRun) {
            Write-Log "[DRYRUN] 将断开 OneDrive 链接" "DRYRUN"
        } else {
            try {
                Start-Process -FilePath $oneDriveExe -ArgumentList "/unlink" -NoNewWindow -Wait -PassThru
                Write-Log "OneDrive 已断开链接" "SUCCESS"
                Write-Log "重要：为防止云端恢复旧文件夹名，请按以下步骤操作：" "HINT"
                Write-Log "  1. 登录 OneDrive 网页版 (onedrive.live.com)" "HINT"
                Write-Log "  2. 找到原来的文件夹（通常为 '$UserName'），将其重命名为 '$NewUserName' 或移动内容到新文件夹" "HINT"
                Write-Log "  3. 重新链接时选择新路径: $NewProfilePath\OneDrive" "HINT"
                Add-RollbackAction -Type "OneDrive" -Description "重新链接 OneDrive" -Action {
                    Start-Process -FilePath $oneDriveExe -ArgumentList "/background" -WindowStyle Hidden
                } -Priority 50
            } catch {
                Write-Log "断开 OneDrive 失败: $_" "WARN"
            }
        }
    }
}

function Resume-OneDrive {
    param([string]$NewUserName, [string]$SID)
    Write-Log "恢复 OneDrive 配置（需手动重新链接）..." "STEP"
    $newProfilePath = Join-Path $UsersBasePath $NewUserName
    Write-Log "OneDrive 已断开。请在脚本完成后，手动运行 OneDrive 并选择新文件夹路径: $newProfilePath\OneDrive" "HINT"
    Write-Log "重新链接时请选择「使用此文件夹」，不要选择「更改位置」。" "HINT"
}

function Prompt-ConvertMsaToLocal {
    param([string]$UserName)
    Write-Log "准备将微软账户 '$UserName' 转换为本地账户..." "STEP"
    Write-Log "重要：转换后请务必注销并重新登录，然后重新运行此脚本。" "HINT"
    Write-Log "请按以下步骤手动操作：" "STEP"
    Write-Log "  1. 打开「设置」→「账户」→「你的信息」" "HINT"
    Write-Log "  2. 点击「改用本地账户登录」" "HINT"
    Write-Log "  3. 输入当前微软账户密码验证身份" "HINT"
    Write-Log "  4. 设置本地账户用户名和密码（可与原密码相同）" "HINT"
    Write-Log "  5. 注销并使用本地账户重新登录" "HINT"
    Write-Log "  6. 重新运行此脚本（无需 -ForceLocalAccount）" "HINT"
    Write-Log "额外建议：在转换前，断开 OneDrive 链接并等待所有文件同步完成。" "HINT"
    if ($DryRun) {
        Write-Log "[DRYRUN] 将打开设置页面引导用户转换" "DRYRUN"
        return
    }
    $choice = Read-Host "是否现在打开设置页面？(输入 YES 打开，NO 退出脚本)"
    if ($choice -ne "YES") {
        throw "用户取消操作"
    }
    Start-Process "ms-settings:yourinfo"
    Write-Log "设置页面已打开。完成转换并注销后，请重新运行此脚本。" "HINT"
    throw "请完成转换后重新运行脚本。"
}

# =====================================================
# 拼音生成函数
# =====================================================
function Get-UniqueEnglishName {
    param([string]$ChineseName)
    if ($script:NameCache.ContainsKey($ChineseName)) {
        return $script:NameCache[$ChineseName]
    }
    $baseName = Convert-ToPinyin $ChineseName
    $finalName = $baseName
    $maxAttempts = 1000
    $attempt = 0
    while (($finalName -in $script:UsedNames -or (Test-Path (Join-Path $UsersBasePath $finalName))) -and $attempt -lt $maxAttempts) {
        $randomSuffix = Get-Random -Minimum 100 -Maximum 9999
        $finalName = "$baseName$randomSuffix"
        $attempt++
    }
    if ($attempt -ge $maxAttempts) {
        $timestamp = Get-Date -Format "mmss"
        $finalName = "User$timestamp$(Get-Random -Minimum 10 -Maximum 99)"
    }
    $script:UsedNames += $finalName
    $script:NameCache[$ChineseName] = $finalName
    return $finalName
}

function Convert-ToPinyin {
    param([string]$ChineseName)
    $surnameMap = @{
        "王"="Wang";"李"="Li";"张"="Zhang";"刘"="Liu";"陈"="Chen";"杨"="Yang";"黄"="Huang";"赵"="Zhao";"周"="Zhou";"吴"="Wu"
        "徐"="Xu";"孙"="Sun";"马"="Ma";"朱"="Zhu";"胡"="Hu";"郭"="Guo";"何"="He";"高"="Gao";"林"="Lin";"罗"="Luo"
        "郑"="Zheng";"梁"="Liang";"谢"="Xie";"宋"="Song";"唐"="Tang";"许"="Xu";"韩"="Han";"冯"="Feng";"邓"="Deng";"曹"="Cao"
        "彭"="Peng";"曾"="Zeng";"肖"="Xiao";"田"="Tian";"董"="Dong";"袁"="Yuan";"潘"="Pan";"于"="Yu";"蒋"="Jiang";"蔡"="Cai"
        "余"="Yu";"杜"="Du";"叶"="Ye";"程"="Cheng";"魏"="Wei";"苏"="Su";"吕"="Lv";"丁"="Ding";"任"="Ren";"沈"="Shen"
        "姚"="Yao";"卢"="Lu";"姜"="Jiang";"崔"="Cui";"钟"="Zhong";"谭"="Tan";"陆"="Lu";"汪"="Wang";"范"="Fan";"金"="Jin"
        "石"="Shi";"廖"="Liao";"贾"="Jia";"夏"="Xia";"韦"="Wei";"付"="Fu";"方"="Fang";"白"="Bai";"邹"="Zou";"孟"="Meng"
        "熊"="Xiong";"秦"="Qin";"邱"="Qiu";"江"="Jiang";"尹"="Yin";"薛"="Xue";"闫"="Yan";"段"="Duan";"雷"="Lei";"季"="Ji"
        "史"="Shi";"陶"="Tao";"贺"="He";"毛"="Mao";"郝"="Hao";"顾"="Gu";"龚"="Gong";"邵"="Shao";"万"="Wan";"钱"="Qian"
        "严"="Yan";"覃"="Qin";"武"="Wu";"戴"="Dai";"莫"="Mo";"孔"="Kong";"向"="Xiang";"汤"="Tang"
    }
    $firstChar = $ChineseName.Substring(0,1)
    $surname = $surnameMap[$firstChar]
    if ($surname) {
        $randomNum = Get-Random -Minimum 100 -Maximum 999
        return "$surname$randomNum"
    } else {
        return "User$(Get-Random -Minimum 1000 -Maximum 9999)"
    }
}

# =====================================================
# 修复函数
# =====================================================
function Repair-JunctionPoints {
    param([string]$OldProfilePath, [string]$NewProfilePath)
    Write-Log "重建交接点（Junction）和挂载点（MountPoint）..." "STEP"
    # 扫描 Junction 和 MountPoint，深度5
    $allItems = Get-ChildItem -Path $OldProfilePath -Recurse -Depth 5 -Force -ErrorAction SilentlyContinue
    $reparsePoints = $allItems | Where-Object { $_.LinkType -in @('Junction', 'MountPoint') }
    foreach ($item in $reparsePoints) {
        $oldTarget = $item.Target
        if (-not $oldTarget) { continue }
        if ($oldTarget -like "$OldProfilePath*") {
            $newTarget = $oldTarget -replace [regex]::Escape($OldProfilePath), $NewProfilePath
            if ($DryRun) {
                Write-Log "[DRYRUN] 将重建 $($item.LinkType): $($item.FullName) -> $newTarget" "DRYRUN"
            } else {
                try {
                    $tempPath = $item.FullName + "_tmp_$(Get-Random)"
                    Rename-Item -Path $item.FullName -NewName (Split-Path $tempPath -Leaf) -Force
                    if ($item.LinkType -eq 'Junction') {
                        New-Item -Path $item.FullName -ItemType Junction -Target $newTarget -Force
                    } else {
                        # MountPoint 需要使用 New-Item -ItemType Junction 也可以？实际上 MountPoint 是目录链接的一种
                        New-Item -Path $item.FullName -ItemType Junction -Target $newTarget -Force
                    }
                    Remove-Item $tempPath -Force -Recurse -ErrorAction SilentlyContinue
                    Write-Log "  已重建 $($item.LinkType): $($item.FullName)" "SUCCESS"
                } catch {
                    Write-Log "  重建失败: $_" "WARN"
                }
            }
        }
    }
}

function Repair-SystemPath {
    param([string]$OldProfilePath, [string]$NewProfilePath)
    Write-Log "修复系统 PATH 环境变量..." "STEP"
    $systemEnvKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    $pathValue = (Get-ItemProperty -Path $systemEnvKey -Name "PATH" -ErrorAction SilentlyContinue).PATH
    if ($pathValue -and $pathValue -like "*$OldProfilePath*") {
        $newPath = $pathValue -replace [regex]::Escape($OldProfilePath), $NewProfilePath
        $oldPath = $pathValue
        Add-RollbackAction -Type "Registry" -Description "恢复系统 PATH" -Action {
            Set-ItemProperty -Path $systemEnvKey -Name "PATH" -Value $oldPath -Type ExpandString
        } -Priority 25
        if ($DryRun) {
            Write-Log "[DRYRUN] 将更新系统 PATH" "DRYRUN"
        } else {
            Set-ItemProperty -Path $systemEnvKey -Name "PATH" -Value $newPath -Type ExpandString
            Write-Log "系统 PATH 已更新" "SUCCESS"
        }
    } else {
        Write-Log "系统 PATH 无需更新" "SKIP"
    }
}

function Repair-Shortcuts {
    param([string]$OldProfilePath, [string]$NewProfilePath, [string]$SID)
    if ($SkipShortcutRepair) {
        Write-Log "跳过快捷方式修复（-SkipShortcutRepair 已指定）" "SKIP"
        return
    }
    Write-Log "修复快捷方式（开始菜单和桌面）..." "STEP"
    $userShellFolders = "Registry::HKEY_USERS\$SID\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    $programsPath = $null
    $desktopPath = $null
    if (Test-Path $userShellFolders) {
        $props = Get-ItemProperty -Path $userShellFolders -ErrorAction SilentlyContinue
        $programsPath = $props.Programs
        $desktopPath = $props.Desktop
    }
    if (-not $programsPath) {
        $programsPath = Join-Path $NewProfilePath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    }
    if (-not $desktopPath) {
        $desktopPath = Join-Path $NewProfilePath "Desktop"
    }
    if ($programsPath -like "*%*") { $programsPath = [Environment]::ExpandEnvironmentVariables($programsPath) }
    if ($desktopPath -like "*%*") { $desktopPath = [Environment]::ExpandEnvironmentVariables($desktopPath) }
    
    $searchPaths = @($programsPath, $desktopPath, "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs")
    $fixedCount = 0
    $shell = $null
    try {
        $shell = New-Object -ComObject WScript.Shell
        foreach ($basePath in $searchPaths) {
            if (-not (Test-Path $basePath)) { continue }
            $shortcuts = Get-ChildItem $basePath -Recurse -Filter "*.lnk" -ErrorAction SilentlyContinue
            foreach ($shortcut in $shortcuts) {
                try {
                    $link = $shell.CreateShortcut($shortcut.FullName)
                    $target = $link.TargetPath
                    if ([string]::IsNullOrEmpty($target)) { continue }
                    if ($target -like "*$OldProfilePath*") {
                        $newTarget = $target -replace [regex]::Escape($OldProfilePath), $NewProfilePath
                        $link.TargetPath = $newTarget
                        $link.Save()
                        $fixedCount++
                        Write-Log "  修复: $($shortcut.Name)" "SUCCESS"
                    }
                } catch {
                    Write-Log "  无法修复快捷方式 $($shortcut.Name): $_" "WARN"
                }
            }
        }
    } finally {
        if ($shell) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null }
    }
    Write-Log "快捷方式修复完成，共修复 $fixedCount 个。" "SUCCESS"
}

function Repair-ShellFolders {
    param([string]$OldName, [string]$NewName, [string]$SID)
    Write-Log "修复 Shell Folders 路径..." "STEP"
    $oldProfilePath = Join-Path $UsersBasePath $OldName
    $newProfilePath = Join-Path $UsersBasePath $NewName
    $shellFolderPaths = @(
        "Registry::HKEY_USERS\$SID\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        "Registry::HKEY_USERS\$SID\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    )
    $foldersToFix = @("AppData","Local AppData","Cache","Cookies","History","Personal","My Pictures","My Music","My Video","Desktop","Favorites","NetHood","PrintHood","Programs","Recent","SendTo","Start Menu","Startup","Templates")
    foreach ($regPath in $shellFolderPaths) {
        if (!(Test-Path $regPath)) { continue }
        try {
            $props = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
            foreach ($folder in $foldersToFix) {
                $currentValue = $props.$folder
                if ($currentValue -and ($currentValue -is [string]) -and ($currentValue -like "*$oldProfilePath*")) {
                    $newValue = $currentValue -replace [regex]::Escape($oldProfilePath), $newProfilePath
                    $localRegPath = $regPath
                    $localFolder = $folder
                    $localCurrentValue = $currentValue
                    Add-RollbackAction -Type "Registry" -Description "恢复 Shell Folder: $folder" -Action {
                        Set-ItemProperty -Path $localRegPath -Name $localFolder -Value $localCurrentValue
                    } -Priority 25
                    if ($DryRun) {
                        Write-Log "[DRYRUN] 将修复 $folder: $newValue" "DRYRUN"
                    } else {
                        Set-ItemProperty -Path $regPath -Name $folder -Value $newValue
                        Write-Log "  修复 $folder: $newValue" "SUCCESS"
                    }
                }
            }
        } catch { Write-Log "  修复 $regPath 失败: $_" "WARN" }
    }
    Write-Log "Shell Folders 修复完成" "SUCCESS"
}

function Repair-StoreApps {
    param([string]$OldName, [string]$NewName, [string]$SID)
    Write-Log "修复 Windows Store 应用路径..." "STEP"
    $oldProfilePath = Join-Path $UsersBasePath $OldName
    $newProfilePath = Join-Path $UsersBasePath $NewName
    $fixedCount = 0
    $appContainerPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppContainer\Mappings"
    if (Test-Path $appContainerPath) {
        try {
            $mappings = Get-ChildItem $appContainerPath -ErrorAction SilentlyContinue
            foreach ($mapping in $mappings) {
                $props = Get-ItemProperty $mapping.PSPath -ErrorAction SilentlyContinue
                if ($props.PackageRootFolder -and $props.PackageRootFolder -like "*$oldProfilePath*") {
                    $oldVal = $props.PackageRootFolder
                    $newPackagePath = $oldVal -replace [regex]::Escape($oldProfilePath), $newProfilePath
                    $localMappingPath = $mapping.PSPath
                    $localOldVal = $oldVal
                    Add-RollbackAction -Type "Registry" -Description "恢复 Store App PackageRootFolder" -Action {
                        Set-ItemProperty -Path $localMappingPath -Name "PackageRootFolder" -Value $localOldVal
                    } -Priority 25
                    if ($DryRun) {
                        Write-Log "[DRYRUN] 将修复 PackageRootFolder: $newPackagePath" "DRYRUN"
                    } else {
                        Set-ItemProperty -Path $mapping.PSPath -Name "PackageRootFolder" -Value $newPackagePath
                        $fixedCount++
                    }
                }
                if ($props.PackageRepository -and $props.PackageRepository -like "*$oldProfilePath*") {
                    $oldVal = $props.PackageRepository
                    $newRepoPath = $oldVal -replace [regex]::Escape($oldProfilePath), $newProfilePath
                    $localMappingPath = $mapping.PSPath
                    $localOldVal = $oldVal
                    Add-RollbackAction -Type "Registry" -Description "恢复 Store App PackageRepository" -Action {
                        Set-ItemProperty -Path $localMappingPath -Name "PackageRepository" -Value $localOldVal
                    } -Priority 25
                    if ($DryRun) {
                        Write-Log "[DRYRUN] 将修复 PackageRepository: $newRepoPath" "DRYRUN"
                    } else {
                        Set-ItemProperty -Path $mapping.PSPath -Name "PackageRepository" -Value $newRepoPath
                        $fixedCount++
                    }
                }
            }
            if ($fixedCount -gt 0 -and !$DryRun) { Write-Log "  修复了 $fixedCount 个应用容器路径" "SUCCESS" }
        } catch { Write-Log "  修复应用容器路径失败: $_" "WARN" }
    }
    $packageStatePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\PackageState"
    if (Test-Path $packageStatePath) {
        try {
            $states = Get-ChildItem $packageStatePath -ErrorAction SilentlyContinue
            $fixedState = 0
            foreach ($state in $states) {
                $userData = Get-ItemProperty $state.PSPath -ErrorAction SilentlyContinue
                foreach ($prop in $userData.PSObject.Properties) {
                    if ($prop.Value -is [string] -and $prop.Value -like "*$oldProfilePath*") {
                        $oldVal = $prop.Value
                        $newValue = $oldVal -replace [regex]::Escape($oldProfilePath), $newProfilePath
                        $localStatePath = $state.PSPath
                        $localPropName = $prop.Name
                        $localOldVal = $oldVal
                        Add-RollbackAction -Type "Registry" -Description "恢复 PackageState $($prop.Name)" -Action {
                            Set-ItemProperty -Path $localStatePath -Name $localPropName -Value $localOldVal
                        } -Priority 25
                        if ($DryRun) {
                            Write-Log "[DRYRUN] 将修复包状态 $($prop.Name): $newValue" "DRYRUN"
                        } else {
                            Set-ItemProperty -Path $state.PSPath -Name $prop.Name -Value $newValue
                            $fixedState++
                        }
                    }
                }
            }
            if ($fixedState -gt 0 -and !$DryRun) { Write-Log "  修复了 $fixedState 个包状态记录" "SUCCESS" }
        } catch { Write-Log "  修复包状态失败: $_" "WARN" }
    }
    Write-Log "提示：极少数 UWP 应用（如计算器、照片）可能需要手动重置（设置 → 应用 → 找到应用 → 高级选项 → 重置）" "HINT"
}

function Repair-UserPath {
    param([string]$OldProfilePath, [string]$NewProfilePath, [string]$SID)
    Write-Log "修复用户 PATH 环境变量..." "STEP"
    $envPath = "Registry::HKEY_USERS\$SID\Environment"
    if (Test-Path $envPath) {
        try {
            $pathValue = (Get-ItemProperty -Path $envPath -Name "PATH" -ErrorAction SilentlyContinue).PATH
            if ($pathValue -and ($pathValue -like "*$OldProfilePath*")) {
                $oldVal = $pathValue
                $newPath = $pathValue -replace [regex]::Escape($OldProfilePath), $NewProfilePath
                $localEnvPath = $envPath
                $localOldVal = $oldVal
                Add-RollbackAction -Type "Registry" -Description "恢复用户 PATH" -Action {
                    Set-ItemProperty -Path $localEnvPath -Name "PATH" -Value $localOldVal -Type ExpandString
                } -Priority 25
                if ($DryRun) {
                    Write-Log "[DRYRUN] 将更新用户 PATH: $newPath" "DRYRUN"
                } else {
                    Set-ItemProperty -Path $envPath -Name "PATH" -Value $newPath -Type ExpandString
                    Write-Log "  用户 PATH 已更新" "SUCCESS"
                }
            } else {
                Write-Log "  用户 PATH 无需更新" "SKIP"
            }
        } catch {
            Write-Log "  更新用户 PATH 失败: $_" "WARN"
        }
    }
}

function Repair-DeepPermissions {
    param([string]$NewUserName, [string]$SID)
    Write-Log "执行深度权限修复（仅目录本身所有权+继承）..." "STEP"
    $newProfilePath = Join-Path $UsersBasePath $NewUserName
    if ($DryRun) {
        Write-Log "[DRYRUN] 将执行: takeown /F $newProfilePath" "DRYRUN"
        Write-Log "[DRYRUN] 将执行: icacls $newProfilePath /grant `"$NewUserName:(OI)(CI)F`" /inheritance:e" "DRYRUN"
    } else {
        try {
            takeown /F $newProfilePath 2>$null | Out-Null
            icacls $newProfilePath /grant "$($NewUserName):(OI)(CI)F" /inheritance:e 2>$null | Out-Null
            Write-Log "  权限修复完成（目录所有者已更新，继承已启用）" "SUCCESS"
        } catch {
            Write-Log "  权限修复失败: $_" "WARN"
        }
    }
    try {
        $userRegPath = "Registry::HKEY_USERS\$SID"
        if (Test-Path $userRegPath) {
            $regAcl = Get-Acl $userRegPath
            $userRule = New-Object System.Security.AccessControl.RegistryAccessRule($NewUserName,"FullControl","ContainerInherit,ObjectInherit","None","Allow")
            $regAcl.SetAccessRule($userRule)
            if ($DryRun) {
                Write-Log "[DRYRUN] 将修复用户注册表权限: $userRegPath" "DRYRUN"
            } else {
                Set-Acl $userRegPath $regAcl
                Write-Log "  用户注册表权限已修复" "SUCCESS"
            }
        }
    } catch {
        Write-Log "  修复注册表权限失败: $_" "WARN"
    }
}

function Repair-WindowsTerminal {
    param([string]$OldProfilePath, [string]$NewProfilePath)
    Write-Log "修复 Windows Terminal 配置..." "STEP"
    $terminalPackages = Get-ChildItem "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_*" -ErrorAction SilentlyContinue
    foreach ($pkg in $terminalPackages) {
        $settingsJson = Join-Path $pkg.FullName "LocalState\settings.json"
        if (Test-Path $settingsJson) {
            try {
                $content = Get-Content $settingsJson -Raw -Encoding UTF8
                if ($content -match [regex]::Escape($OldProfilePath)) {
                    $oldContent = $content
                    $newContent = $content -replace [regex]::Escape($OldProfilePath), $NewProfilePath
                    $localSettingsJson = $settingsJson
                    $localOldContent = $oldContent
                    Add-RollbackAction -Type "File" -Description "恢复 Windows Terminal 配置" -Action {
                        Set-Content -Path $localSettingsJson -Value $localOldContent -Encoding UTF8 -NoNewline
                    } -Priority 30
                    if ($DryRun) {
                        Write-Log "[DRYRUN] 将更新 Windows Terminal 配置: $settingsJson" "DRYRUN"
                    } else {
                        Set-Content -Path $settingsJson -Value $newContent -Encoding UTF8 -NoNewline
                        Write-Log "  Windows Terminal 配置已更新" "SUCCESS"
                    }
                } else {
                    Write-Log "  Windows Terminal 配置无需更新" "SKIP"
                }
            } catch {
                Write-Log "  更新 Windows Terminal 配置失败: $_" "WARN"
            }
        }
    }
}

function Stop-WSearchGracefully {
    $wsearch = Get-Service WSearch -ErrorAction SilentlyContinue
    if ($wsearch -and $wsearch.Status -eq "Running") {
        Write-Log "优雅停止 WSearch 服务..." "STEP"
        if ($DryRun) {
            Write-Log "[DRYRUN] 将停止 WSearch 服务" "DRYRUN"
        } else {
            try {
                Stop-Service WSearch -Force -ErrorAction Stop
                $wsearch.WaitForStatus("Stopped", [TimeSpan]::FromSeconds(10))
                Write-Log "WSearch 服务已停止" "SUCCESS"
            } catch {
                Write-Log "停止 WSearch 失败: $_" "WARN"
            }
        }
    }
}

function Start-WSearchAndRebuildIndex {
    if (-not $RebuildSearchIndex) { 
        Start-Service WSearch -ErrorAction SilentlyContinue
        return 
    }
    Write-Log "重建 Windows Search 索引（增量方式）..." "STEP"
    if ($DryRun) {
        Write-Log "[DRYRUN] 将重建索引" "DRYRUN"
        return
    }
    try {
        $crawler = New-Object -ComObject Microsoft.Search.Crawler
        $crawler.RebuildIndex()
        Write-Log "索引重建已触发，将在后台完成" "SUCCESS"
    } catch {
        Write-Log "触发索引重建失败，尝试删除数据库" "WARN"
        try {
            Stop-Service WSearch -Force
            $indexPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
            if (Test-Path $indexPath) { Remove-Item $indexPath -Force }
            Start-Service WSearch
            Write-Log "索引数据库已重建" "SUCCESS"
        } catch { Write-Log "重建索引失败: $_" "WARN" }
    }
}

function Invoke-SystemRefresh {
    Write-Log "通知系统刷新配置..." "STEP"
    if ($DryRun) {
        Write-Log "[DRYRUN] 将发送 WM_SETTINGCHANGE 刷新环境变量" "DRYRUN"
        return
    }
    try {
        [Environment]::SetEnvironmentVariable("TEMP", [Environment]::GetEnvironmentVariable("TEMP","User"), "Process")
        [Environment]::SetEnvironmentVariable("TMP", [Environment]::GetEnvironmentVariable("TMP","User"), "Process")
        $code = @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
    public static void RefreshEnvironment() {
        IntPtr result;
        SendMessageTimeout((IntPtr)0xFFFF, 0x1A, IntPtr.Zero, "Environment", 2, 5000, out result);
    }
}
"@
        Add-Type -TypeDefinition $code -Language CSharp -ErrorAction SilentlyContinue
        [Win32]::RefreshEnvironment()
        Write-Log "  系统配置刷新已发送" "SUCCESS"
    } catch { Write-Log "  系统刷新失败（非致命）: $_" "WARN" }
}

function Check-ScheduledTasks {
    param([string]$OldProfilePath)
    Write-Log "检查任务计划程序中引用旧路径的任务..." "STEP"
    $tasksPath = "C:\Windows\System32\Tasks"
    if (-not (Test-Path $tasksPath)) {
        Write-Log "  任务计划目录不存在，跳过" "SKIP"
        return
    }
    $affectedTasks = @()
    Get-ChildItem $tasksPath -File | ForEach-Object {
        $content = Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue
        if ($content -match [regex]::Escape($OldProfilePath)) {
            $affectedTasks += $_.Name
            Write-Log "  ⚠️  任务 '$($_.Name)' 包含旧路径引用" "WARN"
        }
    }
    if ($affectedTasks.Count -eq 0) {
        Write-Log "  未发现受影响的任务" "SUCCESS"
    } else {
        Write-Log "  共发现 $($affectedTasks.Count) 个任务可能受影响，请手动检查或重新配置。" "HINT"
    }
}

# =====================================================
# 核心重命名逻辑（整合所有修复）
# =====================================================
function Rename-SingleUser {
    param([string]$OldName, [string]$NewName, [string]$SID, [switch]$Force = $false)
    
    if ($NewName -notmatch '^[a-zA-Z0-9_]+$') {
        throw "新用户名只能包含字母、数字和下划线，当前值: '$NewName'"
    }
    
    if (Test-UserLoggedIn -UserName $OldName) {
        throw "用户 $OldName 当前已登录，请注销后重试。"
    }
    
    $oldProfilePath = Join-Path $UsersBasePath $OldName
    $newProfilePath = Join-Path $UsersBasePath $NewName
    
    Test-PreCheck -UserName $OldName -ProfilePath $oldProfilePath -SID $SID
    
    $isMSA = Test-MicrosoftAccount -UserName $OldName
    
    if ($ForceLocalAccount -and $isMSA) {
        Write-Log "检测到微软账户，且已指定 -ForceLocalAccount，需要先转换为本地账户。" "STEP"
        Prompt-ConvertMsaToLocal -UserName $OldName
    } elseif ($isMSA -and !$ForceLocalAccount) {
        Write-Log "检测到微软账户，但未指定 -ForceLocalAccount。请先手动转换为本地账户，或使用 -ForceLocalAccount 参数引导转换。" "ERROR"
        throw "微软账户需要先断开同步，请添加 -ForceLocalAccount 参数重新运行。"
    }
    
    if (!$Force) {
        $isAlreadyModified = Test-ProfilePathMismatch -UserName $OldName -SID $SID
        if ($isAlreadyModified) {
            Write-Log "用户 $OldName 可能已修改过，跳过处理" "SKIP"
            return [PSCustomObject]@{ User=$OldName; Status="Skipped"; Reason="Already modified" }
        }
    }
    
    Write-Log "原路径: $oldProfilePath"
    Write-Log "新路径: $newProfilePath"
    Write-Log "【重要提示】此操作仅修改用户文件夹名称，不改变您的账户登录名（%USERNAME% 仍是 $OldName）。" "HINT"
    
    if (!(Test-Path $oldProfilePath)) { throw "原用户目录不存在: $oldProfilePath" }
    if (Test-Path $newProfilePath) { throw "新目录已存在: $newProfilePath" }
    
    if (-not $DryRun) {
        try {
            Checkpoint-Computer -Description "RenameUser_$OldName" -RestorePointType MODIFY_SETTINGS -ErrorAction SilentlyContinue
            Write-Log "还原点创建成功" "SUCCESS"
        } catch { Write-Log "无法创建还原点: $_" "WARN" }
    }
    
    Suspend-OneDrive -UserName $OldName -NewUserName $NewName -NewProfilePath $newProfilePath -SID $SID
    
    $backupSuffix = Get-Date -Format "yyyyMMdd_HHmmss_fff" + "_" + [System.Guid]::NewGuid().ToString().Substring(0,8)
    $backupOldPath = "$oldProfilePath.old.$backupSuffix"
    
    $snapshot_oldPath = $oldProfilePath
    $snapshot_newPath = $newProfilePath
    $snapshot_oldName = $OldName
    
    # ========== 关键顺序修改：先修改 ProfileList，再重命名文件夹 ==========
    $profileKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID"
    $originalPath = (Get-ItemProperty -Path $profileKey -Name "ProfileImagePath").ProfileImagePath
    Add-RollbackAction -Type "Registry" -Description "回滚 ProfileImagePath" -Action {
        Set-ItemProperty -Path $profileKey -Name "ProfileImagePath" -Value $originalPath
    } -Priority 20
    
    if ($DryRun) {
        Write-Log "[DRYRUN] 将修改 ProfileImagePath: $newProfilePath" "DRYRUN"
    } else {
        Set-ItemProperty -Path $profileKey -Name "ProfileImagePath" -Value $newProfilePath
        $friendly = Get-ItemProperty -Path $profileKey -Name "FriendlyName" -ErrorAction SilentlyContinue
        if ($friendly) { Set-ItemProperty -Path $profileKey -Name "FriendlyName" -Value $NewName }
        Write-Log "ProfileList 更新成功" "SUCCESS"
    }
    # ================================================================
    
    # 主回滚动作（修复：捕获变量快照）
    Add-RollbackAction -Type "Folder" -Description "回滚文件夹重命名（主回滚）" -Action {
        if (Test-Path $snapshot_oldPath) {
            $item = Get-Item $snapshot_oldPath -ErrorAction SilentlyContinue
            if ($item -and ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                Remove-Item $snapshot_oldPath -Force -ErrorAction SilentlyContinue
                Write-Log "  已删除符号链接: $snapshot_oldPath" "HINT"
            }
        }
        if (Test-Path $snapshot_newPath) {
            Rename-Item -Path $snapshot_newPath -NewName $snapshot_oldName -Force
            Write-Log "  已恢复文件夹: $snapshot_oldPath" "SUCCESS"
        }
        Write-Log "  如果存在备份目录，请手动检查并清理: $backupOldPath" "HINT"
    } -Priority 10
    
    Stop-WSearchGracefully
    Add-RollbackAction -Type "Service" -Description "恢复 WSearch 服务" -Action { Start-Service WSearch -ErrorAction SilentlyContinue } -Priority 40
    
    # 终止占用进程（增加白名单：Docker, VMware, IDE 等）
    $whitelistProcesses = @("MsMpEng", "SecurityHealthService", "WinDefend", "Sophos", "Avast", "Kaspersky",
                            "docker", "vmware", "devenv", "Code", "idea64", "eclipse", "clion64", "rider64")
    for ($i=0; $i -lt 3; $i++) {
        $processes = Get-Process -ErrorAction SilentlyContinue | Where-Object { 
            $_.Path -and $_.Path -like "$oldProfilePath*" -and $whitelistProcesses -notcontains $_.Name
        }
        if ($processes.Count -eq 0) { break }
        $processes | ForEach-Object { 
            if (-not $DryRun) { 
                try { 
                    Write-Log "终止进程: $($_.Name) (PID: $($_.Id))" "WARN"
                    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue 
                } catch {}
            }
        }
        Start-Sleep 1
    }
    
    # 重启 Explorer（强制终止，确保文件不被占用）
    $currentSessionUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
    if ($currentSessionUser -like "*\$OldName") {
        Write-Log "强制重启 Explorer 以释放文件占用..." "STEP"
        Get-Process explorer -ErrorAction SilentlyContinue | ForEach-Object {
            if (-not $DryRun) {
                try {
                    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                } catch {}
            }
        }
        Start-Sleep 2
        Add-RollbackAction -Type "Service" -Description "恢复 Explorer" -Action { Start-Process explorer } -Priority 50
    } else {
        Write-Log "当前登录用户不是目标用户，跳过 Explorer 重启" "SKIP"
    }
    
    # 重命名文件夹（检查备份冲突）
    if (Test-Path $backupOldPath) {
        throw "备份路径已存在，无法继续: $backupOldPath"
    }
    if ($DryRun) {
        Write-Log "[DRYRUN] 将重命名目录: $oldProfilePath -> $newProfilePath" "DRYRUN"
    } else {
        try {
            Rename-Item -Path $oldProfilePath -NewName $NewName -Force
            Write-Log "文件夹重命名成功" "SUCCESS"
        } catch { throw "文件夹重命名失败: $_" }
    }
    Start-Sleep 1
    
    # 加载用户注册表蜂巢，失败则立即停止（因为后续修复依赖它）
    $hiveLoaded = Mount-UserRegistryHive -SID $SID -ProfilePath $newProfilePath
    if (-not $hiveLoaded) {
        throw "无法加载用户注册表蜂巢，后续修复无法进行。请检查 NTUSER.DAT 是否被锁定。"
    }
    Test-FolderRedirection -UserName $OldName -ProfilePath $oldProfilePath -SID $SID
    
    # 修复 UserManager 中的 UserName
    $userManagerPath = "HKLM:\SOFTWARE\Microsoft\UserManager\Users"
    if (Test-Path $userManagerPath) {
        try {
            $keys = Get-ChildItem $userManagerPath
            foreach ($key in $keys) {
                $name = (Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue).UserName
                if ($name -eq $OldName) {
                    $oldUserName = $name
                    $localKeyPath = $key.PSPath
                    Add-RollbackAction -Type "Registry" -Description "恢复 UserManager 中的 UserName" -Action {
                        Set-ItemProperty -Path $localKeyPath -Name "UserName" -Value $oldUserName
                    } -Priority 25
                    if ($DryRun) {
                        Write-Log "[DRYRUN] 将更新 UserManager 中的 UserName: $NewName" "DRYRUN"
                    } else {
                        Set-ItemProperty -Path $key.PSPath -Name "UserName" -Value $NewName
                        Write-Log "UserManager 已更新" "SUCCESS"
                    }
                }
            }
        } catch { Write-Log "UserManager 修改失败 (非致命)" "WARN" }
    }
    
    # 修复环境变量
    $envRegPaths = @(
        "Registry::HKEY_USERS\$SID\Environment",
        "Registry::HKEY_USERS\$SID\Volatile Environment"
    )
    foreach ($userEnvReg in $envRegPaths) {
        if (Test-Path $userEnvReg) {
            try {
                $envProps = Get-ItemProperty $userEnvReg -ErrorAction SilentlyContinue
                $realProps = $envProps.PSObject.Properties | Where-Object { $_.MemberType -eq "NoteProperty" }
                foreach ($prop in $realProps) {
                    if ($prop.Value -is [string] -and $prop.Value -like "*$oldProfilePath*") {
                        $oldVal = $prop.Value
                        $newValue = $oldVal -replace [regex]::Escape($oldProfilePath), $newProfilePath
                        $localEnvReg = $userEnvReg
                        $localPropName = $prop.Name
                        $localOldVal = $oldVal
                        Add-RollbackAction -Type "Registry" -Description "恢复环境变量 $($prop.Name)" -Action {
                            Set-ItemProperty -Path $localEnvReg -Name $localPropName -Value $localOldVal
                        } -Priority 25
                        if ($DryRun) {
                            Write-Log "[DRYRUN] 将更新环境变量 $($prop.Name): $newValue" "DRYRUN"
                        } else {
                            Set-ItemProperty -Path $userEnvReg -Name $prop.Name -Value $newValue
                            Write-Log "环境变量 $($prop.Name) 已更新" "SUCCESS"
                        }
                    }
                }
            } catch { Write-Log "环境变量修改失败 (非致命): $_" "WARN" }
        }
    }
    
    # 调用所有修复函数
    Repair-UserPath -OldProfilePath $oldProfilePath -NewProfilePath $newProfilePath -SID $SID
    Repair-SystemPath -OldProfilePath $oldProfilePath -NewProfilePath $newProfilePath
    Repair-ShellFolders -OldName $OldName -NewName $NewName -SID $SID
    Repair-StoreApps -OldName $OldName -NewName $NewName -SID $SID
    Repair-DeepPermissions -NewUserName $NewName -SID $SID
    Repair-WindowsTerminal -OldProfilePath $oldProfilePath -NewProfilePath $newProfilePath
    Repair-JunctionPoints -OldProfilePath $oldProfilePath -NewProfilePath $newProfilePath
    Repair-Shortcuts -OldProfilePath $oldProfilePath -NewProfilePath $newProfilePath -SID $SID
    
    # 创建符号链接（简化逻辑，因重命名后旧路径已不存在）
    if ($DryRun) {
        Write-Log "[DRYRUN] 将创建符号链接: $oldProfilePath -> $newProfilePath" "DRYRUN"
    } else {
        try {
            if (Test-Path $oldProfilePath) {
                # 防御性代码：正常情况下旧路径应不存在，但若存在则处理
                $item = Get-Item $oldProfilePath -ErrorAction SilentlyContinue
                if ($item -and ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                    Write-Log "符号链接已存在，跳过创建" "SKIP"
                } else {
                    $children = Get-ChildItem $oldProfilePath -Force -ErrorAction SilentlyContinue
                    if ($children.Count -eq 0) {
                        Rename-Item -Path $oldProfilePath -NewName (Split-Path $backupOldPath -Leaf) -Force
                        $args = @('/c','mklink','/J',$oldProfilePath,$newProfilePath)
                        $output = & cmd @args 2>&1
                        if ($LASTEXITCODE -ne 0) { throw "符号链接创建失败: $output" }
                        Write-Log "符号链接创建成功" "SUCCESS"
                    } else {
                        Write-Log "旧路径非空且不是链接，无法自动创建链接。请手动清理或跳过。" "ERROR"
                        Write-Log "您可以跳过创建符号链接（输入 2），或退出脚本（输入 3）。" "WARN"
                        $choice = Read-Host "请输入选项 (2=跳过创建符号链接, 3=退出脚本)"
                        if ($choice -eq "2") {
                            Write-Log "跳过创建符号链接" "SKIP"
                        } else {
                            throw "用户取消操作"
                        }
                    }
                }
            } else {
                $args = @('/c','mklink','/J',$oldProfilePath,$newProfilePath)
                $output = & cmd @args 2>&1
                if ($LASTEXITCODE -ne 0) { throw "符号链接创建失败: $output" }
                Write-Log "符号链接创建成功" "SUCCESS"
            }
        } catch {
            Write-Log "符号链接创建异常: $_" "ERROR"
            throw
        }
    }
    
    Resume-OneDrive -NewUserName $NewName -SID $SID
    Check-ScheduledTasks -OldProfilePath $oldProfilePath
    Start-WSearchAndRebuildIndex
    Invoke-SystemRefresh
    
    return [PSCustomObject]@{ 
        User=$OldName; Status="Success"; NewName=$NewName; IsMSA=$isMSA
    }
}

# =====================================================
# 批量处理（大小写不敏感确认）
# =====================================================
function Start-BatchProcess {
    param([array]$UserList)
    Write-Log "===================================================" "STEP"
    Write-Log "批量处理模式 - 共 $($UserList.Count) 个用户" "STEP"
    Write-Log "注意：微软账户用户将被跳过，请先手动转换为本地账户" "WARN"
    Write-Log "===================================================" "STEP"
    
    Write-Log "即将处理以下用户：" "STEP"
    for ($i=0; $i -lt $UserList.Count; $i++) {
        $suggested = Get-UniqueEnglishName $UserList[$i].Name
        Write-Host "  $($UserList[$i].Name) -> $suggested" -ForegroundColor Cyan
    }
    Write-Log "" "HINT"
    $confirm = Read-Host "请输入 CONFIRM ALL 确认开始批量处理，其他任何输入将退出"
    if ($confirm.ToUpper() -ne "CONFIRM ALL") {
        Write-Log "用户取消批量处理" "ERROR"
        return @()
    }
    
    $results = @()
    $anySuccess = $false
    $anyFailed = $false
    for ($i=0; $i -lt $UserList.Count; $i++) {
        $user = $UserList[$i]
        $suggestedName = Get-UniqueEnglishName $user.Name
        Write-Log ""; Write-Log "[$($i+1)/$($UserList.Count)] 处理: $($user.Name) -> $suggestedName" "STEP"
        $isMsa = Test-MicrosoftAccount -UserName $user.Name
        if ($isMsa) {
            Write-Log "跳过：用户 $($user.Name) 是微软账户，请先手动转换为本地账户。" "SKIP"
            $results += [PSCustomObject]@{ User=$user.Name; Status="Skipped"; Reason="Microsoft Account" }
            continue
        }
        try {
            $result = Rename-SingleUser -OldName $user.Name -NewName $suggestedName -SID $user.SID
            $results += $result
            if ($result.Status -eq "Success") { $anySuccess = $true }
        } catch {
            $errMsg = $_.Exception.Message
            $results += [PSCustomObject]@{ User=$user.Name; Status="Failed"; Error=$errMsg; Reason=$errMsg }
            Write-Log "错误: $errMsg" "ERROR"
            $anyFailed = $true
            if ($StopOnError) { throw "批量处理因错误停止" }
        }
    }
    Write-Log ""; Write-Log "===================================================" "STEP"
    Write-Log "批量处理完成报告" "STEP"
    Write-Log "===================================================" "STEP"
    $success = $results | Where-Object { $_.Status -eq "Success" }
    $skipped = $results | Where-Object { $_.Status -eq "Skipped" }
    $failed = $results | Where-Object { $_.Status -eq "Failed" }
    Write-Log "成功: $($success.Count) 个" "SUCCESS"
    Write-Log "跳过: $($skipped.Count) 个" $(if($skipped.Count -gt 0){"WARN"}else{"SUCCESS"})
    Write-Log "失败: $($failed.Count) 个" $(if($failed.Count -gt 0){"ERROR"}else{"SUCCESS"})
    if ($failed.Count -gt 0) {
        Write-Log ""; Write-Log "失败详情：" "ERROR"
        $failed | ForEach-Object { Write-Log "  $($_.User): $($_.Reason)" "ERROR" }
    }
    $script:AllSkipped = (-not $anySuccess -and -not $anyFailed)
    if ($script:AllSkipped) {
        Write-Log "警告：所有用户均被跳过（可能都是微软账户），未执行任何修改。" "WARN"
    }
    return $results
}

# =====================================================
# 主程序入口
# =====================================================
Enable-LongPaths
Write-Log "===================================================" "STEP"
Write-Log " RenameUserPro.ps1  Version $ScriptVersion (最终稳定版)" "STEP"
Write-Log " Windows 10/11 专用版 - 已修复所有已知问题" "STEP"
if ($DryRun) { Write-Log " **** DRYRUN MODE - 不会实际修改任何文件或注册表 **** " "DRYRUN" }
if ($SkipShortcutRepair) { Write-Log " **** 已跳过快捷方式修复（-SkipShortcutRepair） **** " "HINT" }
Write-Log "===================================================" "STEP"

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "必须使用管理员权限运行脚本" "ERROR"
    exit 1
}

try { Test-SystemCompatibility } catch { Write-Log "系统兼容性检查失败: $_" "ERROR"; exit 1 }

if ($BatchMode -and $NewUserName) {
    Write-Log "错误: -BatchMode 和 -NewUserName 不能同时使用" "ERROR"; exit 1
}

$currentUser = $env:USERNAME
if ($currentUser -ne "Administrator") {
    Write-Log "当前账户: $currentUser" "WARN"
    Write-Log "警告: 建议在 Administrator 账户下运行" "WARN"
    $continue = Read-Host "是否继续在 $currentUser 账户下运行? (输入 RISK 继续，或 EXIT 退出)"
    if ($continue -ne "RISK") { exit 0 }
}

if (-not (Get-Module -ListAvailable -Name Microsoft.PowerShell.LocalAccounts)) {
    Write-Log "LocalAccounts 模块不可用，将使用 WMI 获取用户列表" "SKIP"
} else {
    Import-Module Microsoft.PowerShell.LocalAccounts -ErrorAction SilentlyContinue
}

Write-Log "查找系统普通用户..."
try {
    if (Get-Command Get-LocalUser -ErrorAction SilentlyContinue) {
        $allUsers = Get-LocalUser | Where-Object {
            $_.Enabled -and $_.Name -notmatch "Administrator|Guest|DefaultAccount|WDAGUtilityAccount"
        } | ForEach-Object {
            $sid = (New-Object System.Security.Principal.NTAccount($_.Name)).Translate([System.Security.Principal.SecurityIdentifier]).Value
            [PSCustomObject]@{ Name=$_.Name; SID=$sid; LocalAccount=$true }
        }
    } else {
        $allUsers = Get-CimInstance Win32_UserAccount | Where-Object {
            $_.LocalAccount -and $_.Name -notmatch "Administrator|Guest|DefaultAccount|WDAGUtilityAccount"
        } | ForEach-Object {
            [PSCustomObject]@{ Name=$_.Name; SID=$_.SID; LocalAccount=$true }
        }
    }
} catch {
    Write-Log "获取用户列表失败: $_" "ERROR"; exit 1
}

if (!$allUsers) { Write-Log "未找到可修改的用户账户" "ERROR"; $script:NoWaitFlag=$true; Wait-ForExit; exit 1 }

$usersToProcess = @()
foreach ($user in $allUsers) {
    if ($user.Name -match "[\p{IsCJKUnifiedIdeographs}\p{IsCJKUnifiedIdeographsExtensionA}\p{IsCJKCompatibilityIdeographs}]") {
        $isModified = Test-ProfilePathMismatch -UserName $user.Name -SID $user.SID
        if ($isModified) { Write-Log "跳过（已修改过）: $($user.Name)" "SKIP" }
        else { $usersToProcess += $user; Write-Log "待处理: $($user.Name)" "WARN" }
    } else { Write-Log "已英文: $($user.Name) (无需处理)" "SUCCESS" }
}

if ($usersToProcess.Count -eq 0) {
    Write-Log ""; Write-Log "✓ 没有需要处理的中文用户名" "SUCCESS"
    $script:NoWaitFlag=$true; Wait-ForExit; exit 0
}

Write-Log ""; Write-Log "共 $($usersToProcess.Count) 个用户需要处理" "STEP"

if ($BatchMode) {
    $script:Results = Start-BatchProcess -UserList $usersToProcess
} else {
    if ($usersToProcess.Count -eq 1) {
        $selectedUser = $usersToProcess[0]
        Write-Log ""; Write-Log "目标用户: $($selectedUser.Name)" "STEP"
        $suggestedName = Get-UniqueEnglishName $selectedUser.Name
        Write-Log "建议新名: $suggestedName" "HINT"
        if (-not $NewUserName) {
            $NewUserName = Read-Host "请输入新用户名(直接回车使用建议名 $suggestedName)"
            if (-not $NewUserName) { $NewUserName = $suggestedName }
        }
        try {
            $result = Rename-SingleUser -OldName $selectedUser.Name -NewName $NewUserName -SID $selectedUser.SID
            $script:Results = @($result)
            if ($result.Status -eq "Skipped") { $script:AllSkipped = $true }
        } catch {
            Write-Log "错误: $_" "ERROR"
            if ($AutoRollback) { Invoke-Rollback }
            exit 1
        }
    } else {
        Write-Log ""; Write-Log "请选择要修改的用户：" "STEP"
        for ($i=0; $i -lt $usersToProcess.Count; $i++) {
            $suggested = Get-UniqueEnglishName $usersToProcess[$i].Name
            Write-Host "[$i] $($usersToProcess[$i].Name) (建议: $suggested)" -ForegroundColor Yellow
        }
        Write-Host "[B] 进入批量模式，处理所有用户" -ForegroundColor Cyan
        $selection = Read-Host "请输入编号或B"
        if ($selection -eq "B") {
            $script:Results = Start-BatchProcess -UserList $usersToProcess
        } elseif ($selection -match "^\d+$" -and [int]$selection -lt $usersToProcess.Count) {
            $selectedUser = $usersToProcess[[int]$selection]
            $suggestedName = Get-UniqueEnglishName $selectedUser.Name
            if (-not $NewUserName) {
                $NewUserName = Read-Host "请输入新用户名(直接回车使用 $suggestedName)"
                if (-not $NewUserName) { $NewUserName = $suggestedName }
            }
            try {
                $result = Rename-SingleUser -OldName $selectedUser.Name -NewName $NewUserName -SID $selectedUser.SID
                $script:Results = @($result)
                if ($result.Status -eq "Skipped") { $script:AllSkipped = $true }
            } catch {
                Write-Log "错误: $_" "ERROR"
                if ($AutoRollback) { Invoke-Rollback }
                exit 1
            }
        } else {
            Write-Log "无效选择" "ERROR"; exit 1
        }
    }
}

if ($ReportPath -and $script:Results) {
    $report = @{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ScriptVersion = $ScriptVersion
        DryRun = $DryRun
        Results = $script:Results
    } | ConvertTo-Json -Depth 3
    try {
        $report | Out-File -FilePath $ReportPath -Encoding UTF8
        Write-Log "报告已保存至: $ReportPath" "SUCCESS"
    } catch {
        Write-Log "保存报告失败: $_" "WARN"
    }
}

if (-not $DryRun) {
    Start-Process explorer -ErrorAction SilentlyContinue
}

Write-Log ""; Write-Log "日志文件: $LogFile" "HINT"
Write-Log ""; Write-Log "📌 Windows Hello 提示：重命名后 PIN/指纹/人脸登录可能失效。" "HINT"
Write-Log "   解决方法：以管理员身份运行以下命令，然后重新设置 PIN：" "HINT"
Write-Log "   del /F /Q `"%windir%\ServiceProfiles\LocalService\AppData\Local\Microsoft\Ngc`" 2>nul" "HINT"
Write-Log ""; Write-Log "📌 OneDrive 提示：脚本已断开 OneDrive 链接，请手动重新链接并选择新文件夹路径。" "HINT"
Write-Log "   右键任务栏 OneDrive 图标 → 设置 → 账户 → 添加账户 → 登录并选择新位置。" "HINT"
Write-Log "   同时建议登录网页版重命名旧文件夹，防止云端恢复。" "HINT"
Write-Log ""; Write-Log "📌 权限提示：如果某些文件出现无法访问的情况，请以管理员身份运行以下命令修复继承：" "HINT"
Write-Log "   takeown /F `"$env:USERPROFILE`" /R /D Y  && icacls `"$env:USERPROFILE`" /inheritance:e" "HINT"

# 释放 Mutex（正常结束）
$script:Mutex.ReleaseMutex()

if ($script:AllSkipped) {
    exit 1
} else {
    Wait-ForExit
}
