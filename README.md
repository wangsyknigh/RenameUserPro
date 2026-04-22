# RenameUserPro
安全、可靠地将 Windows 中文用户文件夹重命名为英文名，彻底解决专业软件路径兼容性问题。 无需重装系统，无需新建账户，修改后原有账户名和密码保持不变。

📖 目录
为什么需要这个脚本？

设计思路

功能特性

使用方法

前置要求

参数说明

使用示例

执行过程详解

错误处理与回滚机制

⚠️ 重要注意事项

📋 操作前备份建议

❓ 常见问题 FAQ

贡献与反馈

🤔 为什么需要这个脚本？
当 Windows 使用微软账户登录，且账户显示名为中文时，系统会自动在 C:\Users\ 下创建一个中文名称的用户文件夹（例如 C:\Users\王小明）。

许多专业开发工具和游戏引擎（如 Unreal Engine、Android Studio、Python、Node.js、Docker、MATLAB 等）对包含非 ASCII 字符的路径支持不佳，常导致：

编译失败，报错 UnicodeDecodeError 或 invalid path

调试器无法加载符号

安装程序闪退或静默失败

部分软件直接拒绝启动

传统解决方案的痛点
新建英文账户迁移数据 → 耗时、部分软件需重装、注册表残留旧路径。

直接修改注册表 → 高风险，操作失误会导致无法登录。

本脚本的优势
原地修改：无需新建账户，原有用户名、密码、SID 保持不变。

自动修复：智能更新注册表、环境变量、快捷方式、Shell 文件夹等所有关联位置。

符号链接兼容：创建旧中文路径的目录链接，让硬编码旧路径的旧软件仍能正常工作。

内置安全网：失败自动回滚，支持模拟运行，每一步都可追踪。

🎯 设计思路
检测目标用户
扫描系统本地用户，筛选出中文名称的用户文件夹，并排除内置账户（Administrator、Guest 等）。

预检与风险评估

检查用户是否已登录（强制要求完全注销）。

检测微软账户关联（强制要求转换为本地账户，防止云端同步覆盖修改）。

检查 BitLocker、EFS 加密、漫游配置文件等特殊环境。

核心操作

先改注册表，后重命名文件夹（降低系统崩溃风险）。

修改 HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\<SID> 中的 ProfileImagePath 值为新英文路径。

终止占用进程，强制重启 Explorer，重命名文件夹。

创建目录符号链接：mklink /J C:\Users\中文名 C:\Users\英文名，确保旧路径依然有效。

全面修复关联引用

Shell Folders（桌面、文档、下载等）

用户与系统 PATH 环境变量

开始菜单与桌面快捷方式

Windows Store 应用容器映射

Windows Terminal 配置文件

交接点（Junction）与挂载点

安全兜底

每一步操作均记录回滚动作，出错或用户按 Ctrl+C 时自动恢复。

支持 -DryRun 模拟运行，预览所有修改而不实际执行。

可选创建系统还原点。

✨ 功能特性
✅ 支持 Windows 10 (Build 10240+) 和 Windows 11 (Build 22000+)

✅ 单用户交互模式 & 批量处理模式

✅ 自动生成拼音英文名，支持自定义命名

✅ 智能处理 OneDrive（断开链接并引导手动重链）

✅ 修复 10+ 类常见路径依赖项

✅ 内置进程白名单，避免误杀 Docker、VMware、IDE 等

✅ 详细的彩色日志输出，同时保存至 %TEMP% 目录

✅ 退出时自动释放全局互斥锁，支持并发安全调用

🚀 使用方法
前置要求
项目	要求
操作系统	Windows 10 / 11
PowerShell 版本	5.1 或更高
执行权限	以管理员身份运行
当前登录账户	建议使用内置 Administrator 账户（需先手动启用）或其他管理员账户
目标用户状态	必须完全注销，不能有任何活动进程或后台会话
💡 如何启用 Administrator 账户？
以管理员身份运行命令提示符，执行：
net user Administrator /active:yes
然后注销当前账户，登录 Administrator。

参数说明
参数	类型	说明
-NewUserName	string	指定新的英文用户名（仅单用户模式有效）
-BatchMode	switch	批量处理所有检测到的中文用户
-AutoRollback	switch	发生错误时自动执行回滚（默认已启用）
-SkipConfirm	switch	跳过所有交互确认（谨慎使用）
-NoWait	switch	执行完毕后不等待按键，直接退出
-ForceLocalAccount	switch	检测到微软账户时，引导用户转换为本地账户
-StopOnError	switch	批量模式下遇到错误立即停止
-DryRun	switch	模拟运行，不实际修改任何文件/注册表
-RebuildSearchIndex	switch	完成后重建 Windows Search 索引
-SkipShortcutRepair	switch	跳过快捷方式修复（可加快速度，但部分旧快捷方式可能失效）
-ReportPath	string	将执行结果保存为 JSON 报告到指定路径
使用示例
1️⃣ 普通单用户模式（交互式）
powershell
.\RenameUserPro.ps1
脚本会列出所有中文用户，输入编号选择目标用户。

自动生成建议英文名，回车确认或手动输入新名。

确认后开始执行。

2️⃣ 指定新用户名（非交互）
powershell
.\RenameUserPro.ps1 -NewUserName "ZhangSan"
如果系统只有一个中文用户，直接使用指定名称处理；多个用户时仍需选择。

3️⃣ 批量处理所有中文用户
powershell
.\RenameUserPro.ps1 -BatchMode
自动为每个中文用户生成英文名，需输入 CONFIRM ALL 确认后批量执行。

4️⃣ 模拟运行（安全预览）
powershell
.\RenameUserPro.ps1 -DryRun
查看所有将要修改的项目，不实际更改系统。

5️⃣ 完整参数示例
powershell
.\RenameUserPro.ps1 -BatchMode -ForceLocalAccount -ReportPath "C:\report.json" -NoWait
🔄 执行过程详解
脚本执行的主要阶段如下：

阶段	操作内容
1. 初始化	检查管理员权限、系统版本、PowerShell 版本，启用长路径支持
2. 用户扫描	获取本地用户列表，筛选出中文名称且未修改过的用户
3. 预检	登录状态检测、BitLocker/EFS 检查、微软账户检测与转换引导
4. OneDrive 处理	提示同步状态，断开链接（执行 OneDrive.exe /unlink）
5. 核心重命名	创建还原点 → 更新注册表 ProfileImagePath → 终止占用进程 → 重命名文件夹 → 加载用户注册表蜂巢
6. 路径修复	修复 Shell Folders、环境变量、快捷方式、Store 应用、交接点等
7. 符号链接	执行 mklink /J "旧中文路径" "新英文路径" 创建目录联接
8. 收尾	恢复 OneDrive 提示、重启 Explorer、可选重建索引、输出日志
每一步都有对应的回滚动作，确保失败可恢复。

🔧 错误处理与回滚机制
自动回滚触发条件
脚本任何步骤抛出终止性错误（throw）。

用户按下 Ctrl+C 中断执行。

注册表蜂巢加载失败。

文件夹重命名失败（权限不足、文件被占用等）。

回滚动作顺序
所有回滚动作按优先级执行（数值越小越先恢复）：

恢复注册表 ProfileImagePath（最关键）

恢复文件夹名称

恢复注册表各子项（Shell Folders、环境变量等）

恢复服务状态（WSearch、Explorer）

卸载注册表蜂巢

回滚完成后，系统状态与执行前完全一致。如有备份目录（.old.时间戳），请手动检查并删除。

手动紧急恢复
如果脚本意外中断且自动回滚未完全执行，可参考以下步骤手动恢复：

打开 regedit，导航至 HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList。

找到目标用户的 SID 项，将 ProfileImagePath 改回原中文路径。

重启计算机，使用另一个管理员账户登录，将 C:\Users 下的英文文件夹重命名为中文原名。

删除可能存在的符号链接 C:\Users\中文名（如果已是链接，直接删除即可）。

⚠️ 重要注意事项
类别	说明
账户要求	目标用户必须完全注销！建议使用内置 Administrator 账户执行脚本。
微软账户	必须先将微软账户转换为本地账户，否则云端会恢复中文文件夹名。脚本会引导操作。
OneDrive	脚本会主动断开 OneDrive 链接。完成后务必登录网页版 OneDrive，将云端旧文件夹重命名，再重新链接本地 OneDrive 并指向新路径。
Windows Hello (PIN/指纹/人脸)	修改后可能失效。以管理员身份运行：del /F /Q "%windir%\ServiceProfiles\LocalService\AppData\Local\Microsoft\Ngc"，然后重新设置 PIN。
WSL 用户	若 WSL 配置文件（.wslconfig）引用了旧路径，需手动修改。脚本会给出警告。
杀毒软件	极少数情况下安全软件可能锁定文件导致重命名失败。若遇到，建议临时禁用实时防护后重试。
权限继承	重命名后部分文件可能权限异常。可执行：takeown /F "C:\Users\新英文名" /R /D Y 和 icacls "C:\Users\新英文名" /inheritance:e 修复。
系统还原点	脚本会自动尝试创建还原点（需卷影复制服务正常）。若失败，不影响主流程。
📋 操作前备份建议
🔔 强烈建议在执行脚本前手动备份关键数据！

1. 注册表备份
powershell
# 备份整个 ProfileList 分支
reg export "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" "C:\backup_profilelist.reg"
2. 用户文件夹备份（可选）
powershell
# 使用 Robocopy 复制用户目录到其他盘（需管理员权限）
robocopy "C:\Users\中文名" "D:\Backup\Users\中文名" /E /COPYALL /R:0 /W:0
3. 创建系统还原点（手动）
powershell
Checkpoint-Computer -Description "Before_RenameUserPro" -RestorePointType MODIFY_SETTINGS
4. 记录当前用户 SID
powershell
(New-Object System.Security.Principal.NTAccount("你的用户名")).Translate([System.Security.Principal.SecurityIdentifier]).Value
❓ 常见问题 FAQ
Q1：脚本执行后，桌面/开始菜单快捷方式失效怎么办？
A：脚本已自动修复绝大多数快捷方式。若仍有个别失效，手动删除并重新固定即可。

Q2：修改后某些 UWP 应用（如照片、计算器）打不开？
A：极少数应用需要重置。前往 设置 → 应用 → 找到该应用 → 高级选项 → 重置。

Q3：为什么登录后，C:\Users 下又出现了中文文件夹？
A：通常是微软账户云端同步导致。请确保已转换为本地账户，且 OneDrive 已按提示重新配置。

Q4：脚本提示“另一个实例正在运行”？
A：脚本使用全局互斥锁防止并发。若确认没有其他实例，可重启 PowerShell 窗口后重试。

Q5：可以只修改文件夹名而不创建符号链接吗？
A：不推荐。符号链接用于兼容硬编码旧路径的软件，删除后可能导致部分软件无法运行。若坚持删除，直接删除 C:\Users\中文名 这个链接即可（不影响实际数据）。

Q6：脚本支持 Windows 7 吗？
A：不支持。仅 Windows 10 Build 10240 及 Windows 11 经过完整测试。

Q7：执行过程中断电或强制关机了怎么办？
A：重启后使用另一管理员账户登录，检查 C:\Users 下文件夹名称和注册表 ProfileImagePath 是否一致。如不一致，参考上文“手动紧急恢复”步骤修正。

🤝 贡献与反馈
如果你遇到问题或有改进建议，欢迎通过 Issue 提出。
请提供以下信息以便快速定位：

Windows 版本（winver 命令查看）

PowerShell 版本（$PSVersionTable）

脚本日志文件（位于 %TEMP%\RenameUserPro_*.log）

最后提醒： 该脚本修改系统关键配置，请务必在理解原理的前提下使用。尽管内置多重安全机制，仍建议提前备份重要数据。祝使用顺利！
