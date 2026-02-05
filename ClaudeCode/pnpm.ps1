param(
    [Parameter()]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = 'SilentlyContinue'

#region 配置常量
$script:NodeJsDirectory   = "$env:USERPROFILE\nodejs"
$script:PnpmDirectory     = "$env:USERPROFILE\pnpm"
$script:LogDirectory      = "$env:USERPROFILE\.claude\logs"
$script:NodeJsVersion     = "20.11.0"  # LTS 版本
$script:RequestTimeout    = 30
#endregion

#region 消息资源
$script:Messages = @{
    "HelpText" = @"
Claude Code 安装脚本 (pnpm 方式)

用法: .\pnpm.ps1 [选项]

参数:
  -Help             显示此帮助信息

说明:
  此脚本使用 pnpm 安装 Claude Code。
  如果系统没有 Node.js，会自动下载并安装。
  如果系统没有 pnpm，会自动安装。

示例:
  .\pnpm.ps1                          # 使用 pnpm 安装 Claude Code
"@
    "CheckingNode" = "正在检查 Node.js 环境..."
    "InstallingNode" = "正在安装 Node.js..."
    "NodeInstalled" = "Node.js 安装完成: {0}"
    "CheckingPnpm" = "正在检查 pnpm..."
    "InstallingPnpm" = "正在安装 pnpm..."
    "PnpmInstalled" = "pnpm 安装完成"
    "InstallingClaude" = "正在使用 pnpm 安装 Claude Code..."
    "InstallComplete" = "安装完成！"
    "ErrorInstall" = "安装失败"
    "LogSaved" = "详细日志已保存到: {0}"
}

# 日志内容
$script:LogContent = @()

#region 日志函数

<#
.SYNOPSIS
    获取本地化消息

.DESCRIPTION
    从消息资源字典中获取指定键的消息，支持格式化参数
#>
function Get-Message {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Key,

        [string[]]$FormatArgs
    )

    $message = $script:Messages[$Key]
    if ($FormatArgs) {
        $message = $message -f $FormatArgs
    }
    return $message
}

<#
.SYNOPSIS
    写入日志并输出到控制台

.DESCRIPTION
    将日志写入内存缓冲区，同时根据级别输出到控制台（带颜色）
#>
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry  = "[$timestamp] [$Level] $Message"
    $script:LogContent += $logEntry

    switch ($Level) {
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "WARN"    { Write-Host $Message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        default   { Write-Host $Message }
    }
}

<#
.SYNOPSIS
    保存日志到文件

.DESCRIPTION
    将内存中的日志内容保存到日志目录

.OUTPUTS
    [string] 日志文件的完整路径
#>
function Save-Log {
    [CmdletBinding()]
    param()

    New-Item -ItemType Directory -Force -Path $script:LogDirectory | Out-Null

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $logFile   = Join-Path $script:LogDirectory "pnpm-install-$timestamp.log"

    $script:LogContent | Out-File -FilePath $logFile -Encoding UTF8

    return $logFile
}

#endregion

#region Node.js 函数

<#
.SYNOPSIS
    获取已安装的 Node.js 版本

.DESCRIPTION
    检查系统是否已安装 Node.js 并返回其版本号

.OUTPUTS
    [string] 版本号（如 "20.11.0"），如果未安装则返回 $null
#>
function Get-NodeVersion {
    [CmdletBinding()]
    param()

    try {
        $nodeCommand = Get-Command "node" -ErrorAction SilentlyContinue

        if ($nodeCommand) {
            $versionOutput = & node --version 2>$null
            return $versionOutput.TrimStart('v')
        }

        return $null
    }
    catch {
        return $null
    }
}

<#
.SYNOPSIS
    安装 Node.js

.DESCRIPTION
    下载并安装指定版本的 Node.js，支持下载进度显示

.OUTPUTS
    [bool] 是否安装成功
#>
function Install-NodeJS {
    [CmdletBinding()]
    param()

    Write-Log "INFO" (Get-Message "InstallingNode")

    # 创建目录
    New-Item -ItemType Directory -Force -Path $script:NodeJsDirectory | Out-Null

    # 构建下载 URL 和路径
    $nodeVersion     = $script:NodeJsVersion
    $nodeDownloadUrl = "https://nodejs.org/dist/v$nodeVersion/node-v$nodeVersion-win-x64.zip"
    $zipFilePath     = "$env:TEMP\nodejs-$nodeVersion.zip"

    try {
        # 下载 Node.js（带进度条）
        Write-Log "INFO" "正在下载 Node.js v$nodeVersion..."
        Invoke-DownloadWithProgress -Url $nodeDownloadUrl -OutFile $zipFilePath -Description "正在下载 Node.js"

        # 解压
        Write-Log "INFO" "正在解压..."
        Expand-Archive -Path $zipFilePath -DestinationPath $env:TEMP -Force

        # 复制文件到目标目录
        $extractedDirectory = "$env:TEMP\node-v$nodeVersion-win-x64"
        Copy-Item -Path "$extractedDirectory\*" -Destination $script:NodeJsDirectory -Recurse -Force

        # 清理临时文件
        Remove-Item -Path $zipFilePath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $extractedDirectory -Recurse -Force -ErrorAction SilentlyContinue

        # 添加到用户 PATH
        $currentUserPath = [Environment]::GetEnvironmentVariable("Path", "User")
        $isPathConfigured = $currentUserPath -like "*$script:NodeJsDirectory*"

        if (-not $isPathConfigured) {
            [Environment]::SetEnvironmentVariable("Path", "$currentUserPath;$script:NodeJsDirectory", "User")
        }

        # 更新当前会话的 PATH
        $env:Path = "$env:Path;$script:NodeJsDirectory"

        Write-Log "SUCCESS" (Get-Message "NodeInstalled" $nodeVersion)
        return $true
    }
    catch {
        Write-Log "ERROR" "Node.js 安装失败: $_"
        return $false
    }
}

<#
.SYNOPSIS
    执行带进度条的下载

.DESCRIPTION
    使用 WebClient 实现异步下载并显示进度条
#>
function Invoke-DownloadWithProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$OutFile,

        [string]$Description = "正在下载"
    )

    # 确保输出目录存在
    $outputDirectory = Split-Path -Parent $OutFile
    if (-not (Test-Path $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    # 创建 WebClient
    $webClient = New-Object System.Net.WebClient

    $script:lastReportedPercent = -1

    # 注册进度事件
    $progressHandler = {
        $percent         = $EventArgs.ProgressPercentage
        $downloadedBytes = $EventArgs.BytesReceived
        $totalBytes      = $EventArgs.TotalBytesToReceive

        # 每 5% 更新一次，减少闪烁
        if ($percent -ne $script:lastReportedPercent -and $percent % 5 -eq 0) {
            $script:lastReportedPercent = $percent
            $downloadedMB = [math]::Round($downloadedBytes / 1MB, 2)
            $totalMB      = [math]::Round($totalBytes / 1MB, 2)

            Write-Progress -Activity $Event.MessageData `
                -Status "$percent% 完成" `
                -PercentComplete $percent `
                -CurrentOperation "$downloadedMB MB / $totalMB MB"
        }
    }

    $completedHandler = {
        Write-Progress -Activity $Event.MessageData -Completed
    }

    $progressEvent = Register-ObjectEvent -InputObject $webClient `
        -EventName DownloadProgressChanged `
        -Action $progressHandler `
        -MessageData $Description

    $completedEvent = Register-ObjectEvent -InputObject $webClient `
        -EventName DownloadFileCompleted `
        -Action $completedHandler `
        -MessageData $Description

    try {
        # 开始异步下载
        $webClient.DownloadFileAsync($Url, $OutFile)

        # 等待下载完成
        while ($webClient.IsBusy) {
            Start-Sleep -Milliseconds 100
        }
    }
    finally {
        # 清理事件和对象
        Unregister-Event -SourceIdentifier $progressEvent.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $completedEvent.Name -ErrorAction SilentlyContinue
        $webClient.Dispose()
        Write-Progress -Activity $Description -Completed
    }
}

#endregion

#region pnpm 和 Claude 安装函数

<#
.SYNOPSIS
    安装 pnpm

.DESCRIPTION
    使用 npm 全局安装 pnpm

.OUTPUTS
    [bool] 是否安装成功
#>
function Install-Pnpm {
    [CmdletBinding()]
    param()

    Write-Log "INFO" (Get-Message "InstallingPnpm")

    try {
        # 使用 npm 安装 pnpm
        & npm install -g pnpm --force

        $isSuccessful = ($LASTEXITCODE -eq 0)

        if (-not $isSuccessful) {
            throw "npm 安装 pnpm 失败，退出代码: $LASTEXITCODE"
        }

        Write-Log "SUCCESS" (Get-Message "PnpmInstalled")
        return $true
    }
    catch {
        Write-Log "ERROR" "pnpm 安装失败: $_"
        return $false
    }
}

<#
.SYNOPSIS
    使用 pnpm 安装 Claude Code

.DESCRIPTION
    配置淘宝镜像并使用 pnpm 全局安装 @anthropic-ai/claude-code

.OUTPUTS
    [bool] 是否安装成功
#>
function Install-ClaudeWithPnpm {
    [CmdletBinding()]
    param()

    Write-Log "INFO" (Get-Message "InstallingClaude")

    try {
        # 配置使用淘宝镜像（中国用户优化）
        Write-Log "INFO" "配置 npm 使用淘宝镜像..."
        & npm config set registry https://registry.npmmirror.com/

        # 使用 pnpm 全局安装 @anthropic-ai/claude-code
        Write-Log "INFO" "正在安装 @anthropic-ai/claude-code..."
        & pnpm add -g @anthropic-ai/claude-code --registry https://registry.npmmirror.com/

        $isSuccessful = ($LASTEXITCODE -eq 0)

        if (-not $isSuccessful) {
            throw "pnpm 安装 Claude Code 失败，退出代码: $LASTEXITCODE"
        }

        Write-Log "SUCCESS" (Get-Message "InstallComplete")
        return $true
    }
    catch {
        Write-Log "ERROR" (Get-Message "ErrorInstall")
        Write-Log "ERROR" $_
        return $false
    }
}

#endregion

#region 主安装流程

<#
.SYNOPSIS
    主安装流程

.DESCRIPTION
    执行完整的 Claude Code pnpm 安装流程：检查/安装 Node.js、检查/安装 pnpm、安装 Claude Code
#
function Install-ClaudeCode {
    [CmdletBinding()]
    param()

    #region 检查 Node.js
    Write-Log "INFO" (Get-Message "CheckingNode")
    $installedNodeVersion = Get-NodeVersion
    $isNodeJsInstalled    = [bool]$installedNodeVersion

    if (-not $isNodeJsInstalled) {
        Write-Log "WARN" "未检测到 Node.js，开始自动安装..."

        $nodeJsInstalled = Install-NodeJS

        if (-not $nodeJsInstalled) {
            $logFile = Save-Log
            Write-Log "INFO" (Get-Message "LogSaved" $logFile)
            exit 1
        }
    }
    else {
        Write-Log "INFO" "检测到 Node.js 版本: $installedNodeVersion"
    }
    #endregion

    #region 检查 pnpm
    Write-Log "INFO" (Get-Message "CheckingPnpm")
    $pnpmCommand = Get-Command "pnpm" -ErrorAction SilentlyContinue
    $isPnpmInstalled = [bool]$pnpmCommand

    if (-not $isPnpmInstalled) {
        Write-Log "WARN" "未检测到 pnpm，开始自动安装..."

        $pnpmInstallationSuccessful = Install-Pnpm

        if (-not $pnpmInstallationSuccessful) {
            $logFile = Save-Log
            Write-Log "INFO" (Get-Message "LogSaved" $logFile)
            exit 1
        }
    }
    else {
        Write-Log "INFO" "检测到 pnpm: $($pnpmCommand.Source)"
    }
    #endregion

    #region 安装 Claude Code
    $claudeInstallationSuccessful = Install-ClaudeWithPnpm

    if (-not $claudeInstallationSuccessful) {
        $logFile = Save-Log
        Write-Log "INFO" (Get-Message "LogSaved" $logFile)
        exit 1
    }
    #endregion

    #region 完成
    Write-Log "SUCCESS" ""
    Write-Log "SUCCESS" "Claude Code (pnpm 版) 安装完成！"
    Write-Log "SUCCESS" ""
    Write-Log "INFO" "运行 'claude' 开始使用。"
    #endregion
}

#endregion

#region 主程序入口

if ($Help) {
    Write-Host (Get-Message "HelpText")
    exit 0
}

Install-ClaudeCode

#endregion
