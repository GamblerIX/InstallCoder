param(
    [Parameter()]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = 'SilentlyContinue'

# 配置
$NODEJS_DIR = "$env:USERPROFILE\nodejs"
$PNPM_DIR = "$env:USERPROFILE\pnpm"
$LOG_DIR = "$env:USERPROFILE\.claude\logs"

# 消息资源
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

# 获取消息
function Get-Message {
    param([string]$Key, [string[]]$FormatArgs)
    $message = $script:Messages[$Key]
    if ($FormatArgs) {
        $message = $message -f $FormatArgs
    }
    return $message
}

# 写入日志
function Write-Log {
    param(
        [string]$Level,
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    $script:LogContent += $logEntry

    switch ($Level) {
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "WARN" { Write-Host $Message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message }
    }
}

# 保存日志
function Save-Log {
    New-Item -ItemType Directory -Force -Path $LOG_DIR | Out-Null
    $logFile = Join-Path $LOG_DIR "pnpm-install-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    $script:LogContent | Out-File -FilePath $logFile -Encoding UTF8
    return $logFile
}

# 获取已安装的 Node.js 版本
function Get-NodeVersion {
    try {
        $nodePath = Get-Command "node" -ErrorAction SilentlyContinue
        if ($nodePath) {
            $version = & node --version 2>$null
            return $version.TrimStart('v')
        }
        return $null
    } catch {
        return $null
    }
}

# 安装 Node.js
function Install-NodeJS {
    Write-Log "INFO" (Get-Message "InstallingNode")
    
    # 创建目录
    New-Item -ItemType Directory -Force -Path $NODEJS_DIR | Out-Null
    
    # 下载 Node.js LTS
    $nodeVersion = "20.11.0"  # LTS 版本
    $nodeUrl = "https://nodejs.org/dist/v$nodeVersion/node-v$nodeVersion-win-x64.zip"
    $zipPath = "$env:TEMP\nodejs.zip"
    
    try {
        Invoke-WebRequest -Uri $nodeUrl -OutFile $zipPath -UseBasicParsing
        Expand-Archive -Path $zipPath -DestinationPath $env:TEMP -Force
        
        # 复制文件到目标目录
        $extractedDir = "$env:TEMP\node-v$nodeVersion-win-x64"
        Copy-Item -Path "$extractedDir\*" -Destination $NODEJS_DIR -Recurse -Force
        
        # 清理临时文件
        Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $extractedDir -Recurse -Force -ErrorAction SilentlyContinue
        
        # 添加到 PATH
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
        if ($currentPath -notlike "*$NODEJS_DIR*") {
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$NODEJS_DIR", "User")
        }
        
        # 更新当前会话的 PATH
        $env:Path = "$env:Path;$NODEJS_DIR"
        
        Write-Log "SUCCESS" (Get-Message "NodeInstalled" $nodeVersion)
        return $true
    } catch {
        Write-Log "ERROR" "Node.js 安装失败: $_"
        return $false
    }
}

# 安装 pnpm
function Install-Pnpm {
    Write-Log "INFO" (Get-Message "InstallingPnpm")
    
    try {
        # 使用 npm 安装 pnpm
        & npm install -g pnpm --force
        
        if ($LASTEXITCODE -ne 0) {
            throw "npm 安装 pnpm 失败"
        }
        
        Write-Log "SUCCESS" (Get-Message "PnpmInstalled")
        return $true
    } catch {
        Write-Log "ERROR" "pnpm 安装失败: $_"
        return $false
    }
}

# 使用 pnpm 安装 Claude Code
function Install-ClaudeWithPnpm {
    Write-Log "INFO" (Get-Message "InstallingClaude")
    
    try {
        # 配置使用淘宝镜像（中国用户优化）
        Write-Log "INFO" "配置 pnpm 使用淘宝镜像..."
        & npm config set registry https://registry.npmmirror.com/
        
        # 使用 pnpm 全局安装 @anthropic-ai/claude-code
        # 使用 --registry 参数确保使用淘宝镜像
        & pnpm add -g @anthropic-ai/claude-code --registry https://registry.npmmirror.com/
        
        if ($LASTEXITCODE -ne 0) {
            throw "pnpm 安装 Claude Code 失败"
        }
        
        Write-Log "SUCCESS" (Get-Message "InstallComplete")
        return $true
    } catch {
        Write-Log "ERROR" (Get-Message "ErrorInstall")
        Write-Log "ERROR" $_
        return $false
    }
}

# 主安装流程
function Install-ClaudeCode {
    # 检查 Node.js
    Write-Log "INFO" (Get-Message "CheckingNode")
    $nodeVersion = Get-NodeVersion
    
    if (-not $nodeVersion) {
        Write-Log "WARN" "未检测到 Node.js，开始自动安装..."
        $nodeInstalled = Install-NodeJS
        if (-not $nodeInstalled) {
            $logFile = Save-Log
            Write-Log "INFO" (Get-Message "LogSaved" $logFile)
            exit 1
        }
    } else {
        Write-Log "INFO" "检测到 Node.js 版本: $nodeVersion"
    }
    
    # 检查 pnpm
    Write-Log "INFO" (Get-Message "CheckingPnpm")
    $pnpmPath = Get-Command "pnpm" -ErrorAction SilentlyContinue
    
    if (-not $pnpmPath) {
        Write-Log "WARN" "未检测到 pnpm，开始自动安装..."
        $pnpmInstalled = Install-Pnpm
        if (-not $pnpmInstalled) {
            $logFile = Save-Log
            Write-Log "INFO" (Get-Message "LogSaved" $logFile)
            exit 1
        }
    } else {
        Write-Log "INFO" "检测到 pnpm: $($pnpmPath.Source)"
    }
    
    # 安装 Claude Code (pnpm 方式)
    $claudeInstalled = Install-ClaudeWithPnpm
    
    if (-not $claudeInstalled) {
        $logFile = Save-Log
        Write-Log "INFO" (Get-Message "LogSaved" $logFile)
        exit 1
    }
    
    Write-Log "SUCCESS" ""
    Write-Log "SUCCESS" "Claude Code (pnpm 版) 安装完成！"
    Write-Log "SUCCESS" ""
    Write-Log "INFO" "运行 'claude' 开始使用。"
}

# 主程序
if ($Help) {
    Write-Host (Get-Message "HelpText")
    exit 0
}

Install-ClaudeCode
