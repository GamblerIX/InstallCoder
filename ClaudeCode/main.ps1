param(
    [Parameter(Position=0)]
    [ValidatePattern('^(stable|latest|\d+\.\d+\.\d+(-[^\s]+)?)$')]
    [string]$Target = "latest",

    [Parameter()]
    [string]$Proxy,

    [Parameter()]
    [switch]$Help,

    [Parameter()]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = 'SilentlyContinue'

# 保存脚本路径（用于在函数内部引用）
$script:ScriptPath = $MyInvocation.MyCommand.Definition
$script:ScriptDir = Split-Path -Parent $script:ScriptPath

# 配置常量
$GCS_BUCKET = "https://storage.googleapis.com/claude-code-dist-86c565f3-f756-42ad-8dfa-d59b1c096819/claude-code-releases"
$DOWNLOAD_DIR = "$env:USERPROFILE\.claude\downloads"
$LOG_DIR = "$env:USERPROFILE\.claude\logs"
$platform = "win32-x64"

# 语言字符串资源（仅中文）
$script:Messages = @{
    "HelpText" = @"
Claude Code 安装脚本

用法: .\main.ps1 [版本] [选项]

参数:
  [版本]           要安装的版本: latest, stable, 或具体版本号如 0.2.45 (默认: latest)
  -Proxy <代理地址>  指定代理服务器, 如 http://127.0.0.1:7890
  -Help             显示此帮助信息
  -Force            强制重新安装, 即使已安装最新版本

示例:
  .\main.ps1                          # 安装最新版本
  .\main.ps1 stable                   # 安装稳定版本
  .\main.ps1 0.2.45                   # 安装指定版本
  .\main.ps1 -Proxy "http://127.0.0.1:7890"  # 使用代理安装
  .\main.ps1 -Force                   # 强制重新安装

中国用户特别提示:
  由于 Google Cloud Storage 在国内访问受限, 建议使用代理:
  - Clash 默认端口: http://127.0.0.1:7890
  - v2rayN 默认端口: http://127.0.0.1:10809
"@
    "CheckingNetwork" = "正在检测网络环境..."
    "GCSNotAccessible" = "无法直接访问 Google Cloud Storage"
    "TryingProxy" = "正在尝试使用代理..."
    "ProxyDetected" = "检测到系统代理: {0}"
    "UsingExplicitProxy" = "使用指定的代理: {0}"
    "UsingSystemProxy" = "使用系统代理: {0}"
    "UsingEnvProxy" = "使用环境变量代理: {0}"
    "FetchingVersion" = "正在获取版本信息..."
    "Downloading" = "正在下载 Claude Code..."
    "Verifying" = "正在验证文件完整性..."
    "Installing" = "正在安装 Claude Code..."
    "CleaningUp" = "正在清理临时文件..."
    "InstallComplete" = "安装完成！"
    "AlreadyInstalled" = "Claude Code 已安装且为最新版本 ({0}), 跳过安装。"
    "AlreadyInstalledWithForce" = "Claude Code 已安装且为最新版本 ({0}), 但强制重新安装。"
    "Error32Bit" = "Claude Code 不支持 32 位 Windows, 请使用 64 位版本。"
    "ErrorNetwork" = "网络连接失败"
    "ErrorProxy" = "代理连接失败, 请检查代理设置"
    "ErrorDownload" = "下载失败"
    "ErrorVerify" = "文件校验失败"
    "ErrorInstall" = "安装失败"
    "SuggestionProxy" = "建议: 使用代理服务器, 如 -Proxy 'http://127.0.0.1:7890'"
    "SuggestionVPN" = "建议: 开启 VPN 或代理工具后重试"
    "SuggestionCheckProxy" = "建议: 检查代理地址和端口是否正确"
    "SuggestionForceInstall" = "建议: 如果证书验证失败, 可以尝试直接运行下载的文件并添加 --force 参数"
    "Retrying" = "正在重试 ({0}/{1})..."
    "AllAttemptsFailed" = "所有尝试均失败"
    "LogSaved" = "详细日志已保存到: {0}"
    "BinaryLocation" = "Claude Code 已下载到: {0}"
    "ManualInstall" = "您可以手动运行以下命令安装:"
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

    # 同时输出到控制台
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
    $logFile = Join-Path $LOG_DIR "install-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    $script:LogContent | Out-File -FilePath $logFile -Encoding UTF8
    return $logFile
}

# 获取系统代理
function Get-SystemProxy {
    try {
        $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $proxyUrl = $proxy.GetProxy($GCS_BUCKET)
        if ($proxyUrl -and $proxyUrl.AbsoluteUri -ne $GCS_BUCKET) {
            return $proxyUrl.AbsoluteUri
        }
    } catch {
        Write-Log "WARN" "获取系统代理失败: $_"
    }
    return $null
}

# 获取环境变量代理
function Get-EnvironmentProxy {
    $envProxy = $env:HTTP_PROXY -or $env:http_proxy -or $env:HTTPS_PROXY -or $env:https_proxy
    return $envProxy
}

# 确定要使用的代理
function Get-ProxyToUse {
    # 1. 优先使用参数指定的代理
    if ($Proxy) {
        Write-Log "INFO" (Get-Message "UsingExplicitProxy" $Proxy)
        return $Proxy
    }

    # 2. 检查环境变量代理
    $envProxy = Get-EnvironmentProxy
    if ($envProxy) {
        Write-Log "INFO" (Get-Message "UsingEnvProxy" $envProxy)
        return $envProxy
    }

    # 3. 检查系统代理
    $systemProxy = Get-SystemProxy
    if ($systemProxy) {
        Write-Log "INFO" (Get-Message "UsingSystemProxy" $systemProxy)
        return $systemProxy
    }

    return $null
}

# 创建 Web 请求选项
function New-WebRequestOptions {
    param([string]$ProxyUrl)

    $options = @{
        TimeoutSec = 30
        ErrorAction = "Stop"
    }

    if ($ProxyUrl) {
        $options['Proxy'] = $ProxyUrl
    }

    return $options
}

# 测试 GCS 连通性
function Test-GCSConnectivity {
    param([string]$ProxyUrl)

    try {
        $options = New-WebRequestOptions -ProxyUrl $ProxyUrl
        $null = Invoke-RestMethod -Uri "$GCS_BUCKET/latest" @options
        return $true
    } catch {
        return $false
    }
}

# 获取版本信息
function Get-VersionInfo {
    param([string]$ProxyUrl)

    $options = New-WebRequestOptions -ProxyUrl $ProxyUrl
    $version = Invoke-RestMethod -Uri "$GCS_BUCKET/latest" @options
    return $version
}

# 获取清单文件
function Get-Manifest {
    param(
        [string]$Version,
        [string]$ProxyUrl
    )

    $options = New-WebRequestOptions -ProxyUrl $ProxyUrl
    $manifest = Invoke-RestMethod -Uri "$GCS_BUCKET/$Version/manifest.json" @options
    return $manifest
}

# 下载文件
function Download-File {
    param(
        [string]$Url,
        [string]$OutFile,
        [string]$ProxyUrl
    )

    $options = New-WebRequestOptions -ProxyUrl $ProxyUrl
    Invoke-WebRequest -Uri $Url -OutFile $OutFile @options
}

# 获取已安装的 Claude Code 版本
function Get-InstalledClaudeVersion {
    try {
        # 尝试从常见位置查找 claude 命令
        $claudePath = $null
        
        # 检查 PATH 中的 claude
        $claudeInPath = Get-Command "claude" -ErrorAction SilentlyContinue
        if ($claudeInPath) {
            $claudePath = $claudeInPath.Source
        }
        
        # 检查默认安装位置
        $defaultPaths = @(
            "$env:LOCALAPPDATA\Programs\Claude\claude.exe",
            "$env:ProgramFiles\Claude\claude.exe",
            "$env:ProgramFiles(x86)\Claude\claude.exe",
            "$env:USERPROFILE\AppData\Local\Programs\Claude\claude.exe"
        )
        
        foreach ($path in $defaultPaths) {
            if (Test-Path $path) {
                $claudePath = $path
                break
            }
        }
        
        if (-not $claudePath) {
            return $null
        }
        
        # 尝试获取版本信息
        # 使用 --version 参数或其他方式
        $versionOutput = & $claudePath --version 2>$null
        if ($versionOutput -match '(\d+\.\d+\.\d+)') {
            return $matches[1]
        }
        
        # 如果 --version 失败, 尝试从文件版本获取
        $fileVersion = (Get-ItemProperty $claudePath).VersionInfo.FileVersion
        if ($fileVersion) {
            return $fileVersion
        }
        
        return $null
    } catch {
        return $null
    }
}

# 比较版本号
function Compare-Version {
    param(
        [string]$Version1,
        [string]$Version2
    )
    
    try {
        $v1 = [System.Version]$Version1
        $v2 = [System.Version]$Version2
        
        if ($v1 -gt $v2) {
            return 1
        } elseif ($v1 -lt $v2) {
            return -1
        } else {
            return 0
        }
    } catch {
        # 如果无法解析版本, 假设需要更新
        return -1
    }
}

# 主安装流程
function Install-ClaudeCode {
    # 检查 32 位系统
    if (-not [Environment]::Is64BitProcess) {
        Write-Log "ERROR" (Get-Message "Error32Bit")
        exit 1
    }

    # 创建下载目录
    New-Item -ItemType Directory -Force -Path $DOWNLOAD_DIR | Out-Null

    # 确定要使用的代理
    $proxyToUse = Get-ProxyToUse

    # 检测网络环境
    Write-Log "INFO" (Get-Message "CheckingNetwork")
    $canConnectDirectly = Test-GCSConnectivity

    if (-not $canConnectDirectly) {
        Write-Log "WARN" (Get-Message "GCSNotAccessible")

        if (-not $proxyToUse) {
            # 没有代理可用，显示帮助信息
            Write-Log "ERROR" (Get-Message "ErrorNetwork")
            Write-Log "INFO" (Get-Message "SuggestionProxy")
            Write-Log "INFO" (Get-Message "SuggestionVPN")
            Write-Log "INFO" ""
            Write-Log "INFO" "Clash:    .\main.ps1 -Proxy 'http://127.0.0.1:7890'"
            Write-Log "INFO" "v2rayN:   .\main.ps1 -Proxy 'http://127.0.0.1:10809'"

            $logFile = Save-Log
            Write-Log "INFO" (Get-Message "LogSaved" $logFile)
            exit 1
        }

        Write-Log "INFO" (Get-Message "TryingProxy")
    }

    # 获取版本信息
    Write-Log "INFO" (Get-Message "FetchingVersion")
    $version = $null
    $manifest = $null
    $checksum = $null

    $maxRetries = 3
    $retryCount = 0
    $success = $false

    while ($retryCount -lt $maxRetries -and -not $success) {
        try {
            if ($retryCount -gt 0) {
                Write-Log "INFO" (Get-Message "Retrying" ($retryCount + 1), $maxRetries)
                Start-Sleep -Seconds 2
            }

            $version = Get-VersionInfo -ProxyUrl $proxyToUse
            $manifest = Get-Manifest -Version $version -ProxyUrl $proxyToUse
            $checksum = $manifest.platforms.$platform.checksum

            if (-not $checksum) {
                throw "Platform $platform not found in manifest"
            }

            $success = $true
        } catch {
            $retryCount++
            Write-Log "WARN" "尝试 $retryCount 失败: $_"

            if ($retryCount -eq $maxRetries) {
                Write-Log "ERROR" (Get-Message "ErrorNetwork")
                Write-Log "ERROR" (Get-Message "AllAttemptsFailed")

                if ($proxyToUse) {
                    Write-Log "INFO" (Get-Message "SuggestionCheckProxy")
                } else {
                    Write-Log "INFO" (Get-Message "SuggestionProxy")
                }

                $logFile = Save-Log
                Write-Log "INFO" (Get-Message "LogSaved" $logFile)
                exit 1
            }
        }
    }

    # 检查已安装的版本
    $installedVersion = Get-InstalledClaudeVersion
    if ($installedVersion) {
        $comparison = Compare-Version -Version1 $installedVersion -Version2 $version
        
        if ($comparison -ge 0 -and -not $Force) {
            # 已安装版本相同或更新, 且没有强制安装
            Write-Log "SUCCESS" ""
            Write-Log "SUCCESS" (Get-Message "AlreadyInstalled" $installedVersion)
            Write-Log "SUCCESS" ""
            Write-Log "INFO" "使用 -Force 参数可以强制重新安装。"
            exit 0
        } elseif ($comparison -ge 0 -and $Force) {
            Write-Log "INFO" (Get-Message "AlreadyInstalledWithForce" $installedVersion)
        }
    }

    # 检查本地是否已有该版本的文件
    $binaryPath = "$DOWNLOAD_DIR\claude-$version-$platform.exe"
    $needsDownload = $true
    
    if (Test-Path $binaryPath) {
        # 检查文件校验和
        try {
            $existingChecksum = (Get-FileHash -Path $binaryPath -Algorithm SHA256).Hash.ToLower()
            if ($existingChecksum -eq $checksum) {
                Write-Log "INFO" "本地已存在版本 $version 的文件且校验通过, 跳过下载。"
                $needsDownload = $false
            } else {
                Write-Log "WARN" "本地文件校验和不匹配, 重新下载。"
                Remove-Item -Force $binaryPath
            }
        } catch {
            Write-Log "WARN" "无法验证本地文件, 重新下载。"
            Remove-Item -Force $binaryPath -ErrorAction SilentlyContinue
        }
    }

    # 下载二进制文件（如果需要）
    if ($needsDownload) {
        Write-Log "INFO" (Get-Message "Downloading")
        
        $retryCount = 0
        $success = $false

        while ($retryCount -lt $maxRetries -and -not $success) {
            try {
                if ($retryCount -gt 0) {
                    Write-Log "INFO" (Get-Message "Retrying" ($retryCount + 1), $maxRetries)
                    Start-Sleep -Seconds 2
                }

                Download-File -Url "$GCS_BUCKET/$version/$platform/claude.exe" -OutFile $binaryPath -ProxyUrl $proxyToUse
                $success = $true
            } catch {
                $retryCount++
                Write-Log "WARN" "下载尝试 $retryCount 失败: $_"

                if (Test-Path $binaryPath) {
                    Remove-Item -Force $binaryPath
                }

                if ($retryCount -eq $maxRetries) {
                    Write-Log "ERROR" (Get-Message "ErrorDownload")
                    Write-Log "ERROR" $_

                    $logFile = Save-Log
                    Write-Log "INFO" (Get-Message "LogSaved" $logFile)
                    exit 1
                }
            }
        }
    }

    # 验证校验和
    Write-Log "INFO" (Get-Message "Verifying")
    $actualChecksum = (Get-FileHash -Path $binaryPath -Algorithm SHA256).Hash.ToLower()

    if ($actualChecksum -ne $checksum) {
        Write-Log "ERROR" (Get-Message "ErrorVerify")
        Remove-Item -Force $binaryPath

        $logFile = Save-Log
        Write-Log "INFO" (Get-Message "LogSaved" $logFile)
        exit 1
    }

    # 执行安装 - 使用直接调用方式（与 bootstrap.ps1 一致）
    Write-Log "INFO" (Get-Message "Installing")
    $installSuccess = $false
    
    # 保存当前环境变量状态
    $originalErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    
    try {
        # 始终添加 --force 参数以绕过证书验证
        if ($Target -and $Target -ne "latest") {
            & $binaryPath install $Target --force
        } else {
            & $binaryPath install --force
        }
        
        # 检查最后一个命令的退出代码
        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq $null) {
            $installSuccess = $true
        } else {
            throw "Claude Code 安装程序返回退出代码: $LASTEXITCODE"
        }
    } catch {
        $errorMsg = $_
        Write-Log "ERROR" (Get-Message "ErrorInstall")
        Write-Log "ERROR" $errorMsg
        
        # 检查是否是证书验证错误
        if ($errorMsg -match "certificate verification error" -or 
            $errorMsg -match "unknown certificate" -or
            $LASTEXITCODE -eq 1) {
            
            Write-Log "INFO" ""
            Write-Log "INFO" "检测到证书验证错误，正在尝试使用 pnpm 方式安装..."
            Write-Log "INFO" ""
            
            # 获取当前脚本所在目录（使用保存的变量）
            $pnpmScript = Join-Path $script:ScriptDir "pnpm.ps1"
            
            # 检查 pnpm.ps1 是否存在
            if (Test-Path $pnpmScript) {
                Write-Log "INFO" "正在调用 pnpm 安装脚本..."
                
                # 清理当前下载的文件
                if (Test-Path $binaryPath) {
                    Remove-Item -Force $binaryPath -ErrorAction SilentlyContinue
                }
                
                # 执行 pnpm 安装脚本
                try {
                    & $pnpmScript
                    
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "SUCCESS" ""
                        Write-Log "SUCCESS" "通过 pnpm 方式安装成功！"
                        Write-Log "SUCCESS" ""
                        $installSuccess = $true
                    } else {
                        throw "pnpm 安装失败"
                    }
                } catch {
                    Write-Log "ERROR" "pnpm 安装也失败了: $_"
                    Write-Log "INFO" ""
                    Write-Log "INFO" "建议手动安装："
                    Write-Log "INFO" "1. 安装 Node.js: https://nodejs.org/"
                    Write-Log "INFO" "2. 安装 pnpm: npm install -g pnpm"
                    Write-Log "INFO" "3. 安装 Claude: pnpm add -g @anthropic-ai/claude"
                }
            } else {
                Write-Log "WARN" "未找到 pnpm.ps1 脚本，无法自动回退。"
                Write-Log "INFO" (Get-Message "BinaryLocation" $binaryPath)
            }
        }
        
        if (-not $installSuccess) {
            $logFile = Save-Log
            Write-Log "INFO" (Get-Message "LogSaved" $logFile)
            exit 1
        }
    } finally {
        # 恢复原始错误处理设置
        $ErrorActionPreference = $originalErrorAction
        
        # 清理
        if ($installSuccess) {
            Write-Log "INFO" (Get-Message "CleaningUp")
            try {
                Start-Sleep -Seconds 1
                if (Test-Path $binaryPath) {
                    Remove-Item -Force $binaryPath
                }
            } catch {
                Write-Log "WARN" "无法删除临时文件: $binaryPath"
            }
        }
    }

    # 完成
    if ($installSuccess) {
        Write-Log "SUCCESS" ""
        Write-Log "SUCCESS" (Get-Message "InstallComplete")
        Write-Log "SUCCESS" ""
    }
}

# 主程序
if ($Help) {
    Write-Host (Get-Message "HelpText")
    exit 0
}

Install-ClaudeCode
