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

#region 配置常量
$script:GCS_BUCKET = "https://storage.googleapis.com/claude-code-dist-86c565f3-f756-42ad-8dfa-d59b1c096819/claude-code-releases"
$script:DOWNLOAD_DIR = "$env:USERPROFILE\.claude\downloads"
$script:LOG_DIR = "$env:USERPROFILE\.claude\logs"
$script:Platform = "win32-x64"
$script:MaxRetryAttempts = 3
$script:RequestTimeoutSeconds = 30
#endregion

#region 错误类型枚举
enum InstallErrorType {
    Network_DNS
    Network_Timeout
    Network_Certificate
    Network_Proxy
    Network_Other
    Download_ChecksumMismatch
    Download_Corrupted
    Install_Failed
    System_Not64Bit
    System_NoPermission
    Unknown
}
#endregion

# 错误信息资源
$script:ErrorMessages = @{
    [InstallErrorType]::Network_DNS = "DNS 解析失败，请检查网络连接或更换 DNS 服务器"
    [InstallErrorType]::Network_Timeout = "连接超时，可能是网络不稳定，建议重试或使用代理"
    [InstallErrorType]::Network_Certificate = "证书验证失败，将自动尝试 pnpm 安装方式"
    [InstallErrorType]::Network_Proxy = "代理服务器错误，请检查代理设置"
    [InstallErrorType]::Network_Other = "网络连接失败，请检查网络环境"
    [InstallErrorType]::Download_ChecksumMismatch = "文件校验失败，将重新下载"
    [InstallErrorType]::Download_Corrupted = "文件损坏，将重新下载"
    [InstallErrorType]::Install_Failed = "安装程序执行失败"
    [InstallErrorType]::System_Not64Bit = "Claude Code 不支持 32 位 Windows"
    [InstallErrorType]::System_NoPermission = "权限不足，请以管理员身份运行"
    [InstallErrorType]::Unknown = "未知错误，请查看日志获取详细信息"
}

#region 中文字符串资源
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

#region 错误处理函数

<#
.SYNOPSIS
    分析异常并分类错误类型

.DESCRIPTION
    根据异常消息内容识别错误类型，返回包含错误类型、建议操作和是否可重试的对象

.PARAMETER Exception
    捕获到的异常对象

.OUTPUTS
    [hashtable] 包含 Type, Message, Suggestion, IsRetryable 的错误信息
#>
function Resolve-InstallError {
    [CmdletBinding()]
    param([object]$Exception)

    # 支持 ErrorRecord 或 Exception 对象
    $errorSource = $Exception
    if ($Exception -is [System.Management.Automation.ErrorRecord]) {
        $errorSource = $Exception.Exception
    }

    # 如果 Exception 为 null，使用 ErrorRecord 的 ToString()
    if ($null -eq $errorSource) {
        $errorSource = [PSCustomObject]@{
            Message = $Exception.ToString()
        }
    }

    $errorInfo = @{
        Type        = [InstallErrorType]::Unknown
        Message     = $errorSource.Message
        Suggestion  = $script:ErrorMessages[[InstallErrorType]::Unknown]
        IsRetryable = $false
    }

    $errorMsg = $errorSource.Message.ToLower()

    switch -Regex ($errorMsg) {
        "could not resolve|dns|name resolution" {
            $errorInfo.Type        = [InstallErrorType]::Network_DNS
            $errorInfo.Suggestion  = $script:ErrorMessages[[InstallErrorType]::Network_DNS]
            $errorInfo.IsRetryable = $true
        }
        "timeout|timed out" {
            $errorInfo.Type        = [InstallErrorType]::Network_Timeout
            $errorInfo.Suggestion  = $script:ErrorMessages[[InstallErrorType]::Network_Timeout]
            $errorInfo.IsRetryable = $true
        }
        "certificate|ssl|tls|unknown certificate" {
            $errorInfo.Type        = [InstallErrorType]::Network_Certificate
            $errorInfo.Suggestion  = $script:ErrorMessages[[InstallErrorType]::Network_Certificate]
            $errorInfo.IsRetryable = $false
        }
        "407|proxy authentication|proxy.*failed" {
            $errorInfo.Type        = [InstallErrorType]::Network_Proxy
            $errorInfo.Suggestion  = $script:ErrorMessages[[InstallErrorType]::Network_Proxy]
            $errorInfo.IsRetryable = $true
        }
        "404|not found" {
            $errorInfo.Type        = [InstallErrorType]::Network_Other
            $errorInfo.Suggestion  = "资源不存在，可能是版本号错误或已被移除"
            $errorInfo.IsRetryable = $false
        }
        "checksum|hash|verify" {
            $errorInfo.Type        = [InstallErrorType]::Download_ChecksumMismatch
            $errorInfo.Suggestion  = $script:ErrorMessages[[InstallErrorType]::Download_ChecksumMismatch]
            $errorInfo.IsRetryable = $true
        }
        "32-bit|x86|not supported" {
            $errorInfo.Type        = [InstallErrorType]::System_Not64Bit
            $errorInfo.Suggestion  = $script:ErrorMessages[[InstallErrorType]::System_Not64Bit]
            $errorInfo.IsRetryable = $false
        }
        "access denied|permission|unauthorized" {
            $errorInfo.Type        = [InstallErrorType]::System_NoPermission
            $errorInfo.Suggestion  = $script:ErrorMessages[[InstallErrorType]::System_NoPermission]
            $errorInfo.IsRetryable = $false
        }
    }

    return $errorInfo
}

<#
.SYNOPSIS
    显示格式化的错误报告

.DESCRIPTION
    在控制台显示美观的错误报告，包含错误类型、详情和建议操作
#>
function Write-ErrorReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [InstallErrorType]$ErrorType,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [string]$Suggestion,

        [string]$LogFile
    )

    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║                    安装失败                              ║" -ForegroundColor Red
    Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "错误类型: " -NoNewline -ForegroundColor Yellow
    Write-Host $ErrorType -ForegroundColor White
    Write-Host ""
    Write-Host "错误详情:" -ForegroundColor Yellow
    Write-Host "  $Message" -ForegroundColor Gray
    Write-Host ""
    Write-Host "建议操作:" -ForegroundColor Yellow
    Write-Host "  → $Suggestion" -ForegroundColor Cyan

    if ($LogFile -and (Test-Path $LogFile)) {
        Write-Host ""
        Write-Host "日志文件: $LogFile" -ForegroundColor DarkGray
    }
    Write-Host ""
}

<#
.SYNOPSIS
    执行网络诊断测试

.DESCRIPTION
    测试 DNS 解析、HTTP 连通性和代理连接状态，帮助排查网络问题
#>
function Test-NetworkDiagnostics {
    [CmdletBinding()]
    param(
        [string]$TargetUrl = $script:GCS_BUCKET,
        [string]$ProxyUrl
    )

    Write-Log "INFO" "开始网络诊断..."
    Write-Log "INFO" "目标地址: $TargetUrl"

    $diagnostics = @()

    # 测试 DNS 解析
    try {
        $uri        = [Uri]$TargetUrl
        $hostEntry  = [System.Net.Dns]::GetHostEntry($uri.Host)
        $ipAddress  = $hostEntry.AddressList[0].IPAddressToString
        $diagnostics += "✓ DNS 解析成功: $($uri.Host) -> $ipAddress"
        Write-Log "SUCCESS" "DNS 解析: $($uri.Host) -> $ipAddress"
    }
    catch {
        $diagnostics += "✗ DNS 解析失败: $_"
        Write-Log "ERROR" "DNS 解析失败: $_"
    }

    # 测试直接 HTTP 连接
    try {
        $response        = Invoke-WebRequest -Uri $TargetUrl -Method Head -TimeoutSec 10 -UseBasicParsing
        $statusCode      = $response.StatusCode
        $diagnostics     += "✓ HTTP 连接成功: 状态码 $statusCode"
        Write-Log "SUCCESS" "HTTP 连接: 状态码 $statusCode"
    }
    catch {
        $diagnostics += "✗ HTTP 连接失败: $_"
        Write-Log "ERROR" "HTTP 连接失败: $_"
    }

    # 测试代理连接（如果提供了代理）
    if ($ProxyUrl) {
        try {
            $response        = Invoke-WebRequest -Uri $TargetUrl -Method Head -TimeoutSec 10 -Proxy $ProxyUrl -UseBasicParsing
            $statusCode      = $response.StatusCode
            $diagnostics     += "✓ 代理连接成功 ($ProxyUrl): 状态码 $statusCode"
            Write-Log "SUCCESS" "代理连接 ($ProxyUrl): 状态码 $statusCode"
        }
        catch {
            $diagnostics += "✗ 代理连接失败 ($ProxyUrl): $_"
            Write-Log "ERROR" "代理连接失败 ($ProxyUrl): $_"
        }
    }

    # 测试系统代理
    $systemProxy = Get-SystemProxy
    if ($systemProxy -and $systemProxy -ne $ProxyUrl) {
        try {
            $response        = Invoke-WebRequest -Uri $TargetUrl -Method Head -TimeoutSec 10 -Proxy $systemProxy -UseBasicParsing
            $statusCode      = $response.StatusCode
            $diagnostics     += "✓ 系统代理连接成功 ($systemProxy): 状态码 $statusCode"
            Write-Log "SUCCESS" "系统代理 ($systemProxy): 状态码 $statusCode"
        }
        catch {
            $diagnostics += "✗ 系统代理连接失败 ($systemProxy): $_"
            Write-Log "WARN" "系统代理连接失败 ($systemProxy)"
        }
    }

    Write-Log "INFO" "网络诊断完成"
    return $diagnostics
}

#endregion

#region 日志和工具函数

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

    # 同时输出到控制台
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

    New-Item -ItemType Directory -Force -Path $script:LOG_DIR | Out-Null

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $logFile   = Join-Path $script:LOG_DIR "install-$timestamp.log"

    $script:LogContent | Out-File -FilePath $logFile -Encoding UTF8

    return $logFile
}

<#
.SYNOPSIS
    安全删除文件

.DESCRIPTION
    使用 .NET 方法安全删除文件，避免 PowerShell 别名冲突
#>
function Remove-FileSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    try {
        if (Test-Path $Path) {
            [System.IO.File]::Delete($Path)
        }
    }
    catch {
        Write-Log "WARN" "无法删除文件: $Path - $_"
    }
}

#endregion

#region 环境管理和安装函数

<#
.SYNOPSIS
    刷新当前会话的环境变量

.DESCRIPTION
    从注册表重新加载 Machine 和 User 级别的环境变量到当前会话
#>
function Update-EnvironmentVariables {
    [CmdletBinding()]
    param()

    Write-Log "INFO" "正在刷新环境变量..."

    foreach ($level in "Machine", "User") {
        [Environment]::GetEnvironmentVariables($level).GetEnumerator() | ForEach-Object {
            $envVariableName  = $_.Key
            $envVariableValue = $_.Value
            Set-Item -Path "Env:$envVariableName" -Value $envVariableValue
        }
    }

    Write-Log "INFO" "环境变量已刷新"
}

<#
.SYNOPSIS
    执行原生 Claude Code 安装

.DESCRIPTION
    通过 pnpm 安装的 claude 命令执行原生安装

.OUTPUTS
    [bool] 是否安装成功
#>
function Install-NativeClaude {
    [CmdletBinding()]
    param()

    Write-Log "INFO" "正在安装原生 Claude Code..."

    try {
        # 刷新环境变量确保能找到 claude 命令
        Update-EnvironmentVariables

        # 检查 claude 命令是否可用
        $claudeCommand = Get-Command "claude" -ErrorAction SilentlyContinue

        # 如果 PATH 中找不到，尝试从 pnpm 全局目录查找
        if (-not $claudeCommand) {
            $pnpmGlobalDir = & pnpm root -g 2>$null
            if ($pnpmGlobalDir) {
                $claudeExePath = Join-Path $pnpmGlobalDir ".bin\claude.cmd"
                if (Test-Path $claudeExePath) {
                    $claudeCommand = @{ Source = $claudeExePath }
                }
            }
        }

        if ($claudeCommand) {
            Write-Log "INFO" "找到 Claude: $($claudeCommand.Source)"
            Write-Log "INFO" "尝试安装原生 Claude Code..."

            # 执行 claude install 安装原生版本
            & $claudeCommand.Source install --force

            $isSuccessful = ($LASTEXITCODE -eq 0)

            if ($isSuccessful) {
                Write-Log "SUCCESS" "原生 Claude Code 安装成功！"
            }
            else {
                Write-Log "WARN" "原生安装返回退出代码: $LASTEXITCODE"
            }

            return $isSuccessful
        }
        else {
            Write-Log "WARN" "无法找到 claude 命令，跳过原生安装"
            return $false
        }
    }
    catch {
        Write-Log "WARN" "原生安装失败: $_"
        return $false
    }
}

#region 代理和网络函数

<#
.SYNOPSIS
    获取系统代理设置

.DESCRIPTION
    从 Windows 系统设置中获取代理服务器地址

.OUTPUTS
    [string] 代理地址，如果没有设置则返回 $null
#>
function Get-SystemProxy {
    [CmdletBinding()]
    param()

    try {
        $systemWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $proxyUri       = $systemWebProxy.GetProxy($script:GCS_BUCKET)

        $hasProxy = $proxyUri -and $proxyUri.AbsoluteUri -ne $script:GCS_BUCKET
        if ($hasProxy) {
            return $proxyUri.AbsoluteUri
        }
    }
    catch {
        Write-Log "WARN" "获取系统代理失败: $_"
    }

    return $null
}

<#
.SYNOPSIS
    获取环境变量中的代理设置

.DESCRIPTION
    检查 HTTP_PROXY、HTTPS_PROXY 等环境变量

.OUTPUTS
    [string] 代理地址，如果没有设置则返回 $null
#>
function Get-EnvironmentProxy {
    [CmdletBinding()]
    param()

    $environmentProxy = $env:HTTP_PROXY -or
                        $env:http_proxy -or
                        $env:HTTPS_PROXY -or
                        $env:https_proxy

    return $environmentProxy
}

<#
.SYNOPSIS
    确定要使用的代理（优先级：参数 > 环境变量 > 系统设置）

.DESCRIPTION
    按优先级顺序检测并返回可用的代理服务器地址

.OUTPUTS
    [string] 代理地址，如果没有找到则返回 $null
#>
function Get-ProxyToUse {
    [CmdletBinding()]
    param()

    # 1. 优先使用参数指定的代理
    if ($Proxy) {
        Write-Log "INFO" (Get-Message "UsingExplicitProxy" $Proxy)
        return $Proxy
    }

    # 2. 检查环境变量代理
    $environmentProxy = Get-EnvironmentProxy
    if ($environmentProxy) {
        Write-Log "INFO" (Get-Message "UsingEnvProxy" $environmentProxy)
        return $environmentProxy
    }

    # 3. 检查系统代理
    $systemProxy = Get-SystemProxy
    if ($systemProxy) {
        Write-Log "INFO" (Get-Message "UsingSystemProxy" $systemProxy)
        return $systemProxy
    }

    return $null
}

<#
.SYNOPSIS
    创建 Web 请求的选项哈希表

.DESCRIPTION
    创建包含超时和代理设置的请求选项

.PARAMETER ProxyUrl
    可选的代理地址

.OUTPUTS
    [hashtable] 请求选项
#>
function New-WebRequestOptions {
    [CmdletBinding()]
    param([string]$ProxyUrl)

    $requestOptions = @{
        TimeoutSec  = $script:RequestTimeoutSeconds
        ErrorAction = "Stop"
    }

    if ($ProxyUrl) {
        $requestOptions['Proxy'] = $ProxyUrl
    }

    return $requestOptions
}

<#
.SYNOPSIS
    测试 Google Cloud Storage 连通性

.DESCRIPTION
    测试能否访问 GCS 获取版本信息

.PARAMETER ProxyUrl
    可选的代理地址

.OUTPUTS
    [bool] 是否可连通
#>
function Test-GCSConnectivity {
    [CmdletBinding()]
    param([string]$ProxyUrl)

    try {
        $requestOptions = New-WebRequestOptions -ProxyUrl $ProxyUrl
        $null = Invoke-RestMethod -Uri "$script:GCS_BUCKET/latest" @requestOptions
        return $true
    }
    catch {
        return $false
    }
}

#endregion

#region 版本信息函数

<#
.SYNOPSIS
    获取最新版本号

.DESCRIPTION
    从 GCS 获取 Claude Code 的最新版本号

.PARAMETER ProxyUrl
    可选的代理地址

.OUTPUTS
    [string] 版本号（如 "0.2.45"）
#>
function Get-VersionInfo {
    [CmdletBinding()]
    param([string]$ProxyUrl)

    $requestOptions = New-WebRequestOptions -ProxyUrl $ProxyUrl
    $latestVersion  = Invoke-RestMethod -Uri "$script:GCS_BUCKET/latest" @requestOptions

    return $latestVersion
}

<#
.SYNOPSIS
    获取版本清单文件

.DESCRIPTION
    从 GCS 获取指定版本的清单文件，包含校验和等信息

.PARAMETER Version
    版本号

.PARAMETER ProxyUrl
    可选的代理地址

.OUTPUTS
    [PSCustomObject] 清单文件对象
#>
function Get-Manifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Version,

        [string]$ProxyUrl
    )

    $requestOptions = New-WebRequestOptions -ProxyUrl $ProxyUrl
    $manifestUrl    = "$script:GCS_BUCKET/$Version/manifest.json"
    $manifest       = Invoke-RestMethod -Uri $manifestUrl @requestOptions

    return $manifest
}

#endregion

#region 下载函数

<#
.SYNOPSIS
    下载文件并显示进度条

.DESCRIPTION
    从指定 URL 下载文件，显示下载进度，支持代理和重试

.PARAMETER Url
    要下载的文件 URL

.PARAMETER OutFile
    保存路径

.PARAMETER ProxyUrl
    可选的代理地址

.PARAMETER Description
    进度条显示的描述文本

.PARAMETER MaxRetryAttempts
    最大重试次数（默认3次）
#>
function Download-File {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Url,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutFile,

        [string]$ProxyUrl,

        [string]$Description = "正在下载",

        [ValidateRange(1, 10)]
        [int]$MaxRetryAttempts = $script:MaxRetryAttempts
    )

    $currentRetryAttempt = 0
    $isDownloadSuccessful = $false

    while ($currentRetryAttempt -lt $MaxRetryAttempts -and -not $isDownloadSuccessful) {
        try {
            if ($currentRetryAttempt -gt 0) {
                # 指数退避策略
                $delaySeconds = [Math]::Pow(2, $currentRetryAttempt)
                Write-Log "INFO" "第 $currentRetryAttempt/$MaxRetryAttempts 次重试，等待 ${delaySeconds}秒..."
                Start-Sleep -Seconds $delaySeconds
            }

            Invoke-DownloadWithProgress -Url $Url `
                -OutFile $OutFile `
                -ProxyUrl $ProxyUrl `
                -Description $Description

            $isDownloadSuccessful = $true
            Write-Log "SUCCESS" "下载完成: $([IO.Path]::GetFileName($OutFile))"
        }
        catch {
            $currentRetryAttempt++
            $errorInfo = Resolve-InstallError -Exception $_.Exception

            Write-Log "WARN" "下载失败 ($currentRetryAttempt/$MaxRetryAttempts): $($errorInfo.Message)"

            if (-not $errorInfo.IsRetryable -or $currentRetryAttempt -eq $MaxRetryAttempts) {
                throw $_
            }

            # 清理失败的文件
            Remove-FileSafe -Path $OutFile
        }
    }

    return $isDownloadSuccessful
}

<#
.SYNOPSIS
    内部函数：执行带进度条的下载

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

        [string]$ProxyUrl,

        [string]$Description
    )

    # 确保输出目录存在
    $outputDirectory = Split-Path -Parent $OutFile
    if (-not (Test-Path $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    # 创建 WebClient
    $webClient = New-Object System.Net.WebClient

    # 配置代理
    if ($ProxyUrl) {
        $proxy = New-Object System.Net.WebProxy($ProxyUrl)
        $webClient.Proxy = $proxy
    }

    $script:lastReportedPercent = -1

    # 注册进度事件
    $progressHandler = {
        $percent = $EventArgs.ProgressPercentage
        $downloadedBytes = $EventArgs.BytesReceived
        $totalBytes = $EventArgs.TotalBytesToReceive

        # 每 5% 更新一次，减少闪烁
        if ($percent -ne $script:lastReportedPercent -and $percent % 5 -eq 0) {
            $script:lastReportedPercent = $percent
            $downloadedMB = [math]::Round($downloadedBytes / 1MB, 2)
            $totalMB = [math]::Round($totalBytes / 1MB, 2)

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

#region 版本检测和比较函数

<#
.SYNOPSIS
    获取已安装的 Claude Code 版本

.DESCRIPTION
    从 PATH 或默认安装位置查找已安装的 Claude Code 并获取其版本号

.OUTPUTS
    [string] 版本号，如果未安装则返回 $null
#>
function Get-InstalledClaudeVersion {
    [CmdletBinding()]
    param()

    try {
        # 可能的安装路径
        $possiblePaths = @()

        # 检查 PATH 中的 claude
        $claudeInPath = Get-Command "claude" -ErrorAction SilentlyContinue
        if ($claudeInPath) {
            $possiblePaths += $claudeInPath.Source
        }

        # 检查默认安装位置
        $defaultInstallPaths = @(
            "$env:LOCALAPPDATA\Programs\Claude\claude.exe"
            "$env:ProgramFiles\Claude\claude.exe"
            "$env:USERPROFILE\AppData\Local\Programs\Claude\claude.exe"
        )
        $possiblePaths += $defaultInstallPaths

        # 查找第一个存在的路径
        $claudeExecutablePath = $possiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

        if (-not $claudeExecutablePath) {
            return $null
        }

        # 尝试通过 --version 获取版本
        $versionOutput = & $claudeExecutablePath --version 2>$null
        $versionPattern = '(\d+\.\d+\.\d+)'

        if ($versionOutput -match $versionPattern) {
            return $matches[1]
        }

        # 回退到文件版本信息
        $fileVersionInfo = (Get-ItemProperty $claudeExecutablePath).VersionInfo.FileVersion
        if ($fileVersionInfo) {
            return $fileVersionInfo
        }

        return $null
    }
    catch {
        return $null
    }
}

<#
.SYNOPSIS
    比较两个版本号

.DESCRIPTION
    比较两个版本号，返回比较结果

.PARAMETER InstalledVersion
    已安装的版本

.PARAMETER TargetVersion
    目标版本

.OUTPUTS
    [int] 1 (Installed > Target), 0 (相等), -1 (Installed < Target)
#>
function Compare-Version {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$InstalledVersion,

        [Parameter(Mandatory)]
        [string]$TargetVersion
    )

    try {
        $parsedInstalledVersion = [System.Version]$InstalledVersion
        $parsedTargetVersion    = [System.Version]$TargetVersion

        if ($parsedInstalledVersion -gt $parsedTargetVersion) {
            return 1
        }
        elseif ($parsedInstalledVersion -lt $parsedTargetVersion) {
            return -1
        }
        else {
            return 0
        }
    }
    catch {
        # 如果无法解析版本，假设需要更新（返回 -1）
        Write-Log "WARN" "无法解析版本号，假设需要更新"
        return -1
    }
}

#endregion

<#
.SYNOPSIS
    主安装流程

.DESCRIPTION
    执行完整的 Claude Code 安装流程，包括网络检测、版本检查、下载、校验和安装
#>
function Install-ClaudeCode {
    [CmdletBinding()]
    param()

    #region 系统检查
    # 检查 64 位系统
    $is64BitProcess = [Environment]::Is64BitProcess
    if (-not $is64BitProcess) {
        $errorInfo = Resolve-InstallError -Exception ([System.Exception]::new("32-bit system not supported"))
        $logFile     = Save-Log
        Write-ErrorReport -ErrorType $errorInfo.Type `
            -Message $errorInfo.Message `
            -Suggestion $errorInfo.Suggestion `
            -LogFile $logFile
        exit 1
    }

    # 创建下载目录
    New-Item -ItemType Directory -Force -Path $script:DOWNLOAD_DIR | Out-Null
    #endregion

    #region 网络和代理配置
    $proxyToUse         = Get-ProxyToUse
    $networkDiagnostics = @()

    Write-Log "INFO" (Get-Message "CheckingNetwork")
    $canConnectDirectly = Test-GCSConnectivity

    if (-not $canConnectDirectly) {
        Write-Log "WARN" (Get-Message "GCSNotAccessible")

        if (-not $proxyToUse) {
            # 执行网络诊断
            $networkDiagnostics = Test-NetworkDiagnostics

            $logFile = Save-Log
            Write-ErrorReport -ErrorType ([InstallErrorType]::Network_Other) `
                -Message "无法连接到 Google Cloud Storage" `
                -Suggestion "请配置代理后重试。常用代理: Clash (7890), v2rayN (10809)" `
                -LogFile $logFile

            Write-Log "INFO" ""
            Write-Log "INFO" "使用示例:"
            Write-Log "INFO" "  Clash:  .\main.ps1 -Proxy 'http://127.0.0.1:7890'"
            Write-Log "INFO" "  v2rayN: .\main.ps1 -Proxy 'http://127.0.0.1:10809'"
            exit 1
        }

        Write-Log "INFO" (Get-Message "TryingProxy")
    }
    #endregion

    #region 获取版本信息
    Write-Log "INFO" (Get-Message "FetchingVersion")

    $targetVersion      = $null
    $versionManifest    = $null
    $expectedChecksum   = $null
    $currentAttempt     = 0
    $isVersionFetched   = $false

    while ($currentAttempt -lt $script:MaxRetryAttempts -and -not $isVersionFetched) {
        try {
            if ($currentAttempt -gt 0) {
                Write-Log "INFO" (Get-Message "Retrying" ($currentAttempt + 1), $script:MaxRetryAttempts)
                Start-Sleep -Seconds 2
            }

            $targetVersion    = Get-VersionInfo -ProxyUrl $proxyToUse
            $versionManifest  = Get-Manifest -Version $targetVersion -ProxyUrl $proxyToUse
            $expectedChecksum = $versionManifest.platforms.$script:Platform.checksum

            if (-not $expectedChecksum) {
                throw "Platform $script:Platform not found in manifest"
            }

            $isVersionFetched = $true
            Write-Log "SUCCESS" "获取版本信息成功: $targetVersion"
        }
        catch {
            $currentAttempt++
            $errorInfo = Resolve-InstallError -Exception $_.Exception
            Write-Log "WARN" "获取版本失败 ($currentAttempt/$($script:MaxRetryAttempts)): $($errorInfo.Message)"

            if ($currentAttempt -eq $script:MaxRetryAttempts) {
                $logFile = Save-Log
                Write-ErrorReport -ErrorType $errorInfo.Type `
                    -Message $errorInfo.Message `
                    -Suggestion $errorInfo.Suggestion `
                    -LogFile $logFile
                exit 1
            }
        }
    }
    #endregion

    #region 检查已安装版本
    $installedVersion     = Get-InstalledClaudeVersion
    $shouldSkipInstall    = $false

    if ($installedVersion) {
        $versionComparison    = Compare-Version -Version1 $installedVersion -Version2 $targetVersion
        $isUpToDate           = $versionComparison -ge 0
        $shouldSkipInstall    = $isUpToDate -and -not $Force

        if ($shouldSkipInstall) {
            Write-Log "SUCCESS" ""
            Write-Log "SUCCESS" (Get-Message "AlreadyInstalled" $installedVersion)
            Write-Log "SUCCESS" ""
            Write-Log "INFO" "使用 -Force 参数可以强制重新安装。"
            exit 0
        }
        elseif ($isUpToDate -and $Force) {
            Write-Log "INFO" (Get-Message "AlreadyInstalledWithForce" $installedVersion)
        }
    }
    #endregion

    #region 检查本地缓存
    $installerFileName    = "claude-$targetVersion-$script:Platform.exe"
    $installerPath        = Join-Path $script:DOWNLOAD_DIR $installerFileName
    $needsDownload        = $true

    if (Test-Path $installerPath) {
        try {
            $localChecksum    = (Get-FileHash -Path $installerPath -Algorithm SHA256).Hash.ToLower()
            $isCacheValid     = $localChecksum -eq $expectedChecksum

            if ($isCacheValid) {
                Write-Log "INFO" "本地已缓存版本 $targetVersion 且校验通过，跳过下载。"
                $needsDownload = $false
            }
            else {
                Write-Log "WARN" "本地缓存文件校验失败，将重新下载。"
                Remove-FileSafe -Path $installerPath
            }
        }
        catch {
            Write-Log "WARN" "无法验证本地缓存，将重新下载。"
            Remove-FileSafe -Path $installerPath
        }
    }
    #endregion

    #region 下载安装程序
    if ($needsDownload) {
        Write-Log "INFO" (Get-Message "Downloading")

        $downloadUrl = "$script:GCS_BUCKET/$targetVersion/$script:Platform/claude.exe"

        try {
            Download-File -Url $downloadUrl `
                -OutFile $installerPath `
                -ProxyUrl $proxyToUse `
                -Description "正在下载 Claude Code $targetVersion" `
                -MaxRetryAttempts $script:MaxRetryAttempts
        }
        catch {
            $errorInfo = Resolve-InstallError -Exception $_.Exception.Exception
            $logFile    = Save-Log

            Write-ErrorReport -ErrorType $errorInfo.Type `
                -Message $errorInfo.Message `
                -Suggestion $errorInfo.Suggestion `
                -LogFile $logFile
            exit 1
        }
    }
    #endregion

    #region 验证文件完整性
    Write-Log "INFO" (Get-Message "Verifying")

    $actualChecksum   = (Get-FileHash -Path $installerPath -Algorithm SHA256).Hash.ToLower()
    $isChecksumValid  = $actualChecksum -eq $expectedChecksum

    if (-not $isChecksumValid) {
        Remove-FileSafe -Path $installerPath
        $logFile = Save-Log

        Write-ErrorReport -ErrorType ([InstallErrorType]::Download_ChecksumMismatch) `
            -Message "文件校验失败。期望: $expectedChecksum, 实际: $actualChecksum" `
            -Suggestion "文件可能损坏或被篡改，已删除缓存，请重试" `
            -LogFile $logFile
        exit 1
    }

    Write-Log "SUCCESS" "文件校验通过"
    #endregion

    #region 执行安装
    Write-Log "INFO" (Get-Message "Installing")

    $isInstallSuccessful  = $false
    $originalErrorAction  = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    try {
        # 构建安装参数
        $installArguments = @("install", "--force")
        if ($Target -and $Target -ne "latest") {
            $installArguments = @("install", $Target, "--force")
        }

        # 执行安装程序
        & $installerPath @installArguments

        $isInstallSuccessful = ($LASTEXITCODE -eq 0) -or ($null -eq $LASTEXITCODE)

        if (-not $isInstallSuccessful) {
            throw "安装程序返回退出代码: $LASTEXITCODE"
        }
    }
    catch {
        $installError     = $_
        # 使用 ErrorRecord 本身（Resolve-InstallError 会处理两种情况）
        $errorInfo        = Resolve-InstallError -Exception $installError
        $isCertificateError = $errorInfo.Type -eq [InstallErrorType]::Network_Certificate

        # 如果 Exception 为空，使用错误消息本身
        $errorMessage = $installError.Exception.Message
        if ([string]::IsNullOrEmpty($errorMessage)) {
            $errorMessage = $installError.ToString()
        }

        Write-Log "ERROR" (Get-Message "ErrorInstall")
        Write-Log "ERROR" $errorMessage

        # 证书错误：回退到 pnpm 安装
        if ($isCertificateError) {
            Write-Log "INFO" ""
            Write-Log "INFO" "检测到证书验证错误，正在尝试使用 pnpm 方式安装..."
            Write-Log "INFO" ""

            $isInstallSuccessful = Invoke-PnpmFallback -InstallerPath $installerPath
        }

        # 如果仍然失败，显示错误报告
        if (-not $isInstallSuccessful) {
            $logFile = Save-Log
            Write-ErrorReport -ErrorType $errorInfo.Type `
                -Message $errorMessage `
                -Suggestion $errorInfo.Suggestion `
                -LogFile $logFile
            exit 1
        }
    }
    finally {
        $ErrorActionPreference = $originalErrorAction

        # 清理临时文件
        if ($isInstallSuccessful) {
            Write-Log "INFO" (Get-Message "CleaningUp")
            Start-Sleep -Seconds 1
            Remove-FileSafe -Path $installerPath
        }
    }
    #endregion

    #region 完成
    if ($isInstallSuccessful) {
        Write-Log "SUCCESS" ""
        Write-Log "SUCCESS" (Get-Message "InstallComplete")
        Write-Log "SUCCESS" ""
    }
    #endregion
}

<#
.SYNOPSIS
    执行 pnpm 回退安装

.DESCRIPTION
    当原生安装失败时，回退到 pnpm 方式安装 Claude Code

.PARAMETER InstallerPath
    原生安装程序路径（用于清理）

.OUTPUTS
    [bool] 是否安装成功
#>
function Invoke-PnpmFallback {
    [CmdletBinding()]
    param([string]$InstallerPath)

    $pnpmScriptLocal  = Join-Path $script:ScriptDir "pnpm.ps1"
    $pnpmScriptRemote = "https://raw.githubusercontent.com/GamblerIX/InstallCoder/main/ClaudeCode/pnpm.ps1"

    # 下载 pnpm 脚本（如果不存在）
    if (-not (Test-Path $pnpmScriptLocal)) {
        Write-Log "INFO" "正在下载 pnpm 安装脚本..."
        try {
            Invoke-WebRequest -Uri $pnpmScriptRemote -OutFile $pnpmScriptLocal -UseBasicParsing
            Write-Log "SUCCESS" "pnpm 脚本下载完成"
        }
        catch {
            Write-Log "ERROR" "下载 pnpm 脚本失败: $_"
            return $false
        }
    }

    # 清理原生安装程序
    Remove-FileSafe -Path $InstallerPath

    # 执行 pnpm 安装
    if (Test-Path $pnpmScriptLocal) {
        Write-Log "INFO" "正在使用 pnpm 方式安装..."

        try {
            & $pnpmScriptLocal

            if ($LASTEXITCODE -eq 0) {
                Write-Log "SUCCESS" ""
                Write-Log "SUCCESS" "pnpm 方式安装成功！"
                Write-Log "SUCCESS" ""

                # 尝试安装原生版本
                $nativeInstalled = Install-NativeClaude

                if ($nativeInstalled) {
                    Write-Log "SUCCESS" "原生 Claude Code 已成功安装！"
                }
                else {
                    Write-Log "WARN" "原生安装失败，但 pnpm 版 Claude Code 已可用。"
                    Write-Log "INFO" "您可以稍后手动运行 'claude install' 安装原生版本。"
                }

                return $true
            }
        }
        catch {
            Write-Log "ERROR" "pnpm 安装失败: $_"
        }
    }

    # 显示手动安装指导
    Write-Log "INFO" ""
    Write-Log "INFO" "建议手动安装："
    Write-Log "INFO" "  1. 安装 Node.js: https://nodejs.org/"
    Write-Log "INFO" "  2. 安装 pnpm:    npm install -g pnpm"
    Write-Log "INFO" "  3. 安装 Claude:  pnpm add -g @anthropic-ai/claude-code"

    return $false
}

# 主程序
if ($Help) {
    Write-Host (Get-Message "HelpText")
    exit 0
}

Install-ClaudeCode
