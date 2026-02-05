# InstallCoder

> 通过PowerShell下载远程.ps1脚本执行实现快速安装各种AICode。

## Claude Code

### Claude Code（Windows原生）

```Powershell
irm https://raw.githubusercontent.com/GamblerIX/InstallCoder/main/ClaudeCode/main.ps1 | iex
```

#### 自动代理检测

脚本会自动检测以下代理设置：
1. 参数指定的代理 (`-Proxy`)
2. 环境变量代理 (`$env:HTTP_PROXY`)
3. 系统代理设置

#### 参数说明

```powershell
# 使用指定代理
.\main.ps1 -Proxy "http://127.0.0.1:7890"

# 强制重新安装
.\main.ps1 -Force

# 显示帮助
.\main.ps1 -Help
```

---

### Claude Code（pnpm方式）

> 当原生安装遇到证书问题时，会自动回退到此方式以在pnpm安装完成后尝试升级原生安装。
> 如果Windows没有nodejs环境，将自动下载最新的LTS版本到 `~\nodejs` 中并添加系统级环境Path。

```Powershell
irm https://raw.githubusercontent.com/GamblerIX/InstallCoder/main/ClaudeCode/pnpm.ps1 | iex
```

---

## 已知问题

### 证书验证错误

在某些网络环境下，Claude Code 原生安装可能会遇到证书验证错误：

```Powershell
✘ Installation failed

Failed to fetch version from
https://storage.googleapis.com/claude-code-dist-86c565f3-f756-42ad-8dfa-d59b1c096819/claude-code-releases/latest:
unknown certificate verification error

Try running with --force to override checks
```

**规避解决方案：**

脚本将自动迂回处理此问题：
1. 回退至pnpm安装
2. 安装完成后自动执行`claude install`尝试原生安装
