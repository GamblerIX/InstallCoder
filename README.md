# InstallCoder

> 通过PowerShell下载远程.ps1脚本执行实现快速安装各种AICode。

## Claude Code

### Claude Code（Windows原生）- 推荐

```Powershell
irm https://raw.githubusercontent.com/GamblerIX/InstallCoder/main/ClaudeCode/main.ps1 | iex
```

#### 自动代理检测

脚本会自动检测以下代理设置：
1. 参数指定的代理 (`-Proxy`)
2. 环境变量代理 (`$env:HTTP_PROXY`)
3. 系统代理设置

#### 智能回退机制

当原生安装遇到证书验证错误时，脚本会自动回退到 **pnpm 方式** 安装，无需手动干预。

#### 参数说明

```powershell
# 使用指定代理
.\main.ps1 -Proxy "http://127.0.0.1:7890"

# 强制重新安装
.\main.ps1 -Force

# 显示帮助
.\main.ps1 -Help
```

#### 故障排查

如果安装失败，脚本会：
- 显示中文错误提示
- 自动尝试 pnpm 方式安装
- 生成详细日志到 `~\.claude\logs\`

---

### Claude Code（pnpm方式）

> 当原生安装遇到证书问题时，会自动回退到此方式。
> 如果Windows没有nodejs环境，将自动下载最新的LTS版本到 `~\nodejs` 中并添加系统级环境Path。

```Powershell
irm https://raw.githubusercontent.com/GamblerIX/InstallCoder/main/ClaudeCode/pnpm.ps1 | iex
```

#### 功能特点

- 自动检测并安装 Node.js（如果没有）
- 自动检测并安装 pnpm（如果没有）
- 使用 pnpm 全局安装 Claude Code
- 自动配置环境变量

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

**解决方案：**

脚本已自动处理此问题：
1. 首先尝试原生安装（使用 `--force` 参数）
2. 如果失败，自动回退到 **pnpm 方式** 安装
3. 无需手动干预

如果自动回退也失败，可以手动安装：
1. 安装 Node.js: https://nodejs.org/
2. 安装 pnpm: `npm install -g pnpm`
3. 安装 Claude: `pnpm add -g @anthropic-ai/claude`

---

## 中国用户特别提示

由于 Google Cloud Storage 在国内访问受限，建议使用代理：

- **Clash** 默认端口: `http://127.0.0.1:7890`
- **v2rayN** 默认端口: `http://127.0.0.1:10809`

使用代理安装：
```Powershell
irm https://raw.githubusercontent.com/GamblerIX/InstallCoder/main/ClaudeCode/main.ps1 | iex -Proxy "http://127.0.0.1:7890"
```
