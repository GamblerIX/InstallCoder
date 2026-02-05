# InstallCoder

> 通过PowerShell下载远程.ps1脚本执行实现快速安装各种AICode。

## Claude Code

### Claude Code（Windows原生）

```Powershell
irm https://github.com/GamblerIX/ClaudeCode/main.ps1 | iex
```

#### 自动代理检测

脚本会自动检测以下代理设置：
1. 参数指定的代理 (`-Proxy`)
2. 环境变量代理 (`$env:HTTP_PROXY`)
3. 系统代理设置

#### 故障排查

如果安装失败，脚本会：
- 显示中文错误提示
- 提供解决方案建议
- 生成详细日志到 `~\.claude\logs\`

---

### Claude Code（pnpm）

> 如果Windows没有nodejs环境，将自动下载最新的LTS版本到 `~
odejs` 中并添加系统级环境Path。

```Powershell
irm https://github.com/GamblerIX/ClaudeCode/pnpm.ps1 | iex
```

### Claude Code（npm）

> 如果Windows没有nodejs环境，将自动下载最新的LTS版本到 `~
odejs` 中并添加系统级环境Path。

```Powershell
irm https://github.com/GamblerIX/ClaudeCode/npm.ps1 | iex
```
