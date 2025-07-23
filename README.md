<p align="center">
  <h1>企业邮箱批量获取工具</h1>
</p>

<p align="center">
  <!-- Badges will go here -->
  <img alt="Python" src="https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white">
  <img alt="License: MIT" src="https://img.shields.io/badge/License-MIT-yellow.svg">
  <!-- Add more badges as needed, e.g., GitHub Actions status -->
</p>

这是一个使用 Python 编写的自动化脚本，旨在通过调用 `gemini-cli` 命令行工具，批量查询 Excel 文件中指定公司的官方联系邮箱，并将结果写回原文件。

## 📖 目录

- [✨ 功能特性](#-功能特性)
- [⚙️ 安装与配置](#-安装与配置)
  - [第一部分：安装与配置 `gemini-cli`](#-第一部分安装与配置-gemini-cli)
    - [1. 前提条件：安装 Node.js](#1-前提条件安装-nodejs)
    - [2. 安装 `gemini-cli`](#2-安装-gemini-cli)
    - [3. 配置认证 (重要)](#3-配置认证-重要)
    - [4. 网络注意事项 (重要)](#4-网络注意事项-重要)
  - [第二部分：安装本项目依赖](#-第二部分安装本项目依赖)
- [⚙️ 高级配置](#-高级配置)
- [🚀 使用方法](#-使用方法)
- [📝 注意事项](#-注意事项)
- [📄 许可证](#-许可证)

## ✨ 功能特性

- **自动化处理**: 自动读取 Excel 文件，并遍历公司列表进行查询。
- **智能查询**: 利用 Gemini 的强大能力和优化的搜索策略提示词，最大化提升查找成功率。
- **实时保存**: 每处理完一条记录后立即保存结果到 Excel，确保数据安全。
- **断点续传**: 实现智能进度判断、错误自动恢复、交互式控制和详细状态管理。
    - 智能进度判断：仅当记录包含有效邮箱或"Not Found"时视为已完成。
    - 错误自动恢复：API调用失败的记录会自动重置以便重试。
    - 交互式控制：启动时提供继续处理或重新开始的选项。
    - 详细状态管理：成功结果保留邮箱地址，未找到标记为"Not Found"，临时错误清空状态等待重试。
- **精细化日志**: 提供控制台实时显示、未找到邮箱记录、任务启动报告以及API和系统错误单独记录。
    - 控制台实时显示：在控制台实时显示完整的处理过程。
    - 未找到邮箱记录：未找到邮箱的公司将追加记录到 `not_found_log.log` 中。
    - 任务启动报告：每次任务启动时，日志文件会记录启动时间及当次任务的汇总报告。
    - 错误单独记录：API错误和系统错误单独记录。
- **配额自动恢复**: 当检测到API配额用尽时，自动进入定时重试模式（默认30分钟间隔）。
- **交互式进度管理**: 启动时自动检测已有进度，可选择继续处理或重新开始。
- **双语言支持**: 同时支持英文和中文公司名称查询，优先使用中文名搜索本地资源。

## ⚙️ 安装与配置

在运行本脚本前，您需要完成 **2个部分** 的环境配置：`gemini-cli` 的安装与配置，以及本项目 Python 依赖的安装。

### **第一部分：安装与配置 `gemini-cli`**

本工具依赖于 Google 的 `gemini-cli` 命令行程序。

#### 1. 前提条件：安装 Node.js

`gemini-cli` 是一个 Node.js 包，因此您需要先确保系统中已安装 Node.js (推荐LTS版本) 和 npm。您可以从 [Node.js 官网](https://nodejs.org/) 下载安装。

#### 2. 安装 `gemini-cli`

打开您的终端，执行以下命令进行全局安装：

```bash
npm install -g @google/gemini-cli
```

#### 3. 配置认证 (重要)

安装完成后，您需要授权 `gemini-cli` 访问您的 Google 账户。**推荐方式是交互式登录**：

- 在终端中运行 `gemini` 命令。
- 首次运行时，程序会提示您选择认证方式。
- 按 `回车 (Enter)` 键选择默认的 `"1. Login with Google"`。
- `gemini-cli` 会提供一个链接，在浏览器中打开并完成 Google 账户的登录和授权流程即可。
- 授权成功后，您的登录信息会被保存下来，后续即可正常使用。

#### 4. 网络注意事项 (重要)

如果您所在的地区无法直接访问 Google 服务，您需要使用网络代理。

由于 `gemini-cli` 是一个命令行工具，普通的系统 HTTP/HTTPS 代理可能无法接管其网络流量。因此，**强烈推荐使用支持 TUN 模式（虚拟网卡）的代理工具**，例如 [Clash Verge Rev](https://github.com/clash-verge-rev/clash-verge-rev)。

- **使用方法**: 在运行 `gemini` 命令或本项目的 `main.py` 脚本**之前**，请务必启动您的代理工具并**开启 TUN 模式**。

### **第二部分：安装本项目依赖**

1.  确保您已安装 Python 3.10 或更高版本。
2.  在本项目根目录下，打开终端，运行以下命令来安装所有必需的 Python 库：

    ```bash
    pip install -r requirements.txt
    ```

## ⚙️ 高级配置

您可以通过修改 `main.py` 中的以下常量来定制脚本行为：

```python
# 文件配置
EXCEL_FILE = 'data.xlsx'  # Excel文件名
SHEET_NAME = 'Sheet1'     # 工作表名称

# 列名配置
COMPANY_NAME_EN_COL = 'company_name'    # 公司英文名所在列
COMPANY_NAME_TC_COL = 'company_name_tc' # 公司中文名所在列
EMAIL_COL = 'Email'                     # 结果写入列

# Gemini配置
GEMINI_MODEL = 'gemini-2.5-flash'       # 使用的Gemini模型
RETRY_INTERVAL_MINUTES = 30             # API配额错误重试间隔（分钟）
TASK_INTERVAL_SECONDS = 10              # 任务间隔时间（秒）
GEMINI_TIMEOUT_SECONDS = 180            # Gemini调用超时时间（秒）
```

## 🚀 使用方法

1.  **准备 Excel 文件**:
    *   确保项目根目录下存在名为 `data.xlsx` 的文件。
    *   文件中必须包含一个名为 `Sheet1` 的工作表。
    *   工作表中必须包含 `company_name` (公司英文名) 和/或 `company_name_tc` (公司中文名) 列。

2.  **运行脚本**:
    *   完成所有安装与配置后，在项目根目录下打开终端，执行：

    ```bash
    python main.py
    ```

    启动脚本后，会首先在控制台打印当前加载的所有配置参数，方便您核对。

    **交互式菜单**：
    当检测到已有处理进度时，程序会提供以下选项：
    ```
    1. 继续上次任务（跳过已完成的记录）
    2. 重新开始（清空所有结果）
    ```
    - 输入 `1` 或直接按回车键将继续处理未完成的记录。
    - 输入 `2` 将清除所有已有结果并重新开始。

3.  **处理过程**:
    - 程序会自动处理每条记录，控制台显示实时进度
    - 每条记录处理之间会有一个可配置的等待时间 (`TASK_INTERVAL_SECONDS`)，避免频繁调用。
    - 遇到API配额限制时，会自动等待并重试（默认30分钟）。
    - Gemini API调用会有一个可配置的超时时间 (`GEMINI_TIMEOUT_SECONDS`)，默认为3分钟。
    - 按 `Ctrl+C` 可安全中断程序，进度会自动保存。

4.  **查看结果**:
    *   程序运行期间，处理结果会实时保存在 `data.xlsx` 的 `Email` 列中。
    *   所有未找到邮箱的公司，以及最终的统计报告，都会保存在 `not_found_log.log` 文件中。

## 📝 注意事项

1.  **AI提示词优化**：脚本使用精心设计的提示词模板（[见main.py第44-58行]），优先搜索香港本地资源。建议保持提示词原样以获得最佳效果；如需修改，请确保保留关键搜索策略。

2.  **API配额限制**：Gemini API有每日配额限制。当遇到配额错误时，程序会自动进入重试模式，您可通过修改 `RETRY_INTERVAL_MINUTES` 调整重试间隔。

3.  **中文名称优势**：在处理香港公司时，中文名称通常能获得更好的搜索结果。请确保中文名称列包含准确的繁体中文名称。

## 📄 许可证

本项目采用 MIT 许可证。有关更多信息，请参阅 [LICENSE](LICENSE) 文件。
