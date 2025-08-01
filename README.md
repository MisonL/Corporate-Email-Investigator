<div align="center">
  <h1>企业邮箱调查器</h1>
  <!-- Badges will go here -->
  <img alt="Python" src="https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white">
  <img alt="License: MIT" src="https://img.shields.io/badge/License-MIT-yellow.svg">
  <img alt="Powered by Google Gemini" src="https://img.shields.io/badge/Powered%20by-Google%20Gemini-purple?logo=google&logoColor=white">
  <!-- Add more badges as needed, e.g., GitHub Actions status -->
</div>

一个由 `gemini-cli` 驱动的 Python 程序，化身为你的私人助手，帮你批量调查企业公开的联系邮箱。
Tips：通过修改 AI 提示词可以转变成其他功能用途哦！

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
- [🌟 默认AI提示词](#-默认ai提示词)
- [📝 注意事项](#-注意事项)
- [📄 许可证](#-许可证)

## ✨ 功能特性

- **自动化处理**: 自动读取 Excel 文件，并遍历公司列表进行查询。
- **智能查询**: 利用 Gemini 的强大能力和优化的搜索策略提示词，最大化提升查找成功率。
- **实时保存**: 每处理完一条记录后立即保存结果到 Excel，确保数据安全。
- **断点续传**: 实现智能进度判断、错误自动恢复、交互式控制和详细状态管理。
    - 智能进度判断：仅当记录包含有效邮箱或"Not Found"时视为已完成。
    - 错误自动恢复：API调用失败的记录会自动重置以便重试。
    - 交互式控制：启动时提供继续处理、重新开始或**重试处理失败（Not Found）**的选项。
    - 详细状态管理：成功结果保留邮箱地址，未找到标记为"Not Found"，临时错误清空状态等待重试。
- **数据文件状态统计**: 启动时在终端中详细列出 Excel 文件中已处理成功、处理失败 (Not Found) 和未处理 (空白或错误) 的公司数量。
- **进度显示优化**: 终端中的任务进度提示 `[XXXX/XXXX]` 的总数现在精确表示本次程序启动所有要处理的任务总数（即未处理的数量）。
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

本工具依赖于 Google 的 `gemini-cli` 命令行程序。官方仓库：[`https://github.com/google-gemini/gemini-cli/tree/main`](https://github.com/google-gemini/gemini-cli/tree/main)

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
GEMINI_MODEL = 'gemini-2.5-flash'       # 使用的Gemini模型 (现在支持通过环境变量 GEMINI_MODEL 配置)
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

    启动脚本后，会首先在控制台打印当前加载的所有配置参数，并详细列出数据文件当前的状态（已处理成功、处理失败和未处理数量），方便您核对。

    **交互式菜单**：
    当检测到已有处理进度时，程序会提供以下选项：
    ```
    1. 继续上次任务（跳过已完成的记录，重试失败的记录）
    2. 重新开始（清空所有结果）
    3. 重试处理失败（Not Found）的记录
    ```
    - 输入 `1` 或直接按回车键将继续处理未完成（包括上次失败）的记录。
    - 输入 `2` 将清除所有已有结果并重新开始处理所有记录。
    - 输入 `3` 将仅重试上次处理结果为“Not Found”的记录。

3.  **处理过程**:
    - 程序会自动处理每条记录，控制台显示实时进度（例如 `[1/100] 正在处理: XXX公司`，其中总数 `100` 表示本次程序启动需要处理的任务总数）。
    - 每条记录处理之间会有一个可配置的等待时间 (`TASK_INTERVAL_SECONDS`)，避免频繁调用。
    - 遇到API配额限制时，会自动等待并重试（默认30分钟）。
    - Gemini API调用会有一个可配置的超时时间 (`GEMINI_TIMEOUT_SECONDS`)，默认为3分钟。
    - 按 `Ctrl+C` 可安全中断程序，进度会自动保存。

4.  **查看结果**:
    *   程序运行期间，处理结果会实时保存在 `data.xlsx` 的 `Email` 列中。
    *   所有未找到邮箱的公司，以及最终的统计报告，都会保存在 `not_found_log.log` 文件中。

## 🌟 默认AI提示词

本工具的核心在于其强大的AI提示词，它指导 Gemini 模型进行高效的企业邮箱搜索。以下是 `main.py` 中使用的默认提示词模板：

```python
PROMPT_TEMPLATE = (
    "你是一名顶尖的企业信息调查员，专注于查找香港地区公司的联系方式。你的任务是基于我提供的公司名称，通过联网搜索，不惜一切代价找到该公司的官方联系邮箱。\n\n"
    "--- 公司信息 ---\n"
    "英文名: {company_name}\n"
    "中文名: {company_name_tc}\n\n"
    "--- 建议搜索策略 (请优先使用) ---\n"
    "1.  **首要目标 - 官方网站**: 深度挖掘公司的官方网站，特别是“联系我们”(Contact Us)、“关于我们”(About Us) 或页脚部分。\n"
    "2.  **香港官方数据库**: 重点查询香港公司注册处 (Cyber Search Centre) 和香港贸易发展局 (HKTDC) 的数据库。\n"
    "3.  **本地商业目录**: 搜索香港黄页 (yp.com.hk) 和其他本地商业名录。\n"
    "4.  **专业和社交网络**: 检查公司的 LinkedIn 官方页面，以及它可能所属的行业协会网站（如香港保安业协会、香港物业管理公司协会等）。\n"
    "5.  **善用中文名**: 在搜索香港本地资源时，请充分利用公司的中文名称。\n\n"
    "--- 输出要求 ---\n"
    "请严格按照以下格式返回结果，不要返回任何与邮箱无关的文字、解释或说明。\n"
    "如果找到邮箱，请只返回邮箱地址。如果找不到，请只返回 \"Not Found\"。"
)
```

**Tips**: 您可以根据自己的需求修改 `main.py` 中的 `PROMPT_TEMPLATE`，将其转变为其他功能用途，例如：
- 查找公司电话号码
- 提取公司地址
- 总结公司业务范围
- 甚至可以用于其他领域的自动化信息提取任务，只要您能设计出合适的提示词。

## 📝 注意事项

1.  **AI提示词优化**：脚本使用精心设计的提示词模板，优先搜索香港本地资源。建议保持提示词原样以获得最佳效果；如需修改，请确保保留关键搜索策略。

2.  **API配额限制**：Gemini API有每日配额限制。当遇到配额错误时，程序会自动进入重试模式，您可通过修改 `RETRY_INTERVAL_MINUTES` 调整重试间隔。

3.  **中文名称优势**：在处理香港公司时，中文名称通常能获得更好的搜索结果。请确保中文名称列包含准确的繁体中文名称。

## 📄 许可证

本项目采用 MIT 许可证。有关更多信息，请参阅 [LICENSE](LICENSE) 文件。
