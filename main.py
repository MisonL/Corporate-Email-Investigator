# -*- coding: utf-8 -*-
"""
本脚本用于自动从 Excel 文件中读取公司列表，
通过调用 gemini-cli 工具查询每家公司的联系邮箱，
并将结果写回 Excel 文件。同时，它会生成一份详细的日志报告。
"""
import pandas as pd
import subprocess
import sys
import time
import logging

# --- 全局配置 ---
EXCEL_FILE = 'data.xlsx'  # 要处理的 Excel 文件名
SHEET_NAME = 'Sheet1'  # Excel 文件中的工作表名称
COMPANY_NAME_EN_COL = 'company_name'  # 公司英文名所在的列名
COMPANY_NAME_TC_COL = 'company_name_tc'  # 公司中文名所在的列名
EMAIL_COL = 'Email'  # 用于保存邮箱地址的列名
LOG_FILE = 'not_found_log.log'  # 用于记录未找到邮箱的公司和最终汇总的日志文件名
GEMINI_MODEL = 'gemini-2.5-flash'  # 指定使用的 Gemini 模型

# --- 日志系统配置 ---
# 创建一个名为 'console_logger' 的记录器，用于在控制台输出所有实时信息
console_logger = logging.getLogger('console_logger')
# 防止在某些环境中重复添加 handler
if not console_logger.handlers:
    console_logger.setLevel(logging.INFO)  # 设置日志级别为 INFO
    console_handler = logging.StreamHandler()  # 创建一个流处理器，输出到控制台
    # 设置控制台日志的格式
    console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    console_logger.addHandler(console_handler)  # 将处理器添加到记录器

# 创建一个名为 'file_logger' 的记录器，用于将指定信息保存到文件
file_logger = logging.getLogger('file_logger')
# 防止在某些环境中重复添加 handler
if not file_logger.handlers:
    file_logger.setLevel(logging.INFO)  # 设置日志级别为 INFO
    # 创建一个文件处理器，'w' 模式表示每次运行都覆盖旧文件
    file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
    # 设置文件日志的格式
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    file_logger.addHandler(file_handler)  # 将处理器添加到记录器

# --- AI 提示词模板 ---
# 这是一个经过多轮优化的提示词，旨在引导 AI 更高效、更准确地找到企业邮箱
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

def get_email_from_gemini(company_name_en: str, company_name_tc: str) -> str:
    """
    调用 gemini-cli 来获取指定公司的邮箱地址。

    Args:
        company_name_en (str): 公司的英文名称。
        company_name_tc (str): 公司的中文名称。

    Returns:
        str: 返回获取到的邮箱地址，或 "Not Found"，或错误信息。
    """
    # 将公司名称填入提示词模板
    prompt = PROMPT_TEMPLATE.format(company_name=company_name_en, company_name_tc=company_name_tc)
    # 构建要执行的命令列表，强制使用指定的模型
    command = ['gemini', '-m', GEMINI_MODEL]
    try:
        # 执行外部命令。使用 input 参数通过 stdin 传递 prompt，避免特殊字符问题。
        result = subprocess.run(
            command,
            capture_output=True,  # 捕获标准输出和标准错误
            text=True,  # 以文本模式处理输出
            check=True,  # 如果命令返回非零退出码，则抛出异常
            encoding='utf-8',  # 指定编码
            input=prompt  # 将 prompt 作为标准输入
        )
        # 清理和解析返回结果，通常有效信息在最后一行
        lines = result.stdout.strip().split('\n')
        return lines[-1].strip() if lines else "Error: No output"
    except subprocess.CalledProcessError as e:
        # 如果 gemini-cli 返回错误，则记录到控制台
        console_logger.error(f"Gemini Error: {e.stderr.strip()}")
        return "Error: Gemini call failed"
    except FileNotFoundError:
        # 如果系统找不到 gemini 命令，则记录错误并退出
        console_logger.error("'gemini' 命令未找到。请确保 gemini-cli 已安装并位于系统的 PATH 中。")
        sys.exit(1)

def main():
    """
    脚本的主执行函数，负责整个业务流程的调度。
    """
    console_logger.info("--- 开始处理 ---")
    # 初始化计数器
    not_found_count = 0
    success_count = 0
    
    try:
        # 使用 pandas 读取 Excel 文件
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine='openpyxl')
        total_count = len(df)
        console_logger.info(f"成功读取文件 '{EXCEL_FILE}', 找到 {total_count} 条记录。")

        # 如果用于存放邮箱的列不存在，则创建一个空列
        if EMAIL_COL not in df.columns:
            df[EMAIL_COL] = ""

        # --- 手动重置文件日志处理器 ---
        # 确保每次运行都生成一个全新的日志文件
        if file_logger.hasHandlers():
            for handler in file_logger.handlers[:]:
                handler.close()
                file_logger.removeHandler(handler)
        file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        file_logger.addHandler(file_handler)

        # 遍历 Excel 中的每一行
        for index in df.index:
            # 实现断点续传：如果该行已经有邮箱，则跳过
            if pd.notna(df.at[index, EMAIL_COL]) and df.at[index, EMAIL_COL] != '':
                continue

            # 获取英文和中文公司名，处理空值情况
            company_en = str(df.at[index, COMPANY_NAME_EN_COL]) if COMPANY_NAME_EN_COL in df.columns and pd.notna(df.at[index, COMPANY_NAME_EN_COL]) else ''
            company_tc = str(df.at[index, COMPANY_NAME_TC_COL]) if COMPANY_NAME_TC_COL in df.columns and pd.notna(df.at[index, COMPANY_NAME_TC_COL]) else ''
            
            # 如果两个名称都为空，则跳过
            current_company = company_en or company_tc
            if not current_company:
                console_logger.info(f"[{index + 1}/{total_count}] 跳过空行...")
                continue

            # 在控制台打印当前进度
            console_logger.info(f"[{index + 1}/{total_count}] 正在处理: {current_company}")
            # 调用函数获取邮箱
            email = get_email_from_gemini(company_en, company_tc)
            console_logger.info(f"-> 结果: {email}")

            # 将获取到的结果写入 DataFrame
            df.at[index, EMAIL_COL] = email
            
            # 根据结果更新计数器，并记录到文件日志
            if email == "Not Found":
                not_found_count += 1
                file_logger.info(f"未找到邮箱: {current_company}")
            elif not email.startswith("Error:"):
                success_count += 1
            
            # 实时保存：每处理一条就保存一次 Excel 文件
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
            # 增加延迟，避免请求频率过高
            time.sleep(2)

        console_logger.info("--- 全部处理完成！最终结果已在文件中。 ---")
        
        # --- 在文件日志中写入最终的汇总报告 ---
        file_logger.info("\n" + "="*50)
        file_logger.info("           处理结果汇总")
        file_logger.info("="*50)
        file_logger.info(f"总共处理公司数: {success_count + not_found_count}")
        file_logger.info(f"  - 成功找到邮箱: {success_count} 家")
        file_logger.info(f"  - 未找到邮箱:   {not_found_count} 家")
        file_logger.info("="*50)
        file_logger.info("未找到邮箱的公司列表见上方明细。")

    except FileNotFoundError:
        # 处理文件找不到的异常
        console_logger.error(f"错误：文件 '{EXCEL_FILE}' 未找到。请确保文件在正确的路径下。")
        sys.exit(1)
    except Exception as e:
        # 处理所有其他未知异常
        console_logger.error(f"发生未知错误: {e}")

# 当该脚本被直接执行时，才调用 main() 函数
# 如果该脚本被其他模块导入，则不执行 main()
if __name__ == '__main__':
    main()