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

# --- 自定义异常 ---
class QuotaExceededError(Exception):
    """当检测到API配额用尽时抛出此异常"""
    pass

# --- 配置 ---
EXCEL_FILE = 'data.xlsx'
SHEET_NAME = 'Sheet1'
COMPANY_NAME_EN_COL = 'company_name'
COMPANY_NAME_TC_COL = 'company_name_tc'
EMAIL_COL = 'Email'
LOG_FILE = 'not_found_log.log'
GEMINI_MODEL = 'gemini-2.5-flash'
RETRY_INTERVAL_MINUTES = 30  # 配额错误重试间隔（分钟）
TASK_INTERVAL_SECONDS = 10    # 任务间隔时间（秒）
GEMINI_TIMEOUT_SECONDS = 180  # Gemini调用超时时间（秒）

# --- 日志配置 ---
console_logger = logging.getLogger('console_logger')
if not console_logger.handlers:
    console_logger.setLevel(logging.INFO)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    console_logger.addHandler(console_handler)

file_logger = logging.getLogger('file_logger')
if not file_logger.handlers:
    file_logger.setLevel(logging.INFO)
    file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    file_logger.addHandler(file_handler)

# --- AI 提示词模板 ---
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
    调用 gemini-cli 获取邮箱地址。
    
    Args:
        company_name_en (str): 公司的英文名称
        company_name_tc (str): 公司的中文名称

    Returns:
        str: 邮箱地址或错误信息

    Raises:
        QuotaExceededError: 当API配额用尽时抛出
    """
    prompt = PROMPT_TEMPLATE.format(company_name=company_name_en, company_name_tc=company_name_tc)
    command = ['gemini', '-m', GEMINI_MODEL]
    
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=True,
            encoding='utf-8',
            input=prompt,
            timeout=GEMINI_TIMEOUT_SECONDS  # 设置超时时间
        )
        lines = result.stdout.strip().split('\n')
        return lines[-1].strip() if lines else "Error: No output"
    except subprocess.CalledProcessError as e:
        # 检测配额错误
        if "Quota exceeded" in e.stderr or "RESOURCE_EXHAUSTED" in e.stderr:
            raise QuotaExceededError("API配额已用尽") from e
            
        console_logger.error(f"Gemini Error: {e.stderr.strip()}")
        return "Error: Gemini call failed"
    except subprocess.TimeoutExpired:
        console_logger.error(f"Gemini调用超时（{GEMINI_TIMEOUT_SECONDS}秒），跳过该记录")
        return "Error: Timeout"
    except FileNotFoundError:
        console_logger.error("'gemini' 命令未找到。请确保 gemini-cli 已安装并位于系统的 PATH 中。")
        sys.exit(1)

def main():
    """主函数，处理整个流程"""
    console_logger.info("--- 开始处理 ---")
    df = None  # 显式初始化变量
    
    try:
        # 读取Excel文件
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine='openpyxl')
        total_count = len(df)
        console_logger.info(f"成功读取文件 '{EXCEL_FILE}', 找到 {total_count} 条记录。")

        # 初始化Email列
        if EMAIL_COL not in df.columns:
            df[EMAIL_COL] = ''

        # --- 交互式菜单 ---
        # 检测有效进度（排除错误状态）
        valid_progress_mask = df[EMAIL_COL].apply(lambda x: x not in ['', 'Error: No output', 'Error: Gemini call failed'])
        has_progress = valid_progress_mask.any()
        if has_progress:
            print("\n检测到已有处理进度:")
            print("1. 继续上次任务（跳过已完成的记录）")
            print("2. 重新开始（清空所有结果）")
            choice = input("请选择操作 (默认1): ").strip() or "1"
            
            if choice == "2":
                df[EMAIL_COL] = ''
                console_logger.info("已清空所有结果，重新开始处理。")
            else:
                # 清除错误状态以便重试
                error_mask = df[EMAIL_COL].isin(['Error: No output', 'Error: Gemini call failed'])
                df.loc[error_mask, EMAIL_COL] = ''
                console_logger.info("已重置错误状态记录，将继续处理")

        # 重置文件日志
        # 移除旧的文件处理器（如果有），确保日志文件不会重复写入
        for handler in file_logger.handlers[:]:
            if isinstance(handler, logging.FileHandler):
                file_logger.removeHandler(handler)

        # 配置文件日志，以追加模式写入
        file_handler = logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        file_logger.addHandler(file_handler)

        # 记录本次任务启动时间
        start_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        file_logger.info("\n" + "="*70)
        file_logger.info(f"任务启动时间: {start_time}")
        file_logger.info("="*70)

        # 处理数据
        not_found_count = 0
        success_count = 0
        
        for index in df.index:
            # 跳过已处理的记录（仅当有有效结果时跳过）
            current_email = df.at[index, EMAIL_COL]
            if pd.notna(current_email) and current_email not in ['', 'Error: No output', 'Error: Gemini call failed']:
                continue

            company_en = str(df.at[index, COMPANY_NAME_EN_COL]) if COMPANY_NAME_EN_COL in df.columns and pd.notna(df.at[index, COMPANY_NAME_EN_COL]) else ''
            company_tc = str(df.at[index, COMPANY_NAME_TC_COL]) if COMPANY_NAME_TC_COL in df.columns and pd.notna(df.at[index, COMPANY_NAME_TC_COL]) else ''
            
            current_company = company_en or company_tc
            if not current_company:
                console_logger.info(f"[{index + 1}/{total_count}] 跳过空行...")
                continue

            display_name = ""
            if company_en and company_tc:
                display_name = f"{company_en} ({company_tc})"
            elif company_en:
                display_name = company_en
            elif company_tc:
                display_name = company_tc
            
            console_logger.info(f"[{index + 1}/{total_count}] 正在处理: {display_name}")
            
            try:
                email = get_email_from_gemini(company_en, company_tc)
            except QuotaExceededError as e:
                console_logger.error(f"错误：API每日配额已用尽！将在 {RETRY_INTERVAL_MINUTES} 分钟后重试...")
                console_logger.error("您可以：")
                console_logger.error("1. 等待自动重试")
                console_logger.error("2. 手动中断程序（Ctrl+C）并稍后重新运行")
                
                # 自动重试逻辑
                while True:
                    time.sleep(RETRY_INTERVAL_MINUTES * 60)
                    console_logger.info(f"重试中... ({RETRY_INTERVAL_MINUTES}分钟间隔)")
                    try:
                        email = get_email_from_gemini(company_en, company_tc)
                        console_logger.info("配额已恢复，继续处理！")
                        break
                    except QuotaExceededError:
                        console_logger.error(f"配额仍未恢复，{RETRY_INTERVAL_MINUTES}分钟后再次重试...")
                        continue
                
            console_logger.info(f"  -> 结果: {email}")
            if email.startswith("Error:"):
                console_logger.warning("发生API调用错误，跳过该记录")
            else:
                df.at[index, EMAIL_COL] = email
            
            if email == "Not Found":
                not_found_count += 1
                file_logger.info(f"未找到邮箱: {current_company}")
            elif not email.startswith("Error:") and email != "Error: Timeout":
                success_count += 1
            
            # 实时保存结果
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
            time.sleep(TASK_INTERVAL_SECONDS)

        # 最终报告
        console_logger.info("--- 全部处理完成！最终结果已在文件中。 ---")
        
        end_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        file_logger.info("\n" + "="*70)
        file_logger.info("           本次任务处理结果汇总")
        file_logger.info("="*70)
        file_logger.info(f"任务结束时间: {end_time}")
        file_logger.info(f"总共处理记录数: {total_count}") # 使用total_count表示本次处理了多少行，无论是否跳过
        file_logger.info(f"  - 成功找到邮箱: {success_count} 家")
        file_logger.info(f"  - 未找到邮箱:   {not_found_count} 家")
        file_logger.info("="*70)

    except FileNotFoundError:
        console_logger.error(f"错误：文件 '{EXCEL_FILE}' 未找到。请确保文件在正确的路径下。")
        sys.exit(1)
    except KeyboardInterrupt:
        console_logger.info("\n用户中断操作，已保存当前进度。")
        if df is not None:
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        sys.exit(0)
    except Exception as e:
        console_logger.error(f"发生未知错误: {e}")
        if df is not None:
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        else:
            console_logger.error("错误发生时尚未加载数据文件")

if __name__ == '__main__':
    main()