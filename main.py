# -*- coding: utf-8 -*-
"""
本脚本用于自动从 Excel 文件中读取公司列表，
通过调用 gemini-cli 工具查询每家公司的联系邮箱，
并将结果写回 Excel 文件。同时，它会生成一份详细的日志报告。
"""

# --- 导入模块 ---
import pandas as pd
import subprocess
import sys
import time
import logging
import os

# --- 配置 ---
# Excel文件相关配置
EXCEL_FILE = 'data.xlsx'  # 要处理的Excel文件名
SHEET_NAME = 'Sheet1'  # 工作表名称
COMPANY_NAME_EN_COL = 'company_name'  # 公司英文名列
COMPANY_NAME_TC_COL = 'company_name_tc'  # 公司中文名列
EMAIL_COL = 'Email'  # 邮箱结果列

# 日志配置
LOG_FILE = 'not_found_log.log'  # 日志文件名

# Gemini API相关配置
GEMINI_MODEL = os.getenv('GEMINI_MODEL', 'gemini-2.5-flash')  # 从环境变量获取模型名称，默认为'gemini-2.5-flash'
RETRY_INTERVAL_MINUTES = 30  # 配额错误重试间隔（分钟）
TASK_INTERVAL_SECONDS = 10  # 任务间隔时间（秒）
GEMINI_TIMEOUT_SECONDS = 3600  # Gemini调用超时时间（秒）
MAX_API_CALL_RETRIES = 3  # API调用（非配额）最大重试次数
API_RETRY_DELAY_SECONDS = 5  # API调用重试间隔（秒）

# --- AI 提示词模板 ---
# 定义发送给Gemini的提示词模板，包含详细的搜索策略和输出要求
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

# --- 自定义异常 ---
class QuotaExceededError(Exception):
    """当检测到API配额用尽时抛出此异常"""
    pass

# --- 动画函数 ---
def spinning_cursor(seconds, message=""):
    """
    显示一个旋转光标动画，持续指定的秒数，改善用户等待体验
    
    Args:
        seconds (int): 动画持续的秒数
        message (str): 要显示的消息
        
    功能：
        1. 在终端显示旋转光标动画
        2. 显示剩余等待时间
        3. 在单独线程中运行以避免阻塞主程序
    """
    import threading
    import sys
    
    def spin():
        """动画旋转函数"""
        # 定义旋转字符集
        chars = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
        iterations = 0
        # 计算最大迭代次数（假设每秒10帧）
        max_iterations = seconds * 10
        
        # 循环显示动画字符
        while iterations < max_iterations:
            sys.stdout.write(f'\r{message} {chars[iterations % len(chars)]}')
            sys.stdout.flush()
            time.sleep(0.1)
            iterations += 1
    
    # 在单独的线程中运行动画，避免阻塞主程序
    spinner_thread = threading.Thread(target=spin)
    spinner_thread.start()
    
    # 等待指定的秒数
    time.sleep(seconds)
    
    # 等待动画线程结束
    spinner_thread.join()
    
    # 清除动画并换行，保持终端整洁
    sys.stdout.write('\r' + ' ' * (len(message) + 20) + '\r')
    sys.stdout.flush()
# --- 日志配置 ---
# 配置控制台日志记录器
console_logger = logging.getLogger('console_logger')
if not console_logger.handlers:
    # 设置日志级别为INFO
    console_logger.setLevel(logging.INFO)
    # 创建控制台处理器
    console_handler = logging.StreamHandler()
    # 设置日志格式：时间 - 日志级别 - 消息
    console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    # 将处理器添加到日志记录器
    console_logger.addHandler(console_handler)

# 配置文件日志记录器
file_logger = logging.getLogger('file_logger')
if not file_logger.handlers:
    # 设置日志级别为INFO
    file_logger.setLevel(logging.INFO)
    # 创建文件处理器，以写入模式打开，使用UTF-8编码
    file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
    # 设置日志格式：时间 - 消息
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    # 将处理器添加到日志记录器
    file_logger.addHandler(file_handler)

def get_email_from_gemini(company_name_en: str, company_name_tc: str) -> str:
    """
    调用 gemini-cli 获取公司联系邮箱地址。
    
    该函数通过构造特定提示词并调用Gemini API来搜索公司联系邮箱。
    实现了重试机制以处理网络波动或API临时错误。
    
    Args:
        company_name_en (str): 公司的英文名称
        company_name_tc (str): 公司的中文名称

    Returns:
        str: 邮箱地址或错误信息
             - 成功时返回邮箱地址
             - 未找到时返回 "Not Found"
             - 出错时返回 "Error: ..." 格式的错误信息

    Raises:
        QuotaExceededError: 当API配额用尽时抛出
    """
    # 使用公司名称填充提示词模板
    prompt = PROMPT_TEMPLATE.format(company_name=company_name_en, company_name_tc=company_name_tc)
    # 构造gemini-cli命令
    command = ['gemini', '-m', GEMINI_MODEL]
    
    # 实现重试机制
    for attempt in range(MAX_API_CALL_RETRIES):
        try:
            # 调用gemini-cli执行搜索任务
            result = subprocess.run(
                command,
                capture_output=True,  # 捕获标准输出和错误输出
                text=True,  # 以文本模式处理输出
                check=True,  # 如果返回非零退出码则抛出异常
                encoding='utf-8',  # 指定编码
                input=prompt,  # 将提示词作为输入传递
                timeout=GEMINI_TIMEOUT_SECONDS  # 设置超时时间
            )
            # 解析输出结果，取最后一行作为邮箱地址
            lines = result.stdout.strip().split('\n')
            return lines[-1].strip() if lines else "Error: No output"
        except subprocess.CalledProcessError as e:
            # 检测配额错误，需要特殊处理
            if "Quota exceeded" in e.stderr or "RESOURCE_EXHAUSTED" in e.stderr:
                raise QuotaExceededError("API配额已用尽") from e
            
            # 检测其他API错误，进行重试
            # 包括网络连接问题、Gemini错误和502错误
            if ("Gemini Error" in e.stderr or
                "Error 502" in e.stderr or
                "Client network socket disconnected" in e.stderr or
                "socket hang up" in e.stderr or
                "ECONNRESET" in e.stderr or
                "ETIMEDOUT" in e.stderr or
                "Premature close" in e.stderr or
                "API Error" in e.stderr):
                if attempt < MAX_API_CALL_RETRIES - 1:
                    # 记录警告信息和重试计划
                    console_logger.warning(f"API调用错误 (尝试 {attempt + 1}/{MAX_API_CALL_RETRIES}): {e.stderr.strip()}")
                    console_logger.info(f"等待 {API_RETRY_DELAY_SECONDS} 秒后进行第 {attempt + 2} 次重试...")
                    # 显示等待动画
                    spinning_cursor(API_RETRY_DELAY_SECONDS, f"等待 {API_RETRY_DELAY_SECONDS} 秒后进行第 {attempt + 2} 次重试...")
                    # 添加一个空行，避免动画被后续日志覆盖
                    console_logger.info("")
                    continue
                else:
                    # 达到最大重试次数，记录错误并返回
                    console_logger.error(f"Gemini Error: 达到最大重试次数 ({MAX_API_CALL_RETRIES} 次)，跳过该记录。原始错误: {e.stderr.strip()}")
                    return "Error: Gemini call failed"
            else:
                # 其他Gemini错误，直接记录并返回
                console_logger.error(f"Gemini Error: {e.stderr.strip()}")
                return "Error: Gemini call failed"
        except subprocess.TimeoutExpired:
            # 处理超时情况
            if attempt < MAX_API_CALL_RETRIES - 1:
                # 记录超时警告和重试计划
                console_logger.warning(f"Gemini调用超时 (尝试 {attempt + 1}/{MAX_API_CALL_RETRIES})")
                console_logger.info(f"等待 {API_RETRY_DELAY_SECONDS} 秒后进行第 {attempt + 2} 次重试...")
                # 显示等待动画
                spinning_cursor(API_RETRY_DELAY_SECONDS, f"等待 {API_RETRY_DELAY_SECONDS} 秒后进行第 {attempt + 2} 次重试...")
                # 添加一个空行，避免动画被后续日志覆盖
                console_logger.info("")
                continue
            else:
                # 达到最大重试次数，记录超时错误并返回
                console_logger.error(f"Gemini调用超时（{GEMINI_TIMEOUT_SECONDS}秒），达到最大重试次数 ({MAX_API_CALL_RETRIES} 次)，跳过该记录")
                return "Error: Timeout"
        except FileNotFoundError:
            # 处理gemini-cli未安装的情况
            console_logger.error("'gemini' 命令未找到。请确保 gemini-cli 已安装并位于系统的 PATH 中。")
            sys.exit(1)
    # 理论上不会执行到这里，但作为兜底返回
    return "Error: Unknown error after retries"

def main():
    """
    主函数，处理整个公司邮箱搜索流程。
    
    执行步骤：
    1. 读取Excel文件中的公司列表
    2. 显示当前处理进度和配置信息
    3. 提供交互式菜单选择处理模式
    4. 调用Gemini API逐个搜索公司邮箱
    5. 实时保存结果并生成日志报告
    """
    # 记录开始处理的信息
    console_logger.info("--- 开始处理 ---")
    
    # 打印当前配置信息，方便用户确认设置
    config_info = {
        "Excel 文件": EXCEL_FILE,
        "工作表名称": SHEET_NAME,
        "公司英文名列": COMPANY_NAME_EN_COL,
        "公司中文名列": COMPANY_NAME_TC_COL,
        "邮箱结果列": EMAIL_COL,
        "日志文件": LOG_FILE,
        "Gemini 模型": GEMINI_MODEL,
        "配额重试间隔 (分钟)": RETRY_INTERVAL_MINUTES,
        "任务间隔 (秒)": TASK_INTERVAL_SECONDS,
        "Gemini 调用超时 (秒)": GEMINI_TIMEOUT_SECONDS,
    }
    console_logger.info("\n--- 当前配置 ---")
    for key, value in config_info.items():
        console_logger.info(f"- {key}: {value}")
    console_logger.info("----------------\n")

    # 初始化变量
    df = None  # DataFrame对象，用于存储Excel数据
    tasks_to_process_indices = [] # 存储需要处理的任务索引列表
    current_task_number = 0       # 当前处理的任务编号
    total_tasks_for_run = 0       # 本次运行需要处理的总任务数
    
    try:
        # 读取Excel文件
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine='openpyxl')
        total_count = len(df)
        console_logger.info(f"成功读取文件 '{EXCEL_FILE}', 找到 {total_count} 条记录。")

        # 初始化Email列
        # 如果Excel文件中不存在Email列，则创建一个空的Email列
        if EMAIL_COL not in df.columns:
            df[EMAIL_COL] = ''

        # --- 统计当前文件状态 ---
        # 统计已成功处理的记录数（不包括错误和未找到的结果）
        initial_processed_success_count = df[df[EMAIL_COL].apply(lambda x: pd.notna(x) and x not in ['', 'Error: No output', 'Error: Gemini call failed', 'Not Found'])].shape[0]
        # 统计处理失败（未找到邮箱）的记录数
        initial_not_found_count = df[df[EMAIL_COL] == 'Not Found'].shape[0]
        # 统计尚未处理的记录数（包括空白和错误结果）
        initial_unprocessed_count = df[df[EMAIL_COL].isin(['', 'Error: No output', 'Error: Gemini call failed']) | pd.isna(df[EMAIL_COL])].shape[0]

        console_logger.info("\n--- 当前数据文件状态 ---")
        console_logger.info(f"  - 已处理成功: {initial_processed_success_count} 家")
        console_logger.info(f"  - 处理失败 (Not Found): {initial_not_found_count} 家")
        console_logger.info(f"  - 未处理 (空白或错误): {initial_unprocessed_count} 家")
        console_logger.info(f"  - 总计记录数: {total_count} 家")
        console_logger.info("------------------------\n")

        # --- 交互式菜单 ---
        # 初始化任务列表和任务计数器
        tasks_to_process_indices = []
        total_tasks_for_run = 0

        # 根据当前处理进度显示不同的菜单选项
        if initial_processed_success_count > 0 or initial_not_found_count > 0: # 只要有任何处理进度就显示菜单
            # 显示菜单选项
            print("\n检测到已有处理进度:")
            print("1. 继续上次任务（跳过已完成的记录，重试失败的记录）")
            print("2. 重新开始（清空所有结果）")
            print("3. 重试处理失败（Not Found）的记录")
            # 获取用户选择，如果直接回车则默认选择1
            choice = input("请选择操作 (默认1): ").strip() or "1"
            
            # 根据用户选择执行相应操作
            if choice == "2":
                # 重新开始：清空所有结果并处理所有记录
                df[EMAIL_COL] = ''
                console_logger.info("已清空所有结果，重新开始处理。")
                tasks_to_process_indices = df.index.tolist()
                total_tasks_for_run = len(tasks_to_process_indices)
            elif choice == "3":
                # 重试失败记录：只处理标记为"Not Found"的记录
                tasks_to_process_indices = df[df[EMAIL_COL] == 'Not Found'].index.tolist()
                total_tasks_for_run = len(tasks_to_process_indices)
                console_logger.info("将重试处理失败 (Not Found) 的记录。")
            else: # 默认或选择1
                # 继续上次任务：重置错误状态记录并继续处理
                # 重置错误状态以便重试，同时统计本次要处理的数量
                error_mask = df[EMAIL_COL].isin(['Error: No output', 'Error: Gemini call failed'])
                df.loc[error_mask, EMAIL_COL] = ''
                console_logger.info("已重置错误状态记录，将继续处理。")
                # 重新计算本次要处理的数量
                tasks_to_process_indices = df[df[EMAIL_COL].isin(['', 'Error: No output', 'Error: Gemini call failed']) | pd.isna(df[EMAIL_COL])].index.tolist()
                total_tasks_for_run = len(tasks_to_process_indices)
        else: # 没有处理进度，直接处理所有未处理的
            # 没有任何处理进度时，处理所有未处理的记录
            tasks_to_process_indices = df[df[EMAIL_COL].isin(['', 'Error: No output', 'Error: Gemini call failed']) | pd.isna(df[EMAIL_COL])].index.tolist()
            total_tasks_for_run = len(tasks_to_process_indices)

        # 重置文件日志
        # 移除旧的文件处理器（如果有），确保日志文件不会重复写入
        for handler in file_logger.handlers[:]:
            if isinstance(handler, logging.FileHandler):
                file_logger.removeHandler(handler)

        # 配置文件日志，以追加模式写入
        # 创建新的文件处理器，以追加模式打开，使用UTF-8编码
        file_handler = logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8')
        # 设置日志格式：时间 - 消息
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        # 将处理器添加到日志记录器
        file_logger.addHandler(file_handler)

        # 记录本次任务启动时间
        start_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        file_logger.info("\n" + "="*70)
        file_logger.info(f"任务启动时间: {start_time}")
        file_logger.info("="*70)

        # 处理数据
        # 初始化计数器
        not_found_count = 0  # 未找到邮箱的记录数
        success_count = 0    # 成功找到邮箱的记录数

        # 遍历所有需要处理的任务索引
        for index in tasks_to_process_indices:
            # 更新当前任务编号
            current_task_number += 1

            # 获取公司名称（英文和中文）
            company_en = str(df.at[index, COMPANY_NAME_EN_COL]) if COMPANY_NAME_EN_COL in df.columns and pd.notna(df.at[index, COMPANY_NAME_EN_COL]) else ''
            company_tc = str(df.at[index, COMPANY_NAME_TC_COL]) if COMPANY_NAME_TC_COL in df.columns and pd.notna(df.at[index, COMPANY_NAME_TC_COL]) else ''
            
            # 确定当前公司名称（优先使用英文名，如果没有则使用中文名）
            current_company = company_en or company_tc
            # 如果公司名称为空，则跳过该记录
            if not current_company:
                console_logger.info(f"[{current_task_number}/{total_tasks_for_run}] 跳过空行...")
                continue

            # 构造显示名称（中英文结合）
            display_name = ""
            if company_en and company_tc:
                display_name = f"{company_en} ({company_tc})"
            elif company_en:
                display_name = company_en
            elif company_tc:
                display_name = company_tc
            
            # 记录正在处理的公司信息
            console_logger.info(f"[{current_task_number}/{total_tasks_for_run}] 正在处理: {display_name}")
            
            # 调用Gemini API获取邮箱
            try:
                email = get_email_from_gemini(company_en, company_tc)
            except QuotaExceededError as e:
                # 处理API配额用尽的情况
                console_logger.error(f"错误：API每日配额已用尽！将在 {RETRY_INTERVAL_MINUTES} 分钟后重试...")
                console_logger.error("您可以：")
                console_logger.error("1. 等待自动重试")
                console_logger.error("2. 手动中断程序（Ctrl+C）并稍后重新运行")
                
                # 自动重试逻辑
                while True:
                    # 显示等待动画
                    spinning_cursor(RETRY_INTERVAL_MINUTES * 60, f"配额错误等待中，等待 {RETRY_INTERVAL_MINUTES} 分钟...")
                    # 添加一个空行，避免动画被后续日志覆盖
                    console_logger.info("")
                    console_logger.info(f"重试中... ({RETRY_INTERVAL_MINUTES}分钟间隔)")
                    try:
                        # 重试获取邮箱
                        email = get_email_from_gemini(company_en, company_tc)
                        console_logger.info("配额已恢复，继续处理！")
                        break
                    except QuotaExceededError:
                        # 如果配额仍未恢复，继续等待
                        console_logger.error(f"配额仍未恢复，{RETRY_INTERVAL_MINUTES}分钟后再次重试...")
                        continue
                
            # 记录处理结果
            console_logger.info(f"  -> 结果: {email}")
            # 处理错误结果
            if email.startswith("Error:"):
                # 区分API调用错误和超时错误
                if "Error: Timeout" in email:
                    console_logger.warning("API调用超时，跳过该记录")
                else:
                    console_logger.warning("发生API调用错误，跳过该记录")
            else:
                # 保存成功结果
                df.at[index, EMAIL_COL] = email
            
            # 更新计数器
            if email == "Not Found":
                # 未找到邮箱
                not_found_count += 1
                file_logger.info(f"未找到邮箱: {current_company}")
            elif not email.startswith("Error:"): # 仅当不是错误结果时才计入成功
                # 成功找到邮箱
                success_count += 1
            
            # 实时保存结果到Excel文件
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
            # 显示任务间隔等待动画
            spinning_cursor(TASK_INTERVAL_SECONDS, f"任务间隔中，等待 {TASK_INTERVAL_SECONDS} 秒...")
            # 添加一个空行，避免动画被后续日志覆盖
            console_logger.info("")

        # 最终报告
        # 显示处理完成信息
        console_logger.info("--- 全部处理完成！最终结果已在文件中。 ---")
        
        # 记录任务结束时间和处理结果汇总
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
        # 处理文件未找到的错误
        console_logger.error(f"错误：文件 '{EXCEL_FILE}' 未找到。请确保文件在正确的路径下。")
        sys.exit(1)
    except KeyboardInterrupt:
        # 处理用户中断操作（Ctrl+C）
        console_logger.info("\n用户中断操作，已保存当前进度。")
        if df is not None:
            # 保存当前进度到Excel文件
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        sys.exit(0)
    except Exception as e:
        # 处理其他未预期的错误
        console_logger.error(f"发生未知错误: {e}")
        if df is not None:
            # 保存当前进度到Excel文件
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        else:
            console_logger.error("错误发生时尚未加载数据文件")

if __name__ == '__main__':
    main()