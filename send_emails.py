#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
保研自适应个性化邮件发送工具 (Antigravity Designed)
- 自动调用 DeepSeek API 根据导师信息补充模板中的学术研究背景
- 自动读取 teachers.csv 中的导师列表并支持增量发送
- 自动将生成的邮件内容导出至 out_emails/ 目录下作为本地备份
- 支持 163 邮箱 SMTP 自动发送，带简历附件
"""

import os
import sys
import re
import csv
import json
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
import requests

# 尝试导入 python-dotenv，如果不存在则使用手动解析
try:
    from dotenv import load_dotenv
    HAS_DOTENV = True
except ImportError:
    HAS_DOTENV = False

# ----------------- 配置加载与初始化 -----------------

ENV_FILE_PATH = ".env"
TEMPLATE_FILE_PATH = "模板.md"
TEACHERS_CSV_PATH = "teachers.csv"
OUTPUT_DIR = "out_emails"

def load_env_variables():
    """加载 .env 配置文件"""
    if HAS_DOTENV:
        load_dotenv(ENV_FILE_PATH)
    else:
        # 手动解析 .env 文件作为备用
        if os.path.exists(ENV_FILE_PATH):
            with open(ENV_FILE_PATH, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, val = line.split('=', 1)
                        os.environ[key.strip()] = val.strip()

def save_api_key_to_env(api_key):
    """保存用户输入的 API key 到 .env 文件"""
    content = ""
    if os.path.exists(ENV_FILE_PATH):
        with open(ENV_FILE_PATH, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 替换已有或空白的 DEEPSEEK_API_KEY
        if "DEEPSEEK_API_KEY=" in content:
            content = re.sub(r'DEEPSEEK_API_KEY=.*', f'DEEPSEEK_API_KEY={api_key}', content)
        else:
            content += f"\nDEEPSEEK_API_KEY={api_key}\n"
    else:
        # 创建新的 .env 文件并填入默认配置
        content = f"""# DeepSeek API 密钥
DEEPSEEK_API_KEY={api_key}

# DeepSeek API 接口地址
DEEPSEEK_API_URL=https://api.deepseek.com/v1
DEEPSEEK_MODEL=deepseek-chat

# 邮箱 SMTP 配置 (已根据您的输入预填)
SMTP_SERVER=smtp.163.com
SMTP_PORT=465
SMTP_USER=your_email@example.com
SMTP_PASS=your_smtp_auth_code_here

# 默认简历附件路径
DEFAULT_ATTACHMENT_PATH=resume.pdf

# 发件人个人信息 (用于自动替换自荐信和邮件标题)
SENDER_NAME=您的姓名
SENDER_UNIVERSITY=您的本科学校
EMAIL_SUBJECT=预推免自荐信-{{SENDER_UNIVERSITY}}-{{SENDER_NAME}}
"""
    with open(ENV_FILE_PATH, 'w', encoding='utf-8') as f:
        f.write(content)
    # 重新加载
    if HAS_DOTENV:
        load_dotenv(ENV_FILE_PATH, override=True)
    else:
        os.environ["DEEPSEEK_API_KEY"] = api_key

def check_and_get_config():
    """获取程序运行所需配置"""
    load_env_variables()
    
    api_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        print("="*60)
        print("提示: 未检测到 DEEPSEEK_API_KEY。")
        print("请输入您的 DeepSeek API Key (回车确认，输入后将自动保存至 .env 配置文件中):")
        api_key = input(">> ").strip()
        while not api_key:
            print("API Key 不能为空，请重新输入:")
            api_key = input(">> ").strip()
        save_api_key_to_env(api_key)
        print("API Key 已成功保存至 .env 文件。")
        print("="*60)
    
    config = {
        "api_key": os.environ.get("DEEPSEEK_API_KEY", api_key),
        "api_url": os.environ.get("DEEPSEEK_API_URL", "https://api.deepseek.com/v1").strip(),
        "model": os.environ.get("DEEPSEEK_MODEL", "deepseek-chat").strip(),
        "smtp_server": os.environ.get("SMTP_SERVER", "smtp.163.com").strip(),
        "smtp_port": int(os.environ.get("SMTP_PORT", "465")),
        "smtp_user": os.environ.get("SMTP_USER", "your_email@example.com").strip(),
        "smtp_pass": os.environ.get("SMTP_PASS", "your_smtp_auth_code_here").strip(),
        "attachment": os.environ.get("DEFAULT_ATTACHMENT_PATH", "resume.pdf").strip(),
        "sender_name": os.environ.get("SENDER_NAME", "您的姓名").strip(),
        "sender_university": os.environ.get("SENDER_UNIVERSITY", "您的本科学校").strip(),
        "email_subject": os.environ.get("EMAIL_SUBJECT", "预推免自荐信-{SENDER_UNIVERSITY}-{SENDER_NAME}").strip(),
    }
    return config

# ----------------- 文件与数据读取 -----------------

def read_template():
    """读取模板.md文件"""
    if not os.path.exists(TEMPLATE_FILE_PATH):
        print(f"错误: 模板文件 '{TEMPLATE_FILE_PATH}' 不存在！请创建该文件后再运行本脚本。")
        sys.exit(1)
    with open(TEMPLATE_FILE_PATH, 'r', encoding='utf-8') as f:
        return f.read()

def read_teachers_csv():
    """兼容多种编码读取 teachers.csv"""
    if not os.path.exists(TEACHERS_CSV_PATH):
        print(f"提示: 未找到 '{TEACHERS_CSV_PATH}'。已自动为您创建包含示例数据的模板。")
        create_sample_csv()
        print("请在 teachers.csv 中填入您要发信的导师信息，然后重新运行本脚本。")
        sys.exit(0)
        
    encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312']
    for enc in encodings:
        try:
            with open(TEACHERS_CSV_PATH, 'r', encoding=enc) as f:
                reader = csv.DictReader(f)
                # 校验表头
                headers = reader.fieldnames
                required = ['姓名', '学校', '学院', '邮箱', '研究方向简述', '状态']
                # 兼容表头可能有空格的情况
                cleaned_headers = [h.strip() for h in headers] if headers else []
                
                missing = [req for req in required if req not in cleaned_headers]
                if missing:
                    continue  # 表头不匹配，尝试其他编码或继续抛出错误
                
                # 转换为标准字典列表
                teachers = []
                for row in reader:
                    teacher = {k.strip(): v.strip() for k, v in row.items() if k}
                    teachers.append(teacher)
                return teachers, enc
        except (UnicodeDecodeError, ValueError):
            continue
            
    print("错误: 无法解析 'teachers.csv'。请确保文件是 CSV 格式，且表头包含: 姓名,学校,学院,邮箱,研究方向简述,状态")
    sys.exit(1)

def write_teachers_csv(teachers, encoding):
    """写回更新后的导师状态到 CSV"""
    with open(TEACHERS_CSV_PATH, 'w', encoding=encoding, newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['姓名', '学校', '学院', '邮箱', '研究方向简述', '状态'])
        writer.writeheader()
        writer.writerows(teachers)

def create_sample_csv():
    """生成样例 csv 文件"""
    with open(TEACHERS_CSV_PATH, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['姓名', '学校', '学院', '邮箱', '研究方向简述', '状态'])
        writer.writeheader()
        writer.writerow({
            "姓名": "张老师",
            "学校": "北京航空航天大学",
            "学院": "仪器科学与光电工程学院",
            "邮箱": "teacher_zhang@example.edu.cn",
            "研究方向简述": "惯性导航 / 多传感器组合导航",
            "状态": "待发送"
        })
        writer.writerow({
            "姓名": "王老师",
            "学校": "西北工业大学",
            "学院": "自动化学院",
            "邮箱": "teacher_wang@example.edu.cn",
            "研究方向简述": "自适应卡尔曼滤波 / 智能控制",
            "状态": "待发送"
        })

# ----------------- DeepSeek API 交互与文本生成 -----------------

def clean_and_parse_json(text):
    """去除 Markdown 格式的多余字符并解析 JSON"""
    text = text.strip()
    # 匹配最外层的 { 到 } 之间的内容
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        text = match.group(0)
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        print(f"JSON 解析失败! 原始文本: \n{text}")
        raise e

def query_deepseek_teacher_info(config, name, university, school, research_focus=""):
    """调用 DeepSeek 补充导师研究背景"""
    headers = {
        "Authorization": f"Bearer {config['api_key']}",
        "Content-Type": "application/json"
    }
    
    system_prompt = (
        "你是一个专业的学术助手，专门帮助保研学生研究导师的学术背景。请根据给出的导师姓名、学校、学院以及研究方向（如有），"
        "从你的知识库或事实推理中总结出符合该导师真实学术背景的各项个性化信息。\n"
        "你必须严格按照指定的 JSON 格式输出，不要包含任何额外的解释或 markdown 格式，只返回一个 JSON 对象。"
    )
    
    user_prompt = f"""
导师基本信息：
- 姓名：{name}
- 学校：{university}
- 学院：{school}
- 研究方向/细分领域（参考）：{research_focus}

请为该导师总结并生成以下字段（均使用简体中文，语气严谨专业且学术化）：
1. "teacher_name": 导师姓名（通常与输入一致）。
2. "university": 学校名称（通常与输入一致）。
3. "school": 学院名称（通常与输入一致）。
4. "research_achievement": 该导师的一个代表性学术成就或核心成果总结（约10-25个字，如“多源协同导航与自适应滤波理论”、“深空探测自主控制与鲁棒估计方法”等）。
5. "research_direction": 该导师的代表性科研方向（约10-25个字，需与卡尔曼滤波、多传感器融合、捷联惯导、状态估计等导航控制领域契合，以便与学生背景契合，如“高精度多源融合导航与容错状态估计”、“无人自主系统自适应卡尔曼滤波方法”等）。
6. "project_1": 该导师课题组正在开展或代表性的前沿课题/项目名称1（约10-20个字，如“协同自主导航定位技术”、“非高斯噪声下自适应滤波算法”等）。
7. "project_2": 该导师课题组正在开展或代表性的前沿课题/项目名称2（约10-20个字，如“深海潜器多传感器融合定位”、“多智能体协同感知与控制”等）。

输出格式要求：
必须且仅能输出如下格式的 JSON 字符串，不能有任何其他文字（例如不要带 ```json 等标记，直接以 {{ 开始）：
{{
  "teacher_name": "...",
  "university": "...",
  "school": "...",
  "research_achievement": "...",
  "research_direction": "...",
  "project_1": "...",
  "project_2": "..."
}}
"""

    payload = {
        "model": config["model"],
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"}
    }
    
    url = f"{config['api_url'].rstrip('/')}/chat/completions"
    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()
    
    result = response.json()
    content = result['choices'][0]['message']['content']
    return clean_and_parse_json(content)

def generate_email_content(template, info):
    """用 API 获得的信息填充模板.md 中的 XXX"""
    content = template
    
    # 获取今日日期
    now = datetime.now()
    today_str = f"{now.year}年{now.month}月{now.day}日"
    
    # 按照严格匹配的占位符进行替换
    content = content.replace("尊敬的XXX老师：", f"尊敬的{info['teacher_name']}老师：")
    content = content.replace("近期在关注XXX大学XXX学院的导师信息时", f"近期在关注{info['university']}{info['school']}的导师信息时")
    content = content.replace("我详细阅读了您在XXX的研究成果", f"我详细阅读了您在{info['research_achievement']}的研究成果")
    content = content.replace("XXX大学深厚的数理底蕴和学术氛围令我由衷向往", f"{info['university']}深厚的数理底蕴和学术氛围令我由衷向往")
    content = content.replace("而您在XXX方向的深耕", f"而您在{info['research_direction']}方向的深耕")
    content = content.replace("我了解到您在开展XXX与XXX等前沿课题时", f"我了解到您在开展{info['project_1']}与{info['project_2']}等前沿课题时")
    content = content.replace("XXX大学XXX学院在复杂系统控制与智能导航领域拥有国际一流的科研平台", f"{info['university']}{info['school']}在复杂系统控制与智能导航领域拥有国际一流的科研平台")
    content = content.replace("2026年X月XX日", today_str)
    
    return content

# ----------------- SMTP 邮件发送 -----------------

def send_email_via_smtp(config, to_email, subject, body_text):
    """通过 163 邮箱发送包含附件的邮件"""
    # 检查附件
    attachment_path = config["attachment"]
    has_attachment = os.path.exists(attachment_path)
    
    # 创建邮件容器
    msg = MIMEMultipart('alternative')
    sender_name = config.get("sender_name", "您的姓名").strip()
    msg['From'] = f"{sender_name} <{config['smtp_user']}>"
    msg['To'] = to_email
    msg['Subject'] = subject
    
    # 将 markdown 文本包装为 HTML，以便在邮箱中保留格式与换行
    # 同时将加粗的 ** 渲染为 strong 标签
    html_content = body_text
    html_content = html_content.replace('\n', '<br>')
    html_content = re.sub(
        r'\*\*(.*?)\*\*', 
        r'<strong style="color: #1a365d; font-weight: bold;">\1</strong>', 
        html_content
    )
    
    html_body = f"""
    <html>
    <body>
        <div style="font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif; font-size: 15px; color: #2d3748; line-height: 1.8; max-width: 750px; margin: 0 auto; padding: 10px;">
            {html_content}
        </div>
    </body>
    </html>
    """
    
    # 附加纯文本与 HTML 两种格式的邮件正文
    part_text = MIMEText(body_text, 'plain', 'utf-8')
    part_html = MIMEText(html_body, 'html', 'utf-8')
    msg.attach(part_text)
    msg.attach(part_html)
    
    # 构建最终的混合格式邮件以支持附件
    outer_msg = MIMEMultipart('mixed')
    # 将正文部分（含 text/html 的 alternative 容器）复制到 mixed 容器中
    for part in msg.get_payload():
        outer_msg.attach(part)
        
    # 复制邮件头
    outer_msg['From'] = msg['From']
    outer_msg['To'] = msg['To']
    outer_msg['Subject'] = msg['Subject']
    
    # 处理附件
    if has_attachment:
        filename = os.path.basename(attachment_path)
        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                
                # 对中文文件名进行编码以防止乱码
                encoded_filename = Header(filename, 'utf-8').encode()
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename=\"{encoded_filename}\"",
                )
                outer_msg.attach(part)
                print(f" -> 成功添加附件: {filename}")
        except Exception as e:
            print(f" -> 警告: 附件添加失败 ({e})，将尝试不带附件发送。")
    else:
        print(f" -> 提示: 未在目录下找到简历文件 '{attachment_path}'，将不带附件发送。")
        
    # 建立 SSL 连接并发送
    try:
        server = smtplib.SMTP_SSL(config["smtp_server"], config["smtp_port"], timeout=15)
        server.login(config["smtp_user"], config["smtp_pass"])
        server.sendmail(config["smtp_user"], to_email, outer_msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f" -> 邮件发送异常: {e}")
        return False

# ----------------- 业务逻辑与主流程 -----------------

def create_output_directory():
    """创建本地备份目录"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

def run_generate_previews(config, template, teachers):
    """一键生成所有导师的个性化邮件预览"""
    create_output_directory()
    count = 0
    print("\n" + "="*50)
    print("开始生成个性化邮件预览...")
    print("="*50)
    
    for idx, t in enumerate(teachers, 1):
        name = t.get('姓名', '').strip()
        univ = t.get('学校', '').strip()
        school = t.get('学院', '').strip()
        focus = t.get('研究方向简述', '').strip()
        
        if not name or not univ or not school:
            print(f"第 {idx} 行数据不完整，跳过。")
            continue
            
        print(f"\n[{idx}/{len(teachers)}] 正在调用 DeepSeek 分析 {univ} - {name} 老师的学术背景...")
        try:
            info = query_deepseek_teacher_info(config, name, univ, school, focus)
            email_content = generate_email_content(template, info)
            
            # 保存到本地
            filename = f"{config.get('sender_name', '自荐信')}_致_{univ}_{school}_{name}老师.md"
            filepath = os.path.join(OUTPUT_DIR, filename)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(email_content)
            
            print(f" -> 预览生成成功！已保存至 {filepath}")
            print(f" -> 匹配成果: {info['research_achievement']}")
            print(f" -> 匹配方向: {info['research_direction']}")
            count += 1
        except Exception as e:
            print(f" -> 生成失败: {e}")
            
    print("\n" + "="*50)
    print(f"预览生成结束！共成功生成 {count} 封邮件。")
    print(f"您可在 '{OUTPUT_DIR}' 目录下查看生成的 .md 格式个性化邮件。")
    print("="*50)

def run_interactive_send(config, template, teachers, csv_encoding):
    """互动确认发送模式 (预览并逐个确认发送)"""
    create_output_directory()
    
    print("\n" + "="*50)
    print("进入互动确认发送模式...")
    print("="*50)
    
    pending_teachers = [t for t in teachers if t.get('状态', '').strip() != '已发送']
    if not pending_teachers:
        print("所有导师邮件均已发送（状态为“已发送”），无需处理。")
        return
        
    # 检查简历附件
    attachment_path = config["attachment"]
    if not os.path.exists(attachment_path):
        print(f"⚠️ 警告: 未在当前目录下找到您的简历附件 '{attachment_path}'！")
        print(f"建议您将 PDF 简历命名为“{attachment_path}”放入此目录，或在 .env 中修改路径。")
        confirm = input("是否在【没有附件】的情况下继续？(y/n): ").strip().lower()
        if confirm != 'y':
            print("程序已终止。请准备好您的简历附件后再来。")
            return
            
    send_all_remaining = False
    
    for idx, t in enumerate(teachers, 1):
        name = t.get('姓名', '').strip()
        univ = t.get('学校', '').strip()
        school = t.get('学院', '').strip()
        email = t.get('邮箱', '').strip()
        focus = t.get('研究方向简述', '').strip()
        status = t.get('状态', '').strip()
        
        if status == '已发送':
            continue
            
        if not name or not univ or not school or not email:
            print(f"\n第 {idx} 行数据不完整，跳过。")
            continue
            
        print(f"\n" + "-"*40)
        print(f"正在处理 [{idx}/{len(teachers)}]: {univ} - {school} - {name} 老师 ({email})")
        print(f"参考方向: {focus}")
        
        # 1. 尝试从本地 out_emails 中读取，如果已经生成过就不再重复调用 API 消耗 token
        local_filename = f"{config.get('sender_name', '自荐信')}_致_{univ}_{school}_{name}老师.md"
        local_filepath = os.path.join(OUTPUT_DIR, local_filename)
        email_content = ""
        
        if os.path.exists(local_filepath):
            print(f" -> 检测到本地已存在预览文件，正在读取...")
            with open(local_filepath, 'r', encoding='utf-8') as f:
                email_content = f.read()
        else:
            print(f" -> 正在调用 DeepSeek API 分析导师背景并填充模板...")
            try:
                info = query_deepseek_teacher_info(config, name, univ, school, focus)
                email_content = generate_email_content(template, info)
                # 保存一份在本地
                if not os.path.exists(OUTPUT_DIR):
                    os.makedirs(OUTPUT_DIR)
                with open(local_filepath, 'w', encoding='utf-8') as f:
                    f.write(email_content)
                print(f" -> 个性化填充完成！本地预览已保存。")
            except Exception as e:
                print(f" -> 调用 API 失败: {e}，跳过此导师。")
                continue
        
        # 打印部分内容作为预览
        print("\n--- 邮件内容预览 (前 8 行) ---")
        lines = email_content.split('\n')
        for line in lines[:10]:
            print(line)
        print("...... (其余内容略)\n" + "-"*40)
        
        # 2. 确认发送
        if not send_all_remaining:
            print(f"请确认是否发送邮件给 {name} 老师？")
            print(" [y] 发送这封邮件")
            print(" [n] 暂不发送 (跳过此导师)")
            print(" [a] 全部发送 (后续导师将自动静默发送，不再询问)")
            print(" [q] 退出程序")
            choice = input("请输入选项 [y/n/a/q]: ").strip().lower()
            
            if choice == 'q':
                print("程序已主动退出。")
                break
            elif choice == 'n':
                print(f"跳过 {name} 老师。")
                continue
            elif choice == 'a':
                send_all_remaining = True
                print("已切换至【批量自动发送】模式...")
            elif choice != 'y':
                print("输入无效，默认跳过此导师. ")
                continue
                
        # 3. 发送邮件
        subject_template = config.get("email_subject", "预推免自荐信-{SENDER_UNIVERSITY}-{SENDER_NAME}")
        try:
            subject = subject_template.format(
                SENDER_UNIVERSITY=config.get("sender_university", "您的本科学校"),
                SENDER_NAME=config.get("sender_name", "您的姓名")
            )
        except Exception:
            subject = f"自荐信 - {config.get('sender_university', '您的本科学校')} - {config.get('sender_name', '您的姓名')}"
        print(f" 正在以主题「{subject}」向 {email} 发送邮件...")
        success = send_email_via_smtp(config, email, subject, email_content)
        
        if success:
            print(f" 🎉 发送成功！已向 {name} 老师发送了申请邮件。")
            t['状态'] = '已发送'
            # 实时更新 CSV 状态，防止中途异常退出后无法记录
            write_teachers_csv(teachers, csv_encoding)
        else:
            print(f" ❌ 发送失败，请检查您的 SMTP 配置、网络连接或邮箱限制。")
            
    print("\n" + "="*50)
    print("互动发送流程结束！")
    print("="*50)

def run_silent_batch_send(config, template, teachers, csv_encoding):
    """静默批量发送模式"""
    create_output_directory()
    
    print("\n" + "="*50)
    print("进入静默批量发送模式...")
    print("="*50)
    
    pending_teachers = [t for t in teachers if t.get('状态', '').strip() != '已发送']
    if not pending_teachers:
        print("所有导师邮件均已发送（状态为“已发送”），无需处理。")
        return
        
    print(f"检测到共有 {len(pending_teachers)} 封邮件等待发送。")
    attachment_path = config["attachment"]
    if not os.path.exists(attachment_path):
        print(f"❌ 错误: 未在当前目录下找到您的简历附件 '{attachment_path}'！")
        print("为了安全起见，静默发送模式已拒绝执行。请先准备好您的简历 PDF。")
        return
        
    confirm = input("⚠️ 您确定要无确认静默批量发送这些邮件吗？该操作不可逆！(yes/no): ").strip().lower()
    if confirm != 'yes':
        print("已取消静默批量发送。")
        return
        
    for idx, t in enumerate(teachers, 1):
        name = t.get('姓名', '').strip()
        univ = t.get('学校', '').strip()
        school = t.get('学院', '').strip()
        email = t.get('邮箱', '').strip()
        focus = t.get('研究方向简述', '').strip()
        status = t.get('状态', '').strip()
        
        if status == '已发送':
            continue
            
        if not name or not univ or not school or not email:
            continue
            
        print(f"\n正在处理 [{idx}/{len(teachers)}]: {univ} - {name} 老师")
        
        # 获取邮件内容
        local_filename = f"{config.get('sender_name', '自荐信')}_致_{univ}_{school}_{name}老师.md"
        local_filepath = os.path.join(OUTPUT_DIR, local_filename)
        email_content = ""
        
        if os.path.exists(local_filepath):
            with open(local_filepath, 'r', encoding='utf-8') as f:
                email_content = f.read()
        else:
            try:
                info = query_deepseek_teacher_info(config, name, univ, school, focus)
                email_content = generate_email_content(template, info)
                with open(local_filepath, 'w', encoding='utf-8') as f:
                    f.write(email_content)
            except Exception as e:
                print(f" -> 填充失败 ({e})，跳过该导师。")
                continue
                
        subject_template = config.get("email_subject", "预推免自荐信-{SENDER_UNIVERSITY}-{SENDER_NAME}")
        try:
            subject = subject_template.format(
                SENDER_UNIVERSITY=config.get("sender_university", "您的本科学校"),
                SENDER_NAME=config.get("sender_name", "您的姓名")
            )
        except Exception:
            subject = f"自荐信 - {config.get('sender_university', '您的本科学校')} - {config.get('sender_name', '您的姓名')}"
        print(f" 正在以主题「{subject}」向 {email} 发送邮件...")
        success = send_email_via_smtp(config, email, subject, email_content)
        
        if success:
            print(f" 🎉 发送成功！")
            t['状态'] = '已发送'
            write_teachers_csv(teachers, csv_encoding)
            # 适当延时，防止触发邮箱高频垃圾邮件策略
            import time
            time.sleep(5)
        else:
            print(f" ❌ 发送失败！")
            
    print("\n" + "="*50)
    print("批量静默发送流程结束！")
    print("="*50)

def main():
    # 检测是否使用命令行模式
    if "--cli" not in sys.argv:
        try:
            # 尝试拉起 CustomTkinter GUI
            import customtkinter
            from send_emails_gui import EmailSenderApp
            print("正在为您拉起极具现代感与学术审美的 GUI 可视化窗口...")
            app = EmailSenderApp()
            app.mainloop()
            return
        except ImportError:
            print("💡 提示: 未检测到 customtkinter 界面库，已为您切换至极速命令行交互模式。")
            print(" (如果您想使用现代可视化界面，可运行: pip install customtkinter)")
        except Exception as e:
            print(f"💡 提示: 无法拉起 GUI 界面 ({e})，已自动切换至命令行模式。")

    print("""
================================================================
      🏫 保研/考研推免邮件个性化自动发送工具 (GitHub 开源版) 🏫
            (基于 DeepSeek API 与 163 邮箱 SMTP 服务)
================================================================
    """)
    
    # 1. 检查和加载配置
    config = check_and_get_config()
    
    # 2. 读取模板.md
    template = read_template()
    
    # 3. 读取 teachers.csv
    teachers, csv_encoding = read_teachers_csv()
    print(f"已成功加载导师信息表 ({csv_encoding} 编码)，共检测到 {len(teachers)} 位导师。")
    
    while True:
        print("\n请选择您要执行的操作:")
        print(" [1] 一键生成所有导师的个性化邮件 (仅在本地预览/不发送)")
        print(" [2] 互动确认发送模式 (推荐，逐个预览并人工确认发送)")
        print(" [3] 静默批量发送模式 (极速自动发送，需双重确认)")
        print(" [4] 退出程序")
        
        choice = input("请输入选项数字 [1/2/3/4]: ").strip()
        
        if choice == '1':
            run_generate_previews(config, template, teachers)
        elif choice == '2':
            run_interactive_send(config, template, teachers, csv_encoding)
            # 重新加载 CSV 状态以防后续再次运行
            teachers, csv_encoding = read_teachers_csv()
        elif choice == '3':
            run_silent_batch_send(config, template, teachers, csv_encoding)
            # 重新加载 CSV 状态
            teachers, csv_encoding = read_teachers_csv()
        elif choice == '4':
            print("感谢使用，祝您保研顺利，科研长青！")
            break
        else:
            print("无效选项，请重新选择。")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n用户已强行终止程序。祝您保研顺利！")
        sys.exit(0)
