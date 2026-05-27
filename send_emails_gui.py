#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
哈尔滨工程大学 保研个性化邮件自动发送工具 GUI 版本 (Antigravity Designed)
- 基于 CustomTkinter 打造的极具现代感与学术审美的 GUI 界面
- 多线程后台运行，不卡顿界面，实时滚动日志输出
- 配置中心：支持 DeepSeek API 密钥、SMTP 邮箱设置、以及文件弹窗选择简历附件
- 导师管理：支持可视化增删查改 teachers.csv 导师数据列表
- 模板编辑：支持直接在软件内修改与保存 模板.md 内容
- 批量处理：提供“一键本地预览”与“批量安全发送”功能，带进度条和统计看板
"""

import os
import sys
import re
import csv
import json
import smtplib
import threading
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from tkinter import filedialog, messagebox
import requests

try:
    import customtkinter as ctk
except ImportError:
    print("错误: 缺少 customtkinter 库！请运行 'pip install customtkinter' 安装。")
    sys.exit(1)

# 设置主题与色彩系统
ctk.set_appearance_mode("System")  # 自动匹配 Windows 系统的深色/浅色模式
ctk.set_default_color_theme("blue")  # 经典深蓝学术主题

# ----------------- 数据交互与工具函数 -----------------

ENV_FILE_PATH = ".env"
TEMPLATE_FILE_PATH = "模板.md"
TEACHERS_CSV_PATH = "teachers.csv"
OUTPUT_DIR = "out_emails"

def load_config():
    """读取并解析 .env 配置文件"""
    config = {
        "DEEPSEEK_API_KEY": "",
        "DEEPSEEK_API_URL": "https://api.deepseek.com/v1",
        "DEEPSEEK_MODEL": "deepseek-chat",
        "SMTP_SERVER": "smtp.163.com",
        "SMTP_PORT": "465",
        "SMTP_USER": "your_email@example.com",
        "SMTP_PASS": "your_smtp_auth_code_here",
        "DEFAULT_ATTACHMENT_PATH": "resume.pdf",
        "SENDER_NAME": "您的姓名",
        "SENDER_UNIVERSITY": "您的本科学校",
        "EMAIL_SUBJECT": "预推免自荐信-{SENDER_UNIVERSITY}-{SENDER_NAME}"
    }
    if os.path.exists(ENV_FILE_PATH):
        with open(ENV_FILE_PATH, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, val = line.split('=', 1)
                    config[key.strip()] = val.strip()
    return config

def save_config(config):
    """保存配置至 .env"""
    content = f"""# DeepSeek API 密钥
DEEPSEEK_API_KEY={config.get('DEEPSEEK_API_KEY', '')}

# DeepSeek API 接口地址
DEEPSEEK_API_URL={config.get('DEEPSEEK_API_URL', 'https://api.deepseek.com/v1')}
DEEPSEEK_MODEL={config.get('DEEPSEEK_MODEL', 'deepseek-chat')}

# 邮箱 SMTP 配置 (已根据您的输入预填)
SMTP_SERVER={config.get('SMTP_SERVER', 'smtp.163.com')}
SMTP_PORT={config.get('SMTP_PORT', '465')}
SMTP_USER={config.get('SMTP_USER', 'your_email@example.com')}
SMTP_PASS={config.get('SMTP_PASS', 'your_smtp_auth_code_here')}

# 默认简历附件路径
DEFAULT_ATTACHMENT_PATH={config.get('DEFAULT_ATTACHMENT_PATH', 'resume.pdf')}

# 发件人个人信息 (用于自动替换自荐信和邮件标题)
SENDER_NAME={config.get('SENDER_NAME', '您的姓名')}
SENDER_UNIVERSITY={config.get('SENDER_UNIVERSITY', '您的本科学校')}
EMAIL_SUBJECT={config.get('EMAIL_SUBJECT', '预推免自荐信-{{SENDER_UNIVERSITY}}-{{SENDER_NAME}}')}
"""
    with open(ENV_FILE_PATH, 'w', encoding='utf-8') as f:
        f.write(content)

def read_teachers_csv():
    """读取导师数据"""
    if not os.path.exists(TEACHERS_CSV_PATH):
        return [], 'utf-8-sig'
    
    encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312']
    for enc in encodings:
        try:
            with open(TEACHERS_CSV_PATH, 'r', encoding=enc) as f:
                reader = csv.DictReader(f)
                headers = reader.fieldnames if reader.fieldnames else []
                cleaned_headers = [h.strip() for h in headers]
                
                # 判断表头
                required = ['姓名', '学校', '学院', '邮箱', '研究方向简述', '状态']
                if not all(req in cleaned_headers for req in required):
                    continue
                
                teachers = []
                for row in reader:
                    teachers.append({k.strip(): v.strip() for k, v in row.items() if k})
                return teachers, enc
        except (UnicodeDecodeError, ValueError):
            continue
    return [], 'utf-8-sig'

def save_teachers_csv(teachers, encoding='utf-8-sig'):
    """保存导师数据"""
    with open(TEACHERS_CSV_PATH, 'w', encoding=encoding, newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['姓名', '学校', '学院', '邮箱', '研究方向简述', '状态'])
        writer.writeheader()
        writer.writerows(teachers)

# ----------------- GUI 应用程序类 -----------------

class EmailSenderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 初始化状态
        self.is_running = False
        self.stop_requested = False
        self.config = load_config()
        self.teachers, self.csv_encoding = read_teachers_csv()
        
        # 窗口大小与定位
        self.title("🎓 保研/考研联系导师个性化邮件自动发送工具 v2.0")
        self.geometry("1024x720")
        self.minsize(980, 650)
        
        # 居中显示
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'+{x}+{y}')
        
        # 布局格子配置 (1行2列：左边侧边栏导航，右边主内容区)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # 创建左侧边栏
        self.create_sidebar()
        
        # 创建右侧主框架容器
        self.main_container = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_container.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)
        
        # 实例化四个功能页面
        self.frames = {}
        self.create_settings_page()
        self.create_teachers_page()
        self.create_template_page()
        self.create_send_page()
        
        # 默认展示“配置中心”
        self.select_page("settings")

    # ----------------- UI 搭建 -----------------

    def create_sidebar(self):
        """左侧导航侧边栏"""
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)
        
        # Logo/标题区
        logo_label = ctk.CTkLabel(self.sidebar_frame, text="学术发信助手", font=ctk.CTkFont(size=20, weight="bold"))
        logo_label.grid(row=0, column=0, padx=20, pady=(30, 10))
        sub_logo_label = ctk.CTkLabel(self.sidebar_frame, text="GitHub 开源版", font=ctk.CTkFont(size=12, slant="italic"), text_color="gray")
        sub_logo_label.grid(row=1, column=0, padx=20, pady=(0, 30))
        
        # 导航按钮
        self.btn_settings = ctk.CTkButton(self.sidebar_frame, text="🔑 配置中心", anchor="w", fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command=lambda: self.select_page("settings"))
        self.btn_settings.grid(row=2, column=0, padx=15, pady=8, sticky="ew")
        
        self.btn_teachers = ctk.CTkButton(self.sidebar_frame, text="📝 导师管理", anchor="w", fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command=lambda: self.select_page("teachers"))
        self.btn_teachers.grid(row=3, column=0, padx=15, pady=8, sticky="ew")
        
        self.btn_template = ctk.CTkButton(self.sidebar_frame, text="📄 模板编辑", anchor="w", fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command=lambda: self.select_page("template"))
        self.btn_template.grid(row=4, column=0, padx=15, pady=8, sticky="ew")
        
        self.btn_send = ctk.CTkButton(self.sidebar_frame, text="✉️ 开始发送", anchor="w", fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command=lambda: self.select_page("send"))
        self.btn_send.grid(row=5, column=0, padx=15, pady=8, sticky="ew")
        
        # 底部的外观模式切换
        theme_label = ctk.CTkLabel(self.sidebar_frame, text="外观模式:", anchor="w")
        theme_label.grid(row=6, column=0, padx=20, pady=(10, 0), sticky="w")
        
        self.theme_optionmenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["System", "Dark", "Light"], command=self.change_appearance_mode)
        self.theme_optionmenu.grid(row=7, column=0, padx=20, pady=(5, 30), sticky="ew")

    def select_page(self, page_name):
        """选择并展示页面"""
        # 复原按钮样式
        for btn in [self.btn_settings, self.btn_teachers, self.btn_template, self.btn_send]:
            btn.configure(fg_color="transparent")
        
        # 设置激活态样式
        if page_name == "settings":
            self.btn_settings.configure(fg_color=("gray75", "gray25"))
        elif page_name == "teachers":
            self.btn_teachers.configure(fg_color=("gray75", "gray25"))
            self.load_teachers_to_ui()  # 切换时刷新导师列表
        elif page_name == "template":
            self.btn_template.configure(fg_color=("gray75", "gray25"))
            self.load_template_to_ui()  # 切换时刷新模板内容
        elif page_name == "send":
            self.btn_send.configure(fg_color=("gray75", "gray25"))
            self.update_send_dashboard()  # 切换时刷新仪表盘
            
        # 展现页面框架
        for frame in self.frames.values():
            frame.grid_remove()
        self.frames[page_name].grid(row=0, column=0, sticky="nsew")

    def change_appearance_mode(self, new_mode):
        """主题切换"""
        ctk.set_appearance_mode(new_mode)

    # ----------------- 页面一：配置中心 -----------------

    def create_settings_page(self):
        """构建配置页面"""
        page = ctk.CTkFrame(self.main_container, corner_radius=15)
        self.frames["settings"] = page
        page.grid_rowconfigure(0, weight=1)
        page.grid_columnconfigure(0, weight=1)
        
        # 滚动框架以支持内容过多时的显示
        scroll_frame = ctk.CTkScrollableFrame(page, label_text="系统配置中心", label_font=ctk.CTkFont(size=18, weight="bold"))
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)
        scroll_frame.grid_columnconfigure(1, weight=1)
        
        # 1. DeepSeek 配置区
        section_ai = ctk.CTkLabel(scroll_frame, text="🤖 DeepSeek 大模型配置", font=ctk.CTkFont(size=14, weight="bold"), text_color="#1f6aa5")
        section_ai.grid(row=0, column=0, columnspan=3, pady=(15, 10), sticky="w", padx=10)
        
        ctk.CTkLabel(scroll_frame, text="API Key:").grid(row=1, column=0, padx=15, pady=8, sticky="e")
        self.entry_api_key = ctk.CTkEntry(scroll_frame, placeholder_text="填入您的 sk-xxxxxx 密钥", show="*")
        self.entry_api_key.grid(row=1, column=1, padx=15, pady=8, sticky="ew")
        self.entry_api_key.insert(0, self.config.get("DEEPSEEK_API_KEY", ""))
        
        # 眼睛按钮（明文/密文切换）
        self.btn_show_key = ctk.CTkButton(scroll_frame, text="👁️", width=35, command=self.toggle_api_key_visibility)
        self.btn_show_key.grid(row=1, column=2, padx=(0, 15), pady=8)
        
        ctk.CTkLabel(scroll_frame, text="API URL:").grid(row=2, column=0, padx=15, pady=8, sticky="e")
        self.entry_api_url = ctk.CTkEntry(scroll_frame)
        self.entry_api_url.grid(row=2, column=1, columnspan=2, padx=15, pady=8, sticky="ew")
        self.entry_api_url.insert(0, self.config.get("DEEPSEEK_API_URL", "https://api.deepseek.com/v1"))
        
        ctk.CTkLabel(scroll_frame, text="模型名称:").grid(row=3, column=0, padx=15, pady=8, sticky="e")
        self.combo_model = ctk.CTkComboBox(scroll_frame, values=["deepseek-chat", "deepseek-reasoner"])
        self.combo_model.grid(row=3, column=1, columnspan=2, padx=15, pady=8, sticky="ew")
        self.combo_model.set(self.config.get("DEEPSEEK_MODEL", "deepseek-chat"))
        
        # 2. 邮箱 SMTP 配置区
        section_mail = ctk.CTkLabel(scroll_frame, text="✉️ 发信邮箱配置 (SMTP)", font=ctk.CTkFont(size=14, weight="bold"), text_color="#1f6aa5")
        section_mail.grid(row=4, column=0, columnspan=3, pady=(25, 10), sticky="w", padx=10)
        
        ctk.CTkLabel(scroll_frame, text="SMTP 服务器:").grid(row=5, column=0, padx=15, pady=8, sticky="e")
        self.entry_smtp_server = ctk.CTkEntry(scroll_frame)
        self.entry_smtp_server.grid(row=5, column=1, columnspan=2, padx=15, pady=8, sticky="ew")
        self.entry_smtp_server.insert(0, self.config.get("SMTP_SERVER", "smtp.163.com"))
        
        ctk.CTkLabel(scroll_frame, text="SMTP 端口:").grid(row=6, column=0, padx=15, pady=8, sticky="e")
        self.entry_smtp_port = ctk.CTkEntry(scroll_frame)
        self.entry_smtp_port.grid(row=6, column=1, columnspan=2, padx=15, pady=8, sticky="ew")
        self.entry_smtp_port.insert(0, self.config.get("SMTP_PORT", "465"))
        
        ctk.CTkLabel(scroll_frame, text="邮箱账号:").grid(row=7, column=0, padx=15, pady=8, sticky="e")
        self.entry_smtp_user = ctk.CTkEntry(scroll_frame)
        self.entry_smtp_user.grid(row=7, column=1, columnspan=2, padx=15, pady=8, sticky="ew")
        self.entry_smtp_user.insert(0, self.config.get("SMTP_USER", "your_email@example.com"))
        
        ctk.CTkLabel(scroll_frame, text="邮箱授权码:").grid(row=8, column=0, padx=15, pady=8, sticky="e")
        self.entry_smtp_pass = ctk.CTkEntry(scroll_frame, show="*")
        self.entry_smtp_pass.grid(row=8, column=1, padx=15, pady=8, sticky="ew")
        self.entry_smtp_pass.insert(0, self.config.get("SMTP_PASS", "your_smtp_auth_code_here"))
        
        self.btn_show_pass = ctk.CTkButton(scroll_frame, text="👁️", width=35, command=self.toggle_smtp_pass_visibility)
        self.btn_show_pass.grid(row=8, column=2, padx=(0, 15), pady=8)
        
        # 3. 附件配置区
        section_file = ctk.CTkLabel(scroll_frame, text="📁 个人简历附件", font=ctk.CTkFont(size=14, weight="bold"), text_color="#1f6aa5")
        section_file.grid(row=9, column=0, columnspan=3, pady=(25, 10), sticky="w", padx=10)
        
        ctk.CTkLabel(scroll_frame, text="默认简历路径:").grid(row=10, column=0, padx=15, pady=8, sticky="e")
        self.entry_attachment = ctk.CTkEntry(scroll_frame)
        self.entry_attachment.grid(row=10, column=1, padx=15, pady=8, sticky="ew")
        self.entry_attachment.insert(0, self.config.get("DEFAULT_ATTACHMENT_PATH", "resume.pdf"))
        
        self.btn_browse = ctk.CTkButton(scroll_frame, text="浏览...", width=60, command=self.browse_attachment)
        self.btn_browse.grid(row=10, column=2, padx=(0, 15), pady=8)

        # 3.5 个人自荐信个性化字段
        section_personal = ctk.CTkLabel(scroll_frame, text="👤 发件人个人信息（自动填入信件和标题）", font=ctk.CTkFont(size=14, weight="bold"), text_color="#1f6aa5")
        section_personal.grid(row=11, column=0, columnspan=3, pady=(25, 10), sticky="w", padx=10)
        
        ctk.CTkLabel(scroll_frame, text="姓名:").grid(row=12, column=0, padx=15, pady=8, sticky="e")
        self.entry_sender_name = ctk.CTkEntry(scroll_frame)
        self.entry_sender_name.grid(row=12, column=1, columnspan=2, padx=15, pady=8, sticky="ew")
        self.entry_sender_name.insert(0, self.config.get("SENDER_NAME", "您的姓名"))
        
        ctk.CTkLabel(scroll_frame, text="本科学校:").grid(row=13, column=0, padx=15, pady=8, sticky="e")
        self.entry_sender_university = ctk.CTkEntry(scroll_frame)
        self.entry_sender_university.grid(row=13, column=1, columnspan=2, padx=15, pady=8, sticky="ew")
        self.entry_sender_university.insert(0, self.config.get("SENDER_UNIVERSITY", "您的本科学校"))
        
        ctk.CTkLabel(scroll_frame, text="邮件标题模版:").grid(row=14, column=0, padx=15, pady=8, sticky="e")
        self.entry_email_subject = ctk.CTkEntry(scroll_frame, placeholder_text="支持占位符: {SENDER_UNIVERSITY}, {SENDER_NAME}")
        self.entry_email_subject.grid(row=14, column=1, columnspan=2, padx=15, pady=8, sticky="ew")
        self.entry_email_subject.insert(0, self.config.get("EMAIL_SUBJECT", "预推免自荐信-{SENDER_UNIVERSITY}-{SENDER_NAME}"))
        
        # 4. 保存操作栏
        self.btn_save_config = ctk.CTkButton(scroll_frame, text="💾 保存并应用配置", height=38, font=ctk.CTkFont(size=14, weight="bold"), command=self.action_save_config)
        self.btn_save_config.grid(row=15, column=0, columnspan=3, padx=30, pady=30, sticky="ew")

    def toggle_api_key_visibility(self):
        """明暗显示 API Key"""
        if self.entry_api_key.cget("show") == "*":
            self.entry_api_key.configure(show="")
        else:
            self.entry_api_key.configure(show="*")
            
    def toggle_smtp_pass_visibility(self):
        """明暗显示 SMTP 授权码"""
        if self.entry_smtp_pass.cget("show") == "*":
            self.entry_smtp_pass.configure(show="")
        else:
            self.entry_smtp_pass.configure(show="*")

    def browse_attachment(self):
        """文件选择弹窗选取简历 PDF"""
        filepath = filedialog.askopenfilename(
            title="选择您的简历附件",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if filepath:
            # 如果在当前项目目录下，转换为相对路径，更加美观
            cwd = os.getcwd()
            if filepath.startswith(cwd):
                filepath = os.path.relpath(filepath, cwd)
            self.entry_attachment.delete(0, ctk.END)
            self.entry_attachment.insert(0, filepath)

    def action_save_config(self):
        """读取界面输入并保存配置"""
        self.config["DEEPSEEK_API_KEY"] = self.entry_api_key.get().strip()
        self.config["DEEPSEEK_API_URL"] = self.entry_api_url.get().strip()
        self.config["DEEPSEEK_MODEL"] = self.combo_model.get().strip()
        self.config["SMTP_SERVER"] = self.entry_smtp_server.get().strip()
        self.config["SMTP_PORT"] = self.entry_smtp_port.get().strip()
        self.config["SMTP_USER"] = self.entry_smtp_user.get().strip()
        self.config["SMTP_PASS"] = self.entry_smtp_pass.get().strip()
        self.config["DEFAULT_ATTACHMENT_PATH"] = self.entry_attachment.get().strip()
        self.config["SENDER_NAME"] = self.entry_sender_name.get().strip()
        self.config["SENDER_UNIVERSITY"] = self.entry_sender_university.get().strip()
        self.config["EMAIL_SUBJECT"] = self.entry_email_subject.get().strip()
        
        try:
            save_config(self.config)
            messagebox.showinfo("配置成功", "配置已成功保存至本地 .env 文件！")
        except Exception as e:
            messagebox.showerror("保存失败", f"无法写入配置文件: {e}")

    # ----------------- 页面二：导师列表管理 -----------------

    def create_teachers_page(self):
        """导师管理页面"""
        page = ctk.CTkFrame(self.main_container, corner_radius=15)
        self.frames["teachers"] = page
        page.grid_rowconfigure(2, weight=1)
        page.grid_columnconfigure(0, weight=1)
        
        # 头部说明区
        header_frame = ctk.CTkFrame(page, corner_radius=0, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        
        title_label = ctk.CTkLabel(header_frame, text="📝 导师联系人表 (teachers.csv)", font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(side="left", padx=5)
        
        self.lbl_csv_encoding = ctk.CTkLabel(header_frame, text="编码: utf-8", font=ctk.CTkFont(size=12), text_color="gray")
        self.lbl_csv_encoding.pack(side="left", padx=15)
        
        # AI 导师智能检索工具栏
        self.search_frame = ctk.CTkFrame(page, corner_radius=10, fg_color=("gray90", "gray15"))
        self.search_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        self.search_frame.grid_columnconfigure(1, weight=1)
        
        lbl_ai_search = ctk.CTkLabel(self.search_frame, text="🤖 AI 导师智能检索并录入:", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_ai_search.grid(row=0, column=0, padx=15, pady=10, sticky="w")
        
        self.entry_ai_search_query = ctk.CTkEntry(self.search_frame, placeholder_text="输入老师姓名及所属高校学院线索 (如: 清华大学自动化系周东华)")
        self.entry_ai_search_query.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        self.btn_ai_search = ctk.CTkButton(self.search_frame, text="🔍 智能检索并添加", width=120, fg_color="#9b59b6", hover_color="#8e44ad", command=self.start_ai_search_thread)
        self.btn_ai_search.grid(row=0, column=2, padx=15, pady=10)
        
        # 中部表格区 (基于 ScrollableFrame 实现的可编辑行卡片列表) - 移到第 2 行
        self.teachers_scroll_frame = ctk.CTkScrollableFrame(page)
        self.teachers_scroll_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        self.teachers_scroll_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        
        # 底部操作工具栏 - 移到第 3 行
        toolbar_frame = ctk.CTkFrame(page, height=50, fg_color="transparent")
        toolbar_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=15)
        
        btn_add = ctk.CTkButton(toolbar_frame, text="➕ 添加新导师", width=120, fg_color="#2ecc71", hover_color="#27ae60", command=self.add_teacher_row)
        btn_add.pack(side="left", padx=5)
        
        btn_save = ctk.CTkButton(toolbar_frame, text="💾 保存数据修改", width=120, command=self.save_teachers_ui_data)
        btn_save.pack(side="right", padx=5)
        
        btn_reload = ctk.CTkButton(toolbar_frame, text="🔄 重新加载", width=100, fg_color="gray", hover_color="dimgray", command=self.load_teachers_to_ui)
        btn_reload.pack(side="right", padx=5)
        
        # 用于记录界面中的每一行输入框对象
        self.teacher_ui_rows = []

    def load_teachers_to_ui(self):
        """将 CSV 中的导师加载到界面中"""
        # 清除旧的 UI 行
        for widgets in self.teacher_ui_rows:
            for w in widgets.values():
                if isinstance(w, ctk.CTkBaseClass):
                    w.destroy()
        self.teacher_ui_rows.clear()
        
        # 重新读取
        self.teachers, self.csv_encoding = read_teachers_csv()
        self.lbl_csv_encoding.configure(text=f"本地编码: {self.csv_encoding}")
        
        # 绘制表头
        headers = ["姓名", "学校", "学院", "邮箱", "研究方向简述", "状态", "操作"]
        for col_idx, text in enumerate(headers):
            lbl = ctk.CTkLabel(self.teachers_scroll_frame, text=text, font=ctk.CTkFont(size=12, weight="bold"))
            lbl.grid(row=0, column=col_idx, padx=5, pady=5, sticky="w")
            
        # 渲染行数据
        for idx, t in enumerate(self.teachers, 1):
            self.render_teacher_row(idx, t)

    def render_teacher_row(self, row_idx, data):
        """渲染单行导师输入框"""
        # 输入框设计
        e_name = ctk.CTkEntry(self.teachers_scroll_frame, width=90)
        e_name.grid(row=row_idx, column=0, padx=3, pady=4, sticky="ew")
        e_name.insert(0, data.get("姓名", ""))
        
        e_univ = ctk.CTkEntry(self.teachers_scroll_frame, width=120)
        e_univ.grid(row=row_idx, column=1, padx=3, pady=4, sticky="ew")
        e_univ.insert(0, data.get("学校", ""))
        
        e_school = ctk.CTkEntry(self.teachers_scroll_frame, width=120)
        e_school.grid(row=row_idx, column=2, padx=3, pady=4, sticky="ew")
        e_school.insert(0, data.get("学院", ""))
        
        e_email = ctk.CTkEntry(self.teachers_scroll_frame, width=150)
        e_email.grid(row=row_idx, column=3, padx=3, pady=4, sticky="ew")
        e_email.insert(0, data.get("邮箱", ""))
        
        e_focus = ctk.CTkEntry(self.teachers_scroll_frame, width=180)
        e_focus.grid(row=row_idx, column=4, padx=3, pady=4, sticky="ew")
        e_focus.insert(0, data.get("研究方向简述", ""))
        
        # 状态选择下拉框
        c_status = ctk.CTkComboBox(self.teachers_scroll_frame, values=["待发送", "已发送"], width=90)
        c_status.grid(row=row_idx, column=5, padx=3, pady=4, sticky="ew")
        c_status.set(data.get("状态", "待发送"))
        
        # 删除按钮
        btn_del = ctk.CTkButton(self.teachers_scroll_frame, text="🗑️", width=35, fg_color="#e74c3c", hover_color="#c0392b")
        btn_del.grid(row=row_idx, column=6, padx=5, pady=4)
        
        row_widgets = {
            "姓名": e_name,
            "学校": e_univ,
            "学院": e_school,
            "邮箱": e_email,
            "研究方向简述": e_focus,
            "状态": c_status,
            "del_btn": btn_del
        }
        
        # 绑定删除操作
        btn_del.configure(command=lambda w=row_widgets: self.remove_teacher_row_ui(w))
        
        self.teacher_ui_rows.append(row_widgets)

    def add_teacher_row(self):
        """在 UI 底部追加一个空白行"""
        new_row_idx = len(self.teacher_ui_rows) + 1
        dummy_data = {"姓名": "", "学校": "", "学院": "", "邮箱": "", "研究方向简述": "", "状态": "待发送"}
        self.render_teacher_row(new_row_idx, dummy_data)
        # 滚动到底部
        self.teachers_scroll_frame._parent_canvas.yview_moveto(1.0)

    def remove_teacher_row_ui(self, row_widgets):
        """将选定行从界面销毁"""
        for w in row_widgets.values():
            if isinstance(w, ctk.CTkBaseClass):
                w.destroy()
        self.teacher_ui_rows.remove(row_widgets)

    def save_teachers_ui_data(self):
        """从 UI 框中汇总数据并持久化到 CSV 文件中"""
        new_teachers = []
        for row in self.teacher_ui_rows:
            name = row["姓名"].get().strip()
            univ = row["学校"].get().strip()
            school = row["学院"].get().strip()
            email = row["邮箱"].get().strip()
            focus = row["研究方向简述"].get().strip()
            status = row["状态"].get().strip()
            
            # 至少姓名、学校、学院、邮箱其中之一有值才存入
            if name or univ or school or email:
                new_teachers.append({
                    "姓名": name,
                    "学校": univ,
                    "学院": school,
                    "邮箱": email,
                    "研究方向简述": focus,
                    "状态": status
                })
        
        try:
            save_teachers_csv(new_teachers, self.csv_encoding)
            self.teachers = new_teachers
            messagebox.showinfo("保存成功", f"成功保存 {len(new_teachers)} 位导师数据到 teachers.csv！")
            self.load_teachers_to_ui()  # 刷新重绘
        except Exception as e:
            messagebox.showerror("保存失败", f"无法写入 CSV 文件: {e}")

    # ----------------- 页面三：模板文本编辑 -----------------

    def create_template_page(self):
        """模板编辑器页面"""
        page = ctk.CTkFrame(self.main_container, corner_radius=15)
        self.frames["template"] = page
        page.grid_rowconfigure(1, weight=1)
        page.grid_columnconfigure(0, weight=1)
        
        # 头部
        header_frame = ctk.CTkFrame(page, corner_radius=0, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        
        title_label = ctk.CTkLabel(header_frame, text="📄 邮件模版编辑器 (模板.md)", font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(side="left", padx=5)
        
        lbl_hint = ctk.CTkLabel(header_frame, text="⚠️ 请不要随意删除信中的 [XXX] 占位符", font=ctk.CTkFont(size=12), text_color="#e67e22")
        lbl_hint.pack(side="right", padx=15)
        
        # 文本框 (大输入框)
        self.txt_template = ctk.CTkTextbox(page, font=ctk.CTkFont(family="Consolas", size=14))
        self.txt_template.grid(row=1, column=0, sticky="nsew", padx=15, pady=10)
        
        # 底部栏
        toolbar_frame = ctk.CTkFrame(page, height=50, fg_color="transparent")
        toolbar_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=15)
        
        btn_save = ctk.CTkButton(toolbar_frame, text="💾 保存模板修改", width=150, command=self.save_template_content)
        btn_save.pack(side="right", padx=5)
        
        btn_reset = ctk.CTkButton(toolbar_frame, text="🔄 重新载入", fg_color="gray", hover_color="dimgray", width=120, command=self.load_template_to_ui)
        btn_reset.pack(side="right", padx=5)

    def load_template_to_ui(self):
        """加载本地模板到编辑区"""
        if os.path.exists(TEMPLATE_FILE_PATH):
            with open(TEMPLATE_FILE_PATH, 'r', encoding='utf-8') as f:
                content = f.read()
            self.txt_template.delete("1.0", ctk.END)
            self.txt_template.insert("1.0", content)
        else:
            self.txt_template.delete("1.0", ctk.END)
            self.txt_template.insert("1.0", "错误: 未找到 模板.md 文件，请在本地创建。")

    def save_template_content(self):
        """保存模板"""
        content = self.txt_template.get("1.0", ctk.END).strip()
        try:
            with open(TEMPLATE_FILE_PATH, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("模板保存成功", "保研自荐信模板 (模板.md) 修改成功并已存入本地！")
        except Exception as e:
            messagebox.showerror("模板保存失败", f"无法写入文件: {e}")

    # ----------------- 页面四：任务发送控制台 -----------------

    def create_send_page(self):
        """发送日志及控制页面"""
        page = ctk.CTkFrame(self.main_container, corner_radius=15)
        self.frames["send"] = page
        page.grid_rowconfigure(2, weight=1)
        page.grid_columnconfigure(0, weight=1)
        
        # 1. 顶部统计仪表盘
        self.dashboard_frame = ctk.CTkFrame(page, height=80)
        self.dashboard_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        self.dashboard_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        self.lbl_dash_total = ctk.CTkLabel(self.dashboard_frame, text="总数: -\n加载中", font=ctk.CTkFont(size=14, weight="bold"))
        self.lbl_dash_total.grid(row=0, column=0, pady=15)
        
        self.lbl_dash_pending = ctk.CTkLabel(self.dashboard_frame, text="未发送: -\n加载中", font=ctk.CTkFont(size=14, weight="bold"), text_color="#e67e22")
        self.lbl_dash_pending.grid(row=0, column=1, pady=15)
        
        self.lbl_dash_success = ctk.CTkLabel(self.dashboard_frame, text="已发送: -\n加载中", font=ctk.CTkFont(size=14, weight="bold"), text_color="#2ecc71")
        self.lbl_dash_success.grid(row=0, column=2, pady=15)
        
        self.lbl_dash_failed = ctk.CTkLabel(self.dashboard_frame, text="失败: -\n加载中", font=ctk.CTkFont(size=14, weight="bold"), text_color="#e74c3c")
        self.lbl_dash_failed.grid(row=0, column=3, pady=15)
        
        # 2. 进度条与状态展示区
        progress_container = ctk.CTkFrame(page, fg_color="transparent")
        progress_container.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        progress_container.grid_columnconfigure(0, weight=1)
        
        self.progress_bar = ctk.CTkProgressBar(progress_container)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=5, pady=(10, 5))
        self.progress_bar.set(0)
        
        self.lbl_progress_status = ctk.CTkLabel(progress_container, text="系统闲置中，等待您的指令...", font=ctk.CTkFont(size=12))
        self.lbl_progress_status.grid(row=1, column=0, sticky="w", padx=5)
        
        # 3. 日志终端窗口 (大Console)
        self.txt_console = ctk.CTkTextbox(page, font=ctk.CTkFont(family="Consolas", size=12), fg_color=("black", "black"), text_color=("lightgray", "lightgray"))
        self.txt_console.grid(row=2, column=0, sticky="nsew", padx=15, pady=5)
        self.write_console_log("[系统就绪] 保研发信助手初始化完成。\n -> 首次发信前建议点击“一键本地预览”检查AI生成效果。\n -> 若本地已生成预览，点击发信将直接读取缓存，极速秒发！\n" + "="*80 + "\n")
        
        # 4. 底部执行动作按钮区
        control_frame = ctk.CTkFrame(page, height=60, fg_color="transparent")
        control_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=15)
        
        self.btn_action_preview = ctk.CTkButton(control_frame, text="🔍 一键生成本地预览 (.md)", fg_color="#34495e", hover_color="#2c3e50", height=40, font=ctk.CTkFont(size=13, weight="bold"), command=self.start_preview_thread)
        self.btn_action_preview.pack(side="left", padx=5)
        
        self.btn_action_send = ctk.CTkButton(control_frame, text="🚀 开始批量安全发信", fg_color="#1f6aa5", hover_color="#144d75", height=40, font=ctk.CTkFont(size=13, weight="bold"), command=self.start_send_thread)
        self.btn_action_send.pack(side="left", padx=5)
        
        self.btn_action_stop = ctk.CTkButton(control_frame, text="🛑 中止发送", fg_color="#e74c3c", hover_color="#c0392b", height=40, font=ctk.CTkFont(size=13, weight="bold"), command=self.request_stop_process, state="disabled")
        self.btn_action_stop.pack(side="left", padx=5)
        
        self.btn_clear_console = ctk.CTkButton(control_frame, text="🗑️ 清空日志", fg_color="gray", hover_color="dimgray", height=40, width=90, command=self.clear_console)
        self.btn_clear_console.pack(side="right", padx=5)

    def update_send_dashboard(self):
        """刷新仪表盘数据统计"""
        self.teachers, self.csv_encoding = read_teachers_csv()
        total = len(self.teachers)
        sent = sum(1 for t in self.teachers if t.get("状态") == "已发送")
        pending = total - sent
        
        # 注意：此处未直接统计“失败”，失败是指在发信环节报错的。我们可以在本轮发送统计中动态显示。
        self.lbl_dash_total.configure(text=f"导入总数\n{total} 人")
        self.lbl_dash_pending.configure(text=f"等待发信\n{pending} 人")
        self.lbl_dash_success.configure(text=f"已经发送\n{sent} 人")
        self.lbl_dash_failed.configure(text="本轮失败\n0 人")

    def write_console_log(self, message):
        """向虚拟终端写入一行日志"""
        self.txt_console.insert(ctk.END, message)
        self.txt_console.see(ctk.END)

    def clear_console(self):
        """清空日志"""
        self.txt_console.delete("1.0", ctk.END)

    # ----------------- 后台多线程任务处理 -----------------

    def disable_ui_controls(self):
        """任务开始时锁定界面控件防止误触"""
        self.is_running = True
        self.btn_action_preview.configure(state="disabled")
        self.btn_action_send.configure(state="disabled")
        self.btn_action_stop.configure(state="normal")
        self.btn_save_config.configure(state="disabled")
        
    def enable_ui_controls(self):
        """任务结束时解锁控件"""
        self.is_running = False
        self.btn_action_preview.configure(state="normal")
        self.btn_action_send.configure(state="normal")
        self.btn_action_stop.configure(state="disabled")
        self.btn_save_config.configure(state="normal")
        self.update_send_dashboard()

    def request_stop_process(self):
        """请求中止发信"""
        if self.is_running:
            self.stop_requested = True
            self.write_console_log("\n⚠️ [中止指令] 收到终止任务请求，正在完成当前步骤后安全关闭...\n")
            self.lbl_progress_status.configure(text="正在安全终止当前线程...")

    # 1. 预览任务线程拉起
    def start_preview_thread(self):
        self.disable_ui_controls()
        self.stop_requested = False
        self.progress_bar.set(0)
        self.lbl_progress_status.configure(text="正在准备生成本地预览...")
        
        # 拉起线程
        t = threading.Thread(target=self.bg_generate_previews)
        t.daemon = True
        t.start()

    def bg_generate_previews(self):
        """后台线程：分析并生成所有预览邮件"""
        self.teachers, self.csv_encoding = read_teachers_csv()
        pending_teachers = [t for t in self.teachers if t.get('状态') != '已发送']
        
        if not pending_teachers:
            self.write_console_log("\n💡 提示: 所有导师的信件已全部发送完毕（状态均已设为“已发送”），无需生成预览！\n")
            self.lbl_progress_status.configure(text="无需生成预览。")
            self.root_after_safe(self.enable_ui_controls)
            return

        api_key = self.config.get("DEEPSEEK_API_KEY", "").strip()
        if not api_key:
            messagebox.showerror("配置错误", "DEEPSEEK_API_KEY 不能为空！请先在‘配置中心’填写。")
            self.root_after_safe(self.enable_ui_controls)
            return

        if not os.path.exists(TEMPLATE_FILE_PATH):
            self.write_console_log(f"\n❌ 错误: 本地未找到模板文件 {TEMPLATE_FILE_PATH}！\n")
            self.root_after_safe(self.enable_ui_controls)
            return
            
        with open(TEMPLATE_FILE_PATH, 'r', encoding='utf-8') as f:
            template_content = f.read()

        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)

        total_tasks = len(pending_teachers)
        self.write_console_log(f"\n🚀 开始批量生成个性化预览邮件 (共需分析 {total_tasks} 位导师背景)...\n" + "-"*60 + "\n")
        
        success_count = 0
        
        for idx, t in enumerate(pending_teachers, 1):
            if self.stop_requested:
                self.write_console_log("\n🛑 预览生成任务已被用户强行终止。\n")
                break
                
            name = t.get('姓名', '').strip()
            univ = t.get('学校', '').strip()
            school = t.get('学院', '').strip()
            focus = t.get('研究方向简述', '').strip()
            
            self.lbl_progress_status.configure(text=f"正在分析导师 ({idx}/{total_tasks}): {univ} - {name}...")
            self.progress_bar.set(idx / total_tasks)
            
            self.write_console_log(f"[{idx}/{total_tasks}] 正在调用 DeepSeek API 分析: {univ} - {school} - {name}...\n")
            
            try:
                # 调用 DeepSeek API
                info = query_deepseek_teacher_info(self.config, name, univ, school, focus)
                email_content = generate_email_content(template_content, info)
                
                # 写到本地
                filename = f"{self.config.get('SENDER_NAME', '自荐信')}_致_{univ}_{school}_{name}老师.md"
                filepath = os.path.join(OUTPUT_DIR, filename)
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(email_content)
                
                self.write_console_log(f" -> 💡 生成成功！已导出至 {filepath}\n")
                self.write_console_log(f" -> 成果关键词: {info['research_achievement']}\n")
                self.write_console_log(f" -> 契合方向: {info['research_direction']}\n\n")
                success_count += 1
            except Exception as e:
                self.write_console_log(f" -> ❌ 生成失败: {e}\n\n")
            
            # 短暂喘息，平滑界面
            time.sleep(0.5)

        self.write_console_log("="*60 + f"\n🎉 预览生成结束！共成功生成 {success_count}/{total_tasks} 封个性化邮件。\n")
        self.write_console_log(f"您可以打开本地文件夹 '{OUTPUT_DIR}/' 查看具体的 Markdown 格式邮件文件！\n")
        self.lbl_progress_status.configure(text="预览文件生成完毕！")
        self.root_after_safe(self.enable_ui_controls)

    # 2. 发送任务线程拉起
    def start_send_thread(self):
        # 做一下前置附件检查
        attachment_path = self.config.get("DEFAULT_ATTACHMENT_PATH", "resume.pdf")
        if not os.path.exists(attachment_path):
            confirm = messagebox.askyesno(
                "附件未找到警告", 
                f"系统未在目录下找到您的简历附件 '{attachment_path}'！\n\n您确定要【不带任何简历附件】直接发信给导师吗？"
            )
            if not confirm:
                return
                
        self.disable_ui_controls()
        self.stop_requested = False
        self.progress_bar.set(0)
        self.lbl_progress_status.configure(text="正在连接并校验 API 状态...")
        
        t = threading.Thread(target=self.bg_send_emails)
        t.daemon = True
        t.start()

    def bg_send_emails(self):
        """后台线程：分析并发送邮件"""
        self.teachers, self.csv_encoding = read_teachers_csv()
        pending_teachers = [t for t in self.teachers if t.get('状态') != '已发送']
        
        if not pending_teachers:
            self.write_console_log("\n💡 提示: 所有导师状态均为“已发送”，未发现待发送的导师。\n")
            self.lbl_progress_status.configure(text="没有待发任务。")
            self.root_after_safe(self.enable_ui_controls)
            return

        api_key = self.config.get("DEEPSEEK_API_KEY", "").strip()
        if not api_key:
            messagebox.showerror("配置错误", "DEEPSEEK_API_KEY 不能为空！请先填写。")
            self.root_after_safe(self.enable_ui_controls)
            return

        if not os.path.exists(TEMPLATE_FILE_PATH):
            self.write_console_log(f"\n❌ 错误: 未找到模板文件 {TEMPLATE_FILE_PATH}！\n")
            self.root_after_safe(self.enable_ui_controls)
            return
            
        with open(TEMPLATE_FILE_PATH, 'r', encoding='utf-8') as f:
            template_content = f.read()

        total_tasks = len(pending_teachers)
        self.write_console_log(f"\n🚀 开始批量保研发信流程 (共 {total_tasks} 封待发邮件)...\n" + "="*80 + "\n")
        
        success_count = 0
        failed_count = 0
        
        # 建立网络会话，进行 SMTP 状态校验（可选，这里在循环里处理）
        for idx, t in enumerate(pending_teachers, 1):
            if self.stop_requested:
                self.write_console_log("\n🛑 批量发信流程已被用户强行终止！\n")
                break
                
            name = t.get('姓名', '').strip()
            univ = t.get('学校', '').strip()
            school = t.get('学院', '').strip()
            email = t.get('邮箱', '').strip()
            focus = t.get('研究方向简述', '').strip()
            
            self.lbl_progress_status.configure(text=f"正在处理 ({idx}/{total_tasks}): {univ} - {name}...")
            self.progress_bar.set(idx / total_tasks)
            
            self.write_console_log(f"[{idx}/{total_tasks}] 开始为 {univ} {name} 老师准备邮件 ({email})...\n")
            
            # A. 校验本地是否有 out_emails/ 目录下生成的缓存，如果有，直接读取以省去 DeepSeek 的 Token
            local_filename = f"{self.config.get('SENDER_NAME', '自荐信')}_致_{univ}_{school}_{name}老师.md"
            local_filepath = os.path.join(OUTPUT_DIR, local_filename)
            email_content = ""
            
            if os.path.exists(local_filepath):
                self.write_console_log(" -> ⚡ 检测到本地已存在个性化预览文件，直接读取缓存...\n")
                with open(local_filepath, 'r', encoding='utf-8') as f:
                    email_content = f.read()
            else:
                self.write_console_log(" -> 正在调用 DeepSeek API 补全学术研究方向占位符...\n")
                try:
                    info = query_deepseek_teacher_info(self.config, name, univ, school, focus)
                    email_content = generate_email_content(template_content, info)
                    
                    # 顺手写一份本地备份
                    if not os.path.exists(OUTPUT_DIR):
                        os.makedirs(OUTPUT_DIR)
                    with open(local_filepath, 'w', encoding='utf-8') as f:
                        f.write(email_content)
                    self.write_console_log(" -> 💡 API 补全并替换占位符完成。\n")
                except Exception as e:
                    self.write_console_log(f" -> ❌ API 分析背景失败: {e}，跳过此导师。\n\n")
                    failed_count += 1
                    self.lbl_dash_failed.configure(text=f"本轮失败\n{failed_count} 人")
                    continue
            
            # B. 发送邮件
            subject_template = self.config.get("EMAIL_SUBJECT", "预推免自荐信-{SENDER_UNIVERSITY}-{SENDER_NAME}")
            try:
                subject = subject_template.format(
                    SENDER_UNIVERSITY=self.config.get("SENDER_UNIVERSITY", "您的本科学校"),
                    SENDER_NAME=self.config.get("SENDER_NAME", "您的姓名")
                )
            except Exception:
                subject = f"自荐信 - {self.config.get('SENDER_UNIVERSITY', '您的本科学校')} - {self.config.get('SENDER_NAME', '您的姓名')}"
            self.write_console_log(f" -> 正在建立 SSL 隧道并以主题「{subject}」发送邮件至 {email}...\n")
            
            send_success = send_email_via_smtp(self.config, email, subject, email_content)
            
            if send_success:
                self.write_console_log(f" -> 🎉 发送成功！已向 {name} 老师发送推免自荐信。\n\n")
                success_count += 1
                
                # 寻找并在全局列表中修改该导师状态为“已发送”
                for raw_t in self.teachers:
                    if raw_t.get('姓名') == name and raw_t.get('学校') == univ and raw_t.get('邮箱') == email:
                        raw_t['状态'] = '已发送'
                        break
                        
                # 写入持久化
                try:
                    save_teachers_csv(self.teachers, self.csv_encoding)
                except Exception as ex:
                    self.write_console_log(f" -> ⚠️ 写入 CSV 备份状态失败: {ex}\n")
            else:
                self.write_console_log(f" -> ❌ SMTP 发信异常！请检查您的邮箱授权码或邮件高频限制。\n\n")
                failed_count += 1
                self.lbl_dash_failed.configure(text=f"本轮失败\n{failed_count} 人")
                
            # C. 延时，保护机制防封号
            if idx < total_tasks and not self.stop_requested:
                self.write_console_log(" -> ⏳ 冷却保护：为避免触发高频防垃圾邮件机制，静置 6 秒...\n")
                # 每次静置以 1 秒为粒度方便随时退出
                for _ in range(6):
                    if self.stop_requested:
                        break
                    time.sleep(1)

        self.write_console_log("="*80 + f"\n🏁 批量发信流程全部结束！\n -> 发信成功: {success_count} 封\n -> 发信失败: {failed_count} 封\n")
        self.lbl_progress_status.configure(text="批量发信工作已结束。")
        self.root_after_safe(self.enable_ui_controls)

    def root_after_safe(self, func):
        """线程安全的 GUI 调度回调"""
        self.after(0, func)

    def start_ai_search_thread(self):
        """拉起后台线程调用 DeepSeek API 搜索导师信息"""
        query = self.entry_ai_search_query.get().strip()
        if not query:
            messagebox.showwarning("检索提示", "请输入要检索的导师线索（例如姓名与学校）！")
            return
            
        api_key = self.config.get("DEEPSEEK_API_KEY", "").strip()
        if not api_key:
            messagebox.showerror("配置错误", "请先在“配置中心”页面配置您的 DEEPSEEK_API_KEY！")
            return
            
        # 锁定按钮，防止重复点击
        self.btn_ai_search.configure(state="disabled", text="🔍 检索中...")
        
        t = threading.Thread(target=self.bg_ai_search_teacher, args=(query, api_key))
        t.daemon = True
        t.start()
        
    def bg_ai_search_teacher(self, query, api_key):
        """后台线程：调用 DeepSeek API 查询导师并保存录入"""
        try:
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            
            system_prompt = (
                "你是一个极其专业的学术信息检索与推理助手。请根据用户提供的线索（如姓名、学校、学院等），"
                "利用你的知识库与逻辑事实，检索并推演出该导师的真实信息。你的回答必须严格且仅输出如下指定的 JSON 格式，"
                "不能包含任何额外的解释或 Markdown 格式（严禁使用 ```json 标记），直接以 { 开始并以 } 结束。"
            )
            
            user_prompt = f"""
用户提供的导师检索线索：
{query}

请提取并检索该导师的以下 5 项基本信息：
1. "姓名": 该导师的真实姓名（如无法匹配真实姓名，请直接提取输入中的姓名，如“王老师”）。
2. "学校": 导师所属高校的完整官方名称（例如“西北工业大学”而不要缩写为“西工大”）。
3. "学院": 导师所属学院的完整官方名称（例如“自动化学院”）。
4. "邮箱": 导师的真实官方联系电子邮箱（通常是高校 edu 域名的电子邮箱。如果已知或可以查到真实邮箱，请务必给出真实邮箱！如果实在无法确定真实邮箱，请根据该校的edu域名拼写出其可能的邮箱，例如：姓名拼音@mail.nwpu.edu.cn 等，但必须尽量保证邮箱格式的合理性）。
5. "研究方向简述": 该导师的核心科研细分领域或代表性方向（字数控制在 25-50 个字，例如“捷联惯导/自适应卡尔曼滤波/状态估计”）。

请严格按照如下的 JSON 格式直接返回，不能有其他任何文本：
{{
  "姓名": "...",
  "学校": "...",
  "学院": "...",
  "邮箱": "...",
  "研究方向简述": "..."
}}
"""
            payload = {
                "model": self.config.get("DEEPSEEK_MODEL", "deepseek-chat"),
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                "temperature": 0.2,
                "response_format": {"type": "json_object"}
            }
            
            url = f"{self.config.get('DEEPSEEK_API_URL', 'https://api.deepseek.com/v1').rstrip('/')}/chat/completions"
            response = requests.post(url, headers=headers, json=payload, timeout=25)
            response.raise_for_status()
            
            result = response.json()
            content = result['choices'][0]['message']['content'].strip()
            
            # 清洗并解析 JSON
            match = re.search(r'\{.*\}', content, re.DOTALL)
            if match:
                content = match.group(0)
            
            info = json.loads(content)
            
            # 校对字段
            required = ['姓名', '学校', '学院', '邮箱', '研究方向简述']
            for field in required:
                if field not in info:
                    info[field] = ""
            
            info['状态'] = '待发送'
            
            # 读取当前所有的老师
            teachers, encoding = read_teachers_csv()
            teachers.append(info)
            save_teachers_csv(teachers, encoding)
            
            # 主线程安全刷新
            self.root_after_safe(lambda: self.on_ai_search_success(info))
            
        except Exception as e:
            self.root_after_safe(lambda: self.on_ai_search_failed(e))
            
    def on_ai_search_success(self, info):
        """AI 检索成功后的主线程回调"""
        self.btn_ai_search.configure(state="normal", text="🔍 智能检索并添加")
        self.entry_ai_search_query.delete(0, ctk.END)
        messagebox.showinfo("检索成功", f"🎉 AI 已成功检索并导入该导师！\n\n• 姓名: {info['姓名']}\n• 学校: {info['学校']}\n• 学院: {info['学院']}\n• 邮箱: {info['邮箱']}\n• 研究方向简述: {info['研究方向简述']}\n\n该导师已成功追加到 teachers.csv 中。")
        self.load_teachers_to_ui()  # 重绘表格
        
    def on_ai_search_failed(self, error):
        """AI 检索失败后的主线程回调"""
        self.btn_ai_search.configure(state="normal", text="🔍 智能检索并添加")
        messagebox.showerror("检索失败", f"无法通过 AI 检索该导师信息。\n\n详情: {error}")

# ----------------- DeepSeek API 与 SMTP 支持函数 -----------------

def clean_and_parse_json(text):
    text = text.strip()
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        text = match.group(0)
    return json.loads(text)

def query_deepseek_teacher_info(config, name, university, school, research_focus=""):
    headers = {
        "Authorization": f"Bearer {config['DEEPSEEK_API_KEY']}",
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
        "model": config["DEEPSEEK_MODEL"],
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"}
    }
    
    url = f"{config['DEEPSEEK_API_URL'].rstrip('/')}/chat/completions"
    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()
    
    result = response.json()
    content = result['choices'][0]['message']['content']
    return clean_and_parse_json(content)

def generate_email_content(template, info):
    content = template
    now = datetime.now()
    today_str = f"{now.year}年{now.month}月{now.day}日"
    
    content = content.replace("尊敬的XXX老师：", f"尊敬的{info['teacher_name']}老师：")
    content = content.replace("近期在关注XXX大学XXX学院的导师信息时", f"近期在关注{info['university']}{info['school']}的导师信息时")
    content = content.replace("我详细阅读了您在XXX的研究成果", f"我详细阅读了您在{info['research_achievement']}的研究成果")
    content = content.replace("XXX大学深厚的数理底蕴和学术氛围令我由衷向往", f"{info['university']}深厚的数理底蕴和学术氛围令我由衷向往")
    content = content.replace("而您在XXX方向的深耕", f"而您在{info['research_direction']}方向的深耕")
    content = content.replace("我了解到您在开展XXX与XXX等前沿课题时", f"我了解到您在开展{info['project_1']}与{info['project_2']}等前沿课题时")
    content = content.replace("XXX大学XXX学院在复杂系统控制与智能导航领域拥有国际一流的科研平台", f"{info['university']}{info['school']}在复杂系统控制与智能导航领域拥有国际一流的科研平台")
    content = content.replace("2026年X月XX日", today_str)
    
    return content

def send_email_via_smtp(config, to_email, subject, body_text):
    attachment_path = config["DEFAULT_ATTACHMENT_PATH"]
    has_attachment = os.path.exists(attachment_path)
    
    msg = MIMEMultipart('alternative')
    sender_name = config.get("SENDER_NAME", "您的姓名").strip()
    msg['From'] = f"{sender_name} <{config['SMTP_USER']}>"
    msg['To'] = to_email
    msg['Subject'] = subject
    
    # 转换为 HTML 排版
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
    
    part_text = MIMEText(body_text, 'plain', 'utf-8')
    part_html = MIMEText(html_body, 'html', 'utf-8')
    msg.attach(part_text)
    msg.attach(part_html)
    
    outer_msg = MIMEMultipart('mixed')
    for part in msg.get_payload():
        outer_msg.attach(part)
        
    outer_msg['From'] = msg['From']
    outer_msg['To'] = msg['To']
    outer_msg['Subject'] = msg['Subject']
    
    if has_attachment:
        filename = os.path.basename(attachment_path)
        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                
                # 防止中文附件乱码
                encoded_filename = Header(filename, 'utf-8').encode()
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename=\"{encoded_filename}\"",
                )
                outer_msg.attach(part)
        except Exception:
            pass
        
    try:
        server = smtplib.SMTP_SSL(config["SMTP_SERVER"], int(config["SMTP_PORT"]), timeout=15)
        server.login(config["SMTP_USER"], config["SMTP_PASS"])
        server.sendmail(config["SMTP_USER"], to_email, outer_msg.as_string())
        server.quit()
        return True
    except Exception:
        return False

# ----------------- 启动入口 -----------------

if __name__ == "__main__":
    app = EmailSenderApp()
    app.mainloop()
