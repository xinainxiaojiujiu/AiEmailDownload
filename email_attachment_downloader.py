import win32com.client
import os
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from datetime import datetime, timedelta
import requests
import base64
import re
import json
from pdf2image import convert_from_path
import shutil
import threading
import sys

try:
    import win32timezone
except ImportError:
    pass


AI_PROVIDERS = {
    "阿里百炼 (通义千问)": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
        "models": [
            "qwen3-vl-plus",
            "qwen3-vl-32b",
            "qwen3-vl-flash",
            "qwen-vl-max",
            "qwen-vl-plus",
            "qwen-vl-ocr",
            "qwen3.6-plus",
            "qwen3.5-plus",
        ],
        "key_url": "https://bailian.console.aliyun.com/",
        "key_placeholder": "请输入阿里百炼 API Key (sk-...)",
    },
    "月之暗面 (Kimi)": {
        "base_url": "https://api.moonshot.cn/v1/chat/completions",
        "models": [
            "kimi-k2.6",
            "kimi-k2.5",
        ],
        "key_url": "https://platform.moonshot.cn/console/api-keys",
        "key_placeholder": "请输入 Kimi API Key",
    },
    "智谱AI (GLM)": {
        "base_url": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
        "models": [
            "glm-5v-turbo",
            "glm-4v-plus",
        ],
        "key_url": "https://open.bigmodel.cn/usercenter/apikeys",
        "key_placeholder": "请输入智谱AI API Key",
    },
    "字节豆包 (Doubao)": {
        "base_url": "https://ark.cn-beijing.volces.com/api/v3/chat/completions",
        "models": [
            "doubao-seed-2-0-lite-260215",
        ],
        "key_url": "https://console.volcengine.com/ark/region:ark+cn-beijing/apiKey",
        "key_placeholder": "请输入火山方舟 API Key",
    },
    "MiniMax": {
        "base_url": "https://api.minimax.chat/v1/chat/completions",
        "models": [
            "MiniMax-M3",
            "MiniMax-M2.7",
            "MiniMax-M2.5",
        ],
        "key_url": "https://platform.minimax.io/user-center/basic-information/interface-key",
        "key_placeholder": "请输入 MiniMax API Key",
    },
    "小米 MiMo": {
        "base_url": "https://api.xiaomimimo.com/v1/chat/completions",
        "models": [
            "mimo-v2.5",
            "mimo-v2-flash",
        ],
        "key_url": "https://platform.xiaomimimo.com/",
        "key_placeholder": "请输入 MiMo API Key",
    },
    "百度千帆 (文心)": {
        "base_url": "https://qianfan.baidubce.com/v2/chat/completions",
        "models": [
            "ernie-4.5-8k",
            "ernie-4.0-8k",
        ],
        "key_url": "https://console.bce.baidu.com/qianfan/ais/console/applicationConsole/application",
        "key_placeholder": "请输入百度千帆 API Key",
    },
    "自定义": {
        "base_url": "",
        "models": [],
        "key_url": "",
        "key_placeholder": "请输入 API Key",
    },
}

PROVIDER_NAMES = list(AI_PROVIDERS.keys())


def get_app_base_dir():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_user_config_dir(app_name="OutlookAttachmentDownloader"):
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, app_name)
    os.makedirs(path, exist_ok=True)
    return path


AI_CONFIG_FILE_PATH = os.path.join(get_user_config_dir(), "ai_config.json")
OLD_API_KEY_FILE_PATH = os.path.join(get_user_config_dir(), "api_key.json")


def load_ai_config():
    config = {
        "provider": "阿里通义千问",
        "model": "qwen-vl-max",
        "api_key": "",
        "base_url": "",
    }
    try:
        if os.path.exists(AI_CONFIG_FILE_PATH):
            with open(AI_CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                config["provider"] = data.get("provider", config["provider"])
                config["model"] = data.get("model", config["model"])
                config["api_key"] = data.get("api_key", "")
                config["base_url"] = data.get("base_url", "")
            return config
        if os.path.exists(OLD_API_KEY_FILE_PATH):
            with open(OLD_API_KEY_FILE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                old_key = data.get("api_key", "")
                if old_key and old_key != "sk-xxxxxxxxxxxxxxxxxxxxxxxx":
                    config["api_key"] = old_key
            return config
    except Exception:
        pass
    return config


def save_ai_config(config):
    try:
        with open(AI_CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def mask_api_key(api_key):
    if not api_key or api_key == "sk-xxxxxxxxxxxxxxxxxxxxxxxx":
        return ""
    if len(api_key) <= 8:
        return api_key
    return api_key[:4] + "*" * (len(api_key) - 8) + api_key[-4:]


def resolve_poppler_path():
    env_path = os.environ.get("POPPLER_PATH")
    if env_path:
        return env_path
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundled = os.path.join(sys._MEIPASS, "poppler", "Library", "bin")
        if os.path.exists(bundled):
            return bundled
    exe_dir = get_app_base_dir()
    relative_path = os.path.join(exe_dir, "poppler", "Library", "bin")
    if os.path.exists(relative_path):
        return relative_path
    return r"D:\poppler\Library\bin"


POPPLER_PATH = resolve_poppler_path()


def sanitize_filename(text, max_length=40):
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_length:
        text = text[:max_length]
    if not text:
        text = "未命名"
    return text


def convert_pdf_to_images(pdf_path, poppler_path=None):
    if poppler_path:
        if not os.path.exists(poppler_path):
            raise Exception(f"Poppler路径不存在: {poppler_path}")
        pdftoppm_exe = os.path.join(poppler_path, "pdftoppm.exe")
        if not os.path.exists(pdftoppm_exe):
            raise Exception(f"在 {poppler_path} 下未找到 pdftoppm.exe")
    return convert_from_path(pdf_path, poppler_path=poppler_path)


def call_vision_api(api_key, base_url, model_name, image_paths, prompt):
    if not api_key:
        raise Exception("请先填写 API Key")
    if not base_url:
        raise Exception("请先配置 API 地址")

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    if isinstance(image_paths, str):
        image_paths = [image_paths]

    content = []
    for img_path in image_paths:
        with open(img_path, "rb") as f:
            base64_image = base64.b64encode(f.read()).decode("utf-8")
        mime = "image/png" if img_path.lower().endswith(".png") else "image/jpeg"
        content.append(
            {
                "type": "image_url",
                "image_url": {"url": f"data:{mime};base64,{base64_image}"},
            }
        )
    content.append({"type": "text", "text": prompt})

    payload = {
        "model": model_name,
        "messages": [{"role": "user", "content": content}],
        "max_tokens": 150,
        "temperature": 0.2,
    }

    response = requests.post(base_url, headers=headers, json=payload, timeout=60)
    if response.status_code != 200:
        raise Exception(f"API调用失败 (HTTP {response.status_code}): {response.text[:200]}")

    result = response.json()
    try:
        return result["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError):
        raise Exception(f"API返回格式异常: {result}")


def test_api_connection(api_key, base_url, model_name):
    if not api_key:
        return False, "请先填写 API Key"
    if not base_url:
        return False, "请先配置 API 地址"
    if not model_name:
        return False, "请先选择模型"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model_name,
        "messages": [{"role": "user", "content": "Hi"}],
        "max_tokens": 5,
    }

    response = requests.post(base_url, headers=headers, json=payload, timeout=15)
    if response.status_code == 200:
        return True, f"连接成功 (HTTP 200)"
    else:
        return False, f"连接失败 (HTTP {response.status_code}): {response.text[:100]}"


def generate_filename_with_ai(file_path, api_key, base_url, model_name, max_pages=3):
    ext = os.path.splitext(file_path)[1].lower()
    temp_images = []

    try:
        if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff"]:
            prompt = "请用一句简短的话概括这张图片的核心内容（不超过15个字），不要加任何额外说明，只输出文件名建议。"
            result = call_vision_api(api_key, base_url, model_name, file_path, prompt)
            return sanitize_filename(result)

        elif ext == ".pdf":
            images = convert_pdf_to_images(file_path, poppler_path=POPPLER_PATH)
            if not images:
                raise Exception("PDF无页面")
            images = images[:max_pages]
            temp_dir = os.path.join(os.path.dirname(file_path), "_temp_ai_pages")
            os.makedirs(temp_dir, exist_ok=True)
            image_paths = []
            for i, img in enumerate(images):
                temp_path = os.path.join(temp_dir, f"page_{i + 1}.jpg")
                img.save(temp_path, "JPEG")
                image_paths.append(temp_path)
                temp_images.append(temp_path)
            prompt = "请根据这份PDF文档的主要内容，生成一个简短的文件名（不超过15个字），概括文档主题。直接输出文件名，不要加解释。"
            result = call_vision_api(api_key, base_url, model_name, image_paths, prompt)
            return sanitize_filename(result)

        else:
            return None

    except Exception as e:
        print(f"AI生成文件名失败: {str(e)}")
        return None

    finally:
        for temp_path in temp_images:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
        temp_dir = os.path.join(os.path.dirname(file_path), "_temp_ai_pages")
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass


def start_download_task(
    root, log_text, result_label, status_var, btn_start, entry_widgets, config_vars
):
    def post_to_ui(fn, *args, **kwargs):
        try:
            root.after(0, lambda: fn(*args, **kwargs))
        except Exception:
            pass

    def log_message(msg: str):
        def _append():
            log_text.insert(tk.END, msg)
            log_text.see(tk.END)

        post_to_ui(_append)

    def clear_log():
        post_to_ui(lambda: log_text.delete(1.0, tk.END))

    def update_result(text: str, color: str):
        post_to_ui(lambda: result_label.config(text=text, foreground=color))

    def update_status(text: str):
        post_to_ui(lambda: status_var.set(text))

    def set_button_state(enabled: bool):
        state = "normal" if enabled else "disabled"
        try:
            post_to_ui(lambda: btn_start.configure(state=state))
        except Exception:
            pass

    target_sender = entry_widgets["sender"].get().strip()
    target_subject = entry_widgets["subject"].get().strip()
    save_directory = entry_widgets["save_path"].get().strip()
    active_api_key = config_vars["api_key"]
    active_base_url = config_vars.get("base_url", "")
    ai_enabled = bool(config_vars["ai_enabled"].get())
    selected_model = config_vars["model_combo"].get().strip()

    try:
        search_days = int(entry_widgets["days"].get().strip())
    except Exception:
        update_result("⚠️ 请输入有效数字", COLORS["danger"])
        return

    if not target_sender or not save_directory:
        update_result("⚠️ 发件人名称和保存路径为必填", COLORS["danger"])
        return

    clear_log()
    update_result("", COLORS["text_muted"])
    update_status("正在检索邮件…")
    set_button_state(False)

    def worker():
        download_count = 0
        ai_rename_count = 0

        pythoncom = None
        try:
            try:
                import pythoncom as _pythoncom

                pythoncom = _pythoncom
                pythoncom.CoInitialize()
            except Exception:
                pythoncom = None

            log_message("⏳ 正在连接 Outlook…\n")
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)

            if not os.path.exists(save_directory):
                os.makedirs(save_directory, exist_ok=True)

            items = inbox.Items
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            sender_keyword = target_sender.lower()
            subject_keyword = target_subject.lower() if target_subject else ""

            matched_mails = []
            cutoff_date_naive = datetime.now() - timedelta(days=search_days)
            cutoff_date_aware = None
            cutoff_tz = None

            update_status("正在筛选邮件…")
            for mail in items:
                try:
                    received_time = mail.ReceivedTime
                except Exception:
                    continue

                tz_info = getattr(received_time, "tzinfo", None)
                if tz_info:
                    if cutoff_date_aware is None or tz_info != cutoff_tz:
                        try:
                            cutoff_date_aware = datetime.now(tz_info) - timedelta(
                                days=search_days
                            )
                            cutoff_tz = tz_info
                        except Exception:
                            cutoff_date_aware = None
                            cutoff_tz = None
                    if cutoff_date_aware is not None:
                        effective_cutoff = cutoff_date_aware
                    else:
                        effective_cutoff = cutoff_date_naive
                else:
                    effective_cutoff = cutoff_date_naive

                try:
                    if received_time < effective_cutoff:
                        break
                except Exception:
                    pass

                try:
                    sender_name = str(
                        getattr(mail, "SenderName", "") or ""
                    ).lower()
                    sender_email = str(
                        getattr(mail, "SenderEmailAddress", "") or ""
                    ).lower()
                    if sender_keyword not in sender_name and sender_keyword not in sender_email:
                        continue
                except Exception:
                    continue

                if subject_keyword:
                    try:
                        subject = str(getattr(mail, "Subject", "") or "").lower()
                        if subject_keyword not in subject:
                            continue
                    except Exception:
                        continue

                matched_mails.append(mail)

            if len(matched_mails) == 0:
                log_message("🔍 未找到符合条件的邮件\n")
                update_result("🔍 未找到匹配邮件", COLORS["danger"])
                return

            log_message(f"📩 找到 {len(matched_mails)} 封邮件\n\n")
            has_valid_attachment = False
            update_status("正在保存附件…")

            for mail in matched_mails:
                try:
                    attachments = mail.Attachments
                    att_count = attachments.Count
                except Exception:
                    continue

                if att_count <= 0:
                    continue

                try:
                    received_time = mail.ReceivedTime
                    timestamp = received_time.strftime("%Y年%m月%d日_%H时%M分%S秒")
                except Exception:
                    timestamp = datetime.now().strftime("%Y年%m月%d日_%H时%M分%S秒")

                for att in attachments:
                    try:
                        if att.Size <= 1024 * 10:
                            continue
                    except Exception:
                        continue

                    has_valid_attachment = True
                    try:
                        file_ext = os.path.splitext(att.FileName)[1]
                    except Exception:
                        file_ext = ""

                    original_name = f"{timestamp}{file_ext}"
                    file_path = os.path.join(save_directory, original_name)

                    try:
                        att.SaveAsFile(file_path)
                    except Exception as e:
                        log_message(f"⚠️ 保存失败: {original_name} ({str(e)})\n")
                        continue

                    log_message(f"💾 已保存: {original_name}\n")

                    if ai_enabled and active_api_key and active_api_key != "sk-xxxxxxxxxxxxxxxxxxxxxxxx":
                        log_message("   🤖 正在调用AI识别内容...\n")
                        ai_name = generate_filename_with_ai(
                            file_path, active_api_key, active_base_url, selected_model
                        )
                        if ai_name:
                            new_path = os.path.join(
                                save_directory, f"{ai_name}{file_ext}"
                            )
                            counter = 1
                            while os.path.exists(new_path) and new_path != file_path:
                                new_path = os.path.join(
                                    save_directory, f"{ai_name}({counter}){file_ext}"
                                )
                                counter += 1
                            try:
                                os.rename(file_path, new_path)
                                log_message(
                                    f"   📝 AI重命名成功: {os.path.basename(new_path)}\n"
                                )
                                ai_rename_count += 1
                            except Exception as e:
                                log_message(f"   ⚠️ 重命名失败: {str(e)}\n")
                        else:
                            log_message("   ℹ️ AI未生成有效名称，保留原名\n")
                    else:
                        if not ai_enabled:
                            log_message("   ℹ️ 已关闭AI智能命名，使用原文件名\n")
                        else:
                            log_message("   ℹ️ API Key未配置，使用原文件名\n")

                    download_count += 1

            if not has_valid_attachment:
                log_message("❌ 无有效附件（大小>10KB）\n")
                update_result("❌ 未找到可保存附件", COLORS["danger"])
                return

            update_result(
                f"🎉 完成！共保存 {download_count} 个附件，AI重命名成功 {ai_rename_count} 个",
                COLORS["success"],
            )
            update_status("完成")

        except Exception as e:
            update_result(f"❌ 异常：{str(e)}", COLORS["danger"])
            try:
                import traceback

                log_message(traceback.format_exc())
            except Exception:
                pass
            update_status("异常")
        finally:
            try:
                if pythoncom is not None:
                    pythoncom.CoUninitialize()
            except Exception:
                pass
            set_button_state(True)

    threading.Thread(target=worker, daemon=True).start()


def browse_save_directory(entry_path):
    path = filedialog.askdirectory()
    if path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, path)


def open_directory_in_explorer(entry_path, result_label):
    save_dir = entry_path.get().strip()
    if os.path.exists(save_dir):
        os.startfile(save_dir)
    else:
        result_label.config(text="⚠️ 文件夹不存在", foreground=COLORS["danger"])


COLORS = {
    "bg": "#F8FAFC",
    "bg_secondary": "#F1F5F9",
    "card": "#FFFFFF",
    "text": "#0F172A",
    "text_secondary": "#334155",
    "text_muted": "#64748B",
    "border": "#E2E8F0",
    "border_focus": "#3B82F6",
    "primary": "#3B82F6",
    "primary_hover": "#2563EB",
    "primary_active": "#1D4ED8",
    "primary_light": "#EFF6FF",
    "success": "#10B981",
    "success_hover": "#059669",
    "success_light": "#ECFDF5",
    "warning": "#F59E0B",
    "warning_light": "#FFFBEB",
    "danger": "#EF4444",
    "danger_hover": "#DC2626",
    "danger_light": "#FEF2F2",
    "accent": "#8B5CF6",
    "accent_hover": "#7C3AED",
    "accent_light": "#F5F3FF",
    "info": "#06B6D4",
    "info_light": "#ECFEFF",
    "shadow": "#E2E8F0",
    "highlight": "#DBEAFE",
}


def create_main_window():
    root = tk.Tk()
    root.title("邮箱智能下载助手")
    root.resizable(True, True)

    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    def get_work_area():
        try:
            import ctypes

            class RECT(ctypes.Structure):
                _fields_ = [
                    ("left", ctypes.c_long),
                    ("top", ctypes.c_long),
                    ("right", ctypes.c_long),
                    ("bottom", ctypes.c_long),
                ]

            rect = RECT()
            ctypes.windll.user32.SystemParametersInfoW(
                0x0030, 0, ctypes.byref(rect), 0
            )
            return rect.left, rect.top, rect.right, rect.bottom
        except Exception:
            return None

    def configure_window_geometry():
        work = get_work_area()
        if work:
            wl, wt, wr, wb = work
            screen_w, screen_h = wr - wl, wb - wt
            offset_x, offset_y = wl, wt
        else:
            screen_w = root.winfo_screenwidth()
            screen_h = root.winfo_screenheight()
            offset_x, offset_y = 0, 0
        width = min(1320, max(1000, int(screen_w * 0.85)))
        height = min(860, max(700, int(screen_h * 0.90)))
        pos_x = offset_x + max(0, (screen_w - width) // 2)
        pos_y = offset_y + max(0, (screen_h - height) // 2)
        root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        root.minsize(1000, 750)
        root.configure(bg=COLORS["bg"])

    configure_window_geometry()
    return root


def configure_styles():
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    FONT_FAMILY = "微软雅黑"
    FONT_REGULAR = (FONT_FAMILY, 10)
    FONT_MEDIUM = (FONT_FAMILY, 10, "bold")
    FONT_HEADING = (FONT_FAMILY, 15, "bold")
    FONT_SUBHEADING = (FONT_FAMILY, 10)
    FONT_CAPTION = (FONT_FAMILY, 9)
    FONT_SMALL = (FONT_FAMILY, 8)

    style.configure(".", font=FONT_REGULAR, borderwidth=0)

    style.configure("App.TFrame", background=COLORS["bg"])
    style.configure("Card.TFrame", background=COLORS["card"])
    style.configure("CardInner.TFrame", background=COLORS["card"])

    style.configure(
        "Title.TLabel",
        background=COLORS["bg"],
        foreground=COLORS["text"],
        font=FONT_HEADING,
    )
    style.configure(
        "Subtitle.TLabel",
        background=COLORS["bg"],
        foreground=COLORS["text_muted"],
        font=FONT_SUBHEADING,
    )
    style.configure(
        "Card.TLabel",
        background=COLORS["card"],
        foreground=COLORS["text"],
        font=FONT_REGULAR,
    )
    style.configure(
        "CardMuted.TLabel",
        background=COLORS["card"],
        foreground=COLORS["text_muted"],
        font=FONT_CAPTION,
    )
    style.configure(
        "Status.TLabel",
        background=COLORS["bg"],
        foreground=COLORS["text_muted"],
        font=FONT_CAPTION,
    )
    style.configure(
        "FieldLabel.TLabel",
        background=COLORS["card"],
        foreground=COLORS["text_secondary"],
        font=FONT_MEDIUM,
    )

    style.configure(
        "Card.TLabelframe",
        background=COLORS["card"],
        padding=16,
        borderwidth=1,
        bordercolor=COLORS["border"],
        relief="solid",
    )
    style.configure(
        "Card.TLabelframe.Label",
        background=COLORS["card"],
        foreground=COLORS["text"],
        font=FONT_MEDIUM,
        padding=(8, 0),
    )

    style.configure(
        "TEntry",
        padding=(8, 6),
        fieldbackground="#FFFFFF",
        borderwidth=1,
        bordercolor=COLORS["border"],
        relief="solid",
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", COLORS["border_focus"])],
    )
    style.configure(
        "TCombobox",
        padding=(8, 5),
        fieldbackground="#FFFFFF",
        borderwidth=1,
        bordercolor=COLORS["border"],
        relief="solid",
        arrowsize=12,
    )
    style.map(
        "TCombobox",
        bordercolor=[("focus", COLORS["border_focus"])],
    )

    style.configure(
        "Primary.TButton",
        padding=(20, 10),
        background=COLORS["primary"],
        foreground="#FFFFFF",
        borderwidth=0,
        font=(FONT_FAMILY, 10, "bold"),
    )
    style.map(
        "Primary.TButton",
        background=[
            ("active", COLORS["primary_active"]),
            ("pressed", COLORS["primary_active"]),
            ("disabled", COLORS["border"]),
        ],
        foreground=[("disabled", COLORS["text_muted"])],
    )

    style.configure(
        "Secondary.TButton",
        padding=(10, 6),
        background=COLORS["bg_secondary"],
        foreground=COLORS["text_secondary"],
        borderwidth=1,
        bordercolor=COLORS["border"],
        relief="solid",
    )
    style.map(
        "Secondary.TButton",
        background=[("active", COLORS["border"]), ("pressed", COLORS["border"])],
        bordercolor=[("active", COLORS["text_muted"])],
    )

    style.configure(
        "Accent.TButton",
        padding=(10, 6),
        background=COLORS["accent"],
        foreground="#FFFFFF",
        borderwidth=0,
    )
    style.map(
        "Accent.TButton",
        background=[
            ("active", COLORS["accent_hover"]),
            ("pressed", COLORS["accent_hover"]),
        ],
    )

    style.configure(
        "Success.TButton",
        padding=(10, 6),
        background=COLORS["success"],
        foreground="#FFFFFF",
        borderwidth=0,
    )
    style.map(
        "Success.TButton",
        background=[
            ("active", COLORS["success_hover"]),
            ("pressed", COLORS["success_hover"]),
        ],
    )

    style.configure(
        "Danger.TButton",
        padding=(10, 6),
        background=COLORS["danger"],
        foreground="#FFFFFF",
        borderwidth=0,
    )
    style.map(
        "Danger.TButton",
        background=[
            ("active", COLORS["danger_hover"]),
            ("pressed", COLORS["danger_hover"]),
        ],
    )

    style.configure(
        "Link.TButton",
        padding=(4, 2),
        background=COLORS["card"],
        foreground=COLORS["primary"],
        borderwidth=0,
        font=(FONT_FAMILY, 9),
    )
    style.map(
        "Link.TButton",
        foreground=[("active", COLORS["primary_hover"])],
        background=[("active", COLORS["primary_light"])],
    )

    return style


def build_ui(root):
    configure_styles()

    FONT_FAMILY = "微软雅黑"
    FONT_REGULAR = (FONT_FAMILY, 10)
    FONT_MEDIUM = (FONT_FAMILY, 10, "bold")
    FONT_CAPTION = (FONT_FAMILY, 9)

    app_frame = ttk.Frame(root, style="App.TFrame")
    app_frame.pack(fill=tk.BOTH, expand=True)
    app_frame.columnconfigure(0, weight=1)
    app_frame.rowconfigure(0, weight=0)
    app_frame.rowconfigure(1, weight=1)
    app_frame.rowconfigure(2, weight=0)

    header_frame = ttk.Frame(app_frame, style="App.TFrame", padding=(24, 16, 24, 8))
    header_frame.grid(row=0, column=0, sticky="ew")

    title_frame = ttk.Frame(header_frame, style="App.TFrame")
    title_frame.pack(side=tk.LEFT)

    ttk.Label(title_frame, text="📧 邮箱智能下载助手", style="Title.TLabel").pack(
        anchor="w"
    )
    ttk.Label(
        title_frame,
        text="批量下载 Outlook 邮件附件 · 支持 AI 智能命名",
        style="Subtitle.TLabel",
    ).pack(anchor="w", pady=(2, 0))

    body_frame = ttk.Frame(app_frame, style="App.TFrame", padding=(24, 8, 24, 8))
    body_frame.grid(row=1, column=0, sticky="nsew")
    body_frame.columnconfigure(0, weight=0, minsize=540)
    body_frame.columnconfigure(1, weight=1)
    body_frame.rowconfigure(0, weight=1)

    left_panel = ttk.Frame(body_frame, style="App.TFrame")
    left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
    left_panel.columnconfigure(0, weight=1)
    left_panel.rowconfigure(0, weight=1)
    left_panel.rowconfigure(1, weight=0)

    param_frame = ttk.LabelFrame(
        left_panel, text="  📋  下载参数  ", style="Card.TLabelframe"
    )
    param_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 12))
    param_frame.columnconfigure(1, weight=1)

    row_idx = 0
    FIELD_PAD_Y = (4, 4)

    ttk.Label(param_frame, text="发件人名称", style="FieldLabel.TLabel").grid(
        row=row_idx, column=0, sticky="w", padx=(4, 12), pady=FIELD_PAD_Y
    )
    entry_sender = ttk.Entry(param_frame, font=FONT_REGULAR)
    entry_sender.grid(
        row=row_idx, column=1, columnspan=2, sticky="ew", padx=(0, 8), pady=FIELD_PAD_Y
    )
    entry_sender.insert(0, "绿色办公")
    ttk.Label(param_frame, text="必填，匹配发件人姓名或邮箱", style="CardMuted.TLabel").grid(
        row=row_idx, column=3, sticky="w", pady=FIELD_PAD_Y
    )
    row_idx += 1

    ttk.Label(param_frame, text="主题关键词", style="FieldLabel.TLabel").grid(
        row=row_idx, column=0, sticky="w", padx=(4, 12), pady=FIELD_PAD_Y
    )
    entry_subject = ttk.Entry(param_frame, font=FONT_REGULAR)
    entry_subject.grid(
        row=row_idx, column=1, columnspan=2, sticky="ew", padx=(0, 8), pady=FIELD_PAD_Y
    )
    ttk.Label(param_frame, text="可选，不区分大小写", style="CardMuted.TLabel").grid(
        row=row_idx, column=3, sticky="w", pady=FIELD_PAD_Y
    )
    row_idx += 1

    ttk.Label(param_frame, text="保存路径", style="FieldLabel.TLabel").grid(
        row=row_idx, column=0, sticky="w", padx=(4, 12), pady=FIELD_PAD_Y
    )
    entry_path = ttk.Entry(param_frame, font=FONT_REGULAR)
    entry_path.grid(row=row_idx, column=1, sticky="ew", padx=(0, 8), pady=FIELD_PAD_Y)
    entry_path.insert(0, r"D:\重要文件\Desktop\扫描保存的文件夹")
    path_btn_frame = ttk.Frame(param_frame, style="Card.TFrame")
    path_btn_frame.grid(row=row_idx, column=2, sticky="w", pady=FIELD_PAD_Y)
    ttk.Button(
        path_btn_frame,
        text="浏览…",
        style="Secondary.TButton",
        command=lambda: browse_save_directory(entry_path),
    ).pack(side=tk.LEFT, padx=(0, 4))
    ttk.Button(
        path_btn_frame,
        text="打开目录",
        style="Secondary.TButton",
        command=lambda: open_directory_in_explorer(entry_path, result_label),
    ).pack(side=tk.LEFT)
    ttk.Label(param_frame, text="必填", style="CardMuted.TLabel").grid(
        row=row_idx, column=3, sticky="w", padx=(4, 0), pady=FIELD_PAD_Y
    )
    row_idx += 1

    ttk.Label(param_frame, text="检索天数", style="FieldLabel.TLabel").grid(
        row=row_idx, column=0, sticky="w", padx=(4, 12), pady=FIELD_PAD_Y
    )
    entry_days = ttk.Entry(param_frame, width=10, font=FONT_REGULAR)
    entry_days.grid(row=row_idx, column=1, sticky="w", padx=(0, 8), pady=FIELD_PAD_Y)
    entry_days.insert(0, "1")
    ttk.Label(param_frame, text="默认 1 天，首次使用建议 7 或 30", style="CardMuted.TLabel").grid(
        row=row_idx, column=2, columnspan=2, sticky="w", pady=FIELD_PAD_Y
    )
    row_idx += 1

    ttk.Separator(param_frame).grid(
        row=row_idx, column=0, columnspan=4, sticky="ew", pady=(8, 8)
    )
    row_idx += 1

    enable_ai_var = tk.BooleanVar(value=True)
    ai_status_var = tk.StringVar(
        value="AI 智能命名：已开启（根据附件内容自动重命名）"
    )

    ai_toggle_frame = ttk.Frame(param_frame, style="Card.TFrame")
    ai_toggle_frame.grid(
        row=row_idx, column=0, columnspan=4, sticky="ew", pady=(0, 6)
    )
    ai_toggle_frame.columnconfigure(0, weight=1)
    ttk.Label(ai_toggle_frame, textvariable=ai_status_var, style="CardMuted.TLabel").grid(
        row=0, column=0, sticky="w", padx=(4, 8)
    )
    btn_ai_toggle = ttk.Button(ai_toggle_frame, text="已开启", style="Success.TButton")
    btn_ai_toggle.grid(row=0, column=1, sticky="e")
    row_idx += 1

    ai_config = load_ai_config()
    active_api_key_ref = [ai_config["api_key"]]

    ttk.Label(param_frame, text="厂商选择", style="FieldLabel.TLabel").grid(
        row=row_idx, column=0, sticky="w", padx=(4, 12), pady=FIELD_PAD_Y
    )
    combo_provider = ttk.Combobox(
        param_frame,
        width=20,
        values=PROVIDER_NAMES,
        state="readonly",
        font=FONT_REGULAR,
    )
    combo_provider.grid(row=row_idx, column=1, sticky="w", padx=(0, 8), pady=FIELD_PAD_Y)

    api_key_portal_url = [AI_PROVIDERS.get(ai_config["provider"], {}).get("key_url", "")]

    btn_apply_api = ttk.Button(
        param_frame,
        text="申请 Key",
        style="Secondary.TButton",
    )
    btn_apply_api.grid(row=row_idx, column=2, sticky="w", pady=FIELD_PAD_Y)
    ttk.Label(param_frame, text="", style="CardMuted.TLabel").grid(
        row=row_idx, column=3, pady=FIELD_PAD_Y
    )
    row_idx += 1

    ttk.Label(param_frame, text="模型选择", style="FieldLabel.TLabel").grid(
        row=row_idx, column=0, sticky="w", padx=(4, 12), pady=FIELD_PAD_Y
    )
    combo_model = ttk.Combobox(
        param_frame,
        width=24,
        state="readonly",
        font=FONT_REGULAR,
    )
    combo_model.grid(row=row_idx, column=1, sticky="w", padx=(0, 8), pady=FIELD_PAD_Y)
    model_hint_var = tk.StringVar(value="请选择支持图片识别的视觉模型")
    ttk.Label(param_frame, textvariable=model_hint_var, style="CardMuted.TLabel").grid(
        row=row_idx, column=2, columnspan=2, sticky="w", pady=FIELD_PAD_Y
    )
    row_idx += 1

    ttk.Label(param_frame, text="API 地址", style="FieldLabel.TLabel").grid(
        row=row_idx, column=0, sticky="w", padx=(4, 12), pady=FIELD_PAD_Y
    )
    entry_base_url = ttk.Entry(param_frame, font=FONT_REGULAR)
    entry_base_url.grid(row=row_idx, column=1, columnspan=3, sticky="ew", padx=(0, 8), pady=FIELD_PAD_Y)
    row_idx += 1

    PLACEHOLDER_TEXT = "请输入 API Key"

    ttk.Label(param_frame, text="API Key", style="FieldLabel.TLabel").grid(
        row=row_idx, column=0, sticky="w", padx=(4, 12), pady=FIELD_PAD_Y
    )
    entry_apikey = ttk.Entry(param_frame, font=FONT_REGULAR)
    entry_apikey.grid(row=row_idx, column=1, sticky="ew", padx=(0, 8), pady=FIELD_PAD_Y)

    if active_api_key_ref[0]:
        entry_apikey.insert(0, mask_api_key(active_api_key_ref[0]))
        entry_apikey.configure(foreground=COLORS["text"])
    else:
        entry_apikey.configure(foreground=COLORS["text_muted"])
        entry_apikey.insert(0, PLACEHOLDER_TEXT)

    def on_apikey_focus_in(event):
        current = entry_apikey.get()
        if current == PLACEHOLDER_TEXT:
            entry_apikey.delete(0, tk.END)
            entry_apikey.configure(foreground=COLORS["text"])
            entry_apikey.configure(show="*")
        elif mask_api_key(active_api_key_ref[0]) == current:
            entry_apikey.delete(0, tk.END)
            entry_apikey.configure(show="*")
            entry_apikey.insert(0, active_api_key_ref[0])

    def on_apikey_focus_out(event):
        current = entry_apikey.get()
        if not current:
            entry_apikey.configure(foreground=COLORS["text_muted"])
            entry_apikey.insert(0, PLACEHOLDER_TEXT)
            entry_apikey.configure(show="")
            active_api_key_ref[0] = ""
        else:
            active_api_key_ref[0] = current
            entry_apikey.delete(0, tk.END)
            entry_apikey.configure(show="")
            entry_apikey.insert(0, mask_api_key(active_api_key_ref[0]))

    entry_apikey.bind("<FocusIn>", on_apikey_focus_in)
    entry_apikey.bind("<FocusOut>", on_apikey_focus_out)

    api_btn_frame = ttk.Frame(param_frame, style="Card.TFrame")
    api_btn_frame.grid(row=row_idx, column=2, sticky="w", pady=FIELD_PAD_Y)
    btn_save_api = ttk.Button(api_btn_frame, text="保存配置", style="Accent.TButton")
    btn_save_api.pack(side=tk.LEFT, padx=(0, 4))
    btn_test_api = ttk.Button(api_btn_frame, text="测试连接", style="Secondary.TButton")
    btn_test_api.pack(side=tk.LEFT)
    ttk.Label(param_frame, text="", style="CardMuted.TLabel").grid(
        row=row_idx, column=3, pady=FIELD_PAD_Y
    )
    row_idx += 1

    test_status_var = tk.StringVar(value="")
    test_status_label = ttk.Label(
        param_frame, textvariable=test_status_var, style="CardMuted.TLabel"
    )
    test_status_label.grid(row=row_idx, column=0, columnspan=4, sticky="w", padx=(4, 8), pady=(0, 4))
    row_idx += 1

    def on_provider_change(event=None):
        provider = combo_provider.get()
        if not provider or provider not in AI_PROVIDERS:
            return
        prov = AI_PROVIDERS[provider]
        models = prov["models"]
        combo_model["values"] = models
        if models:
            combo_model.set(models[0])
            combo_model.configure(state="readonly")
        else:
            combo_model.set("")
            combo_model.configure(state="normal")
        entry_base_url.delete(0, tk.END)
        entry_base_url.insert(0, prov["base_url"])
        api_key_portal_url[0] = prov["key_url"]
        current_key = entry_apikey.get()
        if current_key == PLACEHOLDER_TEXT or not active_api_key_ref[0]:
            entry_apikey.delete(0, tk.END)
            entry_apikey.configure(foreground=COLORS["text_muted"])
            entry_apikey.insert(0, prov["key_placeholder"])
        test_status_var.set("")

    combo_provider.bind("<<ComboboxSelected>>", on_provider_change)
    combo_provider.set(ai_config["provider"])
    on_provider_change()

    if ai_config["model"]:
        current_models = list(combo_model["values"])
        if ai_config["model"] in current_models:
            combo_model.set(ai_config["model"])

    def get_base_url():
        return entry_base_url.get().strip()

    def open_api_key_portal():
        import webbrowser
        url = api_key_portal_url[0]
        if url:
            webbrowser.open(url)

    btn_apply_api.configure(command=open_api_key_portal)

    def save_current_config():
        current_text = entry_apikey.get().strip()
        if current_text and current_text != PLACEHOLDER_TEXT and current_text != AI_PROVIDERS.get(
            combo_provider.get(), {}
        ).get("key_placeholder", ""):
            if current_text != mask_api_key(active_api_key_ref[0]):
                active_api_key_ref[0] = current_text

        cfg = {
            "provider": combo_provider.get(),
            "model": combo_model.get(),
            "api_key": active_api_key_ref[0],
            "base_url": get_base_url(),
        }
        if save_ai_config(cfg):
            result_label.config(text="✅ 配置保存成功", foreground=COLORS["success"])
        else:
            result_label.config(text="❌ 配置保存失败", foreground=COLORS["danger"])

    btn_save_api.configure(command=save_current_config)

    def do_test_connection():
        key = active_api_key_ref[0]
        if not key:
            current_text = entry_apikey.get().strip()
            if current_text and current_text != PLACEHOLDER_TEXT:
                key = current_text
        url = get_base_url()
        model = combo_model.get().strip()

        def _run():
            try:
                post_to_ui = lambda fn, *a: root.after(0, lambda: fn(*a))
                post_to_ui(lambda: test_status_var.set("⏳ 正在测试连接..."))
                post_to_ui(lambda: test_status_label.configure(foreground=COLORS["text_muted"]))
                ok, msg = test_api_connection(key, url, model)
                if ok:
                    post_to_ui(lambda: test_status_var.set(f"✅ {msg}"))
                    post_to_ui(lambda: test_status_label.configure(foreground=COLORS["success"]))
                else:
                    post_to_ui(lambda: test_status_var.set(f"❌ {msg}"))
                    post_to_ui(lambda: test_status_label.configure(foreground=COLORS["danger"]))
            except Exception as e:
                post_to_ui(lambda: test_status_var.set(f"❌ 测试异常: {str(e)}"))
                post_to_ui(lambda: test_status_label.configure(foreground=COLORS["danger"]))

        threading.Thread(target=_run, daemon=True).start()

    btn_test_api.configure(command=do_test_connection)

    def update_ai_controls():
        enabled = bool(enable_ai_var.get())
        ai_widgets = [
            combo_provider, combo_model, entry_base_url, entry_apikey,
            btn_save_api, btn_test_api, btn_apply_api,
        ]
        if enabled:
            btn_ai_toggle.configure(text="已开启", style="Success.TButton")
            ai_status_var.set("AI 智能命名：已开启（根据附件内容自动重命名）")
            for w in ai_widgets:
                try:
                    if w == combo_provider:
                        w.configure(state="readonly")
                    elif w == combo_model:
                        provider = combo_provider.get()
                        if provider == "自定义":
                            w.configure(state="normal")
                        else:
                            w.configure(state="readonly")
                    else:
                        w.configure(state="normal")
                except Exception:
                    pass
        else:
            btn_ai_toggle.configure(text="已关闭", style="Danger.TButton")
            ai_status_var.set("AI 智能命名：已关闭（仅按时间保存附件）")
            for w in ai_widgets:
                try:
                    w.configure(state="disabled")
                except Exception:
                    pass

    def toggle_ai():
        enable_ai_var.set(not bool(enable_ai_var.get()))
        update_ai_controls()

    btn_ai_toggle.configure(command=toggle_ai)
    update_ai_controls()

    action_frame = ttk.Frame(left_panel, style="App.TFrame")
    action_frame.grid(row=1, column=0, sticky="sew")
    action_frame.columnconfigure(0, weight=1)

    btn_start = ttk.Button(
        action_frame,
        text="▶  开始下载附件",
        style="Primary.TButton",
    )
    btn_start.grid(row=0, column=0, sticky="ew", pady=(0, 8))

    result_label = ttk.Label(action_frame, text="", style="Status.TLabel")
    result_label.grid(row=1, column=0, sticky="w")

    log_frame = ttk.LabelFrame(
        body_frame, text="  📜  下载日志  ", style="Card.TLabelframe"
    )
    log_frame.grid(row=0, column=1, sticky="nsew")

    log_text = scrolledtext.ScrolledText(
        log_frame,
        font=(FONT_FAMILY, 10),
        bg=COLORS["card"],
        fg=COLORS["text"],
        relief=tk.FLAT,
        padx=12,
        pady=10,
        insertbackground=COLORS["primary"],
        selectbackground=COLORS["highlight"],
        selectforeground=COLORS["text"],
        wrap=tk.WORD,
        borderwidth=0,
    )
    log_text.pack(fill=tk.BOTH, expand=True)

    log_text.tag_configure("success", foreground=COLORS["success"])
    log_text.tag_configure("error", foreground=COLORS["danger"])
    log_text.tag_configure("warning", foreground=COLORS["warning"])
    log_text.tag_configure("info", foreground=COLORS["primary"])

    footer_frame = ttk.Frame(app_frame, style="App.TFrame", padding=(24, 8, 24, 12))
    footer_frame.grid(row=2, column=0, sticky="ew")
    footer_frame.columnconfigure(0, weight=0)
    footer_frame.columnconfigure(1, weight=1)

    status_var = tk.StringVar(value="就绪")
    ttk.Label(footer_frame, textvariable=status_var, style="Status.TLabel").grid(
        row=0, column=0, sticky="w"
    )

    guide_visible = tk.BooleanVar(value=False)

    guide_container = ttk.Frame(footer_frame, style="App.TFrame")
    guide_container.grid(row=0, column=1, sticky="e")

    def toggle_guide():
        if guide_visible.get():
            guide_card_frame.grid_remove()
            guide_visible.set(False)
            btn_guide.configure(text="📖 显示使用指南")
        else:
            guide_card_frame.grid()
            guide_visible.set(True)
            btn_guide.configure(text="📖 收起指南")

    btn_guide = ttk.Button(
        guide_container,
        text="📖 显示使用指南",
        style="Link.TButton",
        command=toggle_guide,
    )
    btn_guide.grid(row=0, column=0, sticky="e")

    guide_card_frame = tk.Frame(
        footer_frame,
        bg=COLORS["primary_light"],
        highlightbackground=COLORS["primary"],
        highlightthickness=1,
        padx=2,
        pady=2,
    )
    guide_card_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

    guide_inner = tk.Frame(guide_card_frame, bg=COLORS["primary_light"])
    guide_inner.pack(fill=tk.X, padx=8, pady=6)

    tk.Label(
        guide_inner,
        text="📖  新手使用指南",
        font=(FONT_FAMILY, 10, "bold"),
        bg=COLORS["primary_light"],
        fg=COLORS["primary"],
        anchor="w",
    ).pack(fill=tk.X, pady=(0, 6))

    guide_steps = [
        ("❶", "发件人名称", "填写邮件发件人名称或邮箱中的关键字（必填）"),
        ("❷", "主题关键词", "可留空，填写后只下载主题包含该词的邮件附件"),
        ("❸", "保存路径", "点击「浏览…」选择附件保存的文件夹"),
        ("❹", "检索天数", "默认 1 天；首次使用建议调到 7 或 30 天"),
        ("❺", "AI 配置", "选择厂商和模型 → 填写 API Key → 点「测试连接」→ 点「保存配置」"),
        ("❻", "开始下载", "点击蓝色按钮「开始下载附件」，右侧日志区实时显示进度"),
    ]

    for step_icon, step_title, step_desc in guide_steps:
        step_line = tk.Frame(guide_inner, bg=COLORS["primary_light"])
        step_line.pack(fill=tk.X, pady=1)
        tk.Label(
            step_line,
            text=f"{step_icon} {step_title}：",
            font=(FONT_FAMILY, 9, "bold"),
            bg=COLORS["primary_light"],
            fg=COLORS["text_secondary"],
            anchor="w",
            width=14,
        ).pack(side=tk.LEFT)
        tk.Label(
            step_line,
            text=step_desc,
            font=(FONT_FAMILY, 9),
            bg=COLORS["primary_light"],
            fg=COLORS["text_muted"],
            anchor="w",
        ).pack(side=tk.LEFT)
    guide_card_frame.grid_remove()

    entry_widgets = {
        "sender": entry_sender,
        "subject": entry_subject,
        "save_path": entry_path,
        "days": entry_days,
    }
    config_vars = {
        "api_key": active_api_key_ref[0],
        "base_url": get_base_url(),
        "ai_enabled": enable_ai_var,
        "model_combo": combo_model,
    }

    def on_start_click():
        config_vars["api_key"] = active_api_key_ref[0]
        config_vars["base_url"] = get_base_url()
        start_download_task(
            root,
            log_text,
            result_label,
            status_var,
            btn_start,
            entry_widgets,
            config_vars,
        )

    btn_start.configure(command=on_start_click)

    return root


def main():
    root = create_main_window()
    build_ui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
