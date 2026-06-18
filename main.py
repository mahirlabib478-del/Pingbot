import os
import time
import threading
import requests
import datetime
import json
import io
import gzip
import uuid
import logging
from flask import Flask
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# ================== LOGGING ==================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ================== CONFIG ==================
BOT_TOKEN = "8808046131:AAHCgB22O9KtwtIKrfXpMOBrPZRzNvN-3oo"
ADMIN_CHAT_ID = "2035024902"
CHANNEL_ID = "-1003903695158"

if not BOT_TOKEN or not ADMIN_CHAT_ID:
    raise RuntimeError("BOT_TOKEN and ADMIN_CHAT_ID must be set")

# ================== FILE PATHS ==================
MOTHER_FILE = "mother_accounts.json"
COOLDOWN_FILE = "user_cooldowns.json"
SUBSCRIBERS_FILE = "subscribers.json"
USER_INFO_FILE = "user_info.json"
ACCOUNTS_FILE = "accounts.json"
BALANCES_FILE = "balances.json"
DEPOSITS_FILE = "deposits.json"
CONFIG_FILE = "config.json"
SELL_REQUESTS_FILE = "sell_requests.json"
WITHDRAW_REQUESTS_FILE = "withdraw_requests.json"

app = Flask(__name__)

# ================== GLOBALS ==================
last_update_id = None
subscribed_users = set()
user_info = {}
mother_accounts = []
user_last_request = {}
submission_sessions = {}
support_sessions = set()
maintenance_mode = False

accounts = []               # each dict contains 'type', 'reuse_link' (for normal)
balances = {}
deposits = []
config = {
    "bkash_number": "",
    "price_2fa_buy": 2.2,
    "price_normal_buy": 1.2,
    "price_2fa_sell": 1.8,
    "price_normal_sell": 0.8,
    "group_chat_id": "",
    "channel_id": str(CHANNEL_ID) if CHANNEL_ID else "",
    "maintenance_mode": False
}
deposit_sessions = {}
buy_sessions = {}
add_stock_sessions = {}
loss_recovery_sessions = {}
sell_sessions = {}
sell_requests = []
withdraw_requests = []
withdraw_sessions = {}
last_backup_message_id = None

data_lock = threading.RLock()
backup_lock = threading.Lock()

# ================== HELPER: BENGALI NUMBER CONVERTER ==================
def normalize_bengali_number(text):
    bengali_digits = "০১২৩৪৫৬৭৮৯"
    english_digits = "0123456789"
    trans_table = str.maketrans(bengali_digits, english_digits)
    return text.translate(trans_table)

# ================== FILE I/O ==================
def load_mother_accounts():
    global mother_accounts
    try:
        with open(MOTHER_FILE, "r") as f:
            mother_accounts = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        mother_accounts = []

def save_mother_accounts():
    with data_lock:
        try:
            with open(MOTHER_FILE, "w") as f:
                json.dump(mother_accounts, f, indent=2)
        except IOError as e:
            logger.error(f"Mother save error: {e}")

def load_user_cooldowns():
    global user_last_request
    try:
        with open(COOLDOWN_FILE, "r") as f:
            user_last_request = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        user_last_request = {}

def save_user_cooldowns():
    with data_lock:
        try:
            with open(COOLDOWN_FILE, "w") as f:
                json.dump(user_last_request, f, indent=2)
        except IOError as e:
            logger.error(f"Cooldown save error: {e}")

def load_subscribers():
    global subscribed_users, user_info
    try:
        with open(SUBSCRIBERS_FILE, "r") as f:
            data = json.load(f)
            subscribed_users = set(data.get("subscribed", []))
    except (FileNotFoundError, json.JSONDecodeError):
        subscribed_users = set()
    try:
        with open(USER_INFO_FILE, "r") as f:
            user_info = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        user_info = {}

def save_subscribers():
    with data_lock:
        try:
            with open(SUBSCRIBERS_FILE, "w") as f:
                json.dump({"subscribed": list(subscribed_users)}, f)
        except IOError as e:
            logger.error(f"Subscribers save error: {e}")

def save_user_info():
    with data_lock:
        try:
            with open(USER_INFO_FILE, "w") as f:
                json.dump(user_info, f, indent=2)
        except IOError as e:
            logger.error(f"User info save error: {e}")

def load_market():
    global accounts, balances, deposits, config, CHANNEL_ID, maintenance_mode
    try:
        with open(ACCOUNTS_FILE, "r") as f:
            accounts = json.load(f)
            for acc in accounts:
                if 'type' not in acc:
                    acc['type'] = '2fa'
                if 'reuse_link' not in acc:
                    acc['reuse_link'] = ''
    except (FileNotFoundError, json.JSONDecodeError):
        accounts = []
    try:
        with open(BALANCES_FILE, "r") as f:
            balances = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        balances = {}
    try:
        with open(DEPOSITS_FILE, "r") as f:
            deposits = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        deposits = []
    try:
        with open(CONFIG_FILE, "r") as f:
            loaded_config = json.load(f)
            for key in config:
                if key not in loaded_config:
                    loaded_config[key] = config[key]
            config = loaded_config
            CHANNEL_ID = int(config.get("channel_id", "0")) if config.get("channel_id") else None
            maintenance_mode = config.get("maintenance_mode", False)
    except (FileNotFoundError, json.JSONDecodeError):
        config["channel_id"] = str(CHANNEL_ID) if CHANNEL_ID else ""
        maintenance_mode = False
        CHANNEL_ID = int(config["channel_id"]) if config["channel_id"] else None

def save_accounts():
    with data_lock:
        try:
            with open(ACCOUNTS_FILE, "w") as f:
                json.dump(accounts, f, indent=2)
        except IOError as e:
            logger.error(f"Accounts save error: {e}")

def save_balances():
    with data_lock:
        try:
            with open(BALANCES_FILE, "w") as f:
                json.dump(balances, f, indent=2)
        except IOError as e:
            logger.error(f"Balances save error: {e}")

def save_deposits():
    with data_lock:
        try:
            with open(DEPOSITS_FILE, "w") as f:
                json.dump(deposits, f, indent=2)
        except IOError as e:
            logger.error(f"Deposits save error: {e}")

def save_config():
    with data_lock:
        config["maintenance_mode"] = maintenance_mode
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(config, f, indent=2)
        except IOError as e:
            logger.error(f"Config save error: {e}")

def load_sell_requests():
    global sell_requests
    try:
        with open(SELL_REQUESTS_FILE, "r") as f:
            sell_requests = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        sell_requests = []

def save_sell_requests():
    with data_lock:
        try:
            with open(SELL_REQUESTS_FILE, "w") as f:
                json.dump(sell_requests, f, indent=2)
        except IOError as e:
            logger.error(f"Sell requests save error: {e}")

def load_withdraw_requests():
    global withdraw_requests
    try:
        with open(WITHDRAW_REQUESTS_FILE, "r") as f:
            withdraw_requests = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        withdraw_requests = []

def save_withdraw_requests():
    with data_lock:
        try:
            with open(WITHDRAW_REQUESTS_FILE, "w") as f:
                json.dump(withdraw_requests, f, indent=2)
        except IOError as e:
            logger.error(f"Withdraw requests save error: {e}")

def save_all():
    save_accounts()
    save_balances()
    save_deposits()
    save_config()
    save_subscribers()
    save_user_info()
    save_mother_accounts()
    save_user_cooldowns()
    save_sell_requests()
    save_withdraw_requests()
    save_data_to_channel()

# ================== TELEGRAM HELPERS ==================
def send_telegram_message(text, chat_id, reply_markup=None, parse_mode=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {"chat_id": chat_id, "text": text}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    if parse_mode:
        payload["parse_mode"] = parse_mode
    for attempt in range(3):
        try:
            resp = requests.post(url, json=payload, timeout=10)
            if resp.status_code == 429:
                retry_after = resp.json().get("parameters", {}).get("retry_after", 2)
                time.sleep(retry_after)
                continue
            return resp
        except Exception as e:
            logger.error(f"Send error to {chat_id} (attempt {attempt+1}): {e}")
            if attempt == 2:
                return None
            time.sleep(1)
    return None

def send_telegram_document(file_bytes, filename, chat_id, caption=""):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    try:
        files = {'document': (filename, file_bytes,
                              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        resp = requests.post(url, data={"chat_id": chat_id, "caption": caption}, files=files, timeout=30)
        return resp.status_code == 200 and resp.json().get("ok", False)
    except Exception as e:
        logger.error(f"Document send error: {e}")
        return False

def delete_telegram_message(chat_id, message_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/deleteMessage"
    try:
        requests.post(url, json={"chat_id": chat_id, "message_id": message_id}, timeout=10)
    except Exception as e:
        logger.error(f"Delete message error: {e}")

def forward_telegram_document(chat_id, from_chat_id, message_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/forwardMessage"
    payload = {"chat_id": chat_id, "from_chat_id": from_chat_id, "message_id": message_id}
    try:
        requests.post(url, json=payload, timeout=10)
    except Exception as e:
        logger.error(f"Forward error: {e}")

def broadcast_message(text):
    to_remove = []
    for chat_id in list(subscribed_users):
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
        try:
            resp = requests.post(url, json={"chat_id": chat_id, "text": text}, timeout=5)
            if resp.status_code == 403:
                to_remove.append(chat_id)
            elif resp.status_code == 429:
                retry_after = resp.json().get("parameters", {}).get("retry_after", 1)
                time.sleep(retry_after)
                requests.post(url, json={"chat_id": chat_id, "text": text}, timeout=5)
        except Exception as e:
            logger.error(f"Broadcast to {chat_id} failed: {e}")
            to_remove.append(chat_id)
        time.sleep(0.05)
    if to_remove:
        with data_lock:
            for uid in to_remove:
                subscribed_users.discard(uid)
                user_info.pop(uid, None)
        save_subscribers()
        save_user_info()

def answer_callback_query(callback_id, text=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/answerCallbackQuery"
    payload = {"callback_query_id": callback_id}
    if text:
        payload["text"] = text
    try:
        requests.post(url, json=payload, timeout=5)
    except Exception as e:
        logger.error(f"Callback answer error: {e}")

# ================== CHANNEL BACKUP ==================
def save_data_to_channel():
    global last_backup_message_id
    if not CHANNEL_ID:
        return
    with backup_lock:
        try:
            with data_lock:
                data = {
                    "accounts": accounts,
                    "balances": balances,
                    "deposits": deposits,
                    "config": config,
                    "subscribed_users": list(subscribed_users),
                    "user_info": user_info,
                    "mother_accounts": mother_accounts,
                    "user_cooldowns": user_last_request,
                    "sell_requests": sell_requests,
                    "withdraw_requests": withdraw_requests,
                    "timestamp": datetime.datetime.now().isoformat()
                }
            json_bytes = json.dumps(data, indent=2, ensure_ascii=False).encode('utf-8')
            compressed = gzip.compress(json_bytes, compresslevel=6)
            max_size = 48 * 1024 * 1024
            if len(compressed) > max_size:
                logger.warning("Compressed backup too large, cannot upload to Telegram")
                return

            if last_backup_message_id:
                try:
                    delete_telegram_message(CHANNEL_ID, last_backup_message_id)
                except:
                    pass

            filename = f"backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json.gz"
            url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
            files = {'document': (filename, compressed, 'application/gzip')}
            resp = requests.post(url, data={"chat_id": CHANNEL_ID}, files=files, timeout=60)

            if resp.status_code == 200 and resp.json().get("ok"):
                last_backup_message_id = resp.json()["result"]["message_id"]
                pin_url = f"https://api.telegram.org/bot{BOT_TOKEN}/pinChatMessage"
                requests.post(pin_url, json={
                    "chat_id": CHANNEL_ID,
                    "message_id": last_backup_message_id,
                    "disable_notification": True
                })
            else:
                logger.error(f"Backup upload failed: {resp.text}")

        except Exception as e:
            logger.error(f"Channel backup error: {e}")

def auto_backup_loop():
    while True:
        time.sleep(86400)
        save_data_to_channel()
        send_telegram_message("🔄 অটো ব্যাকআপ সম্পন্ন হয়েছে", ADMIN_CHAT_ID)

# ================== KEYBOARD ==================
def get_main_keyboard(chat_id):
    keyboard = [
        ["💰 ব্যালেন্স", "💸 ডিপোজিট"],
        ["🛒 একাউন্ট কিনুন", "💰 একাউন্ট বিক্রয় করুন"],
        ["💸 উইথড্র"],
        ["📋 সাবমিট", "🎁 মাদার একাউন্ট"],
        ["📞 সাপোর্ট", "🛑 স্টপ"],
        ["🔄 লস রিকভারি"]
    ]
    if str(chat_id) == ADMIN_CHAT_ID:
        keyboard.append(["📥 ডিপোজিট রিকোয়েস্ট", "➕ স্টক যোগ করুন"])
        keyboard.append(["📦 স্টক দেখুন", "🗑️ স্টক ডিলিট"])
        keyboard.append(["📊 সেল requests", "💳 withdraw requests"])
    return {"keyboard": keyboard, "resize_keyboard": True, "one_time_keyboard": False}

def remove_keyboard():
    return {"remove_keyboard": True}

def send_main_keyboard(chat_id, text="\u200b"):
    send_telegram_message(text, chat_id, reply_markup=get_main_keyboard(chat_id))

# ================== EXCEL GENERATORS (updated for reuse_link) ==================
def generate_submission_excel(usernames, passwords, twofa_list, bkash, telegram_username):
    wb = Workbook()
    ws = wb.active
    ws.title = "Account Submission"
    headers = ["Username", "Password", "2FA Key", "Bkash Number", "Telegram Username"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for i in range(len(usernames)):
        row = [
            usernames[i],
            passwords[i] if i < len(passwords) else "",
            twofa_list[i] if i < len(twofa_list) else "",
            bkash,
            telegram_username
        ]
        ws.append(row)
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

def generate_purchase_excel(bought):
    wb = Workbook()
    ws = wb.active
    ws.title = "Purchased Accounts"
    headers = ["Username", "Password", "2FA Key", "Email Reuse Link"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for acc in bought:
        ws.append([
            acc["username"],
            acc["password"],
            acc.get("fa_key", ""),
            acc.get("reuse_link", "")
        ])
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

def generate_sell_excel(accounts_list, telegram_username):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sell Request"
    headers = ["Username", "Password", "2FA Key", "Email Reuse Link", "Telegram Username"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for acc in accounts_list:
        ws.append([
            acc["username"],
            acc["password"],
            acc.get("fa_key", ""),
            acc.get("reuse_link", ""),
            telegram_username
        ])
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

# ================== SUBMISSION HANDLER ==================
# (unchanged)
def start_submission(chat_id, sender_username):
    submission_sessions[chat_id] = {"step": "username", "data": {}, "username": sender_username}
    msg = ("📋 দয়া করে আপনার **ইউজারনেম** লিস্ট দিন (প্রতি লাইনে একটি করে):\n\n"
           "উদাহরণ:\nuser1\nuser2\nuser3\n\nবাতিল করতে নিচের বাটনে চাপুন বা /cancel টাইপ করুন।")
    inline_kb = {"inline_keyboard": [[{"text": "❌ বাতিল", "callback_data": "cancel_submission"}]]}
    send_telegram_message(msg, chat_id, reply_markup=inline_kb)

def process_submission_step(chat_id, text, sender_username):
    if chat_id not in submission_sessions:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        del submission_sessions[chat_id]
        send_telegram_message("❌ জমা প্রক্রিয়া বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    session = submission_sessions[chat_id]
    step = session["step"]
    if step == "username":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["data"]["usernames"] = lines
        session["step"] = "password"
        send_telegram_message(f"🔑 এখন **পাসওয়ার্ড** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nআপনার ইউজারনেম সংখ্যা: {len(lines)}\nবাতিল করতে /cancel বা উপরের বাটন।", chat_id)
        return True
    elif step == "password":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        usernames = session["data"]["usernames"]
        if len(lines) != len(usernames):
            send_telegram_message(f"❌ পাসওয়ার্ড সংখ্যা ({len(lines)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। পুনরায় সঠিক লিস্ট দিন।", chat_id)
            return True
        session["data"]["passwords"] = lines
        session["step"] = "2fa"
        send_telegram_message("🔐 এখন **2FA কী** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nযদি 2FA না থাকে, লাইন ফাঁকা রাখবেন (শুধু এন্টার দিন)।\nবাতিল করতে /cancel বা উপরের বাটন।", chat_id)
        return True
    elif step == "2fa":
        raw_lines = text.splitlines()
        twofa_list = [l.strip() for l in raw_lines]
        usernames = session["data"]["usernames"]
        while len(twofa_list) > len(usernames) and twofa_list and twofa_list[-1] == '':
            twofa_list.pop()
        if len(twofa_list) != len(usernames):
            send_telegram_message(f"❌ 2FA কী সংখ্যা ({len(twofa_list)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। প্রতিটি ইউজারনেমের জন্য একটি লাইন (খালিও হতে পারে) দিন।", chat_id)
            return True
        session["data"]["twofa"] = twofa_list
        session["step"] = "bkash"
        send_telegram_message("💳 দয়া করে আপনার **বিকাশ নম্বর** দিন:\nবাতিল করতে /cancel বা উপরের বাটন।", chat_id)
        return True
    elif step == "bkash":
        bkash_number = text.strip()
        if not bkash_number:
            send_telegram_message("⚠️ বিকাশ নম্বর খালি রাখা যাবে না। আবার দিন।", chat_id)
            return True
        session["data"]["bkash"] = bkash_number
        excel_bytes = generate_submission_excel(
            session["data"]["usernames"], session["data"]["passwords"],
            session["data"]["twofa"], bkash_number, session["username"]
        )
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"submission_{chat_id}_{timestamp}.xlsx"
        if send_telegram_document(excel_bytes, filename, ADMIN_CHAT_ID):
            send_telegram_message("✅ আপনার অ্যাকাউন্ট সফলভাবে জমা হয়েছে।\n\nঅ্যাডমিন শীঘ্রই যোগাযোগ করবে। ধন্যবাদ! 🙏", chat_id)
        else:
            send_telegram_message("⚠️ জমা দেওয়ার সময় ত্রুটি হয়েছে, অনুগ্রহ করে পরে চেষ্টা করুন।", chat_id)
        del submission_sessions[chat_id]
        send_main_keyboard(chat_id)
        return True
    return False

# ================== LOSS RECOVERY HANDLER ==================
# (unchanged)
def start_loss_recovery(chat_id):
    loss_recovery_sessions[chat_id] = {"step": "usernames", "data": {}}
    msg = ("⚠️ সতর্কতা: ভুল তথ্য দিলে লস রিকভারি পাবেন না। সকল তথ্য ম্যানুয়ালি যাচাই করা হবে।\n\n"
           "অনুগ্রহ করে সঠিক তথ্য দিন।\n\n"
           "আপনার কেনা অ্যাকাউন্টগুলোর ইউজারনেম লিস্ট দিন (প্রতি লাইনে একটি):\n\nবাতিল করতে নিচের বাটনে চাপুন বা /cancel টাইপ করুন।")
    inline_kb = {"inline_keyboard": [[{"text": "❌ বাতিল", "callback_data": "cancel_loss_recovery"}]]}
    send_telegram_message(msg, chat_id, reply_markup=inline_kb)

def process_loss_recovery_step(chat_id, text):
    if chat_id not in loss_recovery_sessions:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        del loss_recovery_sessions[chat_id]
        send_telegram_message("❌ লস রিকভারি বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    session = loss_recovery_sessions[chat_id]
    step = session["step"]
    if step == "usernames":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["data"]["usernames"] = lines
        session["step"] = "cookie_date"
        send_telegram_message("📅 কত তারিখে কুকিজ সাবমিট করেছিলেন? (শুধু তারিখের সংখ্যা লিখুন, যেমন: 13 বা 26)", chat_id)
        return True
    elif step == "cookie_date":
        date_str = text.strip()
        if not date_str.isdigit():
            send_telegram_message("⚠️ দয়া করে শুধু সংখ্যা দিন (13, 26 ইত্যাদি)।", chat_id)
            return True
        session["data"]["cookie_date"] = date_str
        session["step"] = "report_file"
        send_telegram_message("📎 এখন রিপোর্ট ফেইল হওয়ার দিনের Excel Report File (.xlsx/.xls) পাঠান।\n\n⚠️ Screenshot, PDF বা অন্য কোনো ফাইল গ্রহণ করা হবে না।", chat_id)
        return True
    elif step == "report_file":
        send_telegram_message("⚠️ শুধুমাত্র Excel File (.xlsx/.xls) পাঠান।", chat_id)
        return True
    elif step == "bkash":
        bkash = text.strip()
        if not bkash:
            send_telegram_message("⚠️ বিকাশ নম্বর খালি রাখা যাবে না।", chat_id)
            return True
        session["data"]["bkash"] = bkash
        session["step"] = "whatsapp"
        send_telegram_message("📞 আপনার হোয়াটসঅ্যাপ নম্বর দিন:", chat_id)
        return True
    elif step == "whatsapp":
        whatsapp = text.strip()
        if not whatsapp:
            send_telegram_message("⚠️ হোয়াটসঅ্যাপ নম্বর দিন।", chat_id)
            return True
        session["data"]["whatsapp"] = whatsapp
        usernames = session["data"]["usernames"]
        cookie_date = session["data"]["cookie_date"]
        bkash = session["data"]["bkash"]
        file_id = session["data"].get("report_file_id")
        file_message_id = session["data"].get("report_message_id")
        admin_text = (
            "🔄 **নতুন লস রিকভারি রিকোয়েস্ট**\n\n"
            f"👤 ইউজার: {user_info.get(chat_id, chat_id)} (`{chat_id}`)\n"
            f"📅 কুকি সাবমিটের তারিখ: {cookie_date}\n"
            f"💳 বিকাশ: {bkash}\n"
            f"📞 হোয়াটসঅ্যাপ: {whatsapp}\n"
            f"🔑 ইউজারনেম: " + ", ".join(usernames)
        )
        send_telegram_message(admin_text, ADMIN_CHAT_ID, parse_mode="Markdown")
        if file_message_id:
            forward_telegram_document(ADMIN_CHAT_ID, chat_id, file_message_id)
        elif file_id:
            send_telegram_message("⚠️ রিপোর্ট ফাইল ফরওয়ার্ড করা যায়নি, কারণ মেসেজ আইডি পাওয়া যায়নি।", ADMIN_CHAT_ID)
        send_telegram_message("✅ আপনার লস রিকভারি রিকোয়েস্ট জমা হয়েছে। অ্যাডমিন শীঘ্রই আপনার সাথে যোগাযোগ করবে।", chat_id)
        del loss_recovery_sessions[chat_id]
        send_main_keyboard(chat_id)
        return True
    return False

def handle_loss_recovery_file(chat_id, message):
    if chat_id not in loss_recovery_sessions:
        return
    session = loss_recovery_sessions[chat_id]
    if session["step"] != "report_file":
        return
    doc = message.get("document")
    if not doc:
        send_telegram_message("⚠️ শুধুমাত্র Excel File (.xlsx/.xls) পাঠান।", chat_id)
        return
    file_name = doc.get("file_name", "")
    if not file_name.lower().endswith(('.xlsx', '.xls')):
        send_telegram_message("❌ শুধুমাত্র .xlsx বা .xls এক্সটেনশনের ফাইল গ্রহণ করা হবে।", chat_id)
        return
    allowed_mimes = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel"
    ]
    mime = doc.get("mime_type", "")
    if mime not in allowed_mimes:
        send_telegram_message("❌ অবৈধ ফাইল ফরম্যাট। শুধুমাত্র এক্সেল ফাইল গ্রহণ করা হবে।", chat_id)
        return
    file_id = doc.get("file_id")
    message_id = message.get("message_id")
    session["data"]["report_file_id"] = file_id
    session["data"]["report_message_id"] = message_id
    session["step"] = "bkash"
    send_telegram_message("💳 দয়া করে আপনার বিকাশ নম্বর দিন (যেটি ব্যবহার করেছিলেন):", chat_id)

# ================== MOTHER ACCOUNT ==================
# (unchanged)
def handle_addmother(chat_id, args):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    parts = args.split(maxsplit=2)
    if len(parts) < 2:
        send_telegram_message("❌ ফরম্যাট: /addmother username password [2fa_key]", chat_id)
        return
    username = parts[0]
    password = parts[1]
    fa_key = parts[2] if len(parts) == 3 else ""
    with data_lock:
        for acc in mother_accounts:
            if acc["username"] == username and acc["password"] == password:
                send_telegram_message("⚠️ এই অ্যাকাউন্টটি আগেই যোগ করা আছে।", chat_id)
                return
        mother_accounts.append({"username": username, "password": password, "fa_key": fa_key, "assigned_to": None, "assigned_at": None})
        save_mother_accounts()
    save_data_to_channel()
    send_telegram_message(f"✅ মাদার অ্যাকাউন্ট যোগ করা হয়েছে: {username}", chat_id)

def handle_getmother(chat_id):
    now = time.time()
    last = user_last_request.get(str(chat_id), 0)
    cooldown = 600
    if now - last < cooldown:
        wait_sec = cooldown - (now - last)
        wait_min = int(wait_sec // 60)
        wait_sec_rem = int(wait_sec % 60)
        send_telegram_message(f"⏳ অনুগ্রহ করে অপেক্ষা করুন। পরবর্তী অ্যাকাউন্ট {wait_min} মিনিট {wait_sec_rem} সেকেন্ড পর নিতে পারবেন।", chat_id)
        send_main_keyboard(chat_id)
        return
    selected_acc = None
    with data_lock:
        for acc in mother_accounts:
            if acc["assigned_to"] is None:
                selected_acc = acc
                break
        if selected_acc is None:
            send_telegram_message("❌ কোনো মাদার অ্যাকাউন্ট উপলব্ধ নেই। পরে আবার চেষ্টা করুন।", chat_id)
            send_main_keyboard(chat_id)
            return
    msg = f"🎁 আপনার মাদার অ্যাকাউন্ট:\n\n👤 ইউজারনেম: {selected_acc['username']}\n🔑 পাসওয়ার্ড: {selected_acc['password']}"
    if selected_acc["fa_key"]:
        msg += f"\n🔐 2FA Key: {selected_acc['fa_key']}"
    if send_telegram_message(msg, chat_id):
        with data_lock:
            selected_acc["assigned_to"] = str(chat_id)
            selected_acc["assigned_at"] = now
            user_last_request[str(chat_id)] = now
            save_mother_accounts()
            save_user_cooldowns()
        send_main_keyboard(chat_id)
    else:
        send_telegram_message("⚠️ মেসেজ পাঠানো যায়নি, পরে আবার চেষ্টা করুন।", chat_id)
    save_data_to_channel()

def handle_motherlist(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    with data_lock:
        if not mother_accounts:
            send_telegram_message("📭 কোনো মাদার অ্যাকাউন্ট নেই।", chat_id)
            return
        lines = ["🎁 *মাদার অ্যাকাউন্ট লিস্ট:*\n"]
        for i, acc in enumerate(mother_accounts, start=1):
            assigned = "কেহ না"
            if acc["assigned_to"]:
                try:
                    assigned_time = datetime.datetime.fromtimestamp(acc["assigned_at"]).strftime('%d/%m %H:%M')
                except:
                    assigned_time = "কিছুক্ষণ আগে"
                assigned = f"{acc['assigned_to']} ({assigned_time})"
            twofa = "আছে" if acc.get("fa_key") else "নেই"
            lines.append(f"{i}. ইউজার: {acc['username']} | পাস: {acc['password']} | 2FA: {twofa} | বরাদ্দ: {assigned}")
    send_telegram_message("\n".join(lines), chat_id)

def handle_deletemother(chat_id, arg):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    arg = arg.strip()
    with data_lock:
        try:
            idx = int(arg) - 1
            if 0 <= idx < len(mother_accounts):
                deleted = mother_accounts.pop(idx)
                save_mother_accounts()
            else:
                send_telegram_message("❌ ভুল ইনডেক্স। /motherlist দিয়ে নম্বর দেখুন।", chat_id)
                return
        except ValueError:
            for i, acc in enumerate(mother_accounts):
                if acc["username"] == arg:
                    deleted = mother_accounts.pop(i)
                    save_mother_accounts()
                    break
            else:
                send_telegram_message(f"❌ `{arg}` নামে কোনো মাদার অ্যাকাউন্ট পাওয়া যায়নি।", chat_id)
                return
    save_data_to_channel()
    send_telegram_message(f"✅ মাদার অ্যাকাউন্ট `{deleted['username']}` মুছে ফেলা হয়েছে।", chat_id)

# ================== MAINTENANCE MODE ==================
def handle_maintenance(chat_id, args):
    global maintenance_mode
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    args = args.strip().lower()
    if args == "on":
        maintenance_mode = True
        save_config()
        send_telegram_message("🔧 রক্ষণাবেক্ষণ মোড চালু করা হয়েছে। সাধারণ ইউজাররা এখন বট ব্যবহার করতে পারবে না।", chat_id)
    elif args == "off":
        maintenance_mode = False
        save_config()
        send_telegram_message("🔧 রক্ষণাবেক্ষণ মোড বন্ধ করা হয়েছে। বট এখন স্বাভাবিক ভাবে চলবে।", chat_id)
    else:
        status = "চালু" if maintenance_mode else "বন্ধ"
        send_telegram_message(f"🔧 রক্ষণাবেক্ষণ মোড বর্তমানে {status} আছে। /maintenance on/off দিয়ে পরিবর্তন করুন।", chat_id)

# ================== ADMIN BROADCAST & USERS ==================
def handle_admin_users(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return
    with data_lock:
        if not subscribed_users:
            send_telegram_message("কোনো সাবস্ক্রাইবার নেই।", chat_id)
            return
        msg_lines = ["📋 সাবস্ক্রাইবড ইউজার লিস্ট:\n"]
        for uid in subscribed_users:
            name = user_info.get(str(uid), f"ID:{uid}")
            if ' ' not in name:
                name = '@' + name
            msg_lines.append(f"• {name} (ID: {uid})")
    send_telegram_message("\n".join(msg_lines), chat_id)

def handle_admin_broadcast(chat_id, message):
    if str(chat_id) != ADMIN_CHAT_ID:
        return
    if not message.strip():
        send_telegram_message("❌ মেসেজ খালি রাখা যাবে না। ফরম্যাট: /broadcast <মেসেজ>", chat_id)
        return
    broadcast_message(f"📢 অ্যাডমিন থেকে বার্তা:\n\n{message}")
    send_telegram_message("✅ বার্তা সকল সাবস্ক্রাইবারকে পাঠানো হয়েছে।", chat_id)

def handle_admin_send(chat_id, target_id, message):
    if str(chat_id) != ADMIN_CHAT_ID:
        return
    if not target_id.isdigit():
        send_telegram_message("❌ সঠিক ইউজার আইডি দিন।", chat_id)
        return
    if not message.strip():
        send_telegram_message("❌ মেসেজ খালি রাখা যাবে না।", chat_id)
        return
    send_telegram_message(f"📩 অ্যাডমিন থেকে:\n\n{message}", target_id)
    send_telegram_message(f"✅ {target_id} কে মেসেজ পাঠানো হয়েছে।", chat_id)

# ================== BACKUP & RESTORE (with config fix) ==================
def handle_backup(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    save_all()
    with data_lock:
        backup = {
            "subscribed_users": list(subscribed_users),
            "user_info": user_info,
            "mother_accounts": mother_accounts,
            "user_cooldowns": user_last_request,
            "accounts": accounts,
            "balances": balances,
            "deposits": deposits,
            "config": config,
            "sell_requests": sell_requests,
            "withdraw_requests": withdraw_requests,
            "timestamp": datetime.datetime.now().isoformat()
        }
    backup_json = json.dumps(backup, indent=2, ensure_ascii=False).encode('utf-8')
    compressed = gzip.compress(backup_json)
    filename = f"backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json.gz"
    if send_telegram_document(compressed, filename, ADMIN_CHAT_ID, caption="Manual backup"):
        send_telegram_message("✅ ব্যাকআপ ফাইল তৈরি ও পাঠানো হয়েছে।", chat_id)
    else:
        send_telegram_message("⚠️ ব্যাকআপ ফাইল পাঠানো যায়নি।", chat_id)

def handle_restore(chat_id, file_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    if _perform_restore(file_id):
        send_telegram_message("✅ ব্যাকআপ রিস্টোর সম্পন্ন হয়েছে।", chat_id)
    else:
        send_telegram_message("❌ ব্যাকআপ রিস্টোর করতে ব্যর্থ হয়েছে।", chat_id)

def _perform_restore(file_id):
    try:
        get_file_url = f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={file_id}"
        resp = requests.get(get_file_url, timeout=20)
        resp.raise_for_status()
        file_data = resp.json()
        if not file_data.get("ok"):
            logger.error("getFile failed")
            return False
        file_path = file_data["result"]["file_path"]
        download_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
        file_resp = requests.get(download_url, timeout=60)
        file_resp.raise_for_status()
        content = file_resp.content
        try:
            content = gzip.decompress(content)
        except gzip.BadGzipFile:
            pass
        backup = json.loads(content.decode('utf-8'))
    except Exception as e:
        logger.error(f"_perform_restore error: {e}")
        return False

    with data_lock:
        global subscribed_users, user_info, mother_accounts, user_last_request
        global accounts, balances, deposits, config, CHANNEL_ID, maintenance_mode
        global sell_requests, withdraw_requests
        subscribed_users = set(backup.get("subscribed_users", []))
        user_info = backup.get("user_info", {})
        mother_accounts = backup.get("mother_accounts", [])
        user_last_request = backup.get("user_cooldowns", {})
        accounts = backup.get("market_accounts", backup.get("accounts", []))
        for acc in accounts:
            if 'type' not in acc:
                acc['type'] = '2fa'
            if 'reuse_link' not in acc:
                acc['reuse_link'] = ''
        balances = backup.get("balances", {})
        deposits = backup.get("deposits", [])

        # Config merge fix
        default_config = {
            "bkash_number": "",
            "price_2fa_buy": 2.2,
            "price_normal_buy": 1.2,
            "price_2fa_sell": 1.8,
            "price_normal_sell": 0.8,
            "group_chat_id": "",
            "channel_id": str(CHANNEL_ID) if CHANNEL_ID else "",
            "maintenance_mode": False
        }
        restored_config = backup.get("config", {})
        default_config.update(restored_config)
        config = default_config

        CHANNEL_ID = int(config.get("channel_id", "0")) if config.get("channel_id") else None
        maintenance_mode = config.get("maintenance_mode", False)
        sell_requests = backup.get("sell_requests", [])
        withdraw_requests = backup.get("withdraw_requests", [])
        save_all()
    logger.info("Restore completed successfully")
    return True

def auto_restore_from_channel():
    if not CHANNEL_ID:
        logger.info("Auto-restore skipped: No CHANNEL_ID")
        return
    try:
        get_chat_url = f"https://api.telegram.org/bot{BOT_TOKEN}/getChat?chat_id={CHANNEL_ID}"
        resp = requests.get(get_chat_url, timeout=20).json()
        if not resp.get("ok"):
            logger.warning("Could not get channel info for auto-restore")
            return
        pinned = resp["result"].get("pinned_message")
        if not pinned or "document" not in pinned:
            logger.info("No pinned document in channel; skipping auto-restore")
            return
        file_id = pinned["document"]["file_id"]
        logger.info("Auto-restore: Found pinned backup. Restoring...")
        if _perform_restore(file_id):
            logger.info("Auto-restore completed successfully")
        else:
            logger.error("Auto-restore failed")
    except Exception as e:
        logger.error(f"Auto-restore error: {e}")

# ================== ADMIN ADD STOCK FLOW (updated for normal reuse_link) ==================
def start_add_stock(chat_id, acc_type):
    if str(chat_id) != ADMIN_CHAT_ID:
        return
    add_stock_sessions[chat_id] = {"step": "usernames", "type": acc_type}
    type_label = "2FA ON একাউন্ট" if acc_type == "2fa" else "Normal একাউন্ট"
    msg = (f"➕ স্টক যোগ করুন ({type_label})\n\nপ্রথমে **ইউজারনেম** লিস্ট দিন (প্রতি লাইনে একটি করে):\n"
           "উদাহরণ:\nuser1\nuser2\nuser3\n\nবাতিল করতে নিচের বাটনে চাপুন বা /cancel টাইপ করুন।")
    inline_kb = {"inline_keyboard": [[{"text": "❌ বাতিল", "callback_data": "cancel_addstock"}]]}
    send_telegram_message(msg, chat_id, reply_markup=inline_kb)

def process_add_stock_step(chat_id, text):
    if chat_id not in add_stock_sessions:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        del add_stock_sessions[chat_id]
        send_telegram_message("❌ স্টক যোগ বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    session = add_stock_sessions[chat_id]
    step = session["step"]
    acc_type = session["type"]
    if step == "usernames":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["usernames"] = lines
        session["step"] = "passwords"
        send_telegram_message(f"🔑 এখন **পাসওয়ার্ড** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nআপনার ইউজারনেম সংখ্যা: {len(lines)}", chat_id)
        return True
    elif step == "passwords":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        usernames = session["usernames"]
        if len(lines) != len(usernames):
            send_telegram_message(f"❌ পাসওয়ার্ড সংখ্যা ({len(lines)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। পুনরায় সঠিক লিস্ট দিন।", chat_id)
            return True
        session["passwords"] = lines
        if acc_type == "2fa":
            session["step"] = "fa_keys"
            send_telegram_message("🔐 এখন **2FA কী** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nযদি 2FA না থাকে, লাইন ফাঁকা রাখবেন (শুধু এন্টার দিন)।", chat_id)
        else:
            # Normal account: ask for email reuse links
            session["step"] = "reuse_links"
            send_telegram_message(f"📧 এখন **ইমেইল রিইউজ লিংক** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nপ্রতিটি অ্যাকাউন্টের জন্য tmailor.com থেকে প্রাপ্ত রিইউজ লিংক দিন।\n\nআপনার ইউজারনেম সংখ্যা: {len(usernames)}", chat_id)
        return True
    elif step == "reuse_links":
        # Normal account reuse links
        raw_lines = text.splitlines()
        reuse_links = [l.strip() for l in raw_lines]
        usernames = session["usernames"]
        while len(reuse_links) > len(usernames) and reuse_links and reuse_links[-1] == '':
            reuse_links.pop()
        if len(reuse_links) != len(usernames):
            send_telegram_message(f"❌ ইমেইল রিইউজ লিংক সংখ্যা ({len(reuse_links)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। প্রতিটি ইউজারনেমের জন্য একটি করে লিংক দিন।", chat_id)
            return True
        count = len(usernames)
        with data_lock:
            for i in range(count):
                accounts.append({
                    "username": usernames[i],
                    "password": session["passwords"][i],
                    "fa_key": "",
                    "type": "normal",
                    "reuse_link": reuse_links[i]
                })
            save_accounts()
        save_data_to_channel()
        del add_stock_sessions[chat_id]
        broadcast_message(f"📢 নতুন স্টক যোগ করা হয়েছে: {count} টি Normal একাউন্ট।")
        send_telegram_message(f"✅ {count} টি Normal একাউন্ট স্টকে যোগ করা হয়েছে (reuse link সহ)।", chat_id)
        send_main_keyboard(chat_id)
        return True
    elif step == "fa_keys":
        raw_lines = text.splitlines()
        fa_list = [l.strip() for l in raw_lines]
        usernames = session["usernames"]
        while len(fa_list) > len(usernames) and fa_list and fa_list[-1] == '':
            fa_list.pop()
        if len(fa_list) != len(usernames):
            send_telegram_message(f"❌ 2FA কী সংখ্যা ({len(fa_list)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। প্রতিটি ইউজারনেমের জন্য একটি লাইন (খালিও হতে পারে) দিন।", chat_id)
            return True
        count = len(usernames)
        with data_lock:
            for i in range(count):
                accounts.append({
                    "username": usernames[i],
                    "password": session["passwords"][i],
                    "fa_key": fa_list[i],
                    "type": "2fa",
                    "reuse_link": ""
                })
            save_accounts()
        save_data_to_channel()
        del add_stock_sessions[chat_id]
        broadcast_message(f"📢 নতুন স্টক যোগ করা হয়েছে: {count} টি 2FA ON একাউন্ট।")
        send_telegram_message(f"✅ {count} টি 2FA ON একাউন্ট স্টকে যোগ করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    return False

# ================== MARKETPLACE: DEPOSIT ==================
# (unchanged)
def start_deposit(chat_id):
    deposit_sessions[chat_id] = {"step": "amount"}
    bkash = config.get("bkash_number", "")
    if not bkash:
        send_telegram_message("⚠️ অ্যাডমিন এখনও বিকাশ নম্বর সেট করেননি। পরে চেষ্টা করুন।", chat_id)
        send_main_keyboard(chat_id)
        deposit_sessions.pop(chat_id, None)
        return
    msg = f"💸 দয়া করে আপনার জমা করার টাকার পরিমাণ লিখুন (শুধু সংখ্যা, যেমন: 100)\n\nবিকাশ নম্বর: {bkash}\n\nবাতিল করতে নিচের বাটনে চাপুন বা /cancel টাইপ করুন।"
    inline_kb = {"inline_keyboard": [[{"text": "❌ বাতিল", "callback_data": "cancel_deposit"}]]}
    send_telegram_message(msg, chat_id, reply_markup=inline_kb)

def process_deposit_step(chat_id, text):
    if chat_id not in deposit_sessions:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        deposit_sessions.pop(chat_id, None)
        send_telegram_message("❌ ডিপোজিট বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    session = deposit_sessions[chat_id]
    step = session["step"]
    if step == "amount":
        try:
            amount = float(normalize_bengali_number(text.strip()))
            if amount <= 0:
                raise ValueError
        except:
            send_telegram_message("⚠️ সঠিক সংখ্যা দিন (ইংরেজি বা বাংলা)। /cancel দিয়ে বাতিল করুন।", chat_id)
            return True
        session["amount"] = amount
        session["step"] = "trxid"
        send_telegram_message("🔢 এখন আপনার বিকাশ ট্রানজেকশন আইডি লিখুন:", chat_id)
        return True
    elif step == "trxid":
        trxid = text.strip()
        if not trxid:
            send_telegram_message("⚠️ ট্রানজেকশন আইডি খালি রাখা যাবে না।", chat_id)
            return True
        amount = session["amount"]
        deposit_id = uuid.uuid4().hex[:10]
        deposit = {"id": deposit_id, "user_id": chat_id, "amount": amount, "trxid": trxid, "status": "pending", "time": time.time()}
        with data_lock:
            deposits.append(deposit)
            save_deposits()
        save_data_to_channel()
        deposit_sessions.pop(chat_id, None)
        admin_msg = f"📥 নতুন ডিপোজিট রিকোয়েস্ট\nআইডি: {deposit_id}\nইউজার: {user_info.get(chat_id, chat_id)} ({chat_id})\nপরিমাণ: {amount} টাকা\nট্রানজেকশন আইডি: {trxid}\nঅনুমোদন করতে: /approve {deposit_id}\nবাতিল করতে: /reject {deposit_id}"
        send_telegram_message(admin_msg, ADMIN_CHAT_ID)
        send_telegram_message(f"✅ আপনার {amount} টাকার ডিপোজিট রিকোয়েস্ট জমা হয়েছে।\nঅ্যাডমিন অনুমোদন করলেই আপনার ব্যালেন্সে যোগ হবে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    return False

# ================== MARKETPLACE: BUY (reuse_link in Excel) ==================
def start_buy_session(chat_id, acc_type):
    buy_sessions[chat_id] = {"step": "quantity", "type": acc_type}
    type_label = "2FA ON একাউন্ট" if acc_type == "2fa" else "Normal একাউন্ট"
    price = config["price_2fa_buy"] if acc_type == "2fa" else config["price_normal_buy"]
    count = len([a for a in accounts if a.get("type") == acc_type])
    if acc_type == "2fa":
        warning = (
            "🔐 2FA ON একাউন্টের ক্ষেত্রে:\n"
            "কুকিজ সাবমিট দিয়ে যত পিস একাউন্টের রিপোর্ট ব্যাক আসবে, "
            "সব একাউন্টের টাকা ফেরত পাবেন লস রিকভারি অপশন ব্যবহারের মাধ্যমে।"
        )
    else:
        warning = (
            "📄 Normal একাউন্টের ক্ষেত্রে:\n"
            "ইনস্টাগ্রাম কুকিজ সাবমিট দিয়ে যত পিস একাউন্টের রিপোর্ট ব্যাক আসবে, "
            "সব একাউন্টের টাকা ফেরত পাবেন লস রিকভারি অপশন ব্যবহারের মাধ্যমে।\n"
            "Normal একাউন্টে কোনো 2FA কী থাকে না।\n"
            "এক্সেল ফাইলে ইমেইল রিইউজ লিংক দেওয়া থাকবে।"
        )
    msg = (f"🛒 {type_label}\n\n{warning}\n\n"
           f"কতটি একাউন্ট কিনতে চান? (সংখ্যা লিখুন)\n"
           f"প্রতি একাউন্টের মূল্য: {price} টাকা\n"
           f"স্টকে আছে: {count} টি\n\n"
           f"বাতিল করতে নিচের বাটনে চাপুন বা /cancel টাইপ করুন।")
    inline_kb = {"inline_keyboard": [[{"text": "❌ বাতিল", "callback_data": "cancel_buy"}]]}
    send_telegram_message(msg, chat_id, reply_markup=inline_kb)

def handle_buy(chat_id, text):
    session = buy_sessions.get(chat_id)
    if not session:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        buy_sessions.pop(chat_id, None)
        send_telegram_message("❌ কেনা বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    clean_text = normalize_bengali_number(text.strip())
    try:
        qty = int(clean_text)
        if qty <= 0:
            raise ValueError
    except:
        send_telegram_message("⚠️ সঠিক সংখ্যা দিন (ইংরেজি বা বাংলা)।", chat_id)
        return False
    acc_type = session["type"]
    price = config["price_2fa_buy"] if acc_type == "2fa" else config["price_normal_buy"]
    with data_lock:
        available = [a for a in accounts if a.get("type") == acc_type]
        if qty > len(available):
            send_telegram_message(f"❌ পর্যাপ্ত একাউন্ট নেই। বর্তমান স্টক ({acc_type}): {len(available)}", chat_id)
            return False
        total = qty * price
        user_balance = balances.get(str(chat_id), 0)
        if user_balance < total:
            send_telegram_message(f"❌ পর্যাপ্ত ব্যালেন্স নেই।\nপ্রয়োজন: {total} টাকা\nআপনার ব্যালেন্স: {user_balance} টাকা\nদয়া করে প্রথমে ডিপোজিট করুন।", chat_id)
            return False
        to_buy = []
        new_accounts = []
        bought_count = 0
        for acc in accounts:
            if acc.get("type") == acc_type and bought_count < qty:
                to_buy.append(acc)
                bought_count += 1
            else:
                new_accounts.append(acc)
        accounts[:] = new_accounts
        balances[str(chat_id)] = user_balance - total
        save_accounts()
        save_balances()
    excel_bytes = generate_purchase_excel(to_buy)  # now includes reuse_link
    filename = f"purchased_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    if send_telegram_document(excel_bytes, filename, chat_id):
        send_telegram_message(f"✅ {qty} টি একাউন্ট কেনা হয়েছে। মোট মূল্য: {total} টাকা।\nঅবশিষ্ট ব্যালেন্স: {balances[str(chat_id)]} টাকা", chat_id)
        admin_msg = f"🛒 {user_info.get(chat_id, chat_id)} ({chat_id}) {qty} টি {acc_type} একাউন্ট কিনেছে। মোট: {total} টাকা।"
        send_telegram_message(admin_msg, ADMIN_CHAT_ID)
    else:
        with data_lock:
            accounts[:0] = to_buy
            balances[str(chat_id)] = user_balance
            save_accounts()
            save_balances()
        send_telegram_message("⚠️ একাউন্ট ডেলিভারি ব্যর্থ হয়েছে। আপনার টাকা ফেরত দেওয়া হয়েছে এবং একাউন্ট পুনরায় স্টকে যোগ করা হয়েছে। পরে আবার চেষ্টা করুন।", chat_id)
    buy_sessions.pop(chat_id, None)
    save_data_to_channel()
    send_main_keyboard(chat_id)
    return True

# ================== SELL (updated with reuse_link for normal) ==================
def start_sell_session(chat_id, acc_type):
    sell_sessions[chat_id] = {"step": "usernames", "type": acc_type}
    type_label = "2FA ON একাউন্ট" if acc_type == "2fa" else "Normal একাউন্ট"
    price = config["price_2fa_sell"] if acc_type == "2fa" else config["price_normal_sell"]
    if acc_type == "2fa":
        warning = "⚠️ সতর্কতা:\n\nআপনি যে একাউন্ট গুলো সেল দিবেন সেগুলো আমাদের ইউজাররা কুকিজ সাবমিট করে তাদের যতগুলা আইডি ব্যাক আসবে সেই পরিমাণ টাকা আপনার সেল একাউন্ট হতে মাইনাস করা হবে।"
        send_telegram_message(warning, chat_id)
    else:
        warning = "⚠️ সতর্কতা:\n\nআপনি যে একাউন্ট গুলো সেল দিবেন সেগুলো আমাদের ইউজাররা কুকিজ সাবমিট করে তাদের যতগুলা আইডি ব্যাক আসবে সেই পরিমাণ টাকা আপনার সেল একাউন্ট হতে মাইনাস করা হবে।"
        send_telegram_message(warning, chat_id)
        # Extra notice for normal accounts about email reuse link
        extra_notice = (
            "📌 Normal একাউন্ট বিক্রির বিশেষ নির্দেশনা:\n\n"
            "Tmailor.com থেকে মেইল নিয়ে প্রতিটি মাদার অ্যাকাউন্টে মেইল অ্যাড করবেন এবং reuse link সংরক্ষণ করবেন। "
            "মাদার অ্যাকাউন্টের মেইল reuse link প্রতিটি অ্যাড অ্যাকাউন্টের মেইল reuse link হিসেবে গণ্য হয়। "
            "বিক্রয়ের সময় এই মেইল reuse link গুলোও দিতে হবে।"
        )
        send_telegram_message(extra_notice, chat_id)

    msg = (f"💰 {type_label} বিক্রয় করুন\n\nপ্রতি একাউন্টের মূল্য: {price} টাকা\n\n"
           "প্রথমে আপনার **ইউজারনেম** লিস্ট দিন (প্রতি লাইনে একটি করে):\n"
           "উদাহরণ:\nuser1\nuser2\nuser3\n\nবাতিল করতে নিচের বাটনে চাপুন বা /cancel টাইপ করুন।")
    inline_kb = {"inline_keyboard": [[{"text": "❌ বাতিল", "callback_data": "cancel_sell"}]]}
    send_telegram_message(msg, chat_id, reply_markup=inline_kb)

def process_sell_step(chat_id, text):
    if chat_id not in sell_sessions:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        del sell_sessions[chat_id]
        send_telegram_message("❌ বিক্রয় প্রক্রিয়া বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    session = sell_sessions[chat_id]
    step = session["step"]
    acc_type = session["type"]
    if step == "usernames":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["usernames"] = lines
        session["step"] = "passwords"
        send_telegram_message(f"🔑 এখন **পাসওয়ার্ড** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nআপনার ইউজারনেম সংখ্যা: {len(lines)}", chat_id)
        return True
    elif step == "passwords":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        usernames = session["usernames"]
        if len(lines) != len(usernames):
            send_telegram_message(f"❌ পাসওয়ার্ড সংখ্যা ({len(lines)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। পুনরায় সঠিক লিস্ট দিন।", chat_id)
            return True
        session["passwords"] = lines
        if acc_type == "2fa":
            session["step"] = "fa_keys"
            send_telegram_message("🔐 এখন **2FA কী** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nযদি 2FA না থাকে, লাইন ফাঁকা রাখবেন (শুধু এন্টার দিন)।", chat_id)
        else:
            # Normal: ask for reuse links
            session["step"] = "reuse_links"
            send_telegram_message(f"📧 এখন **ইমেইল রিইউজ লিংক** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nপ্রতিটি অ্যাকাউন্টের জন্য tmailor.com থেকে প্রাপ্ত রিইউজ লিংক দিন।\n\nআপনার ইউজারনেম সংখ্যা: {len(usernames)}", chat_id)
        return True
    elif step == "reuse_links":
        # Normal sell: process reuse links
        raw_lines = text.splitlines()
        reuse_links = [l.strip() for l in raw_lines]
        usernames = session["usernames"]
        while len(reuse_links) > len(usernames) and reuse_links and reuse_links[-1] == '':
            reuse_links.pop()
        if len(reuse_links) != len(usernames):
            send_telegram_message(f"❌ ইমেইল রিইউজ লিংক সংখ্যা ({len(reuse_links)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। প্রতিটি ইউজারনেমের জন্য একটি করে লিংক দিন।", chat_id)
            return True
        accounts_list = []
        for i in range(len(usernames)):
            accounts_list.append({
                "username": usernames[i],
                "password": session["passwords"][i],
                "fa_key": "",
                "reuse_link": reuse_links[i]
            })
        sell_id = uuid.uuid4().hex[:10]
        sell_req = {"id": sell_id, "user_id": chat_id, "accounts": accounts_list, "status": "pending", "time": time.time(), "type": "normal"}
        with data_lock:
            sell_requests.append(sell_req)
            save_sell_requests()
        save_data_to_channel()
        del sell_sessions[chat_id]
        total_expected = config["price_normal_sell"] * len(accounts_list)
        send_telegram_message(f"✅ আপনার Normal একাউন্ট বিক্রয় রিকোয়েস্ট জমা হয়েছে।\nআইডি: {sell_id}\nঅ্যাকাউন্ট সংখ্যা: {len(accounts_list)}\nপ্রত্যাশিত মূল্য: {total_expected} টাকা\nরিভিউ সম্পন্ন হতে ২৪ ঘণ্টা পর্যন্ত সময় লাগতে পারে।", chat_id)
        tg_username = user_info.get(str(chat_id), f"ID:{chat_id}")
        excel_bytes = generate_sell_excel(accounts_list, tg_username)  # now includes reuse_link
        filename = f"sell_{sell_id}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caption = f"📊 নতুন Normal সেল রিকোয়েস্ট\nআইডি: {sell_id}\nইউজার: {tg_username} (`{chat_id}`)\nঅ্যাকাউন্ট সংখ্যা: {len(accounts_list)}\nপ্রত্যাশিত মূল্য: {total_expected} টাকা\nঅনুমোদন: /approvesell {sell_id} {total_expected}\nবাতিল: /rejectsell {sell_id}"
        send_telegram_document(excel_bytes, filename, ADMIN_CHAT_ID, caption=caption)
        send_main_keyboard(chat_id)
        return True
    elif step == "fa_keys":
        # 2FA sell (unchanged)
        raw_lines = text.splitlines()
        fa_list = [l.strip() for l in raw_lines]
        usernames = session["usernames"]
        while len(fa_list) > len(usernames) and fa_list and fa_list[-1] == '':
            fa_list.pop()
        if len(fa_list) != len(usernames):
            send_telegram_message(f"❌ 2FA কী সংখ্যা ({len(fa_list)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। প্রতিটি ইউজারনেমের জন্য একটি লাইন (খালিও হতে পারে) দিন।", chat_id)
            return True
        accounts_list = []
        for i in range(len(usernames)):
            accounts_list.append({
                "username": usernames[i],
                "password": session["passwords"][i],
                "fa_key": fa_list[i],
                "reuse_link": ""   # 2FA accounts don't have reuse link
            })
        sell_id = uuid.uuid4().hex[:10]
        sell_req = {"id": sell_id, "user_id": chat_id, "accounts": accounts_list, "status": "pending", "time": time.time(), "type": "2fa"}
        with data_lock:
            sell_requests.append(sell_req)
            save_sell_requests()
        save_data_to_channel()
        del sell_sessions[chat_id]
        total_expected = config["price_2fa_sell"] * len(accounts_list)
        send_telegram_message(f"✅ আপনার 2FA ON একাউন্ট বিক্রয় রিকোয়েস্ট জমা হয়েছে।\nআইডি: {sell_id}\nঅ্যাকাউন্ট সংখ্যা: {len(accounts_list)}\nপ্রত্যাশিত মূল্য: {total_expected} টাকা\nরিভিউ সম্পন্ন হতে ২৪ ঘণ্টা পর্যন্ত সময় লাগতে পারে।", chat_id)
        tg_username = user_info.get(str(chat_id), f"ID:{chat_id}")
        excel_bytes = generate_sell_excel(accounts_list, tg_username)
        filename = f"sell_{sell_id}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caption = f"📊 নতুন 2FA সেল রিকোয়েস্ট\nআইডি: {sell_id}\nইউজার: {tg_username} (`{chat_id}`)\nঅ্যাকাউন্ট সংখ্যা: {len(accounts_list)}\nপ্রত্যাশিত মূল্য: {total_expected} টাকা\nঅনুমোদন: /approvesell {sell_id} {total_expected}\nবাতিল: /rejectsell {sell_id}"
        send_telegram_document(excel_bytes, filename, ADMIN_CHAT_ID, caption=caption)
        send_main_keyboard(chat_id)
        return True
    return False

# ================== WITHDRAW (unchanged) ==================
def start_withdraw(chat_id):
    withdraw_sessions[chat_id] = {"step": "amount"}
    msg = "💸 উইথড্র\n\nআপনার কত টাকা উত্তোলন করতে চান? (শুধু সংখ্যা লিখুন)\n\nবাতিল করতে নিচের বাটনে চাপুন বা /cancel টাইপ করুন।"
    inline_kb = {"inline_keyboard": [[{"text": "❌ বাতিল", "callback_data": "cancel_withdraw"}]]}
    send_telegram_message(msg, chat_id, reply_markup=inline_kb)

def process_withdraw_step(chat_id, text):
    if chat_id not in withdraw_sessions:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        del withdraw_sessions[chat_id]
        send_telegram_message("❌ উইথড্র বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    session = withdraw_sessions[chat_id]
    step = session["step"]
    if step == "amount":
        try:
            amount = float(normalize_bengali_number(text.strip()))
            if amount <= 0:
                raise ValueError
        except:
            send_telegram_message("⚠️ সঠিক সংখ্যা দিন (ইংরেজি বা বাংলা)।", chat_id)
            return True
        with data_lock:
            bal = balances.get(str(chat_id), 0)
        if amount > bal:
            send_telegram_message(f"❌ পর্যাপ্ত ব্যালেন্স নেই। আপনার ব্যালেন্স: {bal} টাকা", chat_id)
            send_main_keyboard(chat_id)
            del withdraw_sessions[chat_id]
            return True
        session["amount"] = amount
        session["step"] = "bkash"
        send_telegram_message("📞 আপনার বিকাশ নম্বর দিন:", chat_id)
        return True
    elif step == "bkash":
        bkash = text.strip()
        if not bkash:
            send_telegram_message("⚠️ বিকাশ নম্বর দিন।", chat_id)
            return True
        amount = session["amount"]
        w_id = uuid.uuid4().hex[:10]
        w_req = {"id": w_id, "user_id": chat_id, "amount": amount, "bkash": bkash, "status": "pending", "time": time.time()}
        with data_lock:
            withdraw_requests.append(w_req)
            save_withdraw_requests()
        save_data_to_channel()
        del withdraw_sessions[chat_id]
        send_telegram_message(f"✅ আপনার {amount} টাকা উত্তোলনের রিকোয়েস্ট জমা হয়েছে।\nআইডি: {w_id}\nঅ্যাডমিন অ্যাপ্রুভ করলে টাকা পাঠানো হবে।", chat_id)
        admin_msg = f"💳 নতুন উইথড্র রিকোয়েস্ট\nআইডি: {w_id}\nইউজার: {user_info.get(chat_id, chat_id)} ({chat_id})\nপরিমাণ: {amount} টাকা\nবিকাশ: {bkash}\nঅনুমোদন: /approvewithdraw {w_id}\nবাতিল: /rejectwithdraw {w_id}"
        send_telegram_message(admin_msg, ADMIN_CHAT_ID)
        send_main_keyboard(chat_id)
        return True
    return False

# ================== ADMIN MARKETPLACE COMMANDS (unchanged except stock list) ==================
def admin_stock_cmd(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        if not accounts:
            send_telegram_message("📭 স্টক খালি।", chat_id)
            return True
        msg_lines = [f"📦 মোট স্টক: {len(accounts)} টি\n"]
        for i, acc in enumerate(accounts, 1):
            acc_type = acc.get("type", "2fa")
            type_tag = "🔐 2FA" if acc_type == "2fa" else "📄 Normal"
            line = f"{i}. [{type_tag}] ইউজার: {acc['username']} | পাস: {acc['password']}"
            if acc_type == "2fa":
                line += f" | 2FA: {acc.get('fa_key', 'N/A')}"
            else:
                reuse = acc.get("reuse_link", "")
                line += f" | Reuse Link: {reuse if reuse else 'N/A'}"
            msg_lines.append(line)
    send_telegram_message("\n".join(msg_lines), chat_id)
    return True

def admin_deletestock_cmd(chat_id, arg):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    try:
        idx = int(arg) - 1
        with data_lock:
            if 0 <= idx < len(accounts):
                deleted = accounts.pop(idx)
                save_accounts()
            else:
                send_telegram_message("❌ ভুল ইনডেক্স। /stock দিয়ে নম্বর দেখুন।", chat_id)
                return True
        save_data_to_channel()
        send_telegram_message(f"✅ স্টক থেকে একাউন্ট `{deleted['username']}` মুছে ফেলা হয়েছে।", chat_id)
    except ValueError:
        with data_lock:
            for i, acc in enumerate(accounts):
                if acc["username"] == arg:
                    deleted = accounts.pop(i)
                    save_accounts()
                    save_data_to_channel()
                    send_telegram_message(f"✅ স্টক থেকে একাউন্ট `{deleted['username']}` মুছে ফেলা হয়েছে।", chat_id)
                    return True
            send_telegram_message(f"❌ `{arg}` নামে কোনো একাউন্ট পাওয়া যায়নি।", chat_id)
    return True

def admin_setprice_cmd(chat_id, args):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    parts = args.split()
    if len(parts) < 2:
        send_telegram_message("❌ ফরম্যাট: /setprice 2fa <মূল্য> বা /setprice normal <মূল্য>", chat_id)
        return True
    acc_type = parts[0].lower()
    try:
        price = float(normalize_bengali_number(parts[1]))
        if price <= 0:
            raise ValueError
    except:
        send_telegram_message("❌ সঠিক মূল্য দিন (ইংরেজি বা বাংলা সংখ্যা)।", chat_id)
        return True
    if acc_type == "2fa":
        config["price_2fa_buy"] = price
    elif acc_type == "normal":
        config["price_normal_buy"] = price
    else:
        send_telegram_message("❌ টাইপ হতে হবে '2fa' বা 'normal'", chat_id)
        return True
    save_config()
    save_data_to_channel()
    send_telegram_message(f"✅ {acc_type} ক্রয় মূল্য {price} টাকা নির্ধারণ করা হয়েছে।", chat_id)
    return True

def admin_setsellprice_cmd(chat_id, args):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    parts = args.split()
    if len(parts) < 2:
        send_telegram_message("❌ ফরম্যাট: /setsellprice 2fa <মূল্য> বা /setsellprice normal <মূল্য>", chat_id)
        return True
    acc_type = parts[0].lower()
    try:
        price = float(normalize_bengali_number(parts[1]))
        if price <= 0:
            raise ValueError
    except:
        send_telegram_message("❌ সঠিক মূল্য দিন (ইংরেজি বা বাংলা সংখ্যা)।", chat_id)
        return True
    if acc_type == "2fa":
        config["price_2fa_sell"] = price
    elif acc_type == "normal":
        config["price_normal_sell"] = price
    else:
        send_telegram_message("❌ টাইপ হতে হবে '2fa' বা 'normal'", chat_id)
        return True
    save_config()
    save_data_to_channel()
    send_telegram_message(f"✅ {acc_type} বিক্রয় মূল্য {price} টাকা নির্ধারণ করা হয়েছে।", chat_id)
    return True

def admin_setbkash_cmd(chat_id, number):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        config["bkash_number"] = number
        save_config()
    save_data_to_channel()
    send_telegram_message(f"✅ বিকাশ নম্বর {number} সেট করা হয়েছে।", chat_id)
    return True

def admin_approve_cmd(chat_id, deposit_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        for dep in deposits:
            if dep["id"] == deposit_id and dep["status"] == "pending":
                dep["status"] = "approved"
                user = dep["user_id"]
                balances[str(user)] = balances.get(str(user), 0) + dep["amount"]
                save_balances()
                save_deposits()
                send_telegram_message(f"✅ ডিপোজিট {deposit_id} অনুমোদিত। ইউজারের ব্যালেন্স আপডেট হয়েছে।", chat_id)
                send_telegram_message(f"✅ আপনার {dep['amount']} টাকার ডিপোজিট অনুমোদিত হয়েছে। বর্তমান ব্যালেন্স: {balances[str(user)]} টাকা", user)
                break
        else:
            send_telegram_message("❌ ডিপোজিট পাওয়া যায়নি বা ইতিমধ্যে প্রসেস করা হয়েছে।", chat_id)
    save_data_to_channel()
    return True

def admin_reject_cmd(chat_id, deposit_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        for dep in deposits:
            if dep["id"] == deposit_id and dep["status"] == "pending":
                dep["status"] = "rejected"
                save_deposits()
                send_telegram_message(f"❌ ডিপোজিট {deposit_id} বাতিল করা হয়েছে।", chat_id)
                send_telegram_message(f"❌ আপনার {dep['amount']} টাকার ডিপোজিট বাতিল করা হয়েছে।", dep["user_id"])
                break
        else:
            send_telegram_message("❌ ডিপোজিট পাওয়া যায়নি।", chat_id)
    save_data_to_channel()
    return True

def admin_deposits_cmd(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        pending = [d for d in deposits if d["status"] == "pending"]
    if not pending:
        send_telegram_message("কোনো পেন্ডিং ডিপোজিট নেই।", chat_id)
    else:
        lines = ["⏳ পেন্ডিং ডিপোজিট:"]
        for d in pending:
            lines.append(f"আইডি: {d['id']} | ইউজার: {d['user_id']} | পরিমাণ: {d['amount']} | ট্রানজেকশন: {d['trxid']}")
        send_telegram_message("\n".join(lines), chat_id)
    return True

def admin_addbalance_cmd(chat_id, uid, amt_str):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    try:
        amt = float(normalize_bengali_number(amt_str))
    except:
        send_telegram_message("❌ সঠিক পরিমাণ দিন (ইংরেজি বা বাংলা সংখ্যা)।", chat_id)
        return True
    with data_lock:
        balances[str(uid)] = balances.get(str(uid), 0) + amt
        save_balances()
    save_data_to_channel()
    send_telegram_message(f"✅ {uid} এর ব্যালেন্সে {amt} টাকা যোগ করা হয়েছে। বর্তমান: {balances[str(uid)]}", chat_id)
    try:
        send_telegram_message(f"💰 অ্যাডমিন আপনার অ্যাকাউন্টে {amt} টাকা যোগ করেছেন। বর্তমান ব্যালেন্স: {balances[str(uid)]} টাকা", uid)
    except:
        pass
    return True

def admin_approvesell_cmd(chat_id, sell_id, amount_str=None):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        for req in sell_requests:
            if req["id"] == sell_id and req["status"] == "pending":
                if amount_str is None:
                    acc_type = req.get("type", "2fa")
                    price = config["price_2fa_sell"] if acc_type == "2fa" else config["price_normal_sell"]
                    amount = price * len(req["accounts"])
                else:
                    try:
                        amount = float(amount_str)
                    except:
                        send_telegram_message("❌ সঠিক পরিমাণ দিন।", chat_id)
                        return True
                req["status"] = "approved"
                user = req["user_id"]
                balances[str(user)] = balances.get(str(user), 0) + amount
                save_balances()
                save_sell_requests()
                send_telegram_message(f"✅ সেল রিকোয়েস্ট {sell_id} অনুমোদিত। {amount} টাকা ইউজারের ব্যালেন্সে যোগ হয়েছে।", chat_id)
                send_telegram_message(f"✅ আপনার বিক্রয় রিকোয়েস্ট অনুমোদিত হয়েছে। {amount} টাকা আপনার ব্যালেন্সে যোগ করা হয়েছে। বর্তমান ব্যালেন্স: {balances[str(user)]} টাকা", user)
                break
        else:
            send_telegram_message("❌ রিকোয়েস্ট পাওয়া যায়নি বা ইতিমধ্যে প্রসেস করা হয়েছে।", chat_id)
    save_data_to_channel()
    return True

def admin_rejectsell_cmd(chat_id, sell_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        for req in sell_requests:
            if req["id"] == sell_id and req["status"] == "pending":
                req["status"] = "rejected"
                save_sell_requests()
                send_telegram_message(f"❌ সেল রিকোয়েস্ট {sell_id} বাতিল করা হয়েছে।", chat_id)
                send_telegram_message(f"❌ আপনার বিক্রয় রিকোয়েস্ট বাতিল করা হয়েছে।", req["user_id"])
                break
        else:
            send_telegram_message("❌ রিকোয়েস্ট পাওয়া যায়নি।", chat_id)
    save_data_to_channel()
    return True

def admin_approvewithdraw_cmd(chat_id, w_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        for req in withdraw_requests:
            if req["id"] == w_id and req["status"] == "pending":
                user = req["user_id"]
                bal = balances.get(str(user), 0)
                if bal < req["amount"]:
                    send_telegram_message(f"❌ ইউজারের ব্যালেন্স {bal} টাকা, প্রয়োজন {req['amount']} টাকা।", chat_id)
                    return True
                req["status"] = "approved"
                balances[str(user)] = bal - req["amount"]
                save_balances()
                save_withdraw_requests()
                send_telegram_message(f"✅ উইথড্র {w_id} অনুমোদিত। ইউজারের {req['amount']} টাকা কেটে নেওয়া হয়েছে।", chat_id)
                send_telegram_message(f"✅ আপনার {req['amount']} টাকার উইথড্র অনুমোদিত হয়েছে। অ্যাডমিন শীঘ্রই বিকাশে টাকা পাঠাবেন।", user)
                break
        else:
            send_telegram_message("❌ রিকোয়েস্ট পাওয়া যায়নি।", chat_id)
    save_data_to_channel()
    return True

def admin_rejectwithdraw_cmd(chat_id, w_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        for req in withdraw_requests:
            if req["id"] == w_id and req["status"] == "pending":
                req["status"] = "rejected"
                save_withdraw_requests()
                send_telegram_message(f"❌ উইথড্র {w_id} বাতিল করা হয়েছে।", chat_id)
                send_telegram_message(f"❌ আপনার {req['amount']} টাকার উইথড্র বাতিল করা হয়েছে।", req["user_id"])
                break
        else:
            send_telegram_message("❌ রিকোয়েস্ট পাওয়া যায়নি।", chat_id)
    save_data_to_channel()
    return True

def admin_sell_requests_cmd(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        pending = [r for r in sell_requests if r["status"] == "pending"]
    if not pending:
        send_telegram_message("কোনো পেন্ডিং সেল রিকোয়েস্ট নেই।", chat_id)
    else:
        lines = ["📊 পেন্ডিং সেল রিকোয়েস্ট:"]
        for r in pending:
            acc_names = ", ".join([a["username"] for a in r["accounts"]])
            acc_type = r.get("type", "2fa")
            lines.append(f"আইডি: {r['id']} | টাইপ: {acc_type} | ইউজার: {r['user_id']} | অ্যাকাউন্ট: {acc_names}")
            lines.append(f"অনুমোদন: /approvesell {r['id']} (অথবা /approvesell {r['id']} <amount>)")
            lines.append(f"বাতিল: /rejectsell {r['id']}\n")
        send_telegram_message("\n".join(lines), chat_id)
    return True

def admin_withdraw_requests_cmd(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    with data_lock:
        pending = [r for r in withdraw_requests if r["status"] == "pending"]
    if not pending:
        send_telegram_message("কোনো পেন্ডিং উইথড্র রিকোয়েস্ট নেই।", chat_id)
    else:
        lines = ["💳 পেন্ডিং উইথড্র রিকোয়েস্ট:"]
        for r in pending:
            lines.append(f"আইডি: {r['id']} | ইউজার: {r['user_id']} | পরিমাণ: {r['amount']} | বিকাশ: {r['bkash']}")
            lines.append(f"অনুমোদন: /approvewithdraw {r['id']}\nবাতিল: /rejectwithdraw {r['id']}\n")
        send_telegram_message("\n".join(lines), chat_id)
    return True

# ================== CENTRAL ADMIN DISPATCHER ==================
def handle_market_admin(chat_id, text):
    parts = text.split()
    if not parts:
        return False
    cmd = parts[0].lower()
    if cmd == "/stock":
        return admin_stock_cmd(chat_id)
    elif cmd == "/deletestock":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /deletestock <ইনডেক্স অথবা ইউজারনেম>", chat_id)
            return True
        return admin_deletestock_cmd(chat_id, parts[1])
    elif cmd == "/setprice":
        if len(parts) < 3:
            send_telegram_message("❌ ফরম্যাট: /setprice 2fa/normal <মূল্য>", chat_id)
            return True
        return admin_setprice_cmd(chat_id, ' '.join(parts[1:]))
    elif cmd == "/setsellprice":
        if len(parts) < 3:
            send_telegram_message("❌ ফরম্যাট: /setsellprice 2fa/normal <মূল্য>", chat_id)
            return True
        return admin_setsellprice_cmd(chat_id, ' '.join(parts[1:]))
    elif cmd == "/setbkash":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /setbkash <বিকাশ নম্বর>", chat_id)
            return True
        return admin_setbkash_cmd(chat_id, parts[1])
    elif cmd == "/approve":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /approve <deposit_id>", chat_id)
            return True
        return admin_approve_cmd(chat_id, parts[1])
    elif cmd == "/reject":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /reject <deposit_id>", chat_id)
            return True
        return admin_reject_cmd(chat_id, parts[1])
    elif cmd == "/deposits":
        return admin_deposits_cmd(chat_id)
    elif cmd == "/addbalance":
        if len(parts) < 3:
            send_telegram_message("❌ ফরম্যাট: /addbalance <user_id> <amount>", chat_id)
            return True
        return admin_addbalance_cmd(chat_id, parts[1], parts[2])
    elif cmd == "/approvesell":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /approvesell <id> [amount]", chat_id)
            return True
        sell_id = parts[1]
        amount_str = parts[2] if len(parts) > 2 else None
        return admin_approvesell_cmd(chat_id, sell_id, amount_str)
    elif cmd == "/rejectsell":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /rejectsell <id>", chat_id)
            return True
        return admin_rejectsell_cmd(chat_id, parts[1])
    elif cmd == "/approvewithdraw":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /approvewithdraw <id>", chat_id)
            return True
        return admin_approvewithdraw_cmd(chat_id, parts[1])
    elif cmd == "/rejectwithdraw":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /rejectwithdraw <id>", chat_id)
            return True
        return admin_rejectwithdraw_cmd(chat_id, parts[1])
    return False

# ================== MAIN COMMAND HANDLER ==================
def handle_telegram_commands():
    global last_update_id, maintenance_mode
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates"
    while True:
        try:
            params = {"timeout": 30, "allowed_updates": ["message", "callback_query"]}
            if last_update_id:
                params["offset"] = last_update_id + 1
            response = requests.get(url, params=params, timeout=35)
            data = response.json()
            if data.get("ok") and data.get("result"):
                for update in data["result"]:
                    last_update_id = update["update_id"]
                    if "callback_query" in update:
                        cb = update["callback_query"]
                        cbid = cb["id"]
                        chat_id = str(cb["message"]["chat"]["id"])
                        data_cb = cb["data"]
                        from_user = cb.get("from", {})
                        sender_username = from_user.get("username") or from_user.get("first_name", f"ID:{chat_id}")
                        user_info[chat_id] = sender_username
                        logger.info(f"Received callback: {data_cb} from {chat_id}")
                        answer_callback_query(cbid)
                        if data_cb == "buy_2fa":
                            start_buy_session(chat_id, "2fa")
                        elif data_cb == "buy_normal":
                            start_buy_session(chat_id, "normal")
                        elif data_cb == "sell_2fa":
                            start_sell_session(chat_id, "2fa")
                        elif data_cb == "sell_normal":
                            start_sell_session(chat_id, "normal")
                        elif data_cb == "addstock_2fa":
                            if str(chat_id) == ADMIN_CHAT_ID:
                                start_add_stock(chat_id, "2fa")
                        elif data_cb == "addstock_normal":
                            if str(chat_id) == ADMIN_CHAT_ID:
                                start_add_stock(chat_id, "normal")
                        elif data_cb == "cancel_buy":
                            if chat_id in buy_sessions:
                                del buy_sessions[chat_id]
                                send_telegram_message("❌ কেনা বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
                        elif data_cb == "cancel_deposit":
                            if chat_id in deposit_sessions:
                                del deposit_sessions[chat_id]
                                send_telegram_message("❌ ডিপোজিট বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
                        elif data_cb == "cancel_withdraw":
                            if chat_id in withdraw_sessions:
                                del withdraw_sessions[chat_id]
                                send_telegram_message("❌ উইথড্র বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
                        elif data_cb == "cancel_submission":
                            if chat_id in submission_sessions:
                                del submission_sessions[chat_id]
                                send_telegram_message("❌ জমা প্রক্রিয়া বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
                        elif data_cb == "cancel_sell":
                            if chat_id in sell_sessions:
                                del sell_sessions[chat_id]
                                send_telegram_message("❌ বিক্রয় বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
                        elif data_cb == "cancel_loss_recovery":
                            if chat_id in loss_recovery_sessions:
                                del loss_recovery_sessions[chat_id]
                                send_telegram_message("❌ লস রিকভারি বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
                        elif data_cb == "cancel_addstock":
                            if chat_id in add_stock_sessions:
                                del add_stock_sessions[chat_id]
                                send_telegram_message("❌ স্টক যোগ বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
                        continue

                    if "message" in update:
                        msg = update["message"]
                        chat_id = str(msg["chat"]["id"])
                        text = msg.get("text", "").strip()
                        from_user = msg.get("from", {})
                        sender_username = from_user.get("username") or from_user.get("first_name", f"ID:{chat_id}")
                        user_info[chat_id] = sender_username

                        if maintenance_mode and chat_id != ADMIN_CHAT_ID:
                            send_telegram_message("🔧 বট রক্ষণাবেক্ষণ মোডে আছে। পরে চেষ্টা করুন।", chat_id)
                            continue

                        if "document" in msg and chat_id in loss_recovery_sessions:
                            handle_loss_recovery_file(chat_id, msg)
                            continue

                        if "document" in msg and str(chat_id) == ADMIN_CHAT_ID:
                            caption = msg.get("caption", "").strip().lower()
                            if caption == "/restore":
                                file_id = msg["document"]["file_id"]
                                handle_restore(chat_id, file_id)
                            continue

                        # Smart Cancel
                        cancel_keywords = ["cancel", "/cancel", "/start"]
                        if text.strip().lower() in cancel_keywords:
                            cancelled = False
                            sessions = [
                                (support_sessions, "সাপোর্ট"),
                                (deposit_sessions, "ডিপোজিট"),
                                (submission_sessions, "জমা"),
                                (add_stock_sessions, "স্টক যোগ"),
                                (loss_recovery_sessions, "লস রিকভারি"),
                                (sell_sessions, "বিক্রয়"),
                                (withdraw_sessions, "উইথড্র"),
                                (buy_sessions, "কেনা")
                            ]
                            for sess_dict, name in sessions:
                                if chat_id in sess_dict:
                                    del sess_dict[chat_id]
                                    send_telegram_message(f"❌ {name} বাতিল করা হয়েছে।", chat_id)
                                    send_main_keyboard(chat_id)
                                    cancelled = True
                                    break
                            if cancelled:
                                continue
                            if text.strip().lower() == "/cancel":
                                send_main_keyboard(chat_id)
                                continue

                        # Active sessions
                        if chat_id in support_sessions:
                            forward = f"📩 সাপোর্ট মেসেজ\nইউজার: {sender_username} ({chat_id})\n\n{text}"
                            send_telegram_message(forward, ADMIN_CHAT_ID)
                            send_telegram_message("মেসেজ পাঠানো হয়েছে।", chat_id)
                            support_sessions.discard(chat_id)
                            send_main_keyboard(chat_id)
                            continue

                        if chat_id in deposit_sessions:
                            process_deposit_step(chat_id, text)
                            continue

                        if chat_id in submission_sessions:
                            process_submission_step(chat_id, text, sender_username)
                            continue

                        if chat_id in add_stock_sessions:
                            process_add_stock_step(chat_id, text)
                            continue

                        if chat_id in loss_recovery_sessions:
                            process_loss_recovery_step(chat_id, text)
                            continue

                        if chat_id in sell_sessions:
                            process_sell_step(chat_id, text)
                            continue

                        if chat_id in withdraw_sessions:
                            process_withdraw_step(chat_id, text)
                            continue

                        if chat_id in buy_sessions:
                            handle_buy(chat_id, text)
                            continue

                        # --- Button Handlers ---
                        if text == "📋 সাবমিট":
                            start_submission(chat_id, sender_username)
                            continue
                        elif text == "🎁 মাদার একাউন্ট":
                            handle_getmother(chat_id)
                            continue
                        elif text == "📞 সাপোর্ট":
                            support_sessions.add(chat_id)
                            send_telegram_message("📞 আপনার সমস্যা বা প্রশ্ন লিখুন। অ্যাডমিন সরাসরি দেখতে পাবেন।\nবাতিল করতে /cancel বা /start।", chat_id)
                            continue
                        elif text == "🛑 স্টপ":
                            with data_lock:
                                subscribed_users.discard(chat_id)
                                save_subscribers()
                            save_data_to_channel()
                            send_telegram_message("আপনার সাবস্ক্রিপশন বন্ধ করা হয়েছে। পুনরায় যুক্ত হতে /start দিন।", chat_id, reply_markup=get_main_keyboard(chat_id))
                            continue
                        elif text == "💰 ব্যালেন্স":
                            bal = balances.get(chat_id, 0)
                            send_telegram_message(f"💰 আপনার বর্তমান ব্যালেন্স: {bal} টাকা", chat_id)
                            send_main_keyboard(chat_id)
                            continue
                        elif text == "💸 ডিপোজিট":
                            start_deposit(chat_id)
                            continue
                        elif text == "🛒 একাউন্ট কিনুন":
                            inline_kb = {
                                "inline_keyboard": [
                                    [{"text": "2FA ON একাউন্ট", "callback_data": "buy_2fa"}],
                                    [{"text": "Normal একাউন্ট", "callback_data": "buy_normal"}]
                                ]
                            }
                            send_telegram_message("কোন ধরনের একাউন্ট কিনতে চান?", chat_id, reply_markup=inline_kb)
                            continue
                        elif text == "💰 একাউন্ট বিক্রয় করুন":
                            inline_kb = {
                                "inline_keyboard": [
                                    [{"text": "2FA ON একাউন্ট", "callback_data": "sell_2fa"}],
                                    [{"text": "Normal একাউন্ট", "callback_data": "sell_normal"}]
                                ]
                            }
                            send_telegram_message("কোন ধরনের একাউন্ট বিক্রয় করতে চান?", chat_id, reply_markup=inline_kb)
                            continue
                        elif text == "📥 ডিপোজিট রিকোয়েস্ট":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ অ্যাডমিন নন।", chat_id)
                            else:
                                admin_deposits_cmd(chat_id)
                            send_main_keyboard(chat_id)
                            continue
                        elif text == "➕ স্টক যোগ করুন":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ অ্যাডমিন নন।", chat_id)
                            else:
                                inline_kb = {
                                    "inline_keyboard": [
                                        [{"text": "2FA ON একাউন্ট", "callback_data": "addstock_2fa"}],
                                        [{"text": "Normal একাউন্ট", "callback_data": "addstock_normal"}]
                                    ]
                                }
                                send_telegram_message("কোন ধরনের স্টক যোগ করবেন?", chat_id, reply_markup=inline_kb)
                            continue
                        elif text == "📦 স্টক দেখুন":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ অ্যাডমিন নন।", chat_id)
                            else:
                                admin_stock_cmd(chat_id)
                            send_main_keyboard(chat_id)
                            continue
                        elif text == "🗑️ স্টক ডিলিট":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ অ্যাডমিন নন।", chat_id)
                            else:
                                send_telegram_message("🗑️ স্টক ডিলিট করতে কমান্ড ব্যবহার করুন:\n/deletestock <ইনডেক্স> বা /deletestock <ইউজারনেম>\nস্টক দেখতে /stock দিন।", chat_id)
                            send_main_keyboard(chat_id)
                            continue
                        elif text == "🔄 লস রিকভারি":
                            start_loss_recovery(chat_id)
                            continue
                        elif text == "💸 উইথড্র":
                            start_withdraw(chat_id)
                            continue
                        elif text == "📊 সেল requests":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ অ্যাডমিন নন।", chat_id)
                            else:
                                admin_sell_requests_cmd(chat_id)
                            send_main_keyboard(chat_id)
                            continue
                        elif text == "💳 withdraw requests":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ অ্যাডমিন নন।", chat_id)
                            else:
                                admin_withdraw_requests_cmd(chat_id)
                            send_main_keyboard(chat_id)
                            continue

                        if text.startswith("/"):
                            if handle_market_admin(chat_id, text):
                                continue
                            if text.startswith("/start"):
                                with data_lock:
                                    subscribed_users.add(chat_id)
                                    save_subscribers()
                                save_data_to_channel()
                                send_telegram_message("✨ আমাদের বটে স্বাগতম! ✨", chat_id, reply_markup=get_main_keyboard(chat_id))
                                continue
                            elif text == "/stop":
                                with data_lock:
                                    subscribed_users.discard(chat_id)
                                    save_subscribers()
                                save_data_to_channel()
                                send_telegram_message("সাবস্ক্রিপশন বন্ধ করা হয়েছে। পুনরায় যুক্ত হতে /start দিন।", chat_id, reply_markup=get_main_keyboard(chat_id))
                                continue
                            elif text.startswith("/addmother"):
                                args = text[len("/addmother"):].strip()
                                handle_addmother(chat_id, args)
                                continue
                            elif text == "/getmother":
                                handle_getmother(chat_id)
                                continue
                            elif text == "/motherlist":
                                handle_motherlist(chat_id)
                                continue
                            elif text.startswith("/deletemother"):
                                args = text[len("/deletemother"):].strip()
                                handle_deletemother(chat_id, args)
                                continue
                            elif text.startswith("/maintenance"):
                                args = text[len("/maintenance"):].strip()
                                handle_maintenance(chat_id, args)
                                continue
                            elif text == "/users":
                                handle_admin_users(chat_id)
                                continue
                            elif text.startswith("/broadcast"):
                                if len(text.split()) < 2:
                                    send_telegram_message("❌ ফরম্যাট: /broadcast <মেসেজ>", chat_id)
                                else:
                                    handle_admin_broadcast(chat_id, text.split(maxsplit=1)[1])
                                continue
                            elif text.startswith("/send"):
                                parts = text.split(maxsplit=2)
                                if len(parts) < 3:
                                    send_telegram_message("❌ ফরম্যাট: /send <user_id> <মেসেজ>", chat_id)
                                else:
                                    handle_admin_send(chat_id, parts[1], parts[2])
                                continue
                            elif text == "/backup":
                                handle_backup(chat_id)
                                continue
                            else:
                                send_telegram_message("❌ অজানা কমান্ড।", chat_id)
                                continue
        except Exception as e:
            logger.exception("Telegram Command Error:")
        time.sleep(1)

# ================== FLASK ROUTE ==================
@app.route("/")
def home():
    return "Bot Running Successfully!"

# ================== WEBHOOK CLEANUP ==================
def ensure_polling():
    try:
        resp = requests.post(f"https://api.telegram.org/bot{BOT_TOKEN}/deleteWebhook?drop_pending_updates=true")
        if resp.json().get("ok"):
            logger.info("Webhook deleted, polling mode active.")
        else:
            logger.warning("Webhook deletion may have failed.")
    except Exception as e:
        logger.error(f"Failed to delete webhook: {e}")

# ================== MAIN ==================
if __name__ == "__main__":
    load_mother_accounts()
    load_user_cooldowns()
    load_subscribers()
    load_market()
    load_sell_requests()
    load_withdraw_requests()

    ensure_polling()

    auto_restore_from_channel()

    threading.Thread(target=handle_telegram_commands, daemon=True).start()
    threading.Thread(target=auto_backup_loop, daemon=True).start()

    save_data_to_channel()

    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
