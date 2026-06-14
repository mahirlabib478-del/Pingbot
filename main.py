import os
import time
import threading
import requests
import datetime
import json
from flask import Flask
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import io

# ================== CONFIG ==================
BOT_TOKEN = "8808046131:AAHG_cjPmnzUnY0vCDcmnhmu3XVPGfMlcug"
ADMIN_CHAT_ID = "2035024902"
CHANNEL_ID = -1003903695158  # আপনার প্রাইভেট চ্যানেলের আইডি

# ================== FILE PATHS ==================
HASH_FILE = "pushed_hashes.json"
MOTHER_FILE = "mother_accounts.json"
COOLDOWN_FILE = "user_cooldowns.json"
SUBSCRIBERS_FILE = "subscribers.json"
USER_INFO_FILE = "user_info.json"
ACCOUNTS_FILE = "accounts.json"
BALANCES_FILE = "balances.json"
DEPOSITS_FILE = "deposits.json"
CONFIG_FILE = "config.json"

# ================== FLASK APP ==================
app = Flask(__name__)

# ================== GLOBALS ==================
last_update_id = None
subscribed_users = set()
user_info = {}
pushed_hashes = set()
mother_accounts = []
user_last_request = {}
submission_sessions = {}
support_sessions = set()
maintenance_mode = False

# Marketplace
accounts = []
balances = {}
deposits = []
config = {
    "bkash_number": "",
    "price_per_account": 1.70,
    "group_chat_id": "",
    "channel_id": str(CHANNEL_ID)
}
deposit_sessions = {}
buy_sessions = set()
add_stock_sessions = {}

# Loss Recovery
loss_recovery_sessions = {}

# Channel backup
last_message_id = None
channel_lock = threading.RLock()

data_lock = threading.RLock()

# ================== FILE I/O ==================
def load_hashes():
    global pushed_hashes
    try:
        with open(HASH_FILE, "r") as f:
            pushed_hashes = set(json.load(f))
    except:
        pushed_hashes = set()

def save_hashes():
    with data_lock:
        try:
            with open(HASH_FILE, "w") as f:
                json.dump(list(pushed_hashes), f)
        except Exception as e:
            print(f"Hash save error: {e}")

def load_mother_accounts():
    global mother_accounts
    try:
        with open(MOTHER_FILE, "r") as f:
            mother_accounts = json.load(f)
    except:
        mother_accounts = []

def save_mother_accounts():
    with data_lock:
        try:
            with open(MOTHER_FILE, "w") as f:
                json.dump(mother_accounts, f, indent=2)
        except Exception as e:
            print(f"Mother save error: {e}")

def load_user_cooldowns():
    global user_last_request
    try:
        with open(COOLDOWN_FILE, "r") as f:
            user_last_request = json.load(f)
    except:
        user_last_request = {}

def save_user_cooldowns():
    with data_lock:
        try:
            with open(COOLDOWN_FILE, "w") as f:
                json.dump(user_last_request, f, indent=2)
        except Exception as e:
            print(f"Cooldown save error: {e}")

def load_subscribers():
    global subscribed_users, user_info
    try:
        with open(SUBSCRIBERS_FILE, "r") as f:
            data = json.load(f)
            subscribed_users = set(data.get("subscribed", []))
    except:
        subscribed_users = set()
    try:
        with open(USER_INFO_FILE, "r") as f:
            user_info = json.load(f)
    except:
        user_info = {}

def save_subscribers():
    with data_lock:
        try:
            with open(SUBSCRIBERS_FILE, "w") as f:
                json.dump({"subscribed": list(subscribed_users)}, f)
        except Exception as e:
            print(f"Subscribers save error: {e}")

def save_user_info():
    with data_lock:
        try:
            with open(USER_INFO_FILE, "w") as f:
                json.dump(user_info, f, indent=2)
        except Exception as e:
            print(f"User info save error: {e}")

def load_market():
    global accounts, balances, deposits, config
    try:
        with open(ACCOUNTS_FILE, "r") as f:
            accounts = json.load(f)
    except:
        accounts = []
    try:
        with open(BALANCES_FILE, "r") as f:
            balances = json.load(f)
    except:
        balances = {}
    try:
        with open(DEPOSITS_FILE, "r") as f:
            deposits = json.load(f)
    except:
        deposits = []
    try:
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
            if "channel_id" not in config:
                config["channel_id"] = str(CHANNEL_ID)
    except:
        config = {"bkash_number": "", "price_per_account": 1.70, "group_chat_id": "", "channel_id": str(CHANNEL_ID)}

def save_accounts():
    with data_lock:
        with open(ACCOUNTS_FILE, "w") as f:
            json.dump(accounts, f, indent=2)

def save_balances():
    with data_lock:
        with open(BALANCES_FILE, "w") as f:
            json.dump(balances, f, indent=2)

def save_deposits():
    with data_lock:
        with open(DEPOSITS_FILE, "w") as f:
            json.dump(deposits, f, indent=2)

def save_config():
    with data_lock:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f, indent=2)

def save_all():
    """সব ডেটা ফাইল ও চ্যানেলে সেভ করে"""
    save_accounts()
    save_balances()
    save_deposits()
    save_config()
    save_subscribers()
    save_user_info()
    save_mother_accounts()
    save_user_cooldowns()
    save_hashes()
    save_data_to_channel()

# ================== TELEGRAM HELPERS ==================
def send_telegram_message(text, chat_id, reply_markup=None, parse_mode=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {"chat_id": chat_id, "text": text}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    if parse_mode:
        payload["parse_mode"] = parse_mode
    try:
        requests.post(url, json=payload, timeout=10)
    except Exception as e:
        print(f"Send error to {chat_id}: {e}")

def send_telegram_document(file_bytes, filename, chat_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    try:
        files = {'document': (filename, file_bytes,
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        requests.post(url, data={"chat_id": chat_id}, files=files, timeout=30)
    except Exception as e:
        print(f"Document send error: {e}")

def forward_telegram_document(chat_id, from_chat_id, message_id):
    """Forward a document by file_id"""
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/forwardMessage"
    payload = {
        "chat_id": chat_id,
        "from_chat_id": from_chat_id,
        "message_id": message_id
    }
    try:
        requests.post(url, json=payload, timeout=10)
    except Exception as e:
        print(f"Forward error: {e}")

def broadcast_message(text):
    for chat_id in list(subscribed_users):
        send_telegram_message(text, chat_id)

# ================== CHANNEL BACKUP ==================
def save_data_to_channel():
    global last_message_id
    with channel_lock:
        try:
            data = {
                "accounts": accounts,
                "balances": balances,
                "deposits": deposits,
                "config": config,
                "subscribed_users": list(subscribed_users),
                "user_info": user_info,
                "mother_accounts": mother_accounts,
                "user_cooldowns": user_last_request,
                "timestamp": datetime.datetime.now().isoformat()
            }
            data_str = json.dumps(data, indent=2)

            if last_message_id:
                try:
                    requests.post(
                        f"https://api.telegram.org/bot{BOT_TOKEN}/editMessageText",
                        json={
                            "chat_id": CHANNEL_ID,
                            "message_id": last_message_id,
                            "text": data_str
                        },
                        timeout=10
                    )
                except:
                    resp = requests.post(
                        f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage",
                        json={"chat_id": CHANNEL_ID, "text": data_str},
                        timeout=10
                    )
                    if resp.status_code == 200:
                        last_message_id = resp.json()["result"]["message_id"]
            else:
                resp = requests.post(
                    f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage",
                    json={"chat_id": CHANNEL_ID, "text": data_str},
                    timeout=10
                )
                if resp.status_code == 200:
                    last_message_id = resp.json()["result"]["message_id"]
        except Exception as e:
            print(f"Channel backup error: {e}")

def auto_backup_loop():
    """প্রতি ২৪ ঘণ্টায় চ্যানেলে ব্যাকআপ আপডেট"""
    while True:
        time.sleep(86400)
        save_data_to_channel()
        send_telegram_message("🔄 অটো ব্যাকআপ সম্পন্ন হয়েছে", ADMIN_CHAT_ID)

# ================== KEYBOARD ==================
def get_keyboard(chat_id):
    keyboard = [
        ["💰 ব্যালেন্স", "💸 ডিপোজিট"],
        ["🛒 একাউন্ট কিনুন"],
        ["📋 সাবমিট", "🎁 মাদার একাউন্ট"],
        ["📞 সাপোর্ট", "🛑 স্টপ"],
        ["🔄 লস রিকভারি"]                   # নতুন বাটন
    ]
    if str(chat_id) == ADMIN_CHAT_ID:
        keyboard.append(["📥 ডিপোজিট রিকোয়েস্ট", "➕ স্টক যোগ করুন"])
        keyboard.append(["📦 স্টক দেখুন", "🗑️ স্টক ডিলিট"])
    return {"keyboard": keyboard, "resize_keyboard": True, "one_time_keyboard": False}

def remove_keyboard():
    return {"remove_keyboard": True}

def send_main_keyboard(chat_id, text=" "):
    send_telegram_message(text, chat_id, reply_markup=get_keyboard(chat_id))

# ================== EXCEL GENERATORS ==================
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
    headers = ["Username", "Password", "2FA Key"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for acc in bought:
        ws.append([acc["username"], acc["password"], acc.get("fa_key", "")])
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
def start_submission(chat_id, sender_username):
    submission_sessions[chat_id] = {
        "step": "username",
        "data": {},
        "username": sender_username
    }
    send_telegram_message(
        "📋 দয়া করে আপনার **ইউজারনেম** লিস্ট দিন (প্রতি লাইনে একটি করে):\n\n"
        "উদাহরণ:\nuser1\nuser2\nuser3\n\n/start দিয়ে আবার শুরু করতে পারেন।",
        chat_id
    )

def process_submission_step(chat_id, text, sender_username):
    if chat_id not in submission_sessions:
        return False
    session = submission_sessions[chat_id]
    step = session["step"]
    if text.strip().lower() == "/start":
        del submission_sessions[chat_id]
        send_telegram_message("❌ জমা প্রক্রিয়া বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True

    if step == "username":
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["data"]["usernames"] = lines
        session["step"] = "password"
        send_telegram_message(
            "🔑 এখন **পাসওয়ার্ড** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\n"
            f"আপনার ইউজারনেম সংখ্যা: {len(lines)}",
            chat_id
        )
        return True

    elif step == "password":
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        usernames = session["data"]["usernames"]
        if len(lines) != len(usernames):
            send_telegram_message(
                f"❌ পাসওয়ার্ড সংখ্যা ({len(lines)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। পুনরায় সঠিক লিস্ট দিন।",
                chat_id
            )
            return True
        session["data"]["passwords"] = lines
        session["step"] = "2fa"
        send_telegram_message(
            "🔐 এখন **2FA কী** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\n"
            "যদি 2FA না থাকে, লাইন ফাঁকা রাখবেন (শুধু এন্টার দিন)।",
            chat_id
        )
        return True

    elif step == "2fa":
        raw_lines = text.split('\n')
        twofa_list = [l.strip() for l in raw_lines]
        usernames = session["data"]["usernames"]
        if len(twofa_list) != len(usernames):
            send_telegram_message(
                f"❌ 2FA কী সংখ্যা ({len(twofa_list)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। প্রতিটি ইউজারনেমের জন্য একটি লাইন (খালিও হতে পারে) দিন।",
                chat_id
            )
            return True
        session["data"]["twofa"] = twofa_list
        session["step"] = "bkash"
        send_telegram_message("💳 দয়া করে আপনার **বিকাশ নম্বর** দিন:", chat_id)
        return True

    elif step == "bkash":
        bkash_number = text.strip()
        if not bkash_number:
            send_telegram_message("⚠️ বিকাশ নম্বর খালি রাখা যাবে না। আবার দিন।", chat_id)
            return True
        session["data"]["bkash"] = bkash_number
        usernames = session["data"]["usernames"]
        passwords = session["data"]["passwords"]
        twofa_list = session["data"]["twofa"]
        bkash = session["data"]["bkash"]
        tg_username = session["username"]
        excel_bytes = generate_submission_excel(usernames, passwords, twofa_list, bkash, tg_username)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"submission_{chat_id}_{timestamp}.xlsx"
        send_telegram_document(excel_bytes, filename, ADMIN_CHAT_ID)
        send_telegram_message(
            "✅ আপনার অ্যাকাউন্ট সফলভাবে জমা হয়েছে।\n\nঅ্যাডমিন শীঘ্রই যোগাযোগ করবে। ধন্যবাদ! 🙏",
            chat_id
        )
        del submission_sessions[chat_id]
        send_main_keyboard(chat_id)
        return True
    return False

# ================== LOSS RECOVERY HANDLER ==================
def start_loss_recovery(chat_id):
    loss_recovery_sessions[chat_id] = {
        "step": "usernames",
        "data": {}
    }
    send_telegram_message(
        "⚠️ সতর্কতা: ভুল তথ্য দিলে লস রিকভারি পাবেন না। সকল তথ্য ম্যানুয়ালি যাচাই করা হবে।\n\n"
        "অনুগ্রহ করে সঠিক তথ্য দিন।\n\n"
        "আপনার কেনা অ্যাকাউন্টগুলোর ইউজারনেম লিস্ট দিন (প্রতি লাইনে একটি):",
        chat_id
    )

def process_loss_recovery_step(chat_id, text):
    if chat_id not in loss_recovery_sessions:
        return False
    session = loss_recovery_sessions[chat_id]
    step = session["step"]

    if text.strip().lower() == "/cancel":
        del loss_recovery_sessions[chat_id]
        send_telegram_message("❌ লস রিকভারি বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True

    if step == "usernames":
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["data"]["usernames"] = lines
        session["step"] = "cookie_date"
        send_telegram_message(
            "📅 কত তারিখে কুকিজ সাবমিট করেছিলেন? (শুধু তারিখের সংখ্যা লিখুন, যেমন: 13 বা 26)",
            chat_id
        )
        return True

    elif step == "cookie_date":
        date_str = text.strip()
        if not date_str.isdigit():
            send_telegram_message("⚠️ দয়া করে শুধু সংখ্যা দিন (13, 26 ইত্যাদি)।", chat_id)
            return True
        session["data"]["cookie_date"] = date_str
        session["step"] = "report_file"
        send_telegram_message(
            "📎 এখন রিপোর্ট ফেইল হওয়ার দিনের রিপোর্ট ফাইলটি দিন (স্ক্রিনশট/ডকুমেন্ট)।",
            chat_id
        )
        return True

    elif step == "report_file":
        # If user sends text instead of file, remind them
        send_telegram_message("⚠️ দয়া করে একটি ফাইল/স্ক্রিনশট পাঠান, টেক্সট নয়।", chat_id)
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

        # সব তথ্য অ্যাডমিনকে পাঠাই
        usernames = session["data"]["usernames"]
        cookie_date = session["data"]["cookie_date"]
        bkash = session["data"]["bkash"]
        file_id = session["data"].get("report_file_id")
        file_message_id = session["data"].get("report_message_id")

        # অ্যাডমিনকে মেসেজ
        admin_text = (
            "🔄 **নতুন লস রিকভারি রিকোয়েস্ট**\n\n"
            f"👤 ইউজার: {user_info.get(chat_id, chat_id)} (`{chat_id}`)\n"
            f"📅 কুকি সাবমিটের তারিখ: {cookie_date}\n"
            f"💳 বিকাশ: {bkash}\n"
            f"📞 হোয়াটসঅ্যাপ: {whatsapp}\n"
            f"🔑 ইউজারনেম: " + ", ".join(usernames)
        )
        send_telegram_message(admin_text, ADMIN_CHAT_ID, parse_mode="Markdown")

        # রিপোর্ট ফাইল ফরওয়ার্ড (যদি থাকে)
        if file_message_id:
            forward_telegram_document(ADMIN_CHAT_ID, chat_id, file_message_id)
        elif file_id:
            # সরাসরি ফাইল আইডি দিয়েও পাঠানো যায়, কিন্তু ফরওয়ার্ডই নিরাপদ
            send_telegram_message("⚠️ রিপোর্ট ফাইল ফরওয়ার্ড করা যায়নি, কারণ মেসেজ আইডি পাওয়া যায়নি।", ADMIN_CHAT_ID)

        # ইউজারকে কনফার্মেশন
        send_telegram_message(
            "✅ আপনার লস রিকভারি রিকোয়েস্ট জমা হয়েছে। অ্যাডমিন শীঘ্রই আপনার সাথে যোগাযোগ করবে।",
            chat_id
        )
        del loss_recovery_sessions[chat_id]
        send_main_keyboard(chat_id)
        return True
    return False

def handle_loss_recovery_file(chat_id, message):
    """যখন ইউজার 'রিপোর্ট ফাইল' স্টেপে ফাইল পাঠায়"""
    if chat_id not in loss_recovery_sessions:
        return
    session = loss_recovery_sessions[chat_id]
    if session["step"] != "report_file":
        return  # এখন ফাইল আসার কথা নয়, ইগনোর

    doc = message.get("document")
    if not doc:
        send_telegram_message("⚠️ দয়া করে একটি ফাইল পাঠান।", chat_id)
        return

    file_id = doc.get("file_id")
    message_id = message.get("message_id")

    session["data"]["report_file_id"] = file_id
    session["data"]["report_message_id"] = message_id
    session["step"] = "bkash"
    send_telegram_message("💳 দয়া করে আপনার বিকাশ নম্বর দিন (যেটি ব্যবহার করেছিলেন):", chat_id)

# ================== FREE MOTHER ACCOUNT ==================
def handle_addmother(chat_id, args):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    if not args:
        send_telegram_message("❌ ফরম্যাট: /addmother username password [2fa_key]", chat_id)
        return
    parts = args.split(maxsplit=2)
    if len(parts) < 2:
        send_telegram_message("❌ username এবং password আবশ্যক।", chat_id)
        return
    username = parts[0]
    password = parts[1]
    fa_key = parts[2] if len(parts) == 3 else ""
    for acc in mother_accounts:
        if acc["username"] == username and acc["password"] == password:
            send_telegram_message("⚠️ এই অ্যাকাউন্টটি আগেই যোগ করা আছে।", chat_id)
            return
    mother_accounts.append({
        "username": username,
        "password": password,
        "fa_key": fa_key,
        "assigned_to": None,
        "assigned_at": None
    })
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
        send_telegram_message(
            f"⏳ অনুগ্রহ করে অপেক্ষা করুন। পরবর্তী অ্যাকাউন্ট {wait_min} মিনিট {wait_sec_rem} সেকেন্ড পর নিতে পারবেন।",
            chat_id)
        send_main_keyboard(chat_id)
        return
    for acc in mother_accounts:
        if acc["assigned_to"] is None:
            acc["assigned_to"] = str(chat_id)
            acc["assigned_at"] = now
            save_mother_accounts()
            user_last_request[str(chat_id)] = now
            save_user_cooldowns()
            save_data_to_channel()
            msg = (
                "🎁 আপনার মাদার অ্যাকাউন্ট:\n\n"
                f"👤 ইউজারনেম: {acc['username']}\n"
                f"🔑 পাসওয়ার্ড: {acc['password']}"
            )
            if acc["fa_key"]:
                msg += f"\n🔐 2FA Key: {acc['fa_key']}"
            send_telegram_message(msg, chat_id)
            send_main_keyboard(chat_id)
            return
    send_telegram_message("❌ কোনো মাদার অ্যাকাউন্ট উপলব্ধ নেই। পরে আবার চেষ্টা করুন।", chat_id)
    send_main_keyboard(chat_id)

def handle_motherlist(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
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
        lines.append(
            f"{i}. ইউজার: {acc['username']} | পাস: {acc['password']} | 2FA: {twofa} | বরাদ্দ: {assigned}"
        )
    send_telegram_message("\n".join(lines), chat_id)

def handle_deletemother(chat_id, arg):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    if not arg:
        send_telegram_message("❌ ফরম্যাট: /deletemother <ইনডেক্স নাম্বার অথবা ইউজারনেম>", chat_id)
        return
    arg = arg.strip()
    try:
        idx = int(arg) - 1
        if 0 <= idx < len(mother_accounts):
            deleted = mother_accounts.pop(idx)
            save_mother_accounts()
            save_data_to_channel()
            send_telegram_message(f"✅ মাদার অ্যাকাউন্ট `{deleted['username']}` মুছে ফেলা হয়েছে।", chat_id)
            return
        else:
            send_telegram_message("❌ ভুল ইনডেক্স। /motherlist দিয়ে নম্বর দেখুন।", chat_id)
            return
    except ValueError:
        for i, acc in enumerate(mother_accounts):
            if acc["username"] == arg:
                mother_accounts.pop(i)
                save_mother_accounts()
                save_data_to_channel()
                send_telegram_message(f"✅ ইউজারনেম `{arg}` যুক্ত মাদার অ্যাকাউন্ট মুছে ফেলা হয়েছে।", chat_id)
                return
        send_telegram_message(f"❌ `{arg}` নামে কোনো মাদার অ্যাকাউন্ট পাওয়া যায়নি।", chat_id)

# ================== MAINTENANCE MODE ==================
def handle_maintenance(chat_id, args):
    global maintenance_mode
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    args = args.strip().lower()
    if args == "on":
        maintenance_mode = True
        send_telegram_message("🔧 রক্ষণাবেক্ষণ মোড চালু করা হয়েছে। সাধারণ ইউজাররা এখন বট ব্যবহার করতে পারবে না।", chat_id)
    elif args == "off":
        maintenance_mode = False
        send_telegram_message("🔧 রক্ষণাবেক্ষণ মোড বন্ধ করা হয়েছে। বট এখন স্বাভাবিক ভাবে চলবে।", chat_id)
    else:
        status = "চালু" if maintenance_mode else "বন্ধ"
        send_telegram_message(f"🔧 রক্ষণাবেক্ষণ মোড বর্তমানে {status} আছে। /maintenance on/off দিয়ে পরিবর্তন করুন।", chat_id)

# ================== ADMIN BROADCAST & USERS ==================
def handle_admin_users(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
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
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    if not message.strip():
        send_telegram_message("❌ মেসেজ খালি রাখা যাবে না। ফরম্যাট: /broadcast <মেসেজ>", chat_id)
        return
    broadcast_message(f"📢 অ্যাডমিন থেকে বার্তা:\n\n{message}")
    send_telegram_message("✅ বার্তা সকল সাবস্ক্রাইবারকে পাঠানো হয়েছে।", chat_id)

def handle_admin_send(chat_id, target_id, message):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    if not target_id.isdigit():
        send_telegram_message("❌ সঠিক ইউজার আইডি দিন।", chat_id)
        return
    if not message.strip():
        send_telegram_message("❌ মেসেজ খালি রাখা যাবে না।", chat_id)
        return
    send_telegram_message(f"📩 অ্যাডমিন থেকে:\n\n{message}", target_id)
    send_telegram_message(f"✅ {target_id} কে মেসেজ পাঠানো হয়েছে।", chat_id)

# ================== BACKUP & RESTORE ==================
def handle_backup(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return

    save_all()

    backup = {
        "subscribed_users": list(subscribed_users),
        "user_info": user_info,
        "pushed_hashes": list(pushed_hashes),
        "mother_accounts": mother_accounts,
        "user_cooldowns": user_last_request,
        "accounts": accounts,
        "balances": balances,
        "deposits": deposits,
        "config": config,
        "timestamp": datetime.datetime.now().isoformat()
    }
    backup_json = json.dumps(backup, indent=2, ensure_ascii=False).encode('utf-8')
    filename = f"backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    send_telegram_document(backup_json, filename, ADMIN_CHAT_ID)
    send_telegram_message("✅ ব্যাকআপ ফাইল তৈরি ও পাঠানো হয়েছে। /restore এর মাধ্যমে এটি ব্যবহার করুন।", chat_id)

def handle_restore(chat_id, file_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        send_telegram_message("❌ আপনি এই কমান্ড ব্যবহার করতে পারবেন না।", chat_id)
        return
    try:
        get_file_url = f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={file_id}"
        resp = requests.get(get_file_url, timeout=15)
        resp.raise_for_status()
        file_data = resp.json()
        if not file_data.get("ok"):
            send_telegram_message("❌ ফাইল পাওয়া যায়নি।", chat_id)
            return
        file_path = file_data["result"]["file_path"]
        download_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
        file_resp = requests.get(download_url, timeout=30)
        file_resp.raise_for_status()
        content = file_resp.text
        backup = json.loads(content)
    except Exception as e:
        print(f"Restore download error: {e}")
        send_telegram_message("❌ ব্যাকআপ ফাইল ডাউনলোড বা পার্স করতে ব্যর্থ।", chat_id)
        return
    with data_lock:
        global subscribed_users, user_info, pushed_hashes, mother_accounts, user_last_request
        global accounts, balances, deposits, config
        subscribed_users = set(backup.get("subscribed_users", []))
        user_info = backup.get("user_info", {})
        pushed_hashes = set(backup.get("pushed_hashes", []))
        mother_accounts = backup.get("mother_accounts", [])
        user_last_request = backup.get("user_cooldowns", {})
        accounts = backup.get("market_accounts", backup.get("accounts", []))
        balances = backup.get("balances", {})
        deposits = backup.get("deposits", [])
        config = backup.get("config", {"bkash_number": "", "price_per_account": 1.70, "group_chat_id": "", "channel_id": str(CHANNEL_ID)})
        save_all()
    send_telegram_message("✅ ব্যাকআপ রিস্টোর সম্পন্ন হয়েছে।", chat_id)

# ================== ADMIN ADD STOCK FLOW ==================
def start_add_stock(chat_id):
    if str(chat_id) != ADMIN_CHAT_ID:
        return
    add_stock_sessions[chat_id] = {"step": "usernames"}
    send_telegram_message(
        "➕ স্টক যোগ করুন\n\n"
        "প্রথমে **ইউজারনেম** লিস্ট দিন (প্রতি লাইনে একটি করে):\n"
        "উদাহরণ:\nuser1\nuser2\nuser3\n\n/start দিয়ে বাতিল করুন।",
        chat_id
    )

def process_add_stock_step(chat_id, text):
    if chat_id not in add_stock_sessions:
        return False
    session = add_stock_sessions[chat_id]
    step = session["step"]
    if text.strip().lower() == "/start":
        del add_stock_sessions[chat_id]
        send_telegram_message("❌ স্টক যোগ বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True

    if step == "usernames":
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["usernames"] = lines
        session["step"] = "passwords"
        send_telegram_message(
            "🔑 এখন **পাসওয়ার্ড** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\n"
            f"আপনার ইউজারনেম সংখ্যা: {len(lines)}",
            chat_id
        )
        return True

    elif step == "passwords":
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        usernames = session["usernames"]
        if len(lines) != len(usernames):
            send_telegram_message(
                f"❌ পাসওয়ার্ড সংখ্যা ({len(lines)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। পুনরায় সঠিক লিস্ট দিন।",
                chat_id
            )
            return True
        session["passwords"] = lines
        session["step"] = "fa_keys"
        send_telegram_message(
            "🔐 এখন **2FA কী** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\n"
            "যদি 2FA না থাকে, লাইন ফাঁকা রাখবেন (শুধু এন্টার দিন)।",
            chat_id
        )
        return True

    elif step == "fa_keys":
        raw_lines = text.split('\n')
        fa_list = [l.strip() for l in raw_lines]
        usernames = session["usernames"]
        if len(fa_list) != len(usernames):
            send_telegram_message(
                f"❌ 2FA কী সংখ্যা ({len(fa_list)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। প্রতিটি ইউজারনেমের জন্য একটি লাইন (খালিও হতে পারে) দিন।",
                chat_id
            )
            return True
        count = len(usernames)
        for i in range(count):
            accounts.append({
                "username": usernames[i],
                "password": session["passwords"][i],
                "fa_key": fa_list[i]
            })
        save_accounts()
        save_data_to_channel()
        del add_stock_sessions[chat_id]
        send_telegram_message(f"✅ {count} টি অ্যাকাউন্ট স্টকে যোগ করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    return False

# ================== MARKETPLACE: DEPOSIT & BUY ==================
def start_deposit(chat_id):
    deposit_sessions[chat_id] = {"step": "amount"}
    bkash = config.get("bkash_number", "")
    if not bkash:
        send_telegram_message("⚠️ অ্যাডমিন এখনও বিকাশ নম্বর সেট করেননি। পরে চেষ্টা করুন।", chat_id)
        send_main_keyboard(chat_id)
        deposit_sessions.pop(chat_id, None)
        return
    send_telegram_message(
        f"💸 দয়া করে আপনার জমা করার টাকার পরিমাণ লিখুন (শুধু সংখ্যা, যেমন: 100)\n\n"
        f"বিকাশ নম্বর: {bkash}\n\n"
        "টাকা পাঠানোর পর ট্রানজেকশন আইডি সহ পুনরায় লিখবেন।",
        chat_id
    )

def process_deposit_step(chat_id, text):
    if chat_id not in deposit_sessions:
        return False
    session = deposit_sessions[chat_id]
    step = session["step"]
    if text.strip().lower() == "/cancel":
        deposit_sessions.pop(chat_id, None)
        send_telegram_message("❌ ডিপোজিট বাতিল করা হয়েছে।", chat_id)
        send_main_keyboard(chat_id)
        return True
    if step == "amount":
        try:
            amount = float(text)
            if amount <= 0:
                raise ValueError
        except:
            send_telegram_message("⚠️ সঠিক সংখ্যা দিন। /cancel দিয়ে বাতিল করুন।", chat_id)
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
        deposit_id = str(int(time.time() * 1000))
        deposit = {
            "id": deposit_id,
            "user_id": chat_id,
            "amount": amount,
            "trxid": trxid,
            "status": "pending",
            "time": time.time()
        }
        deposits.append(deposit)
        save_deposits()
        save_data_to_channel()
        deposit_sessions.pop(chat_id, None)
        admin_msg = (
            f"📥 নতুন ডিপোজিট রিকোয়েস্ট\n"
            f"আইডি: {deposit_id}\n"
            f"ইউজার: {user_info.get(chat_id, chat_id)} ({chat_id})\n"
            f"পরিমাণ: {amount} টাকা\n"
            f"ট্রানজেকশন আইডি: {trxid}\n"
            f"অনুমোদন করতে: /approve {deposit_id}\n"
            f"বাতিল করতে: /reject {deposit_id}"
        )
        send_telegram_message(admin_msg, ADMIN_CHAT_ID)
        send_telegram_message(
            f"✅ আপনার {amount} টাকার ডিপোজিট রিকোয়েস্ট জমা হয়েছে।\n"
            f"অ্যাডমিন অনুমোদন করলেই আপনার ব্যালেন্সে যোগ হবে।",
            chat_id
        )
        send_main_keyboard(chat_id)
        return True
    return False

def handle_buy(chat_id, quantity):
    try:
        qty = int(quantity)
        if qty <= 0:
            raise ValueError
    except:
        send_telegram_message("❌ সঠিক সংখ্যা দিন। যেমন: 3", chat_id)
        return
    if qty > len(accounts):
        send_telegram_message(f"❌ পর্যাপ্ত অ্যাকাউন্ট নেই। বর্তমান স্টক: {len(accounts)}", chat_id)
        return
    price = config.get("price_per_account", 1.70)
    total = qty * price
    user_balance = balances.get(chat_id, 0)
    if user_balance < total:
        send_telegram_message(
            f"❌ পর্যাপ্ত ব্যালেন্স নেই।\n"
            f"প্রয়োজন: {total} টাকা\n"
            f"আপনার ব্যালেন্স: {user_balance} টাকা\n"
            f"দয়া করে প্রথমে ডিপোজিট করুন।",
            chat_id
        )
        return
    bought = []
    for _ in range(qty):
        bought.append(accounts.pop(0))
    save_accounts()
    balances[chat_id] = user_balance - total
    save_balances()
    save_data_to_channel()
    excel_bytes = generate_purchase_excel(bought)
    filename = f"purchased_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    send_telegram_document(excel_bytes, filename, chat_id)
    send_telegram_message(
        f"✅ {qty} টি অ্যাকাউন্ট কেনা হয়েছে। মোট মূল্য: {total} টাকা।\n"
        f"অবশিষ্ট ব্যালেন্স: {balances[chat_id]} টাকা",
        chat_id
    )
    admin_msg = f"🛒 {user_info.get(chat_id, chat_id)} ({chat_id}) {qty} টি অ্যাকাউন্ট কিনেছে। মোট: {total} টাকা।"
    send_telegram_message(admin_msg, ADMIN_CHAT_ID)

# ================== ADMIN MARKETPLACE COMMANDS ==================
def handle_market_admin(chat_id, text):
    if str(chat_id) != ADMIN_CHAT_ID:
        return False
    parts = text.split(maxsplit=2)
    cmd = parts[0].lower()
    if cmd == "/addstock":
        start_add_stock(chat_id)
        return True
    elif cmd == "/stock":
        if not accounts:
            send_telegram_message("📭 কোনো অ্যাকাউন্ট স্টকে নেই।", chat_id)
        else:
            lines = [f"📦 স্টক ({len(accounts)} টি):"]
            for i, acc in enumerate(accounts, 1):
                lines.append(f"{i}. ইউজার: {acc['username']} | পাস: {acc['password']} | 2FA: {acc.get('fa_key', 'N/A')}")
            send_telegram_message("\n".join(lines), chat_id)
        return True
    elif cmd == "/deletestock":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /deletestock <ইনডেক্স নাম্বার অথবা ইউজারনেম>", chat_id)
            return True
        arg = parts[1].strip()
        try:
            idx = int(arg) - 1
            if 0 <= idx < len(accounts):
                deleted = accounts.pop(idx)
                save_accounts()
                save_data_to_channel()
                send_telegram_message(f"✅ স্টক থেকে অ্যাকাউন্ট `{deleted['username']}` মুছে ফেলা হয়েছে।", chat_id)
                return True
            else:
                send_telegram_message("❌ ভুল ইনডেক্স। /stock দিয়ে নম্বর দেখুন।", chat_id)
                return True
        except ValueError:
            for i, acc in enumerate(accounts):
                if acc["username"] == arg:
                    accounts.pop(i)
                    save_accounts()
                    save_data_to_channel()
                    send_telegram_message(f"✅ স্টক থেকে ইউজারনেম `{arg}` যুক্ত অ্যাকাউন্ট মুছে ফেলা হয়েছে।", chat_id)
                    return True
            send_telegram_message(f"❌ `{arg}` নামে কোনো অ্যাকাউন্ট পাওয়া যায়নি।", chat_id)
            return True
    elif cmd == "/bulkdelete":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /bulkdelete <সংখ্যা>\nউদাহরণ: /bulkdelete 5 (প্রথম 5 টি অ্যাকাউন্ট ডিলিট হবে)", chat_id)
            return True
        arg = parts[1].strip()
        try:
            count = int(arg)
            if count <= 0:
                raise ValueError
        except:
            send_telegram_message("❌ সঠিক সংখ্যা দিন (পজিটিভ পূর্ণসংখ্যা)।", chat_id)
            return True
        if count > len(accounts):
            send_telegram_message(f"❌ স্টকে মোট {len(accounts)} টি অ্যাকাউন্ট আছে। আপনি {count} টি ডিলিট করতে পারবেন না।", chat_id)
            return True
        deleted = accounts[:count]
        del accounts[:count]
        save_accounts()
        save_data_to_channel()
        usernames = [acc['username'] for acc in deleted]
        send_telegram_message(f"✅ স্টক থেকে প্রথম {count} টি অ্যাকাউন্ট ডিলিট করা হয়েছে:\n" + "\n".join(usernames), chat_id)
        return True
    elif cmd == "/setprice":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /setprice <মূল্য>", chat_id)
            return True
        try:
            price = float(parts[1])
            if price <= 0:
                raise ValueError
        except:
            send_telegram_message("❌ সঠিক মূল্য দিন (সংখ্যা)।", chat_id)
            return True
        config["price_per_account"] = price
        save_config()
        save_data_to_channel()
        send_telegram_message(f"✅ প্রতি অ্যাকাউন্টের মূল্য {price} টাকা নির্ধারণ করা হয়েছে।", chat_id)
        return True
    elif cmd == "/setbkash":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /setbkash <বিকাশ নম্বর>", chat_id)
            return True
        number = parts[1]
        config["bkash_number"] = number
        save_config()
        save_data_to_channel()
        send_telegram_message(f"✅ বিকাশ নম্বর {number} সেট করা হয়েছে।", chat_id)
        return True
    elif cmd == "/setgroup":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /setgroup <গ্রুপ চ্যাট আইডি>", chat_id)
            return True
        group_id = parts[1]
        config["group_chat_id"] = group_id
        save_config()
        save_data_to_channel()
        send_telegram_message(f"✅ ব্যাকআপ গ্রুপ আইডি {group_id} সেট করা হয়েছে।", chat_id)
        return True
    elif cmd == "/setchannel":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /setchannel <চ্যানেল আইডি>", chat_id)
            return True
        global CHANNEL_ID
        CHANNEL_ID = int(parts[1])
        config["channel_id"] = str(CHANNEL_ID)
        save_config()
        save_data_to_channel()
        send_telegram_message(f"✅ ব্যাকআপ চ্যানেল {CHANNEL_ID} সেট করা হয়েছে।", chat_id)
        return True
    elif cmd == "/approve":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /approve <deposit_id>", chat_id)
            return True
        deposit_id = parts[1]
        for dep in deposits:
            if dep["id"] == deposit_id and dep["status"] == "pending":
                dep["status"] = "approved"
                user = dep["user_id"]
                balances[user] = balances.get(user, 0) + dep["amount"]
                save_balances()
                save_deposits()
                save_data_to_channel()
                send_telegram_message(f"✅ ডিপোজিট {deposit_id} অনুমোদিত। ইউজারের ব্যালেন্স আপডেট হয়েছে।", chat_id)
                send_telegram_message(f"✅ আপনার {dep['amount']} টাকার ডিপোজিট অনুমোদিত হয়েছে। বর্তমান ব্যালেন্স: {balances[user]} টাকা", user)
                return True
        send_telegram_message("❌ ডিপোজিট পাওয়া যায়নি বা ইতিমধ্যে প্রসেস করা হয়েছে।", chat_id)
        return True
    elif cmd == "/reject":
        if len(parts) < 2:
            send_telegram_message("❌ ফরম্যাট: /reject <deposit_id>", chat_id)
            return True
        deposit_id = parts[1]
        for dep in deposits:
            if dep["id"] == deposit_id and dep["status"] == "pending":
                dep["status"] = "rejected"
                save_deposits()
                save_data_to_channel()
                send_telegram_message(f"❌ ডিপোজিট {deposit_id} বাতিল করা হয়েছে।", chat_id)
                send_telegram_message(f"❌ আপনার {dep['amount']} টাকার ডিপোজিট বাতিল করা হয়েছে।", dep["user_id"])
                return True
        send_telegram_message("❌ ডিপোজিট পাওয়া যায়নি।", chat_id)
        return True
    elif cmd == "/deposits":
        pending = [d for d in deposits if d["status"] == "pending"]
        if not pending:
            send_telegram_message("কোনো পেন্ডিং ডিপোজিট নেই।", chat_id)
        else:
            lines = ["⏳ পেন্ডিং ডিপোজিট:"]
            for d in pending:
                lines.append(f"আইডি: {d['id']} | ইউজার: {d['user_id']} | পরিমাণ: {d['amount']} | ট্রানজেকশন: {d['trxid']}")
            send_telegram_message("\n".join(lines), chat_id)
        return True
    elif cmd == "/addbalance":
        if len(parts) < 3:
            send_telegram_message("❌ ফরম্যাট: /addbalance <user_id> <amount>", chat_id)
            return True
        uid = parts[1]
        try:
            amt = float(parts[2])
        except:
            send_telegram_message("❌ সঠিক পরিমাণ দিন।", chat_id)
            return True
        balances[uid] = balances.get(uid, 0) + amt
        save_balances()
        save_data_to_channel()
        send_telegram_message(f"✅ {uid} এর ব্যালেন্সে {amt} টাকা যোগ করা হয়েছে। বর্তমান: {balances[uid]}", chat_id)
        try:
            send_telegram_message(f"💰 অ্যাডমিন আপনার অ্যাকাউন্টে {amt} টাকা যোগ করেছেন। বর্তমান ব্যালেন্স: {balances[uid]} টাকা", uid)
        except:
            pass
        return True
    return False

# ================== MAIN COMMAND HANDLER ==================
def handle_telegram_commands():
    global last_update_id, maintenance_mode
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates"
    while True:
        try:
            params = {"timeout": 30}
            if last_update_id:
                params["offset"] = last_update_id + 1
            response = requests.get(url, params=params, timeout=35)
            data = response.json()
            if data.get("ok") and data.get("result"):
                for update in data["result"]:
                    last_update_id = update["update_id"]
                    if "message" in update:
                        msg = update["message"]
                        chat_id = str(msg["chat"]["id"])
                        text = msg.get("text", "").strip()
                        from_user = msg.get("from", {})
                        sender_username = from_user.get("username") or \
                                         from_user.get("first_name", f"ID:{chat_id}")

                        user_info[chat_id] = sender_username
                        save_user_info()
                        save_data_to_channel()

                        if maintenance_mode and chat_id != ADMIN_CHAT_ID:
                            send_telegram_message("🔧 বট রক্ষণাবেক্ষণ মোডে আছে। পরে চেষ্টা করুন।", chat_id)
                            continue

                        # লস রিকভারি ফাইল হ্যান্ডলিং (ডকুমেন্ট আসলে)
                        if "document" in msg and chat_id in loss_recovery_sessions:
                            handle_loss_recovery_file(chat_id, msg)
                            continue

                        # Restore file handling (admin)
                        if "document" in msg:
                            caption = msg.get("caption", "")
                            if caption.strip().lower() == "/restore":
                                file_id = msg["document"]["file_id"]
                                handle_restore(chat_id, file_id)
                            continue

                        # Support session
                        if chat_id in support_sessions:
                            if text.lower() == "/cancel":
                                support_sessions.discard(chat_id)
                                send_telegram_message("সাপোর্ট বাতিল।", chat_id)
                                send_main_keyboard(chat_id)
                            else:
                                forward = f"📩 সাপোর্ট মেসেজ\nইউজার: {sender_username} ({chat_id})\n\n{text}"
                                send_telegram_message(forward, ADMIN_CHAT_ID)
                                send_telegram_message("মেসেজ পাঠানো হয়েছে।", chat_id)
                                support_sessions.discard(chat_id)
                                send_main_keyboard(chat_id)
                            continue

                        # Deposit flow active?
                        if chat_id in deposit_sessions:
                            process_deposit_step(chat_id, text)
                            continue

                        # Submission flow
                        if chat_id in submission_sessions:
                            process_submission_step(chat_id, text, sender_username)
                            continue

                        # Add stock flow (admin)
                        if chat_id in add_stock_sessions:
                            process_add_stock_step(chat_id, text)
                            continue

                        # Loss recovery text handling
                        if chat_id in loss_recovery_sessions:
                            process_loss_recovery_step(chat_id, text)
                            continue

                        # Buy session (waiting for quantity)
                        if chat_id in buy_sessions:
                            handle_buy(chat_id, text)
                            buy_sessions.discard(chat_id)
                            send_main_keyboard(chat_id)
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
                            send_telegram_message(
                                "📞 আপনার সমস্যা বা প্রশ্ন লিখুন। অ্যাডমিন সরাসরি দেখতে পাবেন।\n"
                                "বাতিল করতে /cancel লিখুন।",
                                chat_id
                            )
                            continue
                        elif text == "🛑 স্টপ":
                            subscribed_users.discard(chat_id)
                            save_subscribers()
                            save_data_to_channel()
                            send_telegram_message("আপনার সাবস্ক্রিপশন বন্ধ করা হয়েছে।", chat_id, reply_markup=remove_keyboard())
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
                            buy_sessions.add(chat_id)
                            price = config.get("price_per_account", 1.70)
                            send_telegram_message(
                                f"🛒 কতটি অ্যাকাউন্ট কিনতে চান? (সংখ্যা লিখুন)\n"
                                f"প্রতি অ্যাকাউন্টের মূল্য: {price} টাকা\n"
                                f"স্টক: {len(accounts)} টি",
                                chat_id
                            )
                            continue
                        elif text == "📥 ডিপোজিট রিকোয়েস্ট":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ আপনি এই বাটন ব্যবহার করতে পারবেন না।", chat_id)
                            else:
                                handle_market_admin(chat_id, "/deposits")
                            send_main_keyboard(chat_id)
                            continue
                        elif text == "➕ স্টক যোগ করুন":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ আপনি এই বাটন ব্যবহার করতে পারবেন না।", chat_id)
                            else:
                                start_add_stock(chat_id)
                            continue
                        elif text == "📦 স্টক দেখুন":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ অ্যাডমিন নন।", chat_id)
                            else:
                                handle_market_admin(chat_id, "/stock")
                            send_main_keyboard(chat_id)
                            continue
                        elif text == "🗑️ স্টক ডিলিট":
                            if str(chat_id) != ADMIN_CHAT_ID:
                                send_telegram_message("❌ অ্যাডমিন নন।", chat_id)
                            else:
                                send_telegram_message(
                                    "🗑️ স্টক ডিলিট করতে কমান্ড ব্যবহার করুন:\n"
                                    "/deletestock <ইনডেক্স> বা /deletestock <ইউজারনেম>\n"
                                    "একাধিক একসাথে ডিলিট: /bulkdelete <সংখ্যা>\n"
                                    "স্টক দেখতে /stock দিন।",
                                    chat_id
                                )
                            send_main_keyboard(chat_id)
                            continue
                        elif text == "🔄 লস রিকভারি":
                            start_loss_recovery(chat_id)
                            continue

                        # --- Text Commands ---
                        if text.startswith("/"):
                            # Market admin commands
                            if handle_market_admin(chat_id, text):
                                continue
                            # General commands
                            if text.startswith("/start"):
                                subscribed_users.add(chat_id)
                                save_subscribers()
                                save_data_to_channel()
                                reply = "✨ আমাদের বটে স্বাগতম! ✨"
                                send_telegram_message(reply, chat_id, reply_markup=get_keyboard(chat_id), parse_mode="Markdown")
                                continue
                            elif text == "/stop":
                                subscribed_users.discard(chat_id)
                                save_subscribers()
                                save_data_to_channel()
                                send_telegram_message("সাবস্ক্রিপশন বন্ধ করা হয়েছে।", chat_id, reply_markup=remove_keyboard())
                                continue
                            elif text.startswith("/addmother"):
                                args = text[len("/addmother"):].strip() if len(text) > len("/addmother") else ""
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
                            elif text.startswith("/users"):
                                handle_admin_users(chat_id)
                                continue
                            elif text.startswith("/broadcast"):
                                if len(text.split()) < 2:
                                    send_telegram_message("❌ ফরম্যাট: /broadcast <মেসেজ>", chat_id)
                                else:
                                    message = text.split(maxsplit=1)[1]
                                    handle_admin_broadcast(chat_id, message)
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
            print("Telegram Command Error:", e)
        time.sleep(1)

# ================== FLASK ROUTE ==================
@app.route("/")
def home():
    return "Bot Running Successfully!"

# ================== MAIN ==================
if __name__ == "__main__":
    load_hashes()
    load_mother_accounts()
    load_user_cooldowns()
    load_subscribers()
    load_market()

    threading.Thread(target=handle_telegram_commands, daemon=True).start()
    threading.Thread(target=auto_backup_loop, daemon=True).start()

    # প্রথমবার চ্যানেলে ব্যাকআপ পাঠাই
    save_data_to_channel()

    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
