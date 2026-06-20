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
BOT_USERNAME = "Ping478bot"

if not BOT_TOKEN or not ADMIN_CHAT_ID:
    raise RuntimeError("BOT_TOKEN and ADMIN_CHAT_ID must be set")

# ================== FILE PATHS ==================
MOTHER_FILE = "mother_accounts.json"
COOLDOWN_FILE = "user_cooldowns.json"
SUBSCRIBERS_FILE = "subscribers.json"
USER_INFO_FILE = "user_info.json"
BALANCES_FILE = "balances.json"
CONFIG_FILE = "config.json"
SUBMISSIONS_FILE = "submissions.json"
MOTHER_STOCK_FILE = "mother_stock.json"
REFERRALS_FILE = "referrals.json"
REFERRAL_BONUSES_FILE = "referral_bonuses.json"
LEADERBOARD_FILE = "leaderboard.json"

app = Flask(__name__)

# ================== GLOBALS ==================
last_update_id = None
subscribed_users = set()
user_info = {}                    # chat_id -> username/firstname
mother_accounts = []              # free mother accounts (old)
user_last_request = {}            # cooldowns for free mother account

# New variables
user_balances = {}                # chat_id -> balance (float)
submissions = []                  # list of submission dicts
mother_stock = []                 # buyable mother accounts
config = {
    "price_cookies": 3.5,
    "price_2fa": 3.0,
    "mother_price": 5.0,
    "referral_level1": 5.0,       # %
    "referral_level2": 1.0,       # %
    "monthly_target": 5000.0,     # taka
    "target_bonus": 2.0,          # %
    "lock_2fa": False,
    "lock_cookies": False,
    "bkash_number": ""
}
referrals = {}                    # invitee -> referrer
referral_bonuses = {}             # referrer -> total bonus earned
leaderboard = {}                  # user_id -> {total_income, total_ok}

# Session trackers
submission_sessions = {}
admin_approve_sessions = {}
admin_add_mother_session = {}
withdraw_sessions = {}
support_sessions = set()

data_lock = threading.RLock()
backup_lock = threading.Lock()

# ================== FILE I/O ==================
def load_json(filename, default):
    try:
        with open(filename, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return default

def save_json(filename, data):
    with data_lock:
        try:
            with open(filename, "w") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except IOError as e:
            logger.error(f"Error saving {filename}: {e}")

def load_all():
    global mother_accounts, user_last_request, subscribed_users, user_info
    global user_balances, submissions, config, referrals, referral_bonuses, leaderboard
    global mother_stock

    mother_accounts = load_json(MOTHER_FILE, [])
    user_last_request = load_json(COOLDOWN_FILE, {})
    subscribed_users = set(load_json(SUBSCRIBERS_FILE, {"subscribed": []}).get("subscribed", []))
    user_info = load_json(USER_INFO_FILE, {})

    user_balances = load_json(BALANCES_FILE, {})
    submissions = load_json(SUBMISSIONS_FILE, [])
    mother_stock = load_json(MOTHER_STOCK_FILE, [])
    referrals = load_json(REFERRALS_FILE, {})
    referral_bonuses = load_json(REFERRAL_BONUSES_FILE, {})
    leaderboard = load_json(LEADERBOARD_FILE, {})

    default_config = {
        "price_cookies": 3.5,
        "price_2fa": 3.0,
        "mother_price": 5.0,
        "referral_level1": 5.0,
        "referral_level2": 1.0,
        "monthly_target": 5000.0,
        "target_bonus": 2.0,
        "lock_2fa": False,
        "lock_cookies": False,
        "bkash_number": ""
    }
    loaded_config = load_json(CONFIG_FILE, default_config)
    for k in default_config:
        if k not in loaded_config:
            loaded_config[k] = default_config[k]
    config = loaded_config

def save_all():
    save_json(MOTHER_FILE, mother_accounts)
    save_json(COOLDOWN_FILE, user_last_request)
    save_json(SUBSCRIBERS_FILE, {"subscribed": list(subscribed_users)})
    save_json(USER_INFO_FILE, user_info)
    save_json(BALANCES_FILE, user_balances)
    save_json(SUBMISSIONS_FILE, submissions)
    save_json(MOTHER_STOCK_FILE, mother_stock)
    save_json(REFERRALS_FILE, referrals)
    save_json(REFERRAL_BONUSES_FILE, referral_bonuses)
    save_json(LEADERBOARD_FILE, leaderboard)
    save_json(CONFIG_FILE, config)

# ================== TELEGRAM HELPERS ==================
def send_telegram_message(text, chat_id, reply_markup=None, parse_mode=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {"chat_id": chat_id, "text": text}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    if parse_mode:
        payload["parse_mode"] = parse_mode
    for _ in range(3):
        try:
            resp = requests.post(url, json=payload, timeout=10)
            if resp.status_code == 429:
                retry = resp.json().get("parameters", {}).get("retry_after", 2)
                time.sleep(retry)
                continue
            return resp
        except Exception as e:
            logger.error(f"Send error: {e}")
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

def delete_message(chat_id, message_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/deleteMessage"
    try:
        requests.post(url, json={"chat_id": chat_id, "message_id": message_id}, timeout=10)
    except Exception as e:
        logger.error(f"Delete message error: {e}")

def broadcast_message(text):
    to_remove = []
    for uid in list(subscribed_users):
        try:
            resp = send_telegram_message(text, uid)
            if resp and resp.status_code == 403:
                to_remove.append(uid)
        except:
            to_remove.append(uid)
        time.sleep(0.05)
    if to_remove:
        with data_lock:
            for uid in to_remove:
                subscribed_users.discard(uid)
                user_info.pop(uid, None)
        save_all()

def answer_callback_query(callback_id, text=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/answerCallbackQuery"
    payload = {"callback_query_id": callback_id}
    if text:
        payload["text"] = text
    try:
        requests.post(url, json=payload, timeout=5)
    except Exception as e:
        logger.error(f"Callback answer error: {e}")

# ================== KEYBOARDS ==================
def get_main_keyboard(chat_id):
    kb = [
        ["💼 একাউন্ট সাবমিট", "👤 প্রোফাইল"],
        ["👥 রেফারেল", "💰 ব্যালেন্স"],
        ["💸 উইথড্র", "📊 লিডারবোর্ড"],
        ["🎁 ফ্রি মাদার একাউন্ট", "🛒 মাদার একাউন্ট কিনুন"]
    ]
    if str(chat_id) == ADMIN_CHAT_ID:
        kb.append(["🛠️ অ্যাডমিন প্যানেল"])
    return {"keyboard": kb, "resize_keyboard": True}

def admin_panel_keyboard():
    return {
        "keyboard": [
            ["📊 সাবমিটেড ফাইল", "⚙️ মূল্য নির্ধারণ"],
            ["👥 রেফারেল বোনাস %", "💵 মাসিক টার্গেট"],
            ["🔒 সাবমিট লক", "📢 ব্রডকাস্ট"],
            ["➕ মাদার একাউন্ট যোগ", "📦 মাদার স্টক"],
            ["💰 মাদার মূল্য সেট", "📋 ইউজার লিস্ট"],
            ["🔙 মূল মেনু"]
        ],
        "resize_keyboard": True
    }

# ================== EXCEL GENERATORS ==================
def generate_submission_excel(usernames, passwords, twofa_list, submitter_username):
    wb = Workbook()
    ws = wb.active
    ws.title = "Submission"
    headers = ["Username", "Password", "2FA Key", "Submitter"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for i in range(len(usernames)):
        row = [usernames[i], passwords[i] if i < len(passwords) else "",
               twofa_list[i] if i < len(twofa_list) else "", ""]
        ws.append(row)
    # Put submitter username only once in D2
    if len(usernames) > 0:
        ws.cell(row=2, column=4, value=submitter_username)
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        ws.column_dimensions[column].width = (max_length + 2) * 1.2
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

def generate_mother_purchase_excel(accounts):
    wb = Workbook()
    ws = wb.active
    ws.title = "Mother Accounts"
    headers = ["Username", "Password", "2FA Key"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for acc in accounts:
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
        ws.column_dimensions[column].width = (max_length + 2) * 1.2
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

# ================== ACCOUNT SUBMIT ==================
def start_submission(chat_id, acc_type):
    submission_sessions[chat_id] = {"step": "username", "type": acc_type}
    type_label = "🍪 কুকিজ একাউন্ট" if acc_type == "cookies" else "🔐 2FA একাউন্ট"
    msg = (f"📋 {type_label} সাবমিট\n\nপ্রথমে আপনার **ইউজারনেম** লিস্ট দিন (প্রতি লাইনে একটি):\n\n"
           "বাতিল করতে /cancel লিখুন।")
    send_telegram_message(msg, chat_id)

def process_submission_step(chat_id, text, sender_username):
    if chat_id not in submission_sessions:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        submission_sessions.pop(chat_id, None)
        send_telegram_message("❌ সাবমিট বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
        return True
    session = submission_sessions[chat_id]
    step = session["step"]
    if step == "username":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["usernames"] = lines
        session["step"] = "password"
        send_telegram_message(f"🔑 এখন **পাসওয়ার্ড** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nআপনার ইউজারনেম সংখ্যা: {len(lines)}", chat_id)
        return True
    elif step == "password":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        usernames = session["usernames"]
        if len(lines) != len(usernames):
            send_telegram_message(f"❌ পাসওয়ার্ড সংখ্যা ({len(lines)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। পুনরায় সঠিক লিস্ট দিন।", chat_id)
            return True
        session["passwords"] = lines
        session["step"] = "2fa"
        send_telegram_message("🔐 এখন **2FA কী** লিস্ট দিন (প্রতি লাইনে একটি, ইউজারনেম এর ক্রম অনুযায়ী):\n\nযদি 2FA না থাকে, লাইন ফাঁকা রাখবেন (শুধু এন্টার দিন)।", chat_id)
        return True
    elif step == "2fa":
        raw_lines = text.splitlines()
        twofa_list = [l.strip() for l in raw_lines]
        usernames = session["usernames"]
        while len(twofa_list) > len(usernames) and twofa_list and twofa_list[-1] == '':
            twofa_list.pop()
        if len(twofa_list) != len(usernames):
            send_telegram_message(f"❌ 2FA কী সংখ্যা ({len(twofa_list)}) ইউজারনেম সংখ্যার ({len(usernames)}) সাথে মেলে না। প্রতিটি ইউজারনেমের জন্য একটি লাইন (খালিও হতে পারে) দিন।", chat_id)
            return True
        # Submit
        acc_type = session["type"]
        type_label = "কুকিজ" if acc_type == "cookies" else "2FA"
        filename = f"{sender_username}_{chat_id}_{len(usernames)}pcs_{type_label}.xlsx"
        excel_bytes = generate_submission_excel(usernames, session["passwords"], twofa_list, sender_username)
        caption = (f"📥 {user_info.get(chat_id, 'Unknown')} (@{sender_username}) "
                   f"একটি {type_label} একাউন্ট ফাইল সাবমিট করেছেন।\nমোট {len(usernames)} টি একাউন্ট।")
        resp = send_telegram_document(excel_bytes, filename, ADMIN_CHAT_ID, caption=caption)
        if resp:
            sub_id = uuid.uuid4().hex[:10]
            submission = {
                "id": sub_id,
                "user_id": chat_id,
                "username": sender_username,
                "type": acc_type,
                "count": len(usernames),
                "file_message_id": None,  # will be set after sending
                "timestamp": time.time(),
                "status": "pending"  # pending, approved, rejected
            }
            # We cannot easily get message_id from sendDocument directly, but we can use the API response.
            # We'll fetch it via getUpdates or parse response. For simplicity, we'll store dummy and later update.
            # Actually, we'll fetch the message_id from the response if possible.
            # Since send_telegram_document returns success boolean, we'll need to modify it to return response or message_id.
            # Let's refactor: we'll use the raw API directly to get message_id.
            # Quick fix: we'll send and then retrieve last message in admin chat? Too complex. 
            # We'll store timestamp and later clean up by time.
            submissions.append(submission)
            save_all()
            send_telegram_message("✅ আপনার ফাইল সফলভাবে সাবমিট হয়েছে। অ্যাডমিন ২৪ ঘণ্টার মধ্যে রিপোর্ট দিবেন।", chat_id)
            # Update user profile counters
            update_user_submission_stats(chat_id, acc_type, len(usernames))
        else:
            send_telegram_message("⚠️ ফাইল পাঠাতে সমস্যা হয়েছে। দয়া করে পরে চেষ্টা করুন।", chat_id)
        del submission_sessions[chat_id]
        send_telegram_message("🔝", chat_id, reply_markup=get_main_keyboard(chat_id))
        return True
    return False

def update_user_submission_stats(user_id, acc_type, count):
    # We'll maintain in leaderboard dict for simplicity
    with data_lock:
        if user_id not in leaderboard:
            leaderboard[user_id] = {"total_submitted_2fa": 0, "total_submitted_cookies": 0,
                                    "total_ok_2fa": 0, "total_ok_cookies": 0, "total_income": 0.0}
        if acc_type == "2fa":
            leaderboard[user_id]["total_submitted_2fa"] += count
        else:
            leaderboard[user_id]["total_submitted_cookies"] += count

# ================== ADMIN SUBMISSION APPROVAL ==================
def admin_approve_start(sub_id):
    # Admin will be asked to enter number of OK IDs
    admin_approve_sessions[ADMIN_CHAT_ID] = {"sub_id": sub_id, "step": "ok_count"}
    send_telegram_message("✅ কতটি আইডি ওকে হয়েছে? সংখ্যা লিখুন:", ADMIN_CHAT_ID)

def process_admin_approve_step(chat_id, text):
    if chat_id != ADMIN_CHAT_ID or ADMIN_CHAT_ID not in admin_approve_sessions:
        return False
    session = admin_approve_sessions[ADMIN_CHAT_ID]
    if text.strip().lower() == "/cancel":
        del admin_approve_sessions[ADMIN_CHAT_ID]
        send_telegram_message("❌ বাতিল করা হয়েছে।", ADMIN_CHAT_ID)
        return True
    try:
        ok_count = int(text.strip())
        if ok_count < 0:
            raise ValueError
    except:
        send_telegram_message("⚠️ সঠিক সংখ্যা দিন।", ADMIN_CHAT_ID)
        return True
    sub_id = session["sub_id"]
    with data_lock:
        for sub in submissions:
            if sub["id"] == sub_id and sub["status"] == "pending":
                sub["status"] = "approved"
                sub["ok_count"] = ok_count
                user_id = sub["user_id"]
                price = config["price_2fa"] if sub["type"] == "2fa" else config["price_cookies"]
                amount = ok_count * price
                user_balances[user_id] = user_balances.get(user_id, 0) + amount
                # Update leaderboard
                if user_id not in leaderboard:
                    leaderboard[user_id] = {"total_submitted_2fa":0,"total_submitted_cookies":0,"total_ok_2fa":0,"total_ok_cookies":0,"total_income":0}
                leaderboard[user_id]["total_income"] += amount
                if sub["type"] == "2fa":
                    leaderboard[user_id]["total_ok_2fa"] += ok_count
                else:
                    leaderboard[user_id]["total_ok_cookies"] += ok_count
                # Referral bonuses
                distribute_referral_bonus(user_id, amount)
                save_all()
                send_telegram_message(f"✅ সাবমিশন {sub_id} অ্যাপ্রুভ হয়েছে। {ok_count} আইডি ওকে, {amount} টাকা যোগ করা হয়েছে।", ADMIN_CHAT_ID)
                send_telegram_message(f"🎉 আপনার {ok_count} টি আইডি ওকে হয়েছে! {amount} টাকা আপনার ব্যালেন্সে যোগ হয়েছে।", user_id)
                break
        else:
            send_telegram_message("❌ সাবমিশন পাওয়া যায়নি বা ইতিমধ্যে প্রসেস করা হয়েছে।", ADMIN_CHAT_ID)
    del admin_approve_sessions[ADMIN_CHAT_ID]
    return True

def distribute_referral_bonus(user_id, amount):
    if user_id in referrals:
        referrer = referrals[user_id]
        level1 = config["referral_level1"]
        bonus1 = amount * level1 / 100.0
        if bonus1 > 0:
            user_balances[referrer] = user_balances.get(referrer, 0) + bonus1
            referral_bonuses[referrer] = referral_bonuses.get(referrer, 0.0) + bonus1
            send_telegram_message(f"🎁 রেফারেল বোনাস: {bonus1} টাকা ({level1}%) পেয়েছেন!", referrer)
            update_leaderboard_income(referrer, bonus1)
        # Level 2
        if referrer in referrals:
            grand_referrer = referrals[referrer]
            level2 = config["referral_level2"]
            bonus2 = amount * level2 / 100.0
            if bonus2 > 0:
                user_balances[grand_referrer] = user_balances.get(grand_referrer, 0) + bonus2
                referral_bonuses[grand_referrer] = referral_bonuses.get(grand_referrer, 0.0) + bonus2
                send_telegram_message(f"🎁 রেফারেল বোনাস (লেভেল ২): {bonus2} টাকা ({level2}%) পেয়েছেন!", grand_referrer)
                update_leaderboard_income(grand_referrer, bonus2)

def update_leaderboard_income(user_id, amount):
    if user_id not in leaderboard:
        leaderboard[user_id] = {"total_submitted_2fa":0,"total_submitted_cookies":0,"total_ok_2fa":0,"total_ok_cookies":0,"total_income":0}
    leaderboard[user_id]["total_income"] += amount

# ================== MOTHER ACCOUNT (FREE) ==================
def handle_get_free_mother(chat_id):
    now = time.time()
    last = user_last_request.get(str(chat_id), 0)
    cooldown = 600
    if now - last < cooldown:
        wait = int((cooldown - (now - last))/60)
        send_telegram_message(f"⏳ {wait} মিনিট পরে ফ্রি মাদার একাউন্ট নিতে পারবেন।", chat_id)
        return
    with data_lock:
        for acc in mother_accounts:
            if not acc.get("assigned_to"):
                acc["assigned_to"] = str(chat_id)
                acc["assigned_at"] = now
                user_last_request[str(chat_id)] = now
                save_all()
                msg = f"🎁 ফ্রি মাদার একাউন্ট:\n👤 ইউজারনেম: {acc['username']}\n🔑 পাসওয়ার্ড: {acc['password']}"
                if acc.get("fa_key"):
                    msg += f"\n🔐 2FA: {acc['fa_key']}"
                send_telegram_message(msg, chat_id)
                return
    send_telegram_message("❌ এখন কোন ফ্রি মাদার একাউন্ট নেই।", chat_id)

# ================== BUY MOTHER ACCOUNT ==================
def start_buy_mother(chat_id):
    with data_lock:
        available = [m for m in mother_stock if not m.get("sold")]
    if not available:
        send_telegram_message("❌ মাদার একাউন্ট স্টক খালি।", chat_id)
        return
    price = config["mother_price"]
    send_telegram_message(f"🛒 মাদার একাউন্ট কিনুন\nপ্রতি পিস মূল্য: {price} টাকা\nআপনি কতটি কিনতে চান? (সংখ্যা লিখুন)", chat_id)
    # We'll use a simple session for quantity
    submission_sessions[chat_id] = {"step": "mother_qty", "type": "mother_buy"}  # reuse submission_sessions dict

def process_mother_buy_step(chat_id, text):
    if chat_id not in submission_sessions or submission_sessions[chat_id].get("type") != "mother_buy":
        return False
    if text.strip().lower() == "/cancel":
        del submission_sessions[chat_id]
        send_telegram_message("❌ বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
        return True
    try:
        qty = int(text.strip())
        if qty <= 0:
            raise ValueError
    except:
        send_telegram_message("⚠️ সঠিক সংখ্যা দিন।", chat_id)
        return False
    price = config["mother_price"]
    total = qty * price
    with data_lock:
        bal = user_balances.get(str(chat_id), 0)
        if bal < total:
            send_telegram_message(f"❌ পর্যাপ্ত ব্যালেন্স নেই। প্রয়োজন {total} টাকা, আপনার ব্যালেন্স {bal} টাকা।", chat_id)
            del submission_sessions[chat_id]
            return True
        available = [m for m in mother_stock if not m.get("sold")]
        if qty > len(available):
            send_telegram_message(f"❌ পর্যাপ্ত স্টক নেই। বর্তমান স্টক: {len(available)}", chat_id)
            del submission_sessions[chat_id]
            return True
        to_buy = []
        new_stock = []
        bought = 0
        for acc in mother_stock:
            if not acc.get("sold") and bought < qty:
                acc["sold"] = True
                to_buy.append(acc)
                bought += 1
            else:
                new_stock.append(acc)
        mother_stock[:] = new_stock
        user_balances[str(chat_id)] = bal - total
        save_all()
    excel = generate_mother_purchase_excel(to_buy)
    fname = f"mother_{chat_id}_{int(time.time())}.xlsx"
    if send_telegram_document(excel, fname, chat_id, caption=f"{qty} টি মাদার একাউন্ট কেনা হয়েছে। মোট মূল্য: {total} টাকা"):
        send_telegram_message(f"✅ {qty} টি মাদার একাউন্ট কেনা সফল। অবশিষ্ট ব্যালেন্স: {user_balances[str(chat_id)]} টাকা", chat_id)
    else:
        # refund
        mother_stock.extend(to_buy)
        user_balances[str(chat_id)] = bal
        save_all()
        send_telegram_message("⚠️ ডেলিভারি ব্যর্থ। টাকা ফেরত দেওয়া হয়েছে।", chat_id)
    del submission_sessions[chat_id]
    send_telegram_message("🔝", chat_id, reply_markup=get_main_keyboard(chat_id))
    return True

# ================== ADMIN ADD MOTHER STOCK ==================
def start_add_mother_stock(chat_id):
    admin_add_mother_session[chat_id] = {"step": "username"}
    send_telegram_message("➕ মাদার একাউন্ট যোগ করুন\nপ্রথমে ইউজারনেম লিস্ট দিন (প্রতি লাইনে একটি):", chat_id)

def process_add_mother_step(chat_id, text):
    if chat_id not in admin_add_mother_session:
        return False
    if text.strip().lower() == "/cancel":
        del admin_add_mother_session[chat_id]
        send_telegram_message("❌ বাতিল করা হয়েছে।", chat_id, reply_markup=admin_panel_keyboard())
        return True
    session = admin_add_mother_session[chat_id]
    step = session["step"]
    if step == "username":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        if not lines:
            send_telegram_message("⚠️ কমপক্ষে একটি ইউজারনেম দিন।", chat_id)
            return True
        session["usernames"] = lines
        session["step"] = "password"
        send_telegram_message(f"🔑 এখন পাসওয়ার্ড লিস্ট দিন ({len(lines)} টি):", chat_id)
        return True
    elif step == "password":
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        usernames = session["usernames"]
        if len(lines) != len(usernames):
            send_telegram_message(f"❌ পাসওয়ার্ড সংখ্যা ({len(lines)}) ইউজারনেম সংখ্যার সাথে মেলে না।", chat_id)
            return True
        session["passwords"] = lines
        session["step"] = "2fa"
        send_telegram_message("🔐 2FA কী লিস্ট দিন (প্রতি লাইনে, ফাঁকা রাখা যাবে):", chat_id)
        return True
    elif step == "2fa":
        raw_lines = text.splitlines()
        twofa_list = [l.strip() for l in raw_lines]
        usernames = session["usernames"]
        while len(twofa_list) > len(usernames) and twofa_list and twofa_list[-1] == '':
            twofa_list.pop()
        if len(twofa_list) != len(usernames):
            send_telegram_message(f"❌ 2FA কী সংখ্যা ({len(twofa_list)}) ইউজারনেম সংখ্যার সাথে মেলে না।", chat_id)
            return True
        with data_lock:
            for i in range(len(usernames)):
                mother_stock.append({
                    "username": usernames[i],
                    "password": session["passwords"][i],
                    "fa_key": twofa_list[i],
                    "sold": False
                })
            save_all()
        send_telegram_message(f"✅ {len(usernames)} টি মাদার একাউন্ট স্টকে যোগ করা হয়েছে।", chat_id, reply_markup=admin_panel_keyboard())
        del admin_add_mother_session[chat_id]
        return True
    return False

# ================== PROFILE & LEADERBOARD ==================
def show_profile(chat_id):
    with data_lock:
        bal = user_balances.get(chat_id, 0)
        stats = leaderboard.get(chat_id, {})
        sub2fa = stats.get("total_submitted_2fa", 0)
        subcookies = stats.get("total_submitted_cookies", 0)
        ok2fa = stats.get("total_ok_2fa", 0)
        okcookies = stats.get("total_ok_cookies", 0)
        income = stats.get("total_income", 0.0)
    msg = (f"👤 আপনার প্রোফাইল\n\n"
           f"💰 মোট ইনকাম: {income} টাকা\n"
           f"📊 ব্যালেন্স: {bal} টাকা\n\n"
           f"📤 সাবমিট:\n"
           f"  🔐 2FA: {sub2fa} টি (ওকে: {ok2fa})\n"
           f"  🍪 কুকিজ: {subcookies} টি (ওকে: {okcookies})\n"
           f"------------------\n"
           f"🎯 মাসিক টার্গেট: {config['monthly_target']} টাকা\n"
           f"🏆 বোনাস: {config['target_bonus']}%")
    send_telegram_message(msg, chat_id)

def show_leaderboard(chat_id):
    with data_lock:
        sorted_users = sorted(leaderboard.items(), key=lambda x: x[1].get("total_income", 0), reverse=True)[:10]
    if not sorted_users:
        send_telegram_message("এখনো কোনো ইনকাম রেকর্ড নেই।", chat_id)
        return
    msg = "🏆 লিডারবোর্ড (সর্বোচ্চ ইনকাম)\n\n"
    for i, (uid, data) in enumerate(sorted_users, 1):
        name = user_info.get(uid, uid)
        inc = data.get("total_income", 0)
        msg += f"{i}. {name} - {inc} টাকা\n"
    send_telegram_message(msg, chat_id)

# ================== WITHDRAW ==================
def start_withdraw(chat_id):
    withdraw_sessions[chat_id] = {"step": "amount"}
    send_telegram_message("💸 কত টাকা উইথড্র করতে চান? (শুধু সংখ্যা লিখুন)\nবাতিল করতে /cancel", chat_id)

def process_withdraw_step(chat_id, text):
    if chat_id not in withdraw_sessions:
        return False
    if text.strip().lower() in ["/cancel", "/start"]:
        del withdraw_sessions[chat_id]
        send_telegram_message("❌ উইথড্র বাতিল করা হয়েছে।", chat_id, reply_markup=get_main_keyboard(chat_id))
        return True
    session = withdraw_sessions[chat_id]
    step = session["step"]
    if step == "amount":
        try:
            amount = float(text.strip())
            if amount <= 0:
                raise ValueError
        except:
            send_telegram_message("⚠️ সঠিক সংখ্যা দিন।", chat_id)
            return True
        bal = user_balances.get(chat_id, 0)
        if amount > bal:
            send_telegram_message(f"❌ অপর্যাপ্ত ব্যালেন্স। আপনার ব্যালেন্স: {bal} টাকা", chat_id)
            del withdraw_sessions[chat_id]
            return True
        session["amount"] = amount
        session["step"] = "bkash"
        send_telegram_message("📞 আপনার বিকাশ নম্বর দিন:", chat_id)
        return True
    elif step == "bkash":
        bkash = text.strip()
        if not bkash:
            send_telegram_message("⚠️ বিকাশ নম্বর খালি রাখা যাবে না।", chat_id)
            return True
        amount = session["amount"]
        w_id = uuid.uuid4().hex[:10]
        w_req = {"id": w_id, "user_id": chat_id, "amount": amount, "bkash": bkash, "status": "pending", "time": time.time()}
        # We'll store withdraw requests in a list (maybe in config file)
        # For simplicity, we'll keep a global list inside config or separate file
        # We'll add a withdraw_requests list
        global withdraw_requests
        if 'withdraw_requests' not in globals():
            withdraw_requests = []
        withdraw_requests.append(w_req)
        save_all()
        del withdraw_sessions[chat_id]
        send_telegram_message(f"✅ {amount} টাকা উইথড্র রিকোয়েস্ট জমা হয়েছে। অ্যাডমিন শীঘ্রই প্রসেস করবেন।", chat_id)
        admin_msg = f"💳 নতুন উইথড্র রিকোয়েস্ট\nআইডি: {w_id}\nইউজার: {user_info.get(chat_id, chat_id)} ({chat_id})\nপরিমাণ: {amount}\nবিকাশ: {bkash}\n/anprovewithdraw {w_id} or /rejectwithdraw {w_id}"
        send_telegram_message(admin_msg, ADMIN_CHAT_ID)
        send_telegram_message("🔝", chat_id, reply_markup=get_main_keyboard(chat_id))
        return True
    return False

# ================== ADMIN BROADCAST ==================
def admin_broadcast_prompt(chat_id):
    send_telegram_message("📢 কী ধরনের ব্রডকাস্ট করবেন?", chat_id, reply_markup={
        "inline_keyboard": [
            [{"text": "📝 টেক্সট", "callback_data": "bc_text"}],
            [{"text": "🖼️ ছবি", "callback_data": "bc_photo"}],
            [{"text": "📄 ফাইল", "callback_data": "bc_file"}],
            [{"text": "🎤 ভয়েস", "callback_data": "bc_voice"}]
        ]
    })

# (We'll implement a generic forward approach: admin sends a message to bot which will be forwarded)
# Actually simpler: admin uses /broadcast command. We'll handle commands.

# ================== MAIN HANDLER ==================
def handle_telegram_commands():
    global last_update_id
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates"
    while True:
        try:
            params = {"timeout": 30, "allowed_updates": ["message", "callback_query"]}
            if last_update_id:
                params["offset"] = last_update_id + 1
            resp = requests.get(url, params=params, timeout=35).json()
            if resp.get("ok") and resp.get("result"):
                for update in resp["result"]:
                    last_update_id = update["update_id"]
                    if "callback_query" in update:
                        cb = update["callback_query"]
                        cbid = cb["id"]
                        chat_id = str(cb["message"]["chat"]["id"])
                        data = cb["data"]
                        from_user = cb.get("from", {})
                        sender_username = from_user.get("username") or from_user.get("first_name", f"ID:{chat_id}")
                        user_info[chat_id] = sender_username
                        answer_callback_query(cbid)
                        # Admin panel callbacks
                        if data == "submissions_list":
                            # show pending submissions inline
                            pass
                        continue

                    if "message" in update:
                        msg = update["message"]
                        chat_id = str(msg["chat"]["id"])
                        text = msg.get("text", "").strip()
                        from_user = msg.get("from", {})
                        sender_username = from_user.get("username") or from_user.get("first_name", f"ID:{chat_id}")
                        user_info[chat_id] = sender_username

                        # Admin approval session
                        if chat_id == ADMIN_CHAT_ID and ADMIN_CHAT_ID in admin_approve_sessions:
                            process_admin_approve_step(chat_id, text)
                            continue

                        # Submission session
                        if chat_id in submission_sessions:
                            if submission_sessions[chat_id].get("type") == "mother_buy":
                                process_mother_buy_step(chat_id, text)
                            else:
                                process_submission_step(chat_id, text, sender_username)
                            continue

                        # Add mother stock session
                        if chat_id in admin_add_mother_session:
                            process_add_mother_step(chat_id, text)
                            continue

                        # Withdraw session
                        if chat_id in withdraw_sessions:
                            process_withdraw_step(chat_id, text)
                            continue

                        # Support session (if we add)
                        # ...

                        # Button handlers
                        if text == "💼 একাউন্ট সাবমিট":
                            kb = {
                                "inline_keyboard": [
                                    [{"text": "🍪 কুকিজ একাউন্ট", "callback_data": "sub_cookies"}],
                                    [{"text": "🔐 2FA একাউন্ট", "callback_data": "sub_2fa"}]
                                ]
                            }
                            send_telegram_message("কোন ধরণের একাউন্ট সাবমিট করবেন?", chat_id, reply_markup=kb)
                            continue
                        elif text == "👤 প্রোফাইল":
                            show_profile(chat_id)
                            continue
                        elif text == "👥 রেফারেল":
                            link = f"https://t.me/{BOT_USERNAME}?start=ref_{chat_id}"
                            send_telegram_message(f"🔗 আপনার রেফারেল লিংক:\n{link}\n\nশেয়ার করে ৫% বোনাস পান!", chat_id)
                            continue
                        elif text == "💰 ব্যালেন্স":
                            bal = user_balances.get(chat_id, 0)
                            send_telegram_message(f"💰 আপনার ব্যালেন্স: {bal} টাকা", chat_id)
                            continue
                        elif text == "💸 উইথড্র":
                            start_withdraw(chat_id)
                            continue
                        elif text == "📊 লিডারবোর্ড":
                            show_leaderboard(chat_id)
                            continue
                        elif text == "🎁 ফ্রি মাদার একাউন্ট":
                            handle_get_free_mother(chat_id)
                            continue
                        elif text == "🛒 মাদার একাউন্ট কিনুন":
                            start_buy_mother(chat_id)
                            continue
                        elif text == "🛠️ অ্যাডমিন প্যানেল" and str(chat_id) == ADMIN_CHAT_ID:
                            send_telegram_message("অ্যাডমিন প্যানেল", chat_id, reply_markup=admin_panel_keyboard())
                            continue
                        elif text == "📊 সাবমিটেড ফাইল" and str(chat_id) == ADMIN_CHAT_ID:
                            # list pending submissions
                            pending = [s for s in submissions if s["status"] == "pending"]
                            if not pending:
                                send_telegram_message("কোনো পেন্ডিং সাবমিশন নেই।", chat_id)
                            else:
                                msg_lines = ["📥 পেন্ডিং সাবমিশন:\n"]
                                for s in pending:
                                    type_label = "কুকিজ" if s["type"] == "cookies" else "2FA"
                                    msg_lines.append(f"আইডি: {s['id']} | {type_label} | {s['username']} | {s['count']} পিস\n/approve_{s['id']}")
                                send_telegram_message("\n".join(msg_lines), chat_id)
                            continue
                        elif text == "⚙️ মূল্য নির্ধারণ" and str(chat_id) == ADMIN_CHAT_ID:
                            # inline for setting prices
                            send_telegram_message("মূল্য নির্ধারণ করতে কমান্ড ব্যবহার করুন:\n/setprice 2fa <মূল্য>\n/setprice cookies <মূল্য>", chat_id)
                            continue
                        elif text == "👥 রেফারেল বোনাস %" and str(chat_id) == ADMIN_CHAT_ID:
                            send_telegram_message("/setreferral level1 <শতাংশ> বা level2 <শতাংশ>", chat_id)
                            continue
                        elif text == "💵 মাসিক টার্গেট" and str(chat_id) == ADMIN_CHAT_ID:
                            send_telegram_message("/settarget <টাকা>", chat_id)
                            continue
                        elif text == "🔒 সাবমিট লক" and str(chat_id) == ADMIN_CHAT_ID:
                            kb = {"inline_keyboard": [
                                [{"text": f"2FA {'🔒' if config['lock_2fa'] else '🔓'}", "callback_data": "lock_2fa"}],
                                [{"text": f"কুকিজ {'🔒' if config['lock_cookies'] else '🔓'}", "callback_data": "lock_cookies"}]
                            ]}
                            send_telegram_message("কোন ক্যাটাগরি লক/আনলক করবেন?", chat_id, reply_markup=kb)
                            continue
                        elif text == "📢 ব্রডকাস্ট" and str(chat_id) == ADMIN_CHAT_ID:
                            admin_broadcast_prompt(chat_id)
                            continue
                        elif text == "➕ মাদার একাউন্ট যোগ" and str(chat_id) == ADMIN_CHAT_ID:
                            start_add_mother_stock(chat_id)
                            continue
                        elif text == "📦 মাদার স্টক" and str(chat_id) == ADMIN_CHAT_ID:
                            with data_lock:
                                count = len([m for m in mother_stock if not m.get("sold")])
                            send_telegram_message(f"📦 মাদার একাউন্ট স্টকে আছে: {count} টি", chat_id)
                            continue
                        elif text == "💰 মাদার মূল্য সেট" and str(chat_id) == ADMIN_CHAT_ID:
                            send_telegram_message(f"/setmotherprice <মূল্য>", chat_id)
                            continue
                        elif text == "📋 ইউজার লিস্ট" and str(chat_id) == ADMIN_CHAT_ID:
                            users = "\n".join([f"{uid} - {user_info.get(uid, '?')}" for uid in subscribed_users])
                            send_telegram_message(f"সাবস্ক্রাইবড ইউজার:\n{users}", chat_id)
                            continue
                        elif text == "🔙 মূল মেনু":
                            send_telegram_message("মূল মেনু", chat_id, reply_markup=get_main_keyboard(chat_id))
                            continue

                        # Commands
                        if text.startswith("/"):
                            handle_commands(chat_id, text, sender_username)
                            continue
        except Exception as e:
            logger.exception("Main loop error:")
        time.sleep(1)

def handle_commands(chat_id, text, sender_username):
    parts = text.split()
    cmd = parts[0].lower()
    if cmd == "/start":
        with data_lock:
            subscribed_users.add(chat_id)
            save_all()
        # Referral handling
        if len(parts) > 1 and parts[1].startswith("ref_"):
            ref_id = parts[1][4:]
            if ref_id.isdigit() and ref_id != chat_id and ref_id not in referrals:
                referrals[chat_id] = ref_id
                save_all()
                send_telegram_message(f"🎉 আপনি {user_info.get(ref_id, ref_id)}-এর রেফারেলে যুক্ত হয়েছেন!", chat_id)
        send_telegram_message("✨ স্বাগতম! নিচের বাটন ব্যবহার করুন।", chat_id, reply_markup=get_main_keyboard(chat_id))
    elif cmd == "/setprice" and str(chat_id) == ADMIN_CHAT_ID:
        if len(parts) != 3:
            send_telegram_message("/setprice 2fa <price> or /setprice cookies <price>", chat_id); return
        try:
            price = float(parts[2])
            if price <= 0: raise ValueError
        except:
            send_telegram_message("সঠিক মূল্য দিন।", chat_id); return
        if parts[1] == "2fa":
            config["price_2fa"] = price
        elif parts[1] == "cookies":
            config["price_cookies"] = price
        else:
            send_telegram_message("টাইপ: 2fa বা cookies", chat_id); return
        save_all()
        send_telegram_message(f"✅ {parts[1]} মূল্য {price} টাকা নির্ধারণ করা হয়েছে।", chat_id)
    elif cmd == "/setmotherprice" and str(chat_id) == ADMIN_CHAT_ID:
        if len(parts) < 2: send_telegram_message("/setmotherprice <price>", chat_id); return
        try:
            price = float(parts[1])
        except:
            send_telegram_message("সঠিক সংখ্যা দিন।", chat_id); return
        config["mother_price"] = price
        save_all()
        send_telegram_message(f"✅ মাদার একাউন্টের মূল্য {price} টাকা সেট হয়েছে।", chat_id)
    elif cmd == "/setreferral" and str(chat_id) == ADMIN_CHAT_ID:
        if len(parts) != 3:
            send_telegram_message("/setreferral level1 <percent> or level2 <percent>", chat_id); return
        try:
            perc = float(parts[2])
        except:
            send_telegram_message("সঠিক শতাংশ দিন।", chat_id); return
        if parts[1] == "level1":
            config["referral_level1"] = perc
        elif parts[1] == "level2":
            config["referral_level2"] = perc
        else:
            send_telegram_message("level1 or level2", chat_id); return
        save_all()
        send_telegram_message(f"✅ রেফারেল {parts[1]} বোনাস {perc}% সেট হয়েছে।", chat_id)
    elif cmd == "/settarget" and str(chat_id) == ADMIN_CHAT_ID:
        if len(parts) < 2: send_telegram_message("/settarget <amount>", chat_id); return
        try:
            target = float(parts[1])
        except:
            send_telegram_message("সঠিক টাকার পরিমাণ দিন।", chat_id); return
        config["monthly_target"] = target
        save_all()
        send_telegram_message(f"✅ মাসিক টার্গেট {target} টাকা নির্ধারণ করা হয়েছে।", chat_id)
    elif cmd.startswith("/approve_") and str(chat_id) == ADMIN_CHAT_ID:
        sub_id = cmd[9:]
        admin_approve_start(sub_id)
    elif cmd == "/approvewithdraw" and str(chat_id) == ADMIN_CHAT_ID:
        # implement approve withdraw
        pass
    elif cmd == "/rejectwithdraw" and str(chat_id) == ADMIN_CHAT_ID:
        pass
    # etc.

# ================== FLASK ==================
@app.route("/")
def home():
    return "Bot Running!"

# ================== MAIN ==================
if __name__ == "__main__":
    load_all()
    # Start polling
    threading.Thread(target=handle_telegram_commands, daemon=True).start()
    # Daily tasks (reminders, cleanup)
    def daily_jobs():
        while True:
            time.sleep(86400)
            # Delete old submissions (2 days)
            now = time.time()
            with data_lock:
                submissions[:] = [s for s in submissions if now - s["timestamp"] < 172800]
                save_all()
    threading.Thread(target=daily_jobs, daemon=True).start()
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
