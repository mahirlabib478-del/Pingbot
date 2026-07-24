import os
import json
import time
import threading
import uuid
import logging
import gzip
import io
from flask import Flask
import requests
from datetime import datetime

# ========== কনফিগারেশন ==========
BOT_TOKEN = "8757980176:AAHnpHA3m67Oz3jGdBqVFAB7GgubcD47pWM"
ADMIN_ID = "2035024902"   # এখানে ADMIN_ID ব্যবহার করুন
CHANNEL_ID = "-1003903695158"   # ব্যাকআপের জন্য
BACKUP_CHANNEL_ID = CHANNEL_ID

if not BOT_TOKEN or not ADMIN_ID:
    raise RuntimeError("BOT_TOKEN and ADMIN_ID must be set")
    
# ========== ফাইল পাথ ==========
DATA_FILE = "data.json"

# ========== ডিফল্ট ডেটা ==========
default_data = {
    "users": {},
    "mother_accounts": [],
    "deposits": [],
    "config": {
        "free_limit": 5,
        "price_per_account": 5.0,
        "bkash_number": "01XXXXXXXXX",
        "nagad_number": "01XXXXXXXXX",
        "backup_channel_id": BACKUP_CHANNEL_ID  # সংরক্ষণ করা হবে
    },
    "backup_meta": {
        "last_message_id": None,
        "part_ids": []
    }
}

# ========== ডেটা লোড/সেভ ==========
def load_data():
    try:
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    except:
        return default_data.copy()

def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)

data = load_data()
data_lock = threading.Lock()

# ========== ব্যাকআপ গ্লোবাল ==========
last_backup_message_id = data.get("backup_meta", {}).get("last_message_id")
last_backup_part_ids = data.get("backup_meta", {}).get("part_ids", [])

# ========== টেলিগ্রাম হেলপার ==========
def send_msg(chat_id, text, reply_markup=None, parse_mode=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {"chat_id": chat_id, "text": text}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    if parse_mode:
        payload["parse_mode"] = parse_mode
    try:
        requests.post(url, json=payload, timeout=10)
    except Exception as e:
        logging.error(f"Send error: {e}")

def delete_msg(chat_id, msg_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/deleteMessage"
    try:
        requests.post(url, json={"chat_id": chat_id, "message_id": msg_id}, timeout=5)
    except:
        pass

def answer_callback(callback_id, text=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/answerCallbackQuery"
    payload = {"callback_query_id": callback_id}
    if text:
        payload["text"] = text
    try:
        requests.post(url, json=payload, timeout=5)
    except:
        pass

def send_document(chat_id, file_bytes, filename, caption=""):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    files = {'document': (filename, file_bytes, 'application/gzip')}
    try:
        resp = requests.post(url, data={"chat_id": chat_id, "caption": caption}, files=files, timeout=60)
        if resp.status_code == 200 and resp.json().get("ok"):
            return resp.json()["result"]
        return None
    except Exception as e:
        logging.error(f"Document send error: {e}")
        return None

def pin_message(chat_id, message_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/pinChatMessage"
    try:
        requests.post(url, json={"chat_id": chat_id, "message_id": message_id, "disable_notification": True}, timeout=10)
    except:
        pass

def unpin_message(chat_id, message_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/unpinChatMessage"
    try:
        requests.post(url, json={"chat_id": chat_id, "message_id": message_id}, timeout=10)
    except:
        pass

# ========== ব্যাকআপ ফাংশন ==========
MAX_PART_SIZE = 45 * 1024 * 1024  # 45MB (নিরাপদ)

def cleanup_old_backup(chat_id):
    global last_backup_message_id, last_backup_part_ids
    try:
        if last_backup_message_id:
            unpin_message(chat_id, last_backup_message_id)
            delete_msg(chat_id, last_backup_message_id)
        for pid in last_backup_part_ids:
            delete_msg(chat_id, pid)
        last_backup_message_id = None
        last_backup_part_ids = []
        # মেটাডেটাও আপডেট করুন
        with data_lock:
            data["backup_meta"]["last_message_id"] = None
            data["backup_meta"]["part_ids"] = []
            save_data(data)
    except Exception as e:
        logging.error(f"Cleanup backup error: {e}")

def save_data_to_channel():
    global last_backup_message_id, last_backup_part_ids
    channel_id = data["config"].get("backup_channel_id")
    if not channel_id:
        logging.warning("Backup channel ID not set. Skipping backup.")
        return

    with data_lock:
        # সম্পূর্ণ ডেটা নিন (কপি)
        backup_data = data.copy()
        # timestamp যোগ করুন
        backup_data["timestamp"] = datetime.now().isoformat()

    json_bytes = json.dumps(backup_data, indent=2, ensure_ascii=False).encode('utf-8')
    compressed = gzip.compress(json_bytes, compresslevel=6)

    # পুরনো ব্যাকআপ মুছুন
    cleanup_old_backup(channel_id)

    new_backup_ids = []
    new_part_ids = []

    if len(compressed) <= MAX_PART_SIZE:
        filename = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json.gz"
        result = send_document(channel_id, compressed, filename, caption="📦 ডেটা ব্যাকআপ")
        if result:
            new_backup_ids = [result["message_id"]]
    else:
        # বড় ফাইল – অংশে বিভক্ত
        chunks = [compressed[i:i+MAX_PART_SIZE] for i in range(0, len(compressed), MAX_PART_SIZE)]
        total = len(chunks)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        part_ids = []
        for idx, chunk in enumerate(chunks, 1):
            part_filename = f"backup_{timestamp}_part{idx}of{total}.json.gz"
            result = send_document(channel_id, chunk, part_filename, caption=f"Part {idx}/{total}")
            if result:
                part_ids.append(result["message_id"])
            else:
                logging.error(f"Failed to send backup part {idx}/{total}")
                return
        # সূচী (index) মেসেজ তৈরি করুন
        index_data = {
            "backup_id": timestamp,
            "parts": part_ids,
            "total_parts": total,
            "timestamp": timestamp
        }
        index_text = json.dumps(index_data)
        resp = send_msg(channel_id, index_text)
        if resp and resp.status_code == 200 and resp.json().get("ok"):
            index_msg_id = resp.json()["result"]["message_id"]
            new_backup_ids = [index_msg_id]
            new_part_ids = part_ids
        else:
            logging.error("Failed to send backup index")
            return

    # নতুন পিন করুন
    if new_backup_ids:
        pin_message(channel_id, new_backup_ids[0])
        last_backup_message_id = new_backup_ids[0]
        last_backup_part_ids = new_part_ids
        # মেটাডেটা সংরক্ষণ
        with data_lock:
            data["backup_meta"]["last_message_id"] = last_backup_message_id
            data["backup_meta"]["part_ids"] = last_backup_part_ids
            save_data(data)
        logging.info("Backup completed and pinned.")
    else:
        logging.error("Backup failed – no message ID.")

def auto_restore_from_channel():
    global last_backup_message_id, last_backup_part_ids
    channel_id = data["config"].get("backup_channel_id")
    if not channel_id:
        logging.warning("Backup channel ID not set. Cannot restore.")
        return

    try:
        # পিন করা মেসেজ খুঁজে বের করি
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/getChat"
        resp = requests.get(url, params={"chat_id": channel_id}, timeout=20).json()
        if not resp.get("ok"):
            logging.error("Cannot get chat info")
            return
        pinned = resp["result"].get("pinned_message")
        if not pinned:
            logging.info("No pinned backup message found.")
            return

        # ডাউনলোড করুন
        if "document" in pinned:
            file_id = pinned["document"]["file_id"]
            file_info = requests.get(f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={file_id}").json()
            if not file_info.get("ok"):
                logging.error("Cannot get file info")
                return
            file_path = file_info["result"]["file_path"]
            content = requests.get(f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}", timeout=60).content
            compressed = content
            last_backup_part_ids = []
        elif "text" in pinned:
            index = json.loads(pinned["text"])
            part_ids = index.get("parts", [])
            if not part_ids:
                logging.error("Invalid backup index")
                return
            combined = bytearray()
            for part_msg_id in part_ids:
                msg_resp = requests.get(
                    f"https://api.telegram.org/bot{BOT_TOKEN}/getMessage?chat_id={channel_id}&message_id={part_msg_id}",
                    timeout=20
                ).json()
                if not msg_resp.get("ok") or "document" not in msg_resp.get("result", {}):
                    logging.error(f"Missing part message {part_msg_id}")
                    return
                file_id = msg_resp["result"]["document"]["file_id"]
                file_info = requests.get(
                    f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={file_id}",
                    timeout=20
                ).json()
                if not file_info.get("ok"):
                    logging.error("Cannot get part file info")
                    return
                file_path = file_info["result"]["file_path"]
                part_content = requests.get(
                    f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}",
                    timeout=60
                ).content
                combined.extend(part_content)
            compressed = bytes(combined)
            last_backup_part_ids = part_ids
        else:
            logging.error("Pinned message has no document or text.")
            return

        decompressed = gzip.decompress(compressed)
        restored_data = json.loads(decompressed.decode('utf-8'))

        # রিস্টোর করুন (সম্পূর্ণ ডেটা প্রতিস্থাপন)
        with data_lock:
            # config ও backup_meta বাদে বাকি সব আপডেট করুন
            for key in restored_data:
                if key not in ["config", "backup_meta"]:
                    data[key] = restored_data[key]
            # কিন্তু config এর কিছু অংশ (যেমন channel_id) আমরা আপডেট করতে চাই? আমরা পুরো config না এনে শুধু প্রাসঙ্গিক অংশ নেব
            # তবে আমরা চাইলে পুরো config ও আপডেট করতে পারি, কিন্তু ব্যাকআপে config থাকলে সেটা পুরনো হতে পারে
            # আমরা শুধু user, mother_accounts, deposits আপডেট করছি
            # তবে config ও backup_meta আমরা রাখবো বর্তমান
            # সুতরাং আমরা শুধু users, mother_accounts, deposits নেব
            # কিন্তু আমরা যদি পুরো ডেটা রিস্টোর করতে চাই, তাহলে সবই নিতে পারি
            # কিন্তু backup_meta ও current config রেখে দিই
            # তাই:
            data["users"] = restored_data.get("users", {})
            data["mother_accounts"] = restored_data.get("mother_accounts", [])
            data["deposits"] = restored_data.get("deposits", [])
            # config সংরক্ষণ থাকবে পুরনো
            save_data(data)

        last_backup_message_id = pinned["message_id"]
        logging.info("Data restored from channel backup successfully")
        send_msg(ADMIN_ID, "✅ ডেটা রিস্টোর সফল হয়েছে (চ্যানেল থেকে)।")

    except Exception as e:
        logging.error(f"Auto-restore error: {e}")
        send_msg(ADMIN_ID, f"❌ রিস্টোর ব্যর্থ: {e}")

def auto_backup_loop():
    while True:
        time.sleep(86400)  # ২৪ ঘন্টা
        save_data_to_channel()

# ========== কীবোর্ড (পূর্বের মতো) ==========
def main_keyboard(user_id):
    kb = [
        ["📦 ফ্রি মাদার", "🛒 মাদার কিনুন"],
        ["💰 ব্যালেন্স", "👤 প্রোফাইল"],
        ["💳 ডিপোজিট", "📞 সাপোর্ট"]
    ]
    if str(user_id) == ADMIN_ID:
        kb.append(["🛠️ অ্যাডমিন প্যানেল"])
    return {"keyboard": kb, "resize_keyboard": True}

def admin_keyboard():
    return {
        "keyboard": [
            ["📊 ইউজার লিস্ট", "⚙️ সেটিংস"],
            ["➕ মাদার যোগ", "🗑️ মাদার মুছুন"],
            ["📥 ডিপোজিট রিকোয়েস্ট", "📢 ব্রডকাস্ট"],
            ["📁 ব্যাকআপ করুন", "📥 রিস্টোর করুন"],
            ["🔙 মূল মেনু"]
        ],
        "resize_keyboard": True
    }

# ========== ইউজার ফাংশন (পূর্বের মতো) ==========
def get_user(user_id):
    with data_lock:
        if str(user_id) not in data["users"]:
            data["users"][str(user_id)] = {
                "balance": 0.0,
                "free_taken": 0,
                "bought": 0,
                "username": "Unknown"
            }
            save_data(data)
        return data["users"][str(user_id)]

def update_user(user_id, key, value):
    with data_lock:
        data["users"][str(user_id)][key] = value
        save_data(data)

def get_free_limit():
    return data["config"].get("free_limit", 5)

def get_price():
    return data["config"].get("price_per_account", 5.0)

def get_free_mother(user_id):
    with data_lock:
        user = get_user(user_id)
        free_taken = user["free_taken"]
        limit = get_free_limit()
        if free_taken >= limit:
            return None, "আপনি ইতিমধ্যে আপনার ফ্রি কোটা পূর্ণ করেছেন।"
        for acc in data["mother_accounts"]:
            if not acc.get("sold", False):
                acc["sold"] = True
                user["free_taken"] += 1
                save_data(data)
                return acc, None
        return None, "বর্তমানে কোনো মাদার একাউন্ট স্টকে নেই।"

def buy_mother(user_id, quantity):
    with data_lock:
        user = get_user(user_id)
        price = get_price()
        total = quantity * price
        if user["balance"] < total:
            return False, "অপর্যাপ্ত ব্যালেন্স।"
        available = [a for a in data["mother_accounts"] if not a.get("sold", False)]
        if len(available) < quantity:
            return False, f"শুধু {len(available)} টি একাউন্ট উপলব্ধ।"
        bought = []
        for acc in available[:quantity]:
            acc["sold"] = True
            bought.append(acc)
        user["balance"] -= total
        user["bought"] += quantity
        save_data(data)
        return True, bought

def create_deposit(user_id, amount, method, trxid):
    dep = {
        "id": uuid.uuid4().hex[:10],
        "user_id": str(user_id),
        "amount": amount,
        "method": method,
        "trxid": trxid,
        "status": "pending",
        "time": time.time()
    }
    with data_lock:
        data["deposits"].append(dep)
        save_data(data)
    return dep["id"]

# ========== সেশন ও হ্যান্ডলার (পূর্বের মতো) ==========
user_sessions = {}

def handle_start(chat_id):
    send_msg(chat_id, "👋 স্বাগতম! নিচের মেনু থেকে অপশন বেছে নিন।", reply_markup=main_keyboard(chat_id))

def handle_profile(chat_id):
    user = get_user(chat_id)
    free_limit = get_free_limit()
    msg = (
        f"👤 **প্রোফাইল**\n\n"
        f"💰 ব্যালেন্স: {user['balance']} টাকা\n"
        f"📦 ফ্রি নেওয়া: {user['free_taken']}/{free_limit}\n"
        f"🛒 কেনা: {user['bought']} টি\n"
        f"📊 মোট একাউন্ট: {user['free_taken'] + user['bought']} টি"
    )
    send_msg(chat_id, msg, parse_mode="Markdown")

def handle_balance(chat_id):
    user = get_user(chat_id)
    send_msg(chat_id, f"💰 আপনার ব্যালেন্স: {user['balance']} টাকা")

def handle_free_mother(chat_id):
    acc, err = get_free_mother(chat_id)
    if err:
        send_msg(chat_id, f"❌ {err}")
        return
    msg = f"🎁 **ফ্রি মাদার একাউন্ট**\n\n👤 ইউজারনেম: `{acc['username']}`\n🔑 পাসওয়ার্ড: `{acc['password']}`"
    if acc.get("2fa"):
        msg += f"\n🔐 2FA: `{acc['2fa']}`"
    send_msg(chat_id, msg, parse_mode="Markdown")

def handle_buy_start(chat_id):
    user_sessions[chat_id] = {"step": "buy_qty"}
    send_msg(chat_id, f"🛒 প্রতি একাউন্টের দাম: {get_price()} টাকা\nকতটি কিনতে চান? (সংখ্যা লিখুন)\nবাতিল করতে /cancel লিখুন।")

def handle_buy_qty(chat_id, text):
    try:
        qty = int(text.strip())
        if qty <= 0:
            raise ValueError
    except:
        send_msg(chat_id, "❌ সঠিক সংখ্যা দিন।")
        return
    success, result = buy_mother(chat_id, qty)
    if not success:
        send_msg(chat_id, f"❌ {result}")
    else:
        msg = f"✅ আপনি {qty} টি একাউন্ট কিনেছেন।\n\n"
        for acc in result:
            msg += f"👤 {acc['username']} | 🔑 {acc['password']}"
            if acc.get("2fa"):
                msg += f" | 2FA: {acc['2fa']}"
            msg += "\n"
        send_msg(chat_id, msg)
    user_sessions.pop(chat_id, None)

def handle_deposit_start(chat_id):
    kb = {
        "inline_keyboard": [
            [{"text": "💸 বিকাশ", "callback_data": "dep_bkash"}],
            [{"text": "💸 নগদ", "callback_data": "dep_nagad"}],
            [{"text": "❌ বাতিল", "callback_data": "dep_cancel"}]
        ]
    }
    send_msg(chat_id, "ডিপোজিটের মাধ্যম বাছাই করুন:", reply_markup=kb)

def handle_deposit_method(chat_id, method):
    user_sessions[chat_id] = {"step": "dep_amount", "method": method}
    send_msg(chat_id, f"কত টাকা ডিপোজিট করতে চান? (শুধু সংখ্যা)\nবাতিল করতে /cancel")

def handle_deposit_amount(chat_id, text):
    try:
        amount = float(text.strip())
        if amount <= 0: raise ValueError
    except:
        send_msg(chat_id, "❌ সঠিক সংখ্যা দিন।")
        return
    session = user_sessions.get(chat_id)
    if not session:
        return
    session["amount"] = amount
    session["step"] = "dep_trxid"
    method = session["method"]
    number = data["config"].get(f"{method}_number", "সেট করা নেই")
    send_msg(chat_id, f"আপনার {method.upper()} নম্বর থেকে **{number}** নম্বরে {amount} টাকা পাঠানোর পর ট্রানজেকশন আইডি (TrxID) দিন:\nবাতিল করতে /cancel")

def handle_deposit_trxid(chat_id, text):
    trxid = text.strip()
    if not trxid:
        send_msg(chat_id, "❌ TrxID খালি রাখা যাবে না।")
        return
    session = user_sessions.pop(chat_id, None)
    if not session:
        return
    dep_id = create_deposit(chat_id, session["amount"], session["method"], trxid)
    send_msg(chat_id, f"✅ আপনার ডিপোজিট রিকোয়েস্ট জমা হয়েছে। আইডি: {dep_id}\nঅ্যাডমিন অনুমোদন দিলে ব্যালেন্স যোগ হবে।")
    admin_msg = (
        f"📥 নতুন ডিপোজিট\nআইডি: {dep_id}\nইউজার: {chat_id}\nপরিমাণ: {session['amount']} টাকা\nমাধ্যম: {session['method'].upper()}\nTrxID: {trxid}\n"
        f"অনুমোদন: /approvedeposit {dep_id}\nবাতিল: /rejectdeposit {dep_id}"
    )
    send_msg(ADMIN_ID, admin_msg)

def handle_support(chat_id):
    send_msg(chat_id, "📞 আপনার মেসেজ লিখুন। অ্যাডমিন দেখতে পাবেন।\nবাতিল করতে /cancel")

def forward_support(chat_id, text):
    send_msg(ADMIN_ID, f"📩 সাপোর্ট মেসেজ\nইউজার: {chat_id}\n\n{text}")
    send_msg(chat_id, "✅ আপনার মেসেজ পাঠানো হয়েছে।")

# ========== অ্যাডমিন কমান্ড (ব্যাকআপ যোগ) ==========
def admin_panel(chat_id):
    send_msg(chat_id, "🛠️ অ্যাডমিন প্যানেল", reply_markup=admin_keyboard())

def admin_user_list(chat_id):
    with data_lock:
        users = data["users"]
    if not users:
        send_msg(chat_id, "কোনো ইউজার নেই।")
        return
    msg = "📊 **ইউজার লিস্ট**\n\n"
    for uid, info in users.items():
        msg += f"• {info.get('username', uid)} ({uid}) - ব্যালেন্স: {info['balance']} টাকা, ফ্রি: {info['free_taken']}, কেনা: {info['bought']}\n"
        if len(msg) > 4000:
            send_msg(chat_id, msg)
            msg = ""
    if msg:
        send_msg(chat_id, msg, parse_mode="Markdown")

def admin_settings(chat_id):
    cfg = data["config"]
    msg = (
        f"⚙️ **বর্তমান সেটিংস**\n"
        f"ফ্রি লিমিট: {cfg.get('free_limit', 5)}\n"
        f"প্রতি একাউন্ট দাম: {cfg.get('price_per_account', 5.0)} টাকা\n"
        f"বিকাশ নম্বর: {cfg.get('bkash_number', 'সেট নেই')}\n"
        f"নগদ নম্বর: {cfg.get('nagad_number', 'সেট নেই')}\n"
        f"ব্যাকআপ চ্যানেল: {cfg.get('backup_channel_id', 'সেট নেই')}\n\n"
        f"পরিবর্তন করতে কমান্ড:\n"
        f"/setfreelimit <সংখ্যা>\n"
        f"/setprice <দাম>\n"
        f"/setbkash <নম্বর>\n"
        f"/setnagad <নম্বর>\n"
        f"/setbackupchannel <চ্যানেল_আইডি>"
    )
    send_msg(chat_id, msg, parse_mode="Markdown")

def admin_add_mother_start(chat_id):
    user_sessions[chat_id] = {"step": "add_username"}
    send_msg(chat_id, "➕ নতুন মাদার একাউন্ট যোগ করুন\nপ্রথমে **ইউজারনেম** দিন:\nবাতিল করতে /cancel")

def admin_add_mother_username(chat_id, text):
    session = user_sessions.get(chat_id)
    if not session:
        return
    session["username"] = text.strip()
    session["step"] = "add_password"
    send_msg(chat_id, "🔑 এখন **পাসওয়ার্ড** দিন:")

def admin_add_mother_password(chat_id, text):
    session = user_sessions.get(chat_id)
    if not session:
        return
    session["password"] = text.strip()
    session["step"] = "add_2fa"
    send_msg(chat_id, "🔐 2FA কী দিন (যদি না থাকে, '0' লিখুন):")

def admin_add_mother_2fa(chat_id, text):
    session = user_sessions.pop(chat_id, None)
    if not session:
        return
    twofa = text.strip()
    if twofa == "0":
        twofa = ""
    with data_lock:
        data["mother_accounts"].append({
            "username": session["username"],
            "password": session["password"],
            "2fa": twofa,
            "sold": False
        })
        save_data(data)
    send_msg(chat_id, f"✅ একাউন্ট `{session['username']}` যোগ করা হয়েছে।", reply_markup=admin_keyboard())

def admin_delete_mother_start(chat_id):
    with data_lock:
        accs = [a for a in data["mother_accounts"] if not a.get("sold", False)]
    if not accs:
        send_msg(chat_id, "স্টকে কোনো অবিক্রিত একাউন্ট নেই।")
        return
    msg = "🗑️ **মাদার একাউন্ট মুছুন**\nনিচের তালিকা থেকে মুছতে ইনডেক্স সংখ্যা দিন (1,2,...):\n\n"
    for i, a in enumerate(accs, 1):
        msg += f"{i}. {a['username']} (পাস: {a['password']})\n"
    send_msg(chat_id, msg, parse_mode="Markdown")
    user_sessions[chat_id] = {"step": "delete_mother", "list": accs}

def admin_delete_mother(chat_id, text):
    session = user_sessions.pop(chat_id, None)
    if not session:
        return
    try:
        idx = int(text.strip()) - 1
        if idx < 0 or idx >= len(session["list"]):
            raise ValueError
    except:
        send_msg(chat_id, "❌ সঠিক ইনডেক্স দিন।")
        return
    to_delete = session["list"][idx]
    with data_lock:
        for a in data["mother_accounts"]:
            if a["username"] == to_delete["username"] and a["password"] == to_delete["password"]:
                data["mother_accounts"].remove(a)
                break
        save_data(data)
    send_msg(chat_id, f"✅ `{to_delete['username']}` মুছে ফেলা হয়েছে।", reply_markup=admin_keyboard())

def admin_deposit_requests(chat_id):
    pending = [d for d in data["deposits"] if d["status"] == "pending"]
    if not pending:
        send_msg(chat_id, "কোনো পেন্ডিং ডিপোজিট নেই।")
        return
    for d in pending:
        msg = (
            f"📥 ডিপোজিট আইডি: {d['id']}\n"
            f"ইউজার: {d['user_id']}\n"
            f"পরিমাণ: {d['amount']} টাকা\n"
            f"মাধ্যম: {d['method'].upper()}\n"
            f"TrxID: {d['trxid']}\n"
            f"অনুমোদন: /approvedeposit {d['id']}\n"
            f"বাতিল: /rejectdeposit {d['id']}"
        )
        send_msg(chat_id, msg)

def admin_broadcast_start(chat_id):
    send_msg(chat_id, "📢 ব্রডকাস্ট মেসেজ লিখুন (সব ইউজারকে পাঠানো হবে):\nবাতিল করতে /cancel")
    user_sessions[chat_id] = {"step": "broadcast"}

def admin_broadcast(chat_id, text):
    user_sessions.pop(chat_id, None)
    with data_lock:
        users = list(data["users"].keys())
    for uid in users:
        send_msg(uid, f"📢 ব্রডকাস্ট:\n\n{text}")
        time.sleep(0.05)
    send_msg(chat_id, f"✅ {len(users)} জনকে মেসেজ পাঠানো হয়েছে।")

# ========== ব্যাকআপ কমান্ড হ্যান্ডলার ==========
def admin_backup(chat_id):
    send_msg(chat_id, "⏳ ব্যাকআপ তৈরি হচ্ছে...")
    # থ্রেডে চালান যাতে ব্লক না হয়
    threading.Thread(target=do_backup_and_notify, args=(chat_id,), daemon=True).start()

def do_backup_and_notify(chat_id):
    try:
        save_data_to_channel()
        send_msg(chat_id, "✅ ব্যাকআপ সফলভাবে তৈরি ও পিন করা হয়েছে।")
    except Exception as e:
        send_msg(chat_id, f"❌ ব্যাকআপ ব্যর্থ: {e}")

def admin_restore(chat_id, msg=None):
    # যদি রিপ্লাই করা মেসেজে ডকুমেন্ট থাকে, তা থেকে রিস্টোর
    if msg and msg.get("reply_to_message") and msg["reply_to_message"].get("document"):
        file_id = msg["reply_to_message"]["document"]["file_id"]
        threading.Thread(target=restore_from_file, args=(chat_id, file_id), daemon=True).start()
    else:
        # না হলে পিন করা থেকে রিস্টোর
        send_msg(chat_id, "⏳ পিন করা ব্যাকআপ থেকে রিস্টোর করা হচ্ছে...")
        threading.Thread(target=do_restore_from_channel, args=(chat_id,), daemon=True).start()

def restore_from_file(chat_id, file_id):
    try:
        file_info = requests.get(f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={file_id}").json()
        if not file_info.get("ok"):
            send_msg(chat_id, "❌ ফাইল তথ্য পাওয়া যায়নি।")
            return
        file_path = file_info["result"]["file_path"]
        content = requests.get(f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}", timeout=60).content
        decompressed = gzip.decompress(content)
        restored_data = json.loads(decompressed.decode('utf-8'))
        with data_lock:
            data["users"] = restored_data.get("users", {})
            data["mother_accounts"] = restored_data.get("mother_accounts", [])
            data["deposits"] = restored_data.get("deposits", [])
            save_data(data)
        send_msg(chat_id, "✅ রিস্টোর সফল (ফাইল থেকে)।")
    except Exception as e:
        send_msg(chat_id, f"❌ রিস্টোর ব্যর্থ: {e}")

def do_restore_from_channel(chat_id):
    try:
        auto_restore_from_channel()  # এটি নিজেই অ্যাডমিনকে মেসেজ পাঠায়
    except Exception as e:
        send_msg(chat_id, f"❌ রিস্টোর ব্যর্থ: {e}")

# ========== কমান্ড প্রসেসিং (সংশোধন) ==========
def handle_command(chat_id, text, msg=None):
    if text.startswith("/"):
        cmd = text.split()[0].lower()
        if cmd == "/start":
            handle_start(chat_id)
        elif cmd == "/cancel":
            user_sessions.pop(chat_id, None)
            send_msg(chat_id, "❌ বর্তমান কাজ বাতিল করা হয়েছে।", reply_markup=main_keyboard(chat_id))
        elif cmd == "/setfreelimit" and str(chat_id) == ADMIN_ID:
            parts = text.split()
            if len(parts) == 2:
                try:
                    limit = int(parts[1])
                    if limit < 0: raise ValueError
                    data["config"]["free_limit"] = limit
                    save_data(data)
                    send_msg(chat_id, f"✅ ফ্রি লিমিট {limit} এ সেট করা হয়েছে।")
                except:
                    send_msg(chat_id, "❌ সঠিক সংখ্যা দিন।")
            else:
                send_msg(chat_id, "/setfreelimit <সংখ্যা>")
        elif cmd == "/setprice" and str(chat_id) == ADMIN_ID:
            parts = text.split()
            if len(parts) == 2:
                try:
                    price = float(parts[1])
                    if price < 0: raise ValueError
                    data["config"]["price_per_account"] = price
                    save_data(data)
                    send_msg(chat_id, f"✅ প্রতি একাউন্টের দাম {price} টাকা সেট করা হয়েছে।")
                except:
                    send_msg(chat_id, "❌ সঠিক দাম দিন।")
            else:
                send_msg(chat_id, "/setprice <দাম>")
        elif cmd == "/setbkash" and str(chat_id) == ADMIN_ID:
            parts = text.split()
            if len(parts) == 2:
                data["config"]["bkash_number"] = parts[1]
                save_data(data)
                send_msg(chat_id, f"✅ বিকাশ নম্বর {parts[1]} সেট করা হয়েছে।")
            else:
                send_msg(chat_id, "/setbkash <নম্বর>")
        elif cmd == "/setnagad" and str(chat_id) == ADMIN_ID:
            parts = text.split()
            if len(parts) == 2:
                data["config"]["nagad_number"] = parts[1]
                save_data(data)
                send_msg(chat_id, f"✅ নগদ নম্বর {parts[1]} সেট করা হয়েছে।")
            else:
                send_msg(chat_id, "/setnagad <নম্বর>")
        elif cmd == "/setbackupchannel" and str(chat_id) == ADMIN_ID:
            parts = text.split()
            if len(parts) == 2:
                data["config"]["backup_channel_id"] = parts[1]
                save_data(data)
                send_msg(chat_id, f"✅ ব্যাকআপ চ্যানেল আইডি {parts[1]} সেট করা হয়েছে।")
            else:
                send_msg(chat_id, "/setbackupchannel <চ্যানেল_আইডি>")
        elif cmd == "/backup" and str(chat_id) == ADMIN_ID:
            admin_backup(chat_id)
        elif cmd == "/restore" and str(chat_id) == ADMIN_ID:
            admin_restore(chat_id, msg)
        elif cmd == "/approvedeposit" and str(chat_id) == ADMIN_ID:
            parts = text.split()
            if len(parts) == 2:
                dep_id = parts[1]
                with data_lock:
                    for d in data["deposits"]:
                        if d["id"] == dep_id and d["status"] == "pending":
                            d["status"] = "approved"
                            user_id = d["user_id"]
                            data["users"][user_id]["balance"] += d["amount"]
                            save_data(data)
                            send_msg(chat_id, f"✅ ডিপোজিট {dep_id} অনুমোদিত। {d['amount']} টাকা যোগ হয়েছে।")
                            send_msg(user_id, f"✅ আপনার {d['amount']} টাকার ডিপোজিট অনুমোদিত হয়েছে। নতুন ব্যালেন্স: {data['users'][user_id]['balance']} টাকা")
                            break
                    else:
                        send_msg(chat_id, "❌ ডিপোজিট পাওয়া যায়নি বা ইতিমধ্যে প্রসেস করা হয়েছে।")
            else:
                send_msg(chat_id, "/approvedeposit <আইডি>")
        elif cmd == "/rejectdeposit" and str(chat_id) == ADMIN_ID:
            parts = text.split()
            if len(parts) == 2:
                dep_id = parts[1]
                with data_lock:
                    for d in data["deposits"]:
                        if d["id"] == dep_id and d["status"] == "pending":
                            d["status"] = "rejected"
                            save_data(data)
                            send_msg(chat_id, f"❌ ডিপোজিট {dep_id} বাতিল করা হয়েছে।")
                            send_msg(d["user_id"], f"❌ আপনার {d['amount']} টাকার ডিপোজিট বাতিল করা হয়েছে।")
                            break
                    else:
                        send_msg(chat_id, "❌ ডিপোজিট পাওয়া যায়নি।")
            else:
                send_msg(chat_id, "/rejectdeposit <আইডি>")
        else:
            send_msg(chat_id, "❌ অজানা কমান্ড।")
        return True
    return False

# ========== মেসেজ হ্যান্ডলার (সংশোধন) ==========
def handle_message(chat_id, text, username=None, full_msg=None):
    if username:
        with data_lock:
            if str(chat_id) in data["users"]:
                data["users"][str(chat_id)]["username"] = username
                save_data(data)

    if chat_id in user_sessions:
        session = user_sessions[chat_id]
        step = session.get("step")
        if step == "buy_qty":
            handle_buy_qty(chat_id, text)
            return
        elif step == "dep_amount":
            handle_deposit_amount(chat_id, text)
            return
        elif step == "dep_trxid":
            handle_deposit_trxid(chat_id, text)
            return
        elif step == "add_username":
            admin_add_mother_username(chat_id, text)
            return
        elif step == "add_password":
            admin_add_mother_password(chat_id, text)
            return
        elif step == "add_2fa":
            admin_add_mother_2fa(chat_id, text)
            return
        elif step == "delete_mother":
            admin_delete_mother(chat_id, text)
            return
        elif step == "broadcast":
            admin_broadcast(chat_id, text)
            return
        else:
            user_sessions.pop(chat_id, None)
            send_msg(chat_id, "❌ সেশন রিসেট করা হয়েছে। আবার চেষ্টা করুন।", reply_markup=main_keyboard(chat_id))
            return

    if handle_command(chat_id, text, full_msg):
        return

    # বাটন হ্যান্ডেল
    if text == "📦 ফ্রি মাদার":
        handle_free_mother(chat_id)
    elif text == "🛒 মাদার কিনুন":
        handle_buy_start(chat_id)
    elif text == "💰 ব্যালেন্স":
        handle_balance(chat_id)
    elif text == "👤 প্রোফাইল":
        handle_profile(chat_id)
    elif text == "💳 ডিপোজিট":
        handle_deposit_start(chat_id)
    elif text == "📞 সাপোর্ট":
        handle_support(chat_id)
    elif text == "🛠️ অ্যাডমিন প্যানেল" and str(chat_id) == ADMIN_ID:
        admin_panel(chat_id)
    elif text == "📊 ইউজার লিস্ট" and str(chat_id) == ADMIN_ID:
        admin_user_list(chat_id)
    elif text == "⚙️ সেটিংস" and str(chat_id) == ADMIN_ID:
        admin_settings(chat_id)
    elif text == "➕ মাদার যোগ" and str(chat_id) == ADMIN_ID:
        admin_add_mother_start(chat_id)
    elif text == "🗑️ মাদার মুছুন" and str(chat_id) == ADMIN_ID:
        admin_delete_mother_start(chat_id)
    elif text == "📥 ডিপোজিট রিকোয়েস্ট" and str(chat_id) == ADMIN_ID:
        admin_deposit_requests(chat_id)
    elif text == "📢 ব্রডকাস্ট" and str(chat_id) == ADMIN_ID:
        admin_broadcast_start(chat_id)
    elif text == "📁 ব্যাকআপ করুন" and str(chat_id) == ADMIN_ID:
        admin_backup(chat_id)
    elif text == "📥 রিস্টোর করুন" and str(chat_id) == ADMIN_ID:
        admin_restore(chat_id, None)  # পিন করা থেকে রিস্টোর
    elif text == "🔙 মূল মেনু" and str(chat_id) == ADMIN_ID:
        send_msg(chat_id, "মূল মেনু", reply_markup=main_keyboard(chat_id))
    else:
        send_msg(chat_id, "❓ বুঝতে পারিনি। মেনু ব্যবহার করুন।", reply_markup=main_keyboard(chat_id))

def handle_callback(chat_id, callback_id, data_str):
    answer_callback(callback_id)
    if data_str.startswith("dep_"):
        if data_str == "dep_cancel":
            user_sessions.pop(chat_id, None)
            send_msg(chat_id, "❌ ডিপোজিট বাতিল।", reply_markup=main_keyboard(chat_id))
            return
        method = data_str.split("_")[1]
        handle_deposit_method(chat_id, method)

# ========== পোলিং লুপ ==========
last_update_id = 0

def poll_updates():
    global last_update_id
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates"
    while True:
        try:
            params = {"timeout": 30, "offset": last_update_id + 1}
            resp = requests.get(url, params=params, timeout=35).json()
            if resp.get("ok") and resp.get("result"):
                for update in resp["result"]:
                    last_update_id = update["update_id"]
                    if "message" in update:
                        msg = update["message"]
                        chat_id = str(msg["chat"]["id"])
                        text = msg.get("text", "").strip()
                        username = msg.get("from", {}).get("username") or msg.get("from", {}).get("first_name", "Unknown")
                        handle_message(chat_id, text, username, msg)
                    elif "callback_query" in update:
                        cb = update["callback_query"]
                        chat_id = str(cb["message"]["chat"]["id"])
                        data_str = cb["data"]
                        handle_callback(chat_id, cb["id"], data_str)
            time.sleep(1)
        except Exception as e:
            logging.error(f"Polling error: {e}")
            time.sleep(5)

# ========== ফ্লাস্ক ==========
app = Flask(__name__)

@app.route("/")
def home():
    return "Bot is running!"

# ========== মেইন ==========
if __name__ == "__main__":
    if not os.path.exists(DATA_FILE):
        save_data(default_data)

    # চ্যানেল আইডি সেট না থাকলে এনভায়রনমেন্ট থেকে নিন (যদি সেট থাকে)
    if BACKUP_CHANNEL_ID and not data["config"].get("backup_channel_id"):
        data["config"]["backup_channel_id"] = BACKUP_CHANNEL_ID
        save_data(data)

    # স্বয়ংক্রিয় রিস্টোর (যদি চ্যানেল আইডি থাকে)
    if data["config"].get("backup_channel_id"):
        auto_restore_from_channel()

    # থ্রেড শুরু করুন
    threading.Thread(target=poll_updates, daemon=True).start()
    threading.Thread(target=auto_backup_loop, daemon=True).start()

    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
