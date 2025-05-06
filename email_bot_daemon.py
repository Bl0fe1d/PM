import imaplib
import email
from email.header import decode_header
import os
import csv
from datetime import datetime
import schedule
import time
import logging

# ========== НАСТРОЙКИ ==========

EMAIL = "your_email@gmail.com"
PASSWORD = "your_app_password"
IMAP_SERVER = "imap.gmail.com"

CHECK_INTERVAL_MINUTES = 5  # интервал проверки

BASE_DIR = "attachments"
LOG_DIR = "logs"
CSV_LOG = os.path.join(LOG_DIR, "email_log.csv")
TXT_LOG = os.path.join(LOG_DIR, "service.log")

# Категории по ключевым словам
CATEGORY_RULES = {
    "Работа": ["job", "вакансия", "resume", "собеседование"],
    "Финансы": ["invoice", "счёт", "оплата", "payment"],
    "Реклама": ["sale", "скидка", "promo", "offer"],
    "Личное": ["друзья", "приглашение", "вечеринка", "поздравление"]
}

# ========== ПОДГОТОВКА ==========

# Создание необходимых папок
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# Настройка текстового логгера
logging.basicConfig(
    filename=TXT_LOG,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding='utf-8'
)

# Создание CSV-лога, если он ещё не создан
if not os.path.exists(CSV_LOG):
    with open(CSV_LOG, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Время", "Отправитель", "Тема", "Категория", "Вложения"])

# ========== ФУНКЦИИ ==========

def get_category(subject):
    subject_lower = subject.lower()
    for category, keywords in CATEGORY_RULES.items():
        for keyword in keywords:
            if keyword in subject_lower:
                return category
    return "Другое"

def save_attachments(msg, category):
    saved_files = []
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        if filename:
            folder_path = os.path.join(BASE_DIR, category)
            os.makedirs(folder_path, exist_ok=True)
            filepath = os.path.join(folder_path, filename)

            # Обработка дубликатов имён
            if os.path.exists(filepath):
                base, ext = os.path.splitext(filename)
                filename = f"{base}_{int(time.time())}{ext}"
                filepath = os.path.join(folder_path, filename)

            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))
            saved_files.append(filepath)
            logging.info(f"Сохранено вложение: {filepath}")
    return saved_files

def log_email(timestamp, sender, subject, category, attachments):
    with open(CSV_LOG, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([timestamp, sender, subject, category, "; ".join(attachments)])

def process_emails():
    logging.info("Начало проверки почты")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        status, messages = mail.search(None, "UNSEEN")
        email_ids = messages[0].split()

        for num in email_ids:
            status, data = mail.fetch(num, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = msg["Subject"] or "Без темы"
            decoded_subject, encoding = decode_header(subject)[0]
            subject = decoded_subject.decode(encoding) if isinstance(decoded_subject, bytes) else decoded_subject
            from_ = msg.get("From")
            category = get_category(subject)

            logging.info(f"[{category}] Письмо от {from_}: {subject}")
            attachments = save_attachments(msg, category)

            log_email(
                timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                sender=from_,
                subject=subject,
                category=category,
                attachments=attachments
            )

        mail.logout()
        logging.info("Проверка завершена\n")
    except Exception as e:
        logging.exception("Ошибка при проверке почты")

# ========== ФОНОВАЯ РАБОТА ==========

def main_loop():
    logging.info("Email-бот запущен в фоне")
    schedule.every(CHECK_INTERVAL_MINUTES).minutes.do(process_emails)

    while True:
        schedule.run_pending()
        time.sleep(1)

# ========== ЗАПУСК ==========

if __name__ == "__main__":
    try:
        main_loop()
    except Exception as e:
        logging.exception("Критическая ошибка в работе бота")
