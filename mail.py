import smtplib
import imaplib
import email
from email.mime.text import MIMEText
import pandas as pd
import google.generativeai as genai
from email.header import decode_header
import time
import logging
from datetime import datetime
import pytz
import re
import os

# Load environment variables
GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS")
APP_PASSWORD = os.getenv("APP_PASSWORD")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
EXCEL_FILE = os.getenv("EXCEL_FILE", "users.xlsx")  # Default to 'users.xlsx' if not set

# Configure logging to a file
logging.basicConfig(
    filename='email_responder.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

ANSWER_TEMPLATES = {
    "greeting": {
        "keywords": ["hello", "hi", "hey"],
        "response": """Dear {sender_name},\n\nThank you for reaching out! How can I assist you today?\n\nBest regards,\nAnsh"""
    },
    "support": {
        "keywords": ["help", "issue", "problem", "support"],
        "response": """Dear {sender_name},\n\nI'm here to help with your concern. Could you please provide more details about the issue you're facing?\n\nBest regards,\nAnsh"""
    },
    "pricing": {
        "keywords": ["price", "cost", "pricing", "how much"],
        "response": """Dear {sender_name},\n\nThank you for your interest in our pricing. Could you please specify which product/service you're interested in?\n\nBest regards,\nAnsh"""
    },
}

def get_thread_history(mail, message_id):
    try:
        status, messages = mail.search(None, f'HEADER Message-ID "{message_id}"')
        thread_ids = messages[0].split()
        status, ref_messages = mail.search(None, f'HEADER References "{message_id}"')
        thread_ids.extend(ref_messages[0].split())
        thread_ids = list(set(thread_ids))
        
        email_history = []
        for msg_id in thread_ids:
            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            raw_email = msg_data[0][1]
            email_message = email.message_from_bytes(raw_email)
            
            subject = decode_header(email_message["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            
            sender = email_message["From"]
            date = email_message["Date"]
            if email_message.is_multipart():
                content = email_message.get_payload(0).get_payload(decode=True)
                if isinstance(content, bytes):
                    content = content.decode()
            else:
                content = email_message.get_payload(decode=True)
                if isinstance(content, bytes):
                    content = content.decode()
            
            email_history.append(f"Date: {date}\nFrom: {sender}\nSubject: {subject}\nContent: {content}\n\n")
        
        return "\n".join(email_history)
    except Exception as e:
        logger.error(f"Error getting thread history: {e}")
        return ""

def generate_response(history, sender_name, original_subject):
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel('gemini-2.0-flash')  # Verify model name
        
        template_info = "\n".join([f"Template '{key}': {value['response']}\nKeywords: {', '.join(value.get('keywords', []))}" 
                                 for key, value in ANSWER_TEMPLATES.items()])
        
        prompt = f"""Given this email thread history:
{history}

And these available response templates:
{template_info}

Generate an appropriate professional response for the latest email in the thread. 
Address the sender as {sender_name}.
You can use one of the templates if it fits (replace {sender_name} with the actual name), 
or generate a custom response if none of the templates are appropriate.
Keep it concise and helpful. Do not include the subject line in your response.
Return ONLY the final response text, without any commentary about template selection or reasoning."""
        
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        logger.error(f"Error generating response: {e}")
        return ANSWER_TEMPLATES["default"]["response"].format(sender_name=sender_name)

def send_reply(to, subject, message_text, message_id):
    try:
        msg = MIMEText(message_text)
        msg['Subject'] = subject
        msg['From'] = GMAIL_ADDRESS
        msg['To'] = to
        msg['In-Reply-To'] = message_id
        msg['References'] = message_id
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_ADDRESS, APP_PASSWORD)
            server.send_message(msg)
            logger.info(f"Reply content sent to {to}:\nSubject: {subject}\n{message_text}")
        return True
    except Exception as e:
        logger.error(f"Error sending reply to {to}: {e}")
        return False

def clean_subject(subject):
    return re.sub(r'^(Re:\s*)+', 'Re: ', subject).strip()

def main():
    try:
        df = pd.read_excel(EXCEL_FILE)
        users = df.set_index('email').to_dict('index')
        valid_emails = set(users.keys())
    except Exception as e:
        logger.error(f"Error loading Excel file: {e}")
        return

    server_start_time = datetime.now(pytz.UTC)
    logger.info(f"Server started at {server_start_time}")

    while True:
        try:
            logger.info("Checking for new emails")
            
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(GMAIL_ADDRESS, APP_PASSWORD)
            mail.select("inbox")
            
            imap_date = server_start_time.strftime("%d-%b-%Y")
            status, messages = mail.search(None, f'SINCE "{imap_date}" UNSEEN')
            message_ids = messages[0].split()
            
            valid_message_count = 0
            for msg_id in message_ids:
                status, msg_data = mail.fetch(msg_id, "(RFC822)")
                raw_email = msg_data[0][1]
                email_message = email.message_from_bytes(raw_email)
                sender = email_message["From"]
                sender_email = sender.split('<')[-1].strip('>').lower()
                if sender_email in valid_emails:
                    valid_message_count += 1
            
            if valid_message_count > 0:
                logger.info(f"Found {valid_message_count} new mail(s) from contact list")
            
            for msg_id in message_ids:
                status, msg_data = mail.fetch(msg_id, "(RFC822)")
                raw_email = msg_data[0][1]
                email_message = email.message_from_bytes(raw_email)
                
                msg_date_str = email_message["Date"]
                if msg_date_str:
                    try:
                        msg_date = email.utils.parsedate_to_datetime(msg_date_str)
                        if msg_date.tzinfo is None:
                            msg_date = pytz.UTC.localize(msg_date)
                            
                        if msg_date > server_start_time:
                            sender = email_message["From"]
                            sender_email = sender.split('<')[-1].strip('>').lower()
                            
                            subject_raw = decode_header(email_message["Subject"])[0][0]
                            original_subject = subject_raw.decode() if isinstance(subject_raw, bytes) else subject_raw
                            subject = clean_subject(original_subject)
                            
                            message_id = email_message["Message-ID"]
                            
                            if sender_email in valid_emails:
                                user_info = users[sender_email]
                                sender_name = user_info['name']
                                
                                thread_history = get_thread_history(mail, message_id)
                                response = generate_response(thread_history, sender_name, original_subject)
                                
                                if send_reply(sender_email, subject, response, message_id):
                                    logger.info(f"Reply successfully sent to: {sender_email}")
                                else:
                                    logger.error(f"Failed to send reply to: {sender_email}")
                                
                                mail.store(msg_id, '+FLAGS', '\\Seen')
                    except Exception as e:
                        logger.error(f"Error processing message {msg_id}: {e}")
                        continue
            
            mail.logout()
            time.sleep(10)
            
        except Exception as e:
            logger.error(f"Server error: {e}")
            logger.info("Server ended")
            break

if __name__ == '__main__':
    main()