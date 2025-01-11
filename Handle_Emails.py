import email.policy
import imaplib
import time
from datetime import datetime, timedelta
import email
from email.message import EmailMessage

class EmailHandler:
    def __init__(self, username, password, server):
        self.username = username
        self.password = password
        self.server = server
        self.emails = []

    def login(self, mailbox="INBOX"):
        try:
            imap = imaplib.IMAP4_SSL(self.server)
            imap.login(self.username, self.password)
            imap.select(mailbox)
            print("Login successful!")
            return imap
        except imaplib.IMAP4.error as e:
            print(f"Authentication failed: {e}")
            return None

    def fetch_emails(self):
        with self.login() as imap:
            if not imap:
                return
            yesterday = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%Y')
            status, messages = imap.search(None, f'SINCE {yesterday} SUBJECT "Schedule"')
            self.emails = messages[0].decode('utf-8').split()
            print(self.emails)

            emails = []
            for UID in self.emails:
                status, data = imap.fetch(str(UID), '(BODY[])')
                if status == 'OK':
                    email_message = email.message_from_bytes(data[0][1], policy = email.policy.default)
                    date_sent = email_message['Date']
                    emails.append([UID, date_sent, email_message]) 
            imap.close()
            imap.logout()
        return emails

    def send_email(self, file_path):
        # Create the email message
        msg = EmailMessage()
        msg['Subject'] = 'Updated Calendar'
        msg['From'] = self.username
        msg['To'] = self.username
        msg.set_content("Please find your updated calendar attached.")

        # Attach the .ics file
        with open(file_path, 'rb') as file:
            file_data = file.read()
            file_name = file_path.split("/")[-1]
            msg.add_attachment(file_data, maintype='text', subtype='calendar', filename=file_name)

        with self.login() as imap:
            if not imap:
                return
            imap.append('INBOX', '',imaplib.Time2Internaldate(time.time()), str(msg).encode('utf-8'))

            imap.close()
            imap.logout()
      
    def save_excel(self, emails):
        emails = sorted(emails, key=lambda x: x[1])
        for email_message in emails:
            if email_message[2].iter_attachments():
                for part in email_message[2].iter_attachments():
                    filename = part.get_filename()
                    if filename and filename.endswith('.xlsx'):
                        attachment_data = part.get_payload(decode=True)
                        with open('Schedule.xlsx', 'wb') as f:
                            f.write(attachment_data)