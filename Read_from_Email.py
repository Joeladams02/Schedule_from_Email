import email.policy
import imaplib
from datetime import datetime, timedelta
import email

class EmailFetcher:
    def __init__(self, username, password, server):
        self.username = username
        self.password = password
        self.server = server

    def login(self, mailbox="INBOX"):
        """Log in to the IMAP server."""
        try:
            imap = imaplib.IMAP4_SSL(self.server)
            imap.login(self.username, self.password)
            imap.select(mailbox)
            print("Login successful!")
            return imap
        except imaplib.IMAP4.error as e:
            print(f"Authentication failed: {e}")
            return None

    def get_numbers(self):
        """Get email numbers with subject 'Schedule' from the last 24 hours."""
        with self.login() as imap:
            if not imap:
                return
            yesterday = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%Y')
            status, messages = imap.search(None, f'SINCE {yesterday} SUBJECT "Schedule"')
            with open('imap_numbers.txt', 'wb') as file:
                file.write(messages[0])
            imap.close()
            imap.logout()

    def fetch_emails(self):
        """Fetch the email content for the email numbers stored in the text file."""
        with self.login() as imap:
            if not imap:
                return
            with open('imap_numbers.txt', 'r') as file:
                messages = file.read().split()
            emails = []
            for num in messages:
                status, data = imap.fetch(num, '(BODY[])')
                if status == 'OK':
                    email_message = email.message_from_bytes(data[0][1], policy = email.policy.default)
                    print(f"Email fetched: {email_message['subject']}")
                    date_sent = email_message['Date']
                    emails.append([num, date_sent, email_message]) 
            imap.close()
            imap.logout()
            return emails
        
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
                        print(f"Saved attachment as {filename}")
