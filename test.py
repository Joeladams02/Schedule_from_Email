from Handle_Emails import EmailHandler

email_handler = EmailHandler(username="Joeladams02@icloud.com", 
                             password="vttg-qutv-cvxe-tkon", 
                             server="imap.mail.me.com")

email_handler.send_email('Schedule.ics')