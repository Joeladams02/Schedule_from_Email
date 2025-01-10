import icalendar
from Read_from_Email import EmailFetcher
from Read_from_Excel import ScheduleProcessor


 
email_fetcher = EmailFetcher(username="Joeladams02@icloud.com", 
                             password="vttg-qutv-cvxe-tkon", 
                             server="imap.mail.me.com")

schedule_emails = email_fetcher.fetch_emails()
email_fetcher.save_excel(schedule_emails)

excel_file = 'Schedule.xlsx'

schedule_processor = ScheduleProcessor(excel_file)

schedule_processor.process_schedule()
dates = schedule_processor.get_scheduled_dates()

