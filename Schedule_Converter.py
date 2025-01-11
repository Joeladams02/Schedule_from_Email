from Handle_Emails import EmailHandler
from Read_from_Excel import ReadExcel
from Create_Calendar import DatesToICS
 
email_handler = EmailHandler(username="Joeladams02@icloud.com", 
                             password="vttg-qutv-cvxe-tkon", 
                             server="imap.mail.me.com")

'''Login into my email and fetch excel file'''

schedule_emails = email_handler.fetch_emails()
email_handler.save_excel(schedule_emails)

try:
    '''Read excel file and output the dates as a dictionary'''

    excel_read = ReadExcel('Schedule.xlsx')
    excel_read.process_Excel()
    dates = excel_read.get_scheduled_dates()

    '''Convert dates to ics file of individual lessons'''
    dates_to_calendar = DatesToICS(dates)
    dates_to_calendar.dates_to_lessons()
    dates_to_calendar.lessons_to_ics()

    '''Send ics file to my email for me to open and add to calendar'''
    email_handler.send_email('Schedule.ics')

except FileNotFoundError:
    print('email handler not working')