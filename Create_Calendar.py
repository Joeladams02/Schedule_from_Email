from icalendar import Calendar, Event
from datetime import datetime, timedelta

class DatesToICS:
    def __init__(self,dates):
        self.dates = dates
        self.scheduled_sessions = {}

    def convert(self,time):
        try:
            if 'noon' in time:
                time = time.replace('noon','').replace('am','').replace('pm','').split('-')
            elif 'am' in time and 'pm' not in time:
                time = time.replace('am','').split('-')
            elif 'am' not in time and 'pm' in time:
                time = time.replace('pm','').split('-')
                if '12' in time[0].split('.'):
                    time = [time[0], str(float(time[1])+12)]
                elif '12' not in time[1].split('.'):
                    time = [str(float(x) + 12) for x in time]
            elif 'am' in time and 'pm' in time:
                try:
                    time = time.replace('am','').replace('pm','').split('-')
                    if '12' not in time[1].split('.'):
                        time = [time[0], str(float(time[1]) + 12)]
                except:
                    time = ['00.00', '01.00']
            else:
                print(f'unidentified{time}')
            return time
        except:
            return ['00.00', '01.00']

    def dates_to_lessons(self):
        for keys,items in self.dates.items():
            lessons = [x.strip() for x in items.split(';') if x.strip()]
            for lesson in lessons:            
                raw_time = [x.strip() for x in lesson.split(' ') if x.strip()]
                time = self.convert(raw_time[0])

                length = round((float(time[1]) - float(time[0])),2)
                time_split = time[0].split('.')
                hour = int(time_split[0])
                mins = int(time_split[1]) if len(time_split) > 1 else 0
                if mins < 10:
                    mins = mins * 10
                time_and_date = datetime(keys.year,keys.month,keys.day, hour, mins)
                self.scheduled_sessions[time_and_date] = [length, lesson.replace(raw_time[0], '')]
                
    def lessons_to_ics(self):
        try:
            with open('Schedule.ics', 'rb') as f:
                cal = Calendar.from_ical(f.read())
        except FileNotFoundError:
            cal = Calendar()

        for keys,items in self.scheduled_sessions.items():
            UID = keys.strftime("%Y%m%dT%H%M")
            
            duplicate = False
            for component in cal.subcomponents:
                if component.get("UID") == UID:
                    duplicate = True
                    break
            if not duplicate:
                print('New event found')
                end_time = keys + timedelta(hours = items[0])
                event = Event()
                event.add('uid', UID)
                event.add('summary', items[1])
                event.add('dtstart', keys)
                event.add('dtend', end_time)
                cal.add_component(event)

        # Save updated file
        with open('Schedule.ics', 'wb') as f:
            f.write(cal.to_ical())

