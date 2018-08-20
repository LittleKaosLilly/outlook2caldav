import sys
import win32com.client
from win32com.client import constants as c
from ics import Calendar, Event
import chardet

from dateutil import tz
from dateutil.parser import parse


class MyOutlookCalendar(object):
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.nsoutlook = self.outlook.GetNamespace("MAPI")
        self.defaultCalendar = 9
        self.verbose = False

        # cal = ns.GetDefaultFolder(win32com.client.constants.olFolderCalendar)
        # inbox = mapi.GetDefaultFolder(win32com.client.constants.olFolderInbox)
        # "6" refers to the index of a folder - in this case, the inbox.
        # "9" is calendar
        '''
        how to reach any default folder not just "Inbox" here's the list:
        3  Deleted Items
        4  Outbox
        5  Sent Items
        6  Inbox
        9  Calendar
        10 Contacts
        11 Journal
        12 Notes
        13 Tasks
        14 Drafts
        
        use print_all_email_boxes or print_all_default_folders
        '''

    def enable_verbose(self):
        self.verbose = True

    def print_all_email_boxes(self):
        ''' Print all email boxes.'''

        for i in range(50):
            try:
                box = self.nsoutlook.Folders(i)
                name = box.Name
                print(i, name)
            except:
                pass

    def print_all_default_folders(self):
        ''' Print all default folders.'''

        for i in range(50):
            try:
                box = self.nsoutlook.GetDefaultFolder(i)
                name = box.Name
                print(i, name)
            except Exception as e:
                print("Exception: %s" % (str(e)))
                pass

    def send_meeting_request(self, to, subject, location, start_time, end_time, body_text, all_day=False):
        '''
            Create item im calendar

            Set myItem = myOlApp.CreateItem(olAppointmentItem)
            myItem.MeetingStatus = olMeeting
            myItem.Subject = "Strategy Meeting"
            myItem.Location = "Conference Room B"
            myItem.Start = #9/24/97 1:30:00 PM#
            myItem.Duration = 90
        '''

        appt = self.nslookup.CreateItem(
            c.olAppointmentItem)  # https://msdn.microsoft.com/en-us/library/office/ff869291.aspx
        appt.MeetingStatus = c.olMeeting  # https://msdn.microsoft.com/EN-US/library/office/ff869427.aspx

        # only after setting the MeetingStatus can we add recipients
        appt.Recipients.Add(to)
        appt.Subject = subject
        appt.Location = location

        appt.Start = start_time

        appt.AllDayEvent = all_day

        end_time_list = end_time.split("/")
        end_time_list[-2] = str(int(end_time_list[-2]) + 1)
        appt.End = "/".join(end_time_list)

        appt.Body = body_text
        # appt.Save()
        # appt.Send()
        appt.Display()
        return True

    def remove_accents(self, str):
        """
        Thanks to MiniQuark:
        http://stackoverflow.com/questions/517923/what-is-the-best-way-to-remove-accents-in-a-python-unicode-string/517974#517974
        """

        nkfd_form = unicodedata.normalize('NFKD', unicode(str))
        return u"".join([c for c in nkfd_form if not unicodedata.combining(c)])

    def remove_accents_bis(self, str):
        """
        remove eszett char
        """
        return str.replace('ß', 'ss')

    def get_my_calendar_event(self, start, end, recurence):

        icscal = Calendar()
        # master_recurrent_events_to_add = []
        known_guid_events = {}

        cal = self.nsoutlook.GetDefaultFolder(self.defaultCalendar)
        events = cal.Items

        # https://msdn.microsoft.com/en-us/library/office/gg619398.aspx
        events.Sort("[Start]")
        print("REC: %s" % (str(recurence)))
        events.IncludeRecurrences = ("%s" % (str(recurence)))

        # https://msdn.microsoft.com/EN-US/library/office/ff869427.aspx
        # Indicates the status of the meeting.
        # Name,                        Value,         Description
        # olMeeting                    1         The meeting has been scheduled.
        # olMeetingCanceled            5         The scheduled meeting has been cancelled.
        # olMeetingReceived            3         The meeting request has been received.
        # olMeetingReceivedAndCanceled 7         The scheduled meeting has been cancelled
        #                                             but still appears on the user's calendar.
        # olNonMeeting                 0         An Appointment item without attendees has been scheduled.
        #                                             This status can be used to set up holidays on a calendar.

        # restrict by date
        restriction = ("[Start] >= '%s' AND [End] < '%s'" % (start, end))
        # restrict by date and type
        restriction = ("([MeetingStatus] = 1 OR [MeetingStatus] = 3 OR [MeetingStatus] = 0) "
                       " AND ([Start] >= '%s' AND [End] <= '%s')"
                       % (start, end))

        restricted_events = events.Restrict(restriction)

        for appointment_item in restricted_events:

            '''
            print("subj " + appointment_item.Subject)
            if appointment_item.Location != "":
                print("loc " + appointment_item.Location)
            print("start " + str(appointment_item.StartUTC))
            print("end " + str(appointment_item.EndUTC))
            print("allday " + str(appointment_item.AllDayEvent))
            print("body " + str(appointment_item.Body))
            print("busy " + str(appointment_item.BusyStatus))
            print("cats " + appointment_item.Categories)
            print("creationtime " + str(appointment_item.CreationTime))
            print("duration " + str(appointment_item.Duration))
            print("importance " + str(appointment_item.Importance))
            print("recurring " + str(appointment_item.IsRecurring))
            print("lastmod " + str(appointment_item.LastModificationTime))
            print("recps " + str(appointment_item.Recipients))
            print("recstate " + str(appointment_item.RecurrenceState))
            print("reminderminb4 " + str(appointment_item.ReminderMinutesBeforeStart))
            print("reqattendees " + appointment_item.RequiredAttendees)
            print("")
            print("")
            '''
            if appointment_item.Subject == 'no meeting':
                continue
            if appointment_item.Subject == 'not available':
                continue
            if appointment_item.Subject == 'Daily meeting - update':
                continue

            event_to_add = True

            e = Event()
            e.name = appointment_item.Subject
            e.uid = appointment_item.EntryID
            e.created = appointment_item.CreationTime

            # TODO: need to fetched the recurrent attribute, as all meeting have the same UID.
            if appointment_item.IsRecurring:
                if not recurence:
                    continue
                '''
                https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.olrecurrencestate?view=outlook-pia
                    appointment_item.RecurrenceState
                 olApptNotRecurring   0   The appointment is not a recurring appointment.
                 olApptMaster         1   The appointment is a master appointment.
                 olApptOccurrence     2   The appointment is an occurrence of a recurring appointment defined by a master appointment. 
                 olApptException      3   The appointment is an exception to a recurrence pattern defined by a master appointment.
                '''
                if self.verbose:
                    print('recuring event name: ' + appointment_item.Subject +
                          " Date:" + str(appointment_item.StartUTC) +
                          " State " + str(appointment_item.RecurrenceState) +
                          " UID: " + str(appointment_item.EntryID) +
                          " RecState " + str(appointment_item.RecurrenceState))

                if recurence:
                    recpat = appointment_item.GetRecurrencePattern()
                    # print(recpat)
                    print('Recurence StartDate: %s' % (str(recpat.PatternStartDate)))
                    print('Recurence EndDate: %s' % (str(recpat.PatternEndDate)))
                    print('Recurence StartTime: %s' % (str(recpat.StartTime)))
                    print('Recurence EndTime: %s' % (str(recpat.EndTime)))
                    print('Recurence Type %s' % (str(recpat.RecurrenceType)))
                    '''
                    https://docs.microsoft.com/en-us/office/vba/api/outlook.olrecurrencetype
                    olRecursDaily       0     Represents a daily recurrence pattern.
                    olRecursMonthly     2     Represents a monthly recurrence pattern.
                    olRecursMonthNth    3     Represents a MonthNth recurrence pattern. See RecurrencePattern.Instance property.
                    olRecursWeekly      1     Represents a weekly recurrence pattern.
                    olRecursYearly      5     Represents a yearly recurrence pattern.
                    olRecursYearNth     6     Represents a YearNth recurrence pattern. See RecurrencePattern.Instance property.
                    '''
                    print('------')

                if appointment_item.RecurrenceState != 1:
                    # Store UID to fetch the master event...
                    # if appointment_item.EntryID not in master_recurrent_events_to_add:
                    #    master_recurrent_events_to_add.append(appointment_item.EntryID)
                    if recurence is False:
                        event_to_add = False

            if event_to_add is False:
                continue

            if appointment_item.EntryID not in known_guid_events:
                known_guid_events[appointment_item.EntryID] = {}
                known_guid_events[appointment_item.EntryID]['count'] = 1
                known_guid_events[appointment_item.EntryID]['objects'] = []
            else:
                # if self.verbose:
                #     print('Duplicate GUID detected: ' + str(appointment_item.EntryID))
                print(
                    'Duplicate GUID detected: ' + str(appointment_item.EntryID) + '\n' + str(appointment_item.Subject))
                if recurence is False:
                    delete(known_guid_events[appointment_item.EntryID])
                    continue

                known_guid_events[appointment_item.EntryID]['count'] += 1

            if appointment_item.AllDayEvent is True:
                # For full day event, this is is stored in localtime in outlook and not UTC...
                # so start date is day - 1 and not day.
                # print('Is all day: Subject: %s' % (appointment_item.Subject))

                from_zone = tz.tzutc()
                to_zone = tz.tzlocal()
                utc = parse(str(appointment_item.StartUTC))
                utc = utc.astimezone(to_zone)
                utc = str(utc).split('+')[0]
                e.begin = utc
                e.make_all_day()

            else:
                e.begin = appointment_item.StartUTC
                e.end = appointment_item.EndUTC
            # e.duration = appointment_item.Duration

            body_detect = chardet.detect(appointment_item.Body.encode('utf-8'))
            if body_detect['encoding'] == 'ascii':
                e.description = appointment_item.Body
            elif body_detect['encoding'] is None:
                e.description = appointment_item.Body
            elif body_detect['encoding'] == 'utf-8':
                e.description = appointment_item.Body.encode('utf-8').decode('utf-8', 'ignore')
            elif body_detect['encoding'] == 'ISO-8859-1':
                e.description = appointment_item.Body.encode('utf-8').decode('iso-8859-1', 'ignore')
            elif body_detect['encoding'] == 'ISO-8859-15':
                e.description = appointment_item.Body.decode('iso8859_15', 'ignore')
            elif body_detect['encoding'] == 'Windows-1252':
                try:
                    e.description = appointment_item.Body.decode('windows-1252', 'ignore')
                except Exception as err:
                    # e.description = ''
                    e.description = appointment_item.Body
                    print('Unknown encoding: %s' % (body_detect))
                    print('Exception: %s' % (str(err)))
                    print('Subject: %s' % (appointment_item.Subject))
            else:
                print('Unknown encoding: %s' % (body_detect))
                e.description = ''
            e.description = e.description.replace('\r', '')

            if appointment_item.Location != "":
                e.location = appointment_item.Location

            # TODO: Alarm
            # TODO print("reminderminb4 " + str(appointment_item.ReminderMinutesBeforeStart))

            if event_to_add:
                icscal.events.add(e)

            known_guid_events[e.uid]['objects'].append(e)

        # if self.verbose:
        #    for event in master_recurrent_events_to_add:
        #        print('need to fetch master event with uid:' + str(event))

        return (icscal.events, known_guid_events)


if __name__ == '__main__':
    sys.exit(0)
