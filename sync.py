import sys
import os
from ics import Calendar, Event
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import configparser
import inspect

runPath = os.path.dirname(os.path.realpath(__file__))
sys.path.append(os.path.join(runPath, "lib"))

from myoutlook import MyOutlookCalendar
from mycaldav import MyCaldavCalendar


def obj_dump(obj, everything=False):
    '''
      Object dumper
    '''
    for attr in dir(obj):
        if '_' in attr and '__module__' not in attr and not everything:
            continue
        try:
            print("obj.%s = %s" % (attr, str(getattr(obj, attr))))
        except:
            print("obj.%s = %s" % (attr, getattr(obj, attr)))

        '''
        # bound method
        try:
            if '<bound method ' in str(getattr(obj, attr)):
                logger.debug("---- Start Dumping Bound ----")
                logger.debug("Boundobj.%s = %s" % (attr, str(getattr(obj, attr).__self__)))

                for battr in dir(getattr(obj, attr).__self__):
                    try:
                        logger.debug("obj.%s = %s" % (battr, str(getattr(obj, battr))))
                    except:
                        logger.debug("obj.%s = %s" % (battr, getattr(obj, battr)))

                logger.debug("---- End Dumping Bound ----")

        except:
            logger.debug("Oups:obj.%s = %s" % (attr, getattr(obj, attr)))
        '''

    for name, data in inspect.getmembers(obj):
        if inspect.isclass(data):
            print('name:%s\ndata:%s' % (name, data))

    # for name in inspect.getmembers(obj.getUrl,
    #        lambda m: inspect.ismethod(m) and m.__func__ in m.im_class.__dict__.values()):
    #    logger.debug('name:%s' % (pp.pformat(name)))


def copy_ia(oev, vcal, cal, my_caldav_by_uid, vcal_header, vcal_footer):
    '''

    Buisness parameter logic to perform the sync of the object
    :param oev:
    :param my_caldav_by_uid:
    :return:
    '''

    no_delete = False
    # print(oev.uid)
    if oev.uid in my_caldav_by_uid:
        # print(' * FOUND %s' % (oev.uid))
        '''
        obj_dump(oev)
        print(oev)
        print('------')
        obj_dump(my_caldav_by_uid[oev.uid])
        print(my_caldav_by_uid[oev.uid])
        '''

        c = Calendar(my_caldav_by_uid[oev.uid].data)
        cev = c.events[0]
        updated = False

        if oev.name != cev.name:
            if verbose:
                print("name do not match: '%s:%s'" % (oev.name, cev.name))
            cev.name = oev.name
            updated = True
        if oev.description != cev.description:
            if verbose:
                print("description do not match")
                # print("description do not match: '%s:%s'" % (oev.description, cev.description))
            cev.description = oev.description
            updated = True
        if oev.location != cev.location:
            if verbose:
                print("location do not match: '%s:%s'" % (oev.location, cev.location))
            cev.location = oev.location
            updated = True

        if oev.all_day:
            # outlook declare an all day
            if oev.begin != cev.begin:
                if verbose:
                    print("begin do not match: '%s:%s'" % (oev.begin, cev.begin))
                cev.begin = oev.begin
                updated = True
            if not cev.all_day:
                if verbose:
                    print("change event to all day format")
                cev.make_all_day()
                updated = True
        else:
            '''
            Need to update end date before start date, if not the exception is raised:
            raise ValueError('Begin must be before end')
            '''
            if oev.end != cev.end:
                if verbose:
                    print("end do not match: '%s:%s'" % (oev.end, cev.end))
                try:
                    cev.end = oev.end
                    updated = True
                except Exception as e:
                    pass
            if oev.begin != cev.begin:
                if verbose:
                    print("begin do not match: '%s:%s'" % (oev.begin, cev.begin))
                cev.begin = oev.begin
                updated = True

        if updated:
            my_caldav_by_uid[oev.uid].data = vcal_header + str(cev) + vcal_footer
            my_caldav_by_uid[oev.uid].save()

    else:
        '''create: Adding new outlook event in Caldav.'''
        print("Creating event name: '%s', '%s'" % (oev.name, oev.uid))
        # print(vcal)
        try:
            cal.add_event(vcal)
        except AuthorisationError as ae:
            print('Couldn\'t add event', ae.reason)
            no_delete = True
            pass
        # my_caldav_by_uid[oev.uid].data = vcal_header + str(cev) + vcal_footer
        # my_caldav_by_uid[oev.uid].save()

    return (my_caldav_by_uid, no_delete)


if __name__ == '__main__':

    cleanup_all_caldav = False
    # cleanup_all_caldav = True

    ''' BASE CONFIG '''
    global_iniFile = os.path.join(os.path.dirname(__file__), './etc/', 'configuration.ini')
    if not os.path.isfile(global_iniFile):
        print('%s' % (global_iniFile))
        sys.exit(1)

    ''' LOAD BASE CONFIG '''
    config_global = configparser.RawConfigParser()
    config_global.optionxform(str())
    config_global.optionxform = str
    config_global.read(global_iniFile)

    http_proxy = config_global.get('global', 'http_proxy', fallback=None)
    verbose = config_global.get('global', 'verbose', fallback=False)

    # For now 1 month back, 2 months forth
    period_before = int(config_global.get('period', 'before', fallback=1))
    period_after = int(config_global.get('period', 'after', fallback=2))

    now = datetime.now()
    first_day_of_month = now.replace(day=1)
    # first_day_of_month = datetime.datetime(now.year, now.month, 1)
    # TODO: include period_before in calculation
    lastMonth = first_day_of_month - timedelta(days=1)
    start = datetime(lastMonth.year, lastMonth.month, 1).strftime('%d/%m/%Y')

    date_after_month = now + relativedelta(months=period_after)
    end = date_after_month.strftime('%d/%m/%Y')

    caldav_username = config_global.get('remote', 'username')
    caldav_password = config_global.get('remote', 'password')
    caldav_type = config_global.get('remote', 'type', fallback='')
    caldav_url = config_global.get('remote', 'url', fallback='')
    caldav_name = config_global.get('remote', 'calendar_name')

    no_delete = False
    vcal_header = """BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Example Corp.//CalDAV Client//EN
"""

    vcal_footer = """
END:VCALENDAR"""

    if caldav_type != 'caldav':
        print('protocal not supported for now...')
        sys.exit(1)

    if cleanup_all_caldav:
        icx = MyCaldavCalendar(caldav_url, caldav_username, caldav_password, http_proxy)
        print('deleting all caldav events')
        icx.delete_all_events(cal)
        sys.exit(0)

    print('Fetching events from Outlook')
    MyOC = MyOutlookCalendar()
    (myoutlookevents_norec, myoutlookguid_norec) = MyOC.get_my_calendar_event(start, end, False)
    (myoutlookevents_rec, myoutlookguid_rec) = MyOC.get_my_calendar_event(start, end, True)

    # known_guid_events[e.uid]['objects'].append(e)
    # https://docs.microsoft.com/en-us/office/vba/api/Outlook.Items.IncludeRecurrences
    # https://www.add-in-express.com/creating-addins-blog/2011/11/15/outlook-restrict-calendar-items/

    '''
    for event in myoutlookevents:
        print(eDaily meeting - updatevent.name)
    '''

    '''
    print('\n===============\n')
    print('start searching...')
    for event in myoutlookevents_rec:

        vcal = vcal_header + str(event) + vcal_footer
        # print(vcal)
        # print('------')

        oics = Calendar(str(vcal))
        oev = oics.events[0]

        if oev.uid not in myoutlookguid_norec:
            print('recursive event only')
        elif len(myoutlookguid_rec[oev.uid]['objects']) == len(myoutlookguid_norec[oev.uid]['objects']):
            pass
            # print('number match')
        else:
            print("Len do not match: %s %s" % (
                str(len(myoutlookguid_rec[oev.uid]['objects'])), str(len(myoutlookguid_norec[oev.uid]['objects']))))
            print(myoutlookguid_norec[oev.uid]['objects'])
            print(myoutlookguid_rec[oev.uid]['objects'])
            print('---')

    sys.exit(0)
    '''

    icx = MyCaldavCalendar(caldav_url, caldav_username, caldav_password, http_proxy)

    # print('all known calendar')
    # icx.print_named_calendar()

    if verbose:
        print('My iCal calendar name: %s.' % (caldav_name))
    cal = icx.get_named_calendar(caldav_name)

    if not cal:
        cal = icx.create_calendar(caldav_name)

    # Get Caldav Calendar event from period
    cal_known_events = icx.get_all_event(cal, start, end)

    if verbose:
        print('Fetching events from Caldav')
    my_caldav_by_uid = {}

    for event in cal_known_events:
        '''
        print(dir(event))
        obj_dump(event)
        print(event)
        print(event.url)
        print(event.data)
        '''
        c = Calendar(event.data)
        # print(c.events[0])
        my_caldav_by_uid[c.events[0].uid] = event

    if verbose:
        print('Matching Outlook against iCal')

    no_delete = False
    my_outlook_uid = []
    for event in myoutlookevents_norec:
        vcal = vcal_header + str(event) + vcal_footer
        # print(vcal)
        # print('------')

        oics = Calendar(str(vcal))
        oev = oics.events[0]

        if oev.uid not in myoutlookguid_norec:
            ''' Those are comming from copy/paste of events, the uid is in some case the same.'''
            print('skipping bastard events %s' % (str(oev.uid)))
            print(oev)
            continue
        '''   
        if oev.uid in myoutlookguid_rec:
            print('fuck %s' % (str(oev.uid)))
            print(oev)
            sys.exit(0)
            # if length is the same in both, then this is the same object, and it as been already done
            continue
        '''

        my_outlook_uid.append(oev.uid)

        (my_caldav_by_uid, local_no_delete) = copy_ia(oev, vcal, cal, my_caldav_by_uid, vcal_header, vcal_footer)
        if local_no_delete is True:
            no_delete = True
        # print('oics:', oev.name, oev.uid)

    print('Starting REC events')
    for mevent in myoutlookevents_rec:

        mvcal = vcal_header + str(mevent) + vcal_footer
        # print(mvcal)
        # print('------')

        moics = Calendar(str(mvcal))
        moev = moics.events[0]

        if moev.uid in myoutlookguid_norec and len(myoutlookguid_rec[moev.uid]['objects']) == len(
                myoutlookguid_norec[moev.uid]['objects']):
            ''' if length is the same in both, then this is the same object, and it as been already done'''
            continue

        print("Len do not match: %s" % (str(len(myoutlookguid_rec[moev.uid]['objects']))))
        #       , str(len(myoutlookguid_norec[moev.uid]['objects']))))
        if moev.uid in myoutlookevents_norec:
            print(myoutlookguid_norec[moev.uid]['objects'])
        else:
            print(0)
        print(myoutlookguid_rec[moev.uid]['objects'])
        print('---')

        for event in myoutlookguid_rec[moev.uid]['objects']:

            vcal = vcal_header + str(event) + vcal_footer
            # print(vcal)
            # print('------')

            oics = Calendar(str(vcal))
            oev = oics.events[0]

            ''' We have some duplicate uid, where we shouldn't, let's redefine the duplicate'''
            ''' Adding the date to the uid to generate a fake uid, so we do not mix-up the events.'''
            # print(oev.uid)
            old_oev = oev.uid
            oev.uid = str(oev.begin).replace('-', '').replace(':', '').replace('+', '') + '-' + str(oev.uid)
            # print(oev.uid)
            my_outlook_uid.append(oev.uid)
            vcal = vcal.replace(old_oev, oev.uid)

            (my_caldav_by_uid, local_no_delete) = copy_ia(oev, vcal, cal, my_caldav_by_uid, vcal_header, vcal_footer)
            if local_no_delete is True:
                no_delete = True
            # print('oics:', oev.name, oev.uid)

    print('no_delete status: %s' % (str(no_delete)))
    # no_delete = True

    if no_delete:
        print('no deletion will occurs, as some errors have been detected.')

    if verbose:
        print('deleting not found caldav events')
    for caldav_event_uid in my_caldav_by_uid:
        if caldav_event_uid not in my_outlook_uid:
            c = Calendar(my_caldav_by_uid[caldav_event_uid].data)
            cev = c.events[0]

            if caldav_event_uid in myoutlookguid_norec:
                print('Detecting some anomalies in the delete process, disable the deletion NOREC')
                no_delete = True

            if caldav_event_uid in myoutlookguid_rec:
                print('Detecting some anomalies in the delete process, disable the deletion REC')
                no_delete = True

            if no_delete:
                print(
                    "caldav event should have been deleted... '%s':'%s':'%s'" % (cev.name, cev.begin, caldav_event_uid))
            else:
                print('caldav event need to be deleted... %s:%s:%s' % (cev.name, cev.begin, caldav_event_uid))
                my_caldav_by_uid[caldav_event_uid].delete()

    sys.exit(0)
