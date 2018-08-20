import sys
from datetime import datetime
from bs4 import BeautifulSoup
import caldav
from caldav.elements import dav
import requests
from requests.auth import HTTPBasicAuth


class MyCaldavCalendar(object):
    propfind_principal = (
        u'''<?xml version="1.0" encoding="utf-8"?><propfind xmlns='DAV:'>'''
        u'''<prop><current-user-principal/></prop></propfind>'''
    )
    propfind_calendar_home_set = (
        u'''<?xml version="1.0" encoding="utf-8"?><propfind xmlns='DAV:' '''
        u'''xmlns:cd='urn:ietf:params:xml:ns:caldav'><prop>'''
        u'''<cd:calendar-home-set/></prop></propfind>'''
    )

    def __init__(self, url, username, password, http_proxy):
        self.url = url
        self.username = username
        self.password = password

        self.proxyDict = {}

        if http_proxy is not None:
            self.http_proxy = http_proxy
            self.proxyDict = {
                "http": self.http_proxy,
                "https": self.http_proxy
            }

        self.discover()
        self.get_calendars()

    # discover: connect to caldav using the provided credentials and discover
    # 1. The principal URL
    # 2  The calendar home URL
    #
    # These URL's vary from user to user
    # once doscivered, these  can then be used to manage calendars

    def discover(self):
        # Build and dispatch a request to discover the principal us for the
        # given credentials
        headers = {
            'Depth': '1',
        }
        auth = HTTPBasicAuth(self.username, self.password)
        principal_response = requests.request(
            'PROPFIND',
            self.url,
            auth=auth,
            headers=headers,
            data=self.propfind_principal.encode('utf-8'),
            proxies=self.proxyDict
        )
        if principal_response.status_code != 207:
            print('Failed to retrieve Principal: ',
                  principal_response.status_code)
            exit(-1)
        # Parse the resulting XML response
        print('principal:\n', principal_response.text, '\n')

        soup = BeautifulSoup(principal_response.text, 'lxml')

        try:
            self.principal_path = soup.find(
                'current-user-principal'
            ).find('href').get_text()
            discovery_url = self.url + self.principal_path
        except Exception as e:
            print(soup)
            print(str(e))
            discovery_url = self.url + self.username + '/'

        # Next use the discovery URL to get more detailed properties - such as
        # the calendar-home-set
        home_set_response = requests.request(
            'PROPFIND',
            discovery_url,
            auth=auth,
            headers=headers,
            data=self.propfind_calendar_home_set.encode('utf-8'),
            proxies=self.proxyDict
        )
        if home_set_response.status_code != 207:
            print('Failed to retrieve calendar-home-set',
                  home_set_response.status_code)
            exit(-1)
        # And then extract the calendar-home-set URL
        print('home_set_response:\n', home_set_response.text, '\n')
        soup = BeautifulSoup(home_set_response.text, 'lxml')
        print(soup)

        try:
            self.calendar_home_set_url = soup.find(
                'href',
                attrs={'xmlns': 'DAV:'}
            ).get_text()
        except Exception as e:
            print(e)
            self.calendar_home_set_url = self.url + self.username + '/Calendar/'

    # get_calendars
    # Having discovered the calendar-home-set url
    # we can create a local object to control calendars (thin wrapper around
    # CALDAV library)
    def get_calendars(self):
        self.caldav = caldav.DAVClient(self.calendar_home_set_url,
                                       username=self.username,
                                       password=self.password,
                                       proxy=self.http_proxy)
        self.principal = self.caldav.principal()
        self.calendars = self.principal.calendars()

    def print_named_calendar(self):

        if len(self.calendars) > 0:
            for calendar in self.calendars:
                properties = calendar.get_properties([dav.DisplayName(), ])
                display_name = properties['{DAV:}displayname']
                print('calendar: ', display_name)

    def get_named_calendar(self, name):

        if len(self.calendars) > 0:
            for calendar in self.calendars:
                properties = calendar.get_properties([dav.DisplayName(), ])
                display_name = properties['{DAV:}displayname']
                if display_name == name:
                    return calendar
        return None

    def get_all_event(self, calendar, start, stop):

        start_split = start.split('/')
        stop_split = stop.split('/')
        # print("Looking for events in %s to %s" % (start, stop))
        results = calendar.date_search(
            datetime(int(start_split[2]), int(start_split[1]), int(start_split[0])),
            datetime(int(stop_split[2]), int(stop_split[1]), int(stop_split[0]), 23, 59, 59))

        # print(results[0])
        # print(calendar.event_by_url(results[0]))

        return results

    def create_calendar(self, name):
        return self.principal.make_calendar(name=name)

    def delete_all_events(self, calendar):
        for event in calendar.events():
            event.delete()
        return True

    def create_events_from_ical(self, ical):
        # to do
        pass

    def create_simple_timed_event(self, start_datetime, end_datetime, summary,
                                  description):
        # to do
        pass

    def create_simple_dated_event(self, start_datetime, end_datetime, summary,
                                  description):
        # to do
        pass


if __name__ == '__main__':
    sys.exit(0)
