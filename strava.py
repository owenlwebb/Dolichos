""" Strava API definition

    Author: Owen Webb (owebb@umich.edu)
    Date: 1/11/2020
"""
import json
import os
import time

import requests


class StravaError(BaseException):
    pass


class AuthenticationError(StravaError):
    pass


class Strava(object):
    def __init__(self):
        """Set url constants and load authentication data."""
        self.base_url = 'https://www.strava.com/api/v3/'
        self.auth_url = 'oauth/token'
        self.activities_url = 'athlete/activities?per_page={}'
        self.detailed_activity_url = 'activities/'

        with open('strava_auth.json', 'r') as fin:
            self.auth = json.load(fin)

    def authenticate(self):
        """Authenticates the user whose auth data is in strava_auth.json."""
        # refresh the access token if expired
        if time.time() >= self.auth['expires']:
            # build dict with data needed to get new access_token
            auth_data = {
                            'client_id' : self.auth['client_id'],
                            'client_secret' : self.auth['client_secret'],
                            'grant_type' : 'refresh_token',
                            'refresh_token' : self.auth['refresh_token']
                        }

            # make request and error check
            r = requests.post(self.base_url + self.auth_url, data=auth_data)
            if r.status_code != 200:
                msg = (f'Authentication failure! Status code {r.status_code} '
                       f'Remember to authenticate manually first time.')
                raise AuthenticationError(msg)

            # update auth dict with new auth information
            new_auth = json.loads(r.text)
            self.auth['access_token'] = new_auth['access_token']
            self.auth['refresh_token'] = new_auth['refresh_token']
            self.auth['expires'] = new_auth['expires_at']

            with open('strava_auth.json', 'w') as fout:
                fout.write(json.dumps(self.auth))

        return self.auth['access_token']

    def make_request(self, url):
        """Make get request to given url."""
        token = self.authenticate()
        r = requests.get(self.base_url + url, 
                         headers={'Authorization' : f'Bearer {token}'})
        if r.status_code != 200:
            raise StravaError(f'Request to {url} failed. Code: {r.status_code}')
        return json.loads(r.text)

    def activity_list(self, n=30):
        """Returns the n most recent activities. Max 30."""
        return self.make_request(self.activities_url.format(n))

    def get_detailed_activity(self, activity_id):
        """Returns the full, detailed activity corresponding to activity_id."""
        return self.make_request(self.detailed_activity_url + str(activity_id))


if __name__ == '__main__':
    # TESTING
    s = Strava()
    s.get_activities()
