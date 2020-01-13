"""
    Strava API definition

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
        self.base_url = 'https://www.strava.com/'
        self.auth_url = 'oauth/token'
        self.activity_url = 'api/v3/athlete/activities'

        with open('strava_auth.json', 'r') as fin:
            self.auth = json.load(fin)

    def authenticate(self):
        if time.time() < self.auth['expires']:
            return

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
            msg = 'Authentication failure! Status code {r.status_code} \
                    Remember to authenticate manually first time.'
            raise AuthenticationError(msg)

        # update auth dict with new auth information
        new_auth = json.loads(r.text)
        self.auth['access_token'] = new_auth['access_token']
        self.auth['refresh_token'] = new_auth['refresh_token']
        self.auth['expires'] = new_auth['expires_at']

        with open('strava_auth.json', 'w') as fout:
            fout.write(json.dumps(self.auth))

    def get_activity(self):
        self.authenticate()

        auth_header = {'Authorization' : f"Bearer {self.auth['access_token']}"}
        r = requests.get(self.base_url + self.activity_url, headers=auth_header)
        print(json.dumps(json.loads(r.text)[0], indent=4))



if __name__ == '__main__':
    # TESTING
    s = Strava()
    s.get_activity()
