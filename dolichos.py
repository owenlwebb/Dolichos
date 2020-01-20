""" Dolichos

    Synchronize Strava activities with "The Doc"

    Author: Owen Webb (owebb@umich.edu)
    Date: 1/13/2020
"""
from dataclasses import dataclass
from datetime import datetime
from string import ascii_uppercase

import strava
import sheets


def main():
    # init Strava and GoogleSheets API
    s = strava.Strava()
    the_doc = sheets.GoogleSheets('the_doc.identity')

    # iter over activities and upload each
    print('New activity: ', end='')
    for act in s.get_activity_list():
        # set vars and make conversions
        act_date = tstamp_to_dt(act['start_date'])
        act_name = act['name']
        act_dist = meters_to_miles(act['distance'])
        act_id = act['id']

        # skip activities from before the Winter 2020 semester
        if act_date < datetime(2020, 1, 6):
            continue

        # progress
        print(f'{act_name} - {act_date}...', end='')

        # get activity decription from detailed activity data
        act = s.get_detailed_activity(act_id)
        act_desc = act['description']

        # calculate coords of target cell
        row, col = date_to_cell(act_date)
        details_cell = f'Owen!{col}{row}'
        distance_cell = f'Owen!{col}{row + 1}'

        # send to doc
        details = f'{act_name}: {act_desc}' if act_desc else act_name
        the_doc.set_cell(details_cell, details)
        the_doc.set_cell(distance_cell, act_dist)

        print('Uploaded!')


def meters_to_miles(meters):
    """Convert meters to miles. Conversion factor = 0.00062137."""
    return round(meters * 0.00062137, 2)


def tstamp_to_dt(timestamp):
    """Convert a Strava timestamp to a Python datetime object."""
    return datetime.strptime(timestamp, "%Y-%m-%dT%H:%M:%SZ")


def date_to_cell(timestamp):
    """Convert a date to the proper cell on The Doc."""
    start_date = datetime(2020, 1, 6)
    days_after = (timestamp - start_date).days

    row_num = ((days_after // 7) * 2) + 3
    col_num = (days_after % 7) + 2
    col_alpha = ascii_uppercase[col_num]

    return row_num, col_alpha


if __name__ == '__main__':
    main()
