""" Dolichos

    Synchronize Strava activities with "The Doc"

    Author: Owen Webb (owebb@umich.edu)
    Date: 1/13/2020
"""
from dataclasses import dataclass
from datetime import datetime, date
from string import ascii_uppercase
from collections import defaultdict

import strava
import sheets


SEMESTER_START = date(2020, 1, 6)
SHEET_NAME = 'Owen'


@dataclass
class Activity(object):
    """Keep track of activity-specific data from Strava before being synced
    with The Doc."""
    title: str
    description: str
    distance: float
    is_race: bool
    timestamp: datetime
    identifier: int


def main():
    # init Strava and GoogleSheets API
    s = strava.Strava()
    the_doc = sheets.GoogleSheets('the_doc.identity')

    # maps a day to a list of Activities (see dataclass above)
    date_act_map = defaultdict(list)

    # iter over activities
    for act in s.activity_list():
        # extract vars and make conversions
        act_dt = tstamp_to_dt(act['start_date_local'])

        # skip activities from before the Winter 2020 semester
        if act_dt.date() < SEMESTER_START:
            continue

        # get rest of activty data
        act_title = act['name']
        act_dist = meters_to_miles(act['distance'])
        act_id = act['id']
        act = s.get_detailed_activity(act_id)
        act_desc = act['description']
        act_is_race = act['workout_type'] == 1

        date_act_map[act_dt.date()].append(Activity(act_title,
                                                    act_desc,
                                                    act_dist,
                                                    act_is_race,
                                                    act_dt,
                                                    act_id))

    for act_date in sorted(date_act_map):
        # sort by time began that day
        act_lst = sorted(date_act_map[act_date], key=lambda a: a.timestamp)

        # init datastructures needed to process this day's data
        total_dist = 0
        full_desc = []
        race_day = any(a.is_race for a in act_lst)
        primary_act = None
        max_dist = 0

        # iter over this day's activities and aggregate data appropriately
        for act in act_lst:
            total_dist += act.distance 

            # compile the aggregate description for this day
            full_desc.append(f'{act.timestamp.strftime("%I:%M %p")} \
                               {act.title} - {act.distance}mi')
            if act.description: full_desc[-1] += f': {act.description}'

            # update "primary" activity which will be displayed in the "details"
            # cell of The Doc. Primary activity is simply the longest run. 
            # If it's a race day only a race can be the primary activity. 
            if (act.is_race == race_day) and (act.distance > max_dist):
                max_dist = act.distance
                primary_act = act

        full_desc = "\n\n".join(full_desc)

        # calculate coords of target cell
        row, col = date_to_cell(act_date)
        title_cell = f'{SHEET_NAME}!{col}{row}'
        distance_cell = f'{SHEET_NAME}!{col}{row + 1}'

        # send to doc
        the_doc.set_cell(title_cell, primary_act.title)
        the_doc.set_cell(distance_cell, total_dist)
        the_doc.set_cell_note(title_cell, full_desc)

        print(f'Synchronized {act_date.strftime("%x")} - {primary_act.title}')


def meters_to_miles(meters):
    """Convert meters to miles. Conversion factor = 0.00062137."""
    return round(meters * 0.00062137, 2)


def tstamp_to_dt(timestamp):
    """Convert a Strava timestamp to a Python datetime object."""
    return datetime.strptime(timestamp, "%Y-%m-%dT%H:%M:%SZ")


def date_to_cell(timestamp):
    """Convert a date to the proper cell on The Doc."""
    days_after = (timestamp - SEMESTER_START).days

    row_num = ((days_after // 7) * 2) + 3
    col_num = (days_after % 7) + 2
    col_alpha = ascii_uppercase[col_num]

    return row_num, col_alpha


if __name__ == '__main__':
    main()
