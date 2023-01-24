"""
=======================================
timetable program By Nshuti Alliance ["http:/www.alliancenshuti.herokuapp.com"]

problem solved:
    - to assign randomly preachers more than 50
    classes to preach
    -assign avoid repetition of preacher on a single day
    for more than one class

input file preachersDb.xlsx attributes:
    - name
    - class
    - inconvenient class
output timetable.docx
"""

import random
from xlsxInput import fetch_data
from wordWriter import output_timetable
from openpyxl import Workbook
import os.path as path

classes = [{"className": "y8", "streams": ["y8a", "y8b", "y8c", "y8d", "y8e", "y8g"]},
           {"className": "y9", "streams": ["y9a", "y9b", "y9c", "y9d", "y9e", "y9g"]},
           {"className": "y10", "streams": ["y10a", "y10b", "y10c", "y10d", "y10e"]},
           {"className": "y11", "streams": ["y11a", "y11b", "y11c", "y11d", "y11e"]},
           {"className": "y12", "streams": ["y12ScienceA", "y12ScienceB", "y12ArtsA", "y12ArtsB"]},
           {"className": "y13", "streams": ["y13ScienceA", "y13ScienceB", "y13ArtsA", "y13ArtsB"]}]
days = ["tuesday", "wednesday", "thursday"]

# filter the streams from the class list
streams = []
for i in range(len(classes)):
    for k in range(len(classes[i]["streams"])):
        streams.append(classes[i]["streams"][k])


def Initialise():
    # create an empty data base if it does not exist
    wb = Workbook()
    dest_filename = 'preachersDb.xlsx'
    ws1 = wb.active
    ws1.title = "preachers"
    wb.save(filename=dest_filename)


def Shuffle():
    # sort fo those with more inconvenient classes
    global record, preachers
    preachers = fetch_data()
    for i in range(len(preachers)):
        for k in range(len(preachers)):
            if len(preachers[i]["inconvenient"]) > len(preachers[k]["inconvenient"]):
                temp = preachers[i]
                preachers[i] = preachers[k]
                preachers[k] = temp

    # loop and place preachers on specific days
    table = []
    daily = {"tuesday": [], "wednesday": [], "thursday": []}  # to store daily preachers for avoiding repetition
    # loops for every class and stream to assign randomly preachers
    # to remove repetition you must repeat through the list again

    # loop for each and every class
    for i in range(len(classes)):
        # for each stream in a class
        for k in range(len(classes[i]["streams"])):
            record = {"classname": classes[i]["streams"][k]}
            # loop for each day in a stream
            for j in range(len(days)):

                """
                the code below is jumping some classes empty 
                for a specific preacher .
                therefore we need to loop and assign them  again randomly
                """
                while True:
                    randomDay = random.randint(0, len(days) - 1)
                    if days[randomDay] not in record.keys():
                        break

                while True:
                    randomPreacher = random.randint(0, len(preachers) - 1)
                    # the class is not in the inconvenient list and preacher is not preaching that day
                    if record["classname"] not in preachers[randomPreacher]["inconvenient"] and \
                            preachers[randomPreacher]["name"] not in daily[days[randomDay]]:
                        break

                # append the preacher to the list of people preaching that day
                daily[days[randomDay]].append(preachers[randomPreacher]["name"])
                # assign a random day to the preacher in that stream
                record[days[randomDay]] = preachers[randomPreacher]["name"]
            # add the record to the table list
            table.append(record)
    return table


if __name__ == "__main__":
    if not path.exists("preachersDb.xlsx"):
        Initialise()
    output_timetable(Shuffle(), streams)
    print("executed")