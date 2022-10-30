#!/usr/bin/python
# ----------------------------------------------------------------------------------------------
# syncCalender.py
#
# This script synchronizes kimai2 with outlook calender 
#
# Parameters:
#
# -c : config file with credintials and API url of the kimai2 server
# -b : start date in UTC (optional) 
# -e : end date in UTC (optional)
# -h : displays this help
# 
# ----------------------------------------------------------------------------------------------


from ctypes import sizeof
import os
import sys
import json
import getopt
import logging
import datetime
import requests
import numpy as nm
import pandas as pd
import win32com.client


# This will get the name of this file
script_name = os.path.basename(__file__)
default_loglevel = 'info'

# Output formatting constants
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'    


api_config = json.loads('''
{
    "url":"http://kimai2/api",
    "token":"your-kimai-api-token",
    "user":"user",
    "user id":2,
    "work id":4,
    "work activity":1,
    "off id":9,
    "off activity":15
}
'''
)

config_file ='config.json'
begin_date = datetime.date.today
end_date = datetime.date.today() + datetime.timedelta(days=1)

def progressBar(iterable, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iterable    - Required  : iterable object (Iterable)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    total = len(iterable)
    # Progress Bar Printing Function
    def printProgressBar (iteration):
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + '-' * (length - filledLength)
        print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Initial Call
    printProgressBar(0)
    # Update Progress Bar
    for i, item in enumerate(iterable):
        yield item
        printProgressBar(i + 1)
    # Print New Line on Complete
    print()

def readConfigData(fname:str):
    """Load configuration from the config file.

    :param fname: The filename of the file containing the configuration.
    :return: The public key contained in the file.
    """
    try:
        with open(fname, "r") as file:
            api_config = json.loads(file.read())
            return api_config
    except:  # Yes, we  do really want this on _every_ exception that might occur.
        print("Couldn't load configuration file.")
        sys.exit(2)
    return None

##
# @brief Help document for this script. 
#
help_str = f''' 
    {bcolors.BOLD}{script_name} [-c | --config] <config file json> [-b | --begin] <date> [-e | --end] <date> [--loglevel level]{bcolors.ENDC}
    
    synchronizes outlook calender with kimai server
    
    examples:
        synchronize all events from Oct 10th 2022 08:00 till Oct 15th 2022 18:00
            {script_name} -c kimai.api.json -b 2022-10-12T08:00:00 -e 2022-10-15T18:00:00
            
        synchronize all events from Oct 10th 2022 08:00 till now
            {script_name} -c kimai.api.json -b 2022-10-12T08:00:00 
        
        synchronize all events from today
            {script_name} -c kimai.api.json  

    -c configuration file
    --config file
        file path to the configuration file formatted in JSON, 
        which holds information about kimai server and API credentials
        
        example config.json:
        {{
            "url":"http://kimai2-server.com/api",
            "token":"kzTQ98BG_H8!~e3s",
            "user":"username",
            "user id":2,
            "work id":4,
            "work activity":1,
            "off id":9,
            "off activity":14
        }}

        -url        - url of the Kimai API
        -token      - API token generated for the user on the kimai server
        -user       - username on the kimai server
        -user id    - user id on the kimai server
        -work id    - id of the project to which the regular meetings will be assigned
        -work activity - activity id to which the regular meetings will be assigned
        -off id     - id of the project to which the private meetings will be assigned 
        -off activity - activity id to which the private meetings will be assigned
        
    -b date in UTC
    --begin  date in UTC
        If specified the events starting with the given date will be synchronized
        The date is formatted in UTC.
        example: 
            2022-10-15T13:45:30
    
    -b date in UTC
    --end date in UTC
        If specified the events till the date will be synchronized
        The date is formatted in UTC.
        example: 
            2022-10-15T13:45:30
         
    --loglevel critical | error | warning | info | debug | notset
        Control the verbosity of the script by setting the log level.  Critical
        is the least verbose and notset is the most verbose.
        
        The default loglevel is {default_loglevel}.
        
        These values correspond directly to the python logging module levels. 
        (i.e. https://docs.python.org/3/howto/logging.html#logging-levels)
   
    -h 
    --help 
        print this message
    
'''

def print_help():
    print(help_str, file=sys.stderr)

def help():
    """prints the usage information
    """
    print(bcolors.BOLD +'usage: python ' + script_name +' [-c | --config] <file> [-b | --begin] <date> [-e | --end] <date> [-h | --help]'+bcolors.ENDC)
    print('')
    print('-c, --config\tconfiguration file with API credentials.')
    print('-b, --begin\tbegin date in UTC')
    print('-e, --end\tend date in UTC')
    print('-h, --help\tprints detailed help message')
    print('')
    print('Synchronizes oulook calender timesheets to kimai2')
    print('')

class RequiredOptions:
    '''Just something to keep track of required options'''
    
    def __init__(self, options=[]):
        self.required_options = options
        
    def add(self, option):
        if option not in self.required_options:
            self.required_options.append(option)
            
    def resolve(self, option):
        if option in self.required_options:
            self.required_options.remove(option)
            
    def optionsResolved(self):
        if len(self.required_options):
            return False
        else:
            return True
        

def kimaiGet(api:str, payload:str):
    global api_config
    headers={'Content-Type':'application/json', 
             'Accept': 'application/json',
             'X-AUTH-TOKEN': api_config['token'],
             'X-AUTH-USER': api_config['user']}
    r = requests.get(api_config['url']+api, headers=headers, params=payload)
    return r.json()

def kimaiPost(api:str,data:str):
    global api_config
    headers={'Content-Type':'application/json', 
             'Accept': 'application/json',
             'X-AUTH-TOKEN': api_config['token'],
             'X-AUTH-USER': api_config['user']}
    r = requests.post(api_config['url']+api, headers=headers, data=data)
    return r.status_code

def kimaiPatch(api:str, data:str, id:str):
    global api_config
    headers={'Content-Type':'application/json', 
             'Accept': 'application/json',
             'X-AUTH-TOKEN': api_config['token'],
             'X-AUTH-USER': api_config['user']}
    r = requests.patch(api_config['url']+api+f'/{id}', headers=headers, data=data)
    return r.status_code

def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    #restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

def get_appointments(calendar,subject_kw = None,exclude_subject_kw = None, body_kw = None):
    if subject_kw == None:
        appointments = [app for app in calendar]    
    else:
        appointments = [app for app in calendar if subject_kw in app.subject]
    if exclude_subject_kw != None:
        appointments = [app for app in appointments if exclude_subject_kw not in app.subject]

    # exclude appointments marked as free
    appointments = [app for app in appointments if app.busystatus >0 ]
    
    df = '['
    for app in appointments:
        if '}' in df:
            df+= ","
        bd = f'{app.start}'
        #TODO - FIX time saving zones 
        # currently hardcoded for my time zone: +2000
        bd = bd.replace('+00:00','+0200').replace(' ','T').replace(' ','')
        ed = f'{app.end}'
        ed = ed.replace('+00:00','+0200').replace(' ','T').replace(' ','')
        df += f'''{{
            "begin": "{bd}",
            "end": "{ed}",
            "project": {api_config['work id'] if app.sensitivity != 2 else api_config['off id']},
            "activity": {api_config['work activity'] if app.sensitivity != 2 else api_config['off activity']},
            "description": "{app.subject}",
            "fixedRate":0,
            "hourlyRate":0,
            "user": {api_config['user id']},
            "exported": false,
            "billable": true,
            "tags":""
            }}'''

    df += ']'
    
    return df

def executeIt(begin:datetime,end:datetime):
    print(f'Collecting outlook appointments between {begin.strftime("%Y-%m-%d")} and {end.strftime("%Y-%m-%d")}')
    cal = get_calendar(begin, end)
    # exclude outlook appointments called "Mittagspause" 
    appointments = get_appointments(cal, subject_kw = None, exclude_subject_kw = 'Mittagspause')
    apps = json.loads(appointments)
    if len(apps) > 0:
        print(f'Collecting kimai appointments between {begin.strftime("%Y-%m-%d")} and {end.strftime("%Y-%m-%d")}')
        # get timesheets within the given time range from kimai2
        payload = {'begin': f'{begin.strftime("%Y-%m-%d")}T00:00:01', 'end': f'{end.strftime("%Y-%m-%d")}T23:59:59'}
        timesheets = kimaiGet('/timesheets',payload)
        for item in progressBar(apps, prefix = 'Progress:', suffix = 'Complete', length = 50):
            found = False
            item_id = 0
            for timesheet in timesheets:
                if(timesheet['description'] == item['description'] and timesheet['begin'] == item['begin']):
                    # appointement not exist 
                    found = True
                    item_id = timesheet['id']
                    break
            if(found == False):
                # apointment does not exist - create a new appointment on the kimai server
                payload = json.dumps(item, separators=(',', ':'))
                res = kimaiPost('/timesheets',payload)
                if(res != 200): 
                    print(f'''{bcolors.FAIL}Error code {bcolors.BOLD}{res}{bcolors.ENDC}{bcolors.FAIL} received while creating {item['description']} {bcolors.ENDC} begin:{item['begin']}''')
                    sys.exit(-2)
            else:
                # update existing appointment
                res = kimaiPatch(f'/timesheets',json.dumps(item, separators=(',', ':')),item_id)
                if(res != 200): 
                    print(f'''{bcolors.FAIL}Error code {bcolors.BOLD}{res}{bcolors.ENDC}{bcolors.FAIL} received while updating {item['description']} begin:{item['begin']} {bcolors.ENDC}''')
                    sys.exit(-2)
        
    print("outlook to kimai sync done")


def main(argv):
    logging.getLogger().setLevel(default_loglevel.upper())
    try:
        opts, args = getopt.getopt(argv,"hc:b:e:",["help","config=","begin=","end="])
    except getopt.GetoptError as e:
        print_help() #help()
        logging.exception(e)
        sys.exit(2)
    
    cfg_file =""
    b_date = ""
    e_date = ""
    global begin_date
    global end_date
    global api_config
    
    required_options = RequiredOptions(['begin' ])

    for opt, arg in opts:
        if opt == '-h':
            print_help()
            sys.exit(0)
        elif opt in ("-c", "--config"):
            cfg_file = arg
            
        elif opt in ("-b", "--begin"):
            b_date = arg
            begin_date = datetime.date.fromisoformat(b_date)
            required_options.resolve('begin')
        elif opt in ("-e", "--end"):
            end_date = datetime.date.fromisoformat(arg)
        else:
            help()

    if cfg_file != "":
        api_config = readConfigData(cfg_file)
    else:
        print("no config file provided... using default settings")
    
    #if begin_date > end_date and end_date != None:
    #    print(bcolors.FAIL +'cannot proceed, end date cannot be before begin date.'+bcolors.ENDC)
    #    print_help()
    #    sys.exit(1)

    # Verify that all of the required options have been specified
    #
    if not required_options.optionsResolved():
        print(bcolors.FAIL +'cannot proceed, some parameteres are missing'+bcolors.ENDC)
        help()
        sys.exit(1)

    executeIt(begin_date, end_date)
   
   

if __name__ == "__main__":
   main(sys.argv[1:])
