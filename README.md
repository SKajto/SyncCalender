# SyncOutlook2Kimai

synchronizes [kimai2](https://www.kimai.org) with outlook calender.
Works only on windows

## Requires
- MS Outlook installed on the host
- Python
- KIMAI2 API Token ([see here](https://www.kimai.org/documentation/rest-api.html) for how to create an API token)

## Usage
```bash
python syncCalender.py [-c | --config] <config file json> [-b | --begin] <date> [-e | --end] <date> 
```

examples: 

    synchronize all events from Oct 10th 2022 till Oct 15th 2022  
    syncCalender.py -c config.json -b 2022-10-12 -e 2022-10-16

    synchronize all events from Oct 10th 2022 08:00 till now  
    syncCalender.py -c config.json -b 2022-10-12 
  
    synchronize all events from today   
    syncCalender.py -c config.json   

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
        2022-10-15  

-b date in UTC  
--end date in UTC  
    If specified the events till the date will be synchronized  
    The date is formatted in UTC.  
    example: 
        2022-10-30  

-h  
--help   
    print this message  
