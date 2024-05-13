#!/usr/bin/env python3

import csv
import json
import logging
from uuid import uuid4
import re
from thefuzz import fuzz
from datetime import datetime,timedelta
import xlsxwriter
import numpy as np
import matplotlib.colors as mcolors

INPUT_FILE = '/Users/jpohl/Downloads/events_2024-05-13.csv'
LOG_LEVEL = 'DEBUG'

FUZZ_THRESHOLD = 90

REQUIRED_MATCHES = [
    'Event Type',
    'Game System',
    'Rules Edition',
    'Group',
    'Title',
    'Duration',
    'Minimum Players',
    'Maximum Players',
    'Age Required',
    'Experience Required',
    'Materials Required',
    'Tournament?',
    'Round Number',
    'Total Rounds',
    'Minimum Play Time',
    'Attendee Registration?',
    'Cost $'
]
SANITIZE_HEADERS = ['Title','Game System','Rules Edition','Group','Short Description','Long Description','Materials Required Details','Website','Email']
PARSED_EVENTS = 'parsed_events.csv'
XLSX_OUTPUT = 'parsed_events.xlsx'
DICTIONARY = 'dictionary.json'
MISFIT_HEADERS = [ 'Group', 'Title']
MISFIT_EVENTS = {
    'BLD - Blood on the Clocktower' : re.compile('Blood on the clocktower',re.IGNORECASE),
    'ESC - Escape Room' : re.compile('escape room',re.IGNORECASE),
    'FIR - First Exposure' : re.compile('first exposure',re.IGNORECASE),
    'KOS - Kosmos Family Table' : re.compile('kosmos family table',re.IGNORECASE),
    'LAS - Laser Tag' : re.compile('laser tag',re.IGNORECASE),
    'PUB - Pub Event' : re.compile('pub night|pedal & drink',re.IGNORECASE),
    'AUC - Gen Con Auction' : re.compile('Gen Con Auction',re.IGNORECASE),
    'MEG - MegaGame' : re.compile('Megagame', re.IGNORECASE),
    'WRI - Gen Con Writers Symposium' : re.compile('Gen Con Writers Symposium',re.IGNORECASE),
    'LIB - Gen Con Games Library' : re.compile('Gen Con Games Library',re.IGNORECASE)

}
NON_ALPHA_NUMERIC = re.compile(r'[^a-zA-Z0-9]')
DATE_FORMAT = "%m/%d/%Y %I:%M %p"
TITLE_PATTERN = re.compile(r'^(The|An|A)\s')
def main():
    # Read dictionary
    with open(INPUT_FILE,'r',errors="ignore") as f:
        csvDict = [
            dict(row.items())
            for row in csv.DictReader(f)
        ]
    # Read existing config DB
    try:
        with open(DICTIONARY,'r',errors="ignore") as f:
            eventDatabase = json.load(f)
            convert_datetime_strings(eventDatabase)
            eventDatabase['Events'] = {}
            eventDatabase['Time Blocks'] = set()
            eventDatabase['Estimate Confidence'] = {}
    except Exception:
        eventDatabase = {m: {} for m in REQUIRED_MATCHES}
        eventDatabase['Events'] = {}
        eventDatabase['Time Blocks'] = set()
        eventDatabase['Estimate Confidence'] = {}
    parseList(csvDict,eventDatabase)
    with open (DICTIONARY,'w') as f:
        f.write(json.dumps(eventDatabase,default=set_default,indent=2))
    #log.debug(json.dumps(csvDict,default=set_default,indent=2))
    cookedData = cookData(eventDatabase)
    writeResults(cookedData)
    write_excel(XLSX_OUTPUT,cookedData)
    # Create time table
    # When printing, get the info needed for REQUIRED_MATCHES from the database by UUID
    # Write parsed events
    # Write dictionary

def write_excel(filename,data):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()


    headers = []
    # Write header row
    header_format = workbook.add_format({'bold': True})
    for col_num, header in enumerate(data[0]):
        worksheet.write(0, col_num, header, header_format)
        headers.append({'header':str(header)})

    # Get cell lengths
    cell_lengths = set()
    for row_num,row_data in enumerate(data[1:],start=1):
        for col_num,cell_value in enumerate(row_data[18:],start=19):
            if len(cell_value) > 0:
                cell_lengths.add(len('\n'.join(cell_value)))
            else:
                cell_lengths.add(0)

    colors = generate_color_dict(sorted(list(cell_lengths)))
    log.debug(json.dumps(colors,indent=2))

    
    # Write data to the worksheet
    for row_num, row_data in enumerate(data[1:], start=1):
        for col_num, cell_value in enumerate(row_data):
            if isinstance(cell_value,list):
                if len(cell_value) > 0:
                    clen = len(cell_value)
                    v = '\n'.join(cell_value)
                    val = f'{clen}\n{v}'
                    vlen = len(v)
                    color = colors[vlen]
                    log.debug(f'{v}: {vlen}: {color}')
                    cell_format = workbook.add_format({'bg_color':color,'text_wrap':True})
                    worksheet.write(row_num,col_num,val,cell_format)
                else:
                    worksheet.write_blank(row_num,col_num,None)
                
            if isinstance(cell_value, (int, float)):
                worksheet.write_number(row_num, col_num, cell_value)
            elif isinstance(cell_value, str):
                if cell_value.isdigit():
                    worksheet.write_number(row_num, col_num, int(cell_value))
                elif is_numeric_string(cell_value):
                    worksheet.write_number(row_num, col_num, float(cell_value))
                else:
                    worksheet.write(row_num, col_num, cell_value)

    num_rows = len(data)-1
    num_cols = len(data[0])-1
    '''
 


    worksheet.conditional_format(1,18,num_rows,num_cols, {
        'type': '3_color_scale',
        'min_color': "#FFFFFF",  # white color for min
        'mid_color': "#FFFFFF",  # white color for mid
        'max_color': "#FFFFFF",  # white color for max
        'min_type': 'num',
        'mid_type': 'num',
        'max_type': 'num',
        'min_value': 0,
        'mid_value': (len(str(min(cell_lengths))) + len(str(max(cell_lengths)))) / 2,  # Calculate mid-value
        'max_value': max(map(len, map(str, cell_lengths))),
        'format_min': format1,
        'format_mid': format2,
        'format_max': format3

    })
    '''

    worksheet.add_table(0,0,num_rows,num_cols,{'columns':headers})
    worksheet.autofit()
    

    # Freeze the first 5 columns
    worksheet.freeze_panes(0, 5)

    #worksheet.conditional_format(1, 18, num_rows, num_cols - 1,{'type': '3_color_scale'})

    workbook.close()

def generate_color_dict(sorted_list):
    num_items = len(sorted_list)
    color_dict = {}

    # Define RGB values for the colors
    start_color = (255, 0, 0)  # Red
    middle_color = (255, 255, 0)  # Yellow
    end_color = (0, 255, 0)  # Green

    # Convert RGB values to range 0-1
    start_color_norm = tuple(x / 255.0 for x in start_color)
    middle_color_norm = tuple(x / 255.0 for x in middle_color)
    end_color_norm = tuple(x / 255.0 for x in end_color)

    # Calculate the index for the middle color
    middle_index = num_items // 2

    # Create colormap segments
    red_segment = np.linspace(start_color_norm[0], middle_color_norm[0], middle_index)
    green_segment = np.linspace(start_color_norm[1], middle_color_norm[1], middle_index)
    blue_segment = np.linspace(start_color_norm[2], middle_color_norm[2], middle_index)
    red_segment2 = np.linspace(middle_color_norm[0], end_color_norm[0], num_items - middle_index)
    green_segment2 = np.linspace(middle_color_norm[1], end_color_norm[1], num_items - middle_index)
    blue_segment2 = np.linspace(middle_color_norm[2], end_color_norm[2], num_items - middle_index)

    # Interpolate colors along the colormap segments
    for i, item in enumerate(sorted_list):
        if i < middle_index:
            color = mcolors.rgb2hex((red_segment[i], green_segment[i], blue_segment[i]))
        else:
            color = mcolors.rgb2hex((red_segment2[i - middle_index], green_segment2[i - middle_index], blue_segment2[i - middle_index]))
        color_dict[item] = color

    return color_dict

def generate_html_colors(d):
    n = len(d)
    colors = {}
    step_size = 255 // (n-1)
    for i,v in enumerate(d):
        red = 255 - (i * step_size)
        green = i * step_size
        blue = 0
        colors[v] = "#{:02X}{:02X}{:02X}".format(red, green, blue)
    return colors 

def is_numeric_string(s):
    """Check if a string represents a numeric value."""
    return re.match(r'^-?\d*\.?\d+$', s) is not None

def writeResults(cookedData):
    with open(PARSED_EVENTS,mode='w',newline='') as f:
        writer = csv.writer(f)
        writer.writerows(cookedData)
    return

def sort_by_secondary_values(entry):
   # log.debug(entry)
    s = entry[1]['Earliest Block']
    e = entry[1]['Block Duration']
    t = entry[1]['Event Type']
    y = entry[1]['Game System']
    e = entry[1]['Rules Edition']
    g = entry[1]['Group']
    l = entry[1]['Title']
    d = entry[1]['Short Description']
    return (s,e,t,y,e,g,l,d)

def cookData(db):
    header_row = REQUIRED_MATCHES
    header_row.extend(['Block Duration'])
    time_blocks = sorted(db['Time Blocks'])
    eventDatabase = db['Events']
    data_rows = []
    # Iterate over all events
    for uuid,event in sorted(eventDatabase.items(), key=sort_by_secondary_values):
        #log.debug(json.dumps(event,default=set_default,indent=2))
        this_event = []
        this_event_time_blocks = {}
        # Iterate over sessions in event sorted by start block, end block, and then start time/end time within the block.        
        for gameid,session in sorted(event['Events'].items(),key=lambda x: (x[1]['Start Block'],x[1]['End Block'], x[1]['Start Date & Time'],x[1]['End Date & Time'],x[0])):
        #for gameid,session in event['Events'].items():
        # Build lists of sessions running in a given hour
            for tb in time_blocks:
                if tb >= session['Start Block'] and tb < session['End Block']:
                    if tb not in this_event_time_blocks:
                        this_event_time_blocks[tb] = []
                    this_event_time_blocks[tb].append(f'{gameid}: {session["Start Date & Time"]} to {session["End Date & Time"]}')
        for h in REQUIRED_MATCHES:
            this_event.append(event[h])
        # Append the list inside the list, append an empty list if there is nothing.
        for tb in time_blocks:
            if tb in this_event_time_blocks:
                this_event.append(this_event_time_blocks[tb])
            else:
                this_event.append([])
        data_rows.append(this_event)
    header_row.extend(time_blocks)
    return [header_row, *data_rows]

def convert_datetime_strings(data):
    for key, value in data.items():
        if isinstance(value, dict):
            convert_datetime_strings(value)
        elif isinstance(value, list):
            for i in range(len(value)):
                if isinstance(value[i], (dict, list)):
                    convert_datetime_strings(value[i])
        elif isinstance(value, str):
            try:
                data[key] = datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
            except ValueError:
                pass
        elif isinstance(value,list):
            if list(set(value)) == value:
                data[key] = list(set(value))
            

def set_default(obj):
    if isinstance(obj, set):
        return sorted(list(obj))
    if isinstance(obj,datetime):
        return obj.strftime("%Y-%m-%d %H:%M")
    raise TypeError


'''
MUST MATCH
Title, Event Type, Duration, Group, Game System, Rules Edition must match
Min Players
Max Players
Age
Experience
Materials
Duration
Tournament
Round Number
Total Rounds
Minimum Play Time
Attendee Registration
Cost

- Get longest short tile
- Get longest long title
- Get longest Materials Required Details

Store per event:
- Everything
- Calculate start and end block and block duration
- Start Time
- Start Block
- End Time
- End Block
- GM Names
- Website
- Email



    event: {
        'Title' : 'foo',
        'Earliest Block' : '2024-08-01 10:00:00',
        'Duration' : 1
    }
'''
    
def parseList(d,database=None):
    ''' use a uuid for the key (random)? 
        event name
    '''

    # Iterate over the list sorted by the title
    for row in sorted(d, key=lambda x: (x['Event Type'], x['Game System'],x['Rules Edition'],x['Group'],x['Title'], x['Short Description'])):
        for k,v in row.items():
            if k in SANITIZE_HEADERS:
                #log.debug(f'{k}: {v}')
                if k == 'Title':
                    v = TITLE_PATTERN.sub('', v)
                v = sanitizeVal(v)
                row[k] = v
        # Check for misfit events:
        if row['Event Type'] == 'ZED - Isle of Misfit Events':
            checkForCustomMisfit(row)
        row['uuid'] = {}
        row['uuidstr'] = ''
        for m in REQUIRED_MATCHES:
          #  log.debug(f'Checking {row} against {m}')
            thisUUID = findUUID(row[m],database[m],database['Estimate Confidence'])
            row['uuid'][m] = thisUUID
            row['uuidstr'] += thisUUID
        addOrUpdateEventData(row,database)

    return database

def checkForCustomMisfit(row):
    for k,v in MISFIT_EVENTS.items():
        for h in MISFIT_HEADERS:
            if v.search(row[h]):
                row['Event Type'] = k
                return
    # No match
    return


def addOrUpdateEventData(row,database):
    db = database['Events']
    tb = database['Time Blocks']
    # Store block start
    # block end
    # block duration
    # Add or update earliest block start
    uuidstr = row['uuidstr']
    # Don't get required match (top level) data now.  get it from the databases by UUID later

    startTime = datetime.strptime(row['Start Date & Time'], DATE_FORMAT)
    endTime = datetime.strptime(row['End Date & Time'], DATE_FORMAT)
    startBlock = startTime.replace(minute=0,second=0)
    tmpEndBlock = endTime
    if tmpEndBlock.minute > 0 or tmpEndBlock.second > 0:
        tmpEndBlock += timedelta(hours=1)
    endBlock = tmpEndBlock.replace(minute=0,second=0)
    blockDuration = (endBlock - startBlock).total_seconds() / 3600
    tb.add(startBlock)
    tb.add(endBlock)
    # Check if startBlock is earlier than already recorded  
    
    if uuidstr not in db:
        db[uuidstr] = {
            'Earliest Block' : startBlock,
            'Block Duration' : blockDuration,        
            'Events' : {}
        }
        eventData = db[uuidstr]
        # Copy the gathered UUIDs on initialization, will be used
        #eventData['uuid'] = row['uuid']
        # Copy some basic data from the first event, useful for creating CSV file
        for h in SANITIZE_HEADERS:
            if h not in db[uuidstr] and h not in REQUIRED_MATCHES:
                eventData[h] = row[h]
    else:
        eventData = db[uuidstr]        
        if startBlock < eventData['Earliest Block']:
            eventData['Earliest Block'] = startBlock

    for k,v in row['uuid'].items():
        eventData[k] = database[k][v]['canonical']

    # Populate the event
    gameID = row['Game ID']
    eventData['Events'][gameID] = {
        'Start Block' : startBlock,
        'End Block' : endBlock,
        'Block Duration' : blockDuration,
        'Start Date & Time' : startTime,
        'End Date & Time' : endTime,
        'Duration' : float(row['Duration'])
    }
    return
    
def findUUID(search_name,name_dict,est_dict):
    # Check for literal match
    #log.debug(f'Looking for match for {m}')

    def sanitizeForDiff(m):
        return sanitizeVal(NON_ALPHA_NUMERIC.sub(' ',m)).lower()

    def areTheyTheSame(search_name,sanitized_name,name_dict,found_name,fuzz_ratio,estimate_dict):
        canonical = name_dict['canonical']
        print( '----------------')
        print(f'SEEKING   : "{search_name}"')
        print(f'SEARCH    : "{sanitized_name}"')
        print(f'FOUND     : "{found_name}"')
        print(f'CANONICAL : "{canonical}"')
        print(f'CONFIDENCE: {fuzz_ratio}')
        print(f'1. Same, update canonical to "{search_name}"')
        print(f'2. Same, keep canonical name "{canonical}')
        user_input = input('3. Not the Same\n')
        while user_input not in ['1','2','3']:
            user_input = input('Sorry, that is an invalid option. Please try again\n')
        # It's the same, but the new entry is more correct than the previous
        
        if fuzz_ratio not in estimate_dict:
            estimate_dict[fuzz_ratio] = {}
        if user_input not in estimate_dict[fuzz_ratio]:
            estimate_dict[fuzz_ratio][user_input] = 0
        estimate_dict[fuzz_ratio][user_input] += 1
        return user_input

     
    sanitized_name = sanitizeForDiff(search_name)
    for uuid,this_name_dict in name_dict.items():
        # Found exact match with the sanitized value
        if 'sanitized' in this_name_dict and this_name_dict['sanitized'] == sanitized_name:
            log.debug(f'Exact Match:  Found {search_name} as the canonical name for {this_name_dict["canonical"]}')
            try:
                this_name_dict['names'][sanitized_name] += 1
            except Exception:
                this_name_dict['names'][sanitized_name] = 1
            return uuid
        # The sanitized value is in the list of known aliases
        elif 'names' in this_name_dict and sanitized_name in this_name_dict['names'].keys():
            log.debug(f'Exact Match: Found {search_name} in name list for {this_name_dict["canonical"]}')
            this_name_dict['names'][sanitized_name] += 1
            return uuid
    # No match - check for fuzzy interior matches.  Canonical name will always be in name list as well
    log.debug(f'No exact matches found, checking fuzzy matches for {search_name}')
    for uuid,dict_for_this_name in name_dict.items():
        user_input = None
        if 'names' in dict_for_this_name:
            for name in dict_for_this_name['names']:
                #log.debug(fuzz.ratio(m,name))
                fr = fuzz.ratio(sanitized_name,name)
                if fr >= FUZZ_THRESHOLD:
                    # It looks similar, but we need user verification

                    user_input = areTheyTheSame(search_name,sanitized_name,dict_for_this_name,name,fr,est_dict)
            
                    # User confirms they are the same, return the uuid 
                    if user_input != '3':
                        break
        if user_input == '1':
            log.debug(name_dict)
            dict_for_this_name['canonical'] = search_name
            dict_for_this_name['sanitized'] = sanitized_name
            dict_for_this_name['names'][sanitized_name] = 1
            log.debug(name_dict)
            return uuid
        # It's the same
        elif user_input == '2':
            dict_for_this_name['names'][sanitized_name] = 1
            return uuid

                    # Else, continue checking
    # No match, add a new entry
    log.debug(f'No match for {search_name}, adding a new one')
    uuid = str(uuid4())
    name_dict[uuid] = {
        'canonical' : search_name,
        'names' : { sanitized_name: 1 },
        'sanitized' : sanitized_name
    }
    return uuid
   






    



    #log.debug(f'It looks like {this_group} is the same as {name}, part of {group_name_dict["Canonical Name"]}')
    






def makeIndex(v):
    return re.sub(r'[^a-zA-Z0-9]', '', v).lower()

def sanitizeVal(v):
    v = v.strip().replace('  ',' ').replace(' : ',': ')
    return v


################################################################################
# SET UP LOGGING
################################################################################
def startLogger ():
    logFormat = '%(asctime)s - %(levelname)s - %(message)s'
    logging.basicConfig(format=logFormat)
    log = logging.getLogger()
    log.setLevel(LOG_LEVEL)
    return log

if __name__ == "__main__":
    log = startLogger()
    main()