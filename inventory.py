#!/usr/bin/env python3

import sys
import re
import csv
import xlsxwriter

# Input and output files
inputCSV = 'master.csv'
outputExcel = 'output.xlsx'

# inputCST headers:
HISTORY = 'History'
ID = 'File/Folder id'
NAME = 'Name'
PATH = 'Path'
URL = 'WebViewLink'
NUM = 'permissions'

# Output column headers:
columnHeaders = ('History', 'Type', 'Item', 'Path', 'caseIndex', 'Disposition strategy', '#perms', 'owner', 'writers', 'readers', 'commenters', 'missingWriters', 'missingCommenters')
columnCount = len(columnHeaders)
columnLast = columnCount-1

# output column header index
colHistory,     \
colType,        \
colItem,        \
colPath,        \
colCaseIndex,   \
colDisposition, \
colPermCount,   \
colOwner,       \
colWriters,     \
colReaders,     \
colCommenters,  \
colMissingWriters, \
colMissingCommenters, \
    = range(columnCount)

# Some permission sets:
stdWriters = set((
    'alan_ward', 'annie_olsen', 'becca_spalinger', 'bob_hallissy', 
    'bobby_devos', 'david_raymond', 'david_rowe', 'dawson_tennant', 
    'emily_roth', 'jon_coblentz', 'kim_rasmussen', 'lorna_evans',       # At present, omit eric_macleod
    'martin_hosken', 'martin_raymond', 'michael_cochran', 'nicolas_spalinger', # omit nrsi.old.gdocs
    'peter_martin', 'sharon_correll', 'steven_dyk', 'tim_eves', 'victor_gaultney',
    'Director')) 
stdCommenters = set(('jim_brase',))
# stdReaders = set()

# Folder URLs 
folderURLs = {'WSTech Team/': 'https://drive.google.com/drive/folders/0B-U3EYurfq2sQjF3c05IdTJDdHc?resourcekey=0-_aJvzSliP5q2ZQdSP8jspQ'}

# Keep track of how many unique sets of non-standard permissions we have
# Index is f'{missingWriters}|{missingCommenters}'; valuse an index. 
permissionCaseIndex = {'|': ''}

with open(inputCSV, newline='', encoding='utf-8-sig') as csvfile:
    reader = csv.DictReader(csvfile)
    with xlsxwriter.Workbook(outputExcel) as workbook:  
        # define/update formats:
        hCenterFormat = workbook.add_format({'valign': 'vcenter','align': 'center'})    # Horizontal center
        wrapFormat = workbook.add_format({'valign': 'vcenter', 'text_wrap': True})  # wrap
        # Modify the url format to do vertical center.
        urlFormat = workbook.get_default_url_format()
        urlFormat.set_align('vcenter')
        urlFormat.set_text_wrap(True)

        worksheet = workbook.add_worksheet('Inventory')

        # Set default cell format and wrapping
        worksheet.set_column(0, columnLast, None, wrapFormat)
        for col in (colType, colPermCount):
            worksheet.set_column(col, col, None, hCenterFormat)
        
        worksheet.write_row(0,0, columnHeaders)
        row = 1
        for line in reader:
            print(row, end='\r')
            path = line[PATH]   
            name = line[NAME]
            # Two of the folders actually end with a space but the Inventory spreadsheet omits it.
            if re.search(r'(?:Emily\'s SLDR Change Research|Retirement Events/David Raymond)$', path):
                path += ' '
                name += ' '
            if '/' not in name:
                # Some item names consist of digits only and the 'name' field is thus treated as numeric
                # To work around this if there is no '/' in the name then we can peel the name from the full path:
                name = path.rsplit('/',1)[-1]
            url = line[URL].removesuffix('?usp=drivesdk')
            if url.startswith('https://drive.google.com/drive/folders'):
                itemType = 'F' 
                folderURLs[path+'/'] = url
            else:
                itemType = 'I'
            path = path.removesuffix(name)

            permCount = int('0'+line[NUM])

            worksheet.write_string(row, colHistory, line[HISTORY])
            worksheet.write_string(row, colType, itemType)
            worksheet.write_url(   row, colItem, url, string=name)
            try:
                worksheet.write_url(row, colPath, folderURLs[path], string=path)
            except KeyError:
                print(f'\nrow {row}: bad URL for path - "{path}"\n', file=sys.stderr)
                worksheet.write_string(   row, colPath, path)
            worksheet.write_number(row, colPermCount, permCount)
            
            # process permissions, if any.
            roles = {}
            for p in range(permCount):
                # check for link-sharing
                fileDiscovery = line.get(f'permissions.{p}.allowFileDiscovery', '')
                if len(fileDiscovery):
                    # this is some kind of link sharing -- we're ignoring for now
                    continue
                role = line[f'permissions.{p}.role']
                email = line[f'permissions.{p}.emailAddress'].split('@')[0]
                # Special case: treat director_nrsi and director_wstech as equivalent:
                if re.match(r'director_(?:wstech|nrsi)$', email):
                    email = 'Director'
                roles.setdefault(role,set()).add(email)
            # add permissions to output
            for col,role in zip(range(colOwner,colOwner+4),('owner','writer', 'reader', 'commenter')):
                worksheet.write_string(row, col, ', '.join(sorted(roles.setdefault(role,set()))))
            
            # Do some checks:
            try:
                owner = roles['owner'].pop()
                if len(roles['owner']):
                    print(f'\nrow {row}: more than 1 owner\n', file=sys.stderr)
            except:
                print(f'row {row}: cannot find owner', file=sys.stderr)
                owner = 'unknown'
            
            missingWriters = stdWriters - roles['writer']
            missingWriters.discard(owner)
            missingWriters = ' '.join(sorted(missingWriters))
            worksheet.write_string(row, colMissingWriters, missingWriters)
            
            missingCommenters = stdCommenters - roles['writer'] - roles['commenter']
            missingCommenters.discard(owner)
            missingCommenters = ' '.join(sorted(missingCommenters))
            worksheet.write_string(row, colMissingCommenters, missingCommenters)

            # Keep track of all unique cases:

            caseIndex = permissionCaseIndex.setdefault(f'{missingWriters}|{missingCommenters}', len(permissionCaseIndex))
            if caseIndex:
                worksheet.write_number(row, colCaseIndex, caseIndex)

            row += 1
            
        worksheet.autofilter(0, 0, row, columnLast)
        worksheet.autofit(300)
        worksheet.freeze_panes(1, 1)

print("\nFinished")