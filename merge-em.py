#!/usr/bin/env python3

import csv

# Input csv files
inputCSVs = [
    ['', 'Inventory2026-01-15.csv'],
    ['bh', 'inventory-bh.csv'],
    ['pm', 'inventory-pm.csv'],
    ['vg', 'inventory-vg.csv'],
]

outputCSV = 'master.csv'

# compsite of all 
master = dict()

# Added column for tracing history:
colHistory = "History"

for user, inputCSV in inputCSVs:
    print(f'Processing {inputCSV}')
    with open(inputCSV, newline='') as inputCSVfile:
        reader = csv.DictReader(inputCSVfile)
        outputFieldNames = list(reader.fieldnames)
        # Add .view fields
        for n in range(1,24):
             outputFieldNames.insert(outputFieldNames.index(f'permissions.{n}.type')+1, f'permissions.{n}.view')
        # Add history field at the left edge
        outputFieldNames.insert(0, colHistory)
            
        for row in reader:
            # Strip leading slash from Path if present
            row['Path'] = row['Path'].removeprefix('/')
            # Remove anything URL params ('?' and following)
            row['WebViewLink'] = row['WebViewLink'].split('?')[0]
            # Add source info
            row[colHistory] = user

            # merge into resuting "master" list:
            uid = row['File/Folder id']
            try:
                m_row = master[uid]
                # resolve differences:
                if row['Name'] != m_row['Name']:
                    # File renamed
                    m_row['Name'] = row['Name']
                    m_row[colHistory] = f'{m_row[colHistory]}; {user}: rename'.strip('; ')
                p = int('0'+row['permissions'])
                m_p = int('0'+m_row['permissions'])
                if p > m_p:
                    # Expanded permissions:
                    row[colHistory] = f'{m_row[colHistory]}; {user}: perm {m_p}->{p}'.strip('; ')
                    master[uid] = row
            except KeyError:
                # Not in master list -- add it
                master[uid] = row
            
print(f'Found {len(master)} records')

with open(outputCSV, 'w', newline='', encoding='utf-8-sig') as outputCSVfile:
    writer = csv.DictWriter(outputCSVfile, fieldnames=outputFieldNames)
    writer.writeheader()
    for uid in sorted(master.keys(), key = lambda id: master[id]['Path']):
        writer.writerow(master[uid])

print("\nFinished")