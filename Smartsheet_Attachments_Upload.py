import collections
import json
import os
import requests
import smartsheet
import time

headers = {} # Smartsheet API token
params = {'includeAll': 'True'}

SheetID = # Sheet ID on Smartsheet


# =======================================================================================================
# Get all column ids from sheet on Smartsheet and store entries in dict as {column title: column id}
# =======================================================================================================


SS_columns_raw = requests.get(f'https://api.smartsheet.com/2.0/sheets/{SheetID}/columns', params=params, headers=headers)
SS_columns = json.loads(SS_columns_raw.text)['data']

SS_columnIDs = {column['title']:column['id'] for column in SS_columns}


# ==================================================================
# Get list of all rows and attachments from sheet on Smartsheet
# ==================================================================


params = {'include': 'attachments','columnIds': f"{SS_columnIDs['ItemID']},{SS_columnIDs['Status']}"}

SS_rows_raw = requests.get(f'https://api.smartsheet.com/2.0/sheets/{SheetID}', params=params, headers=headers)
SS_rows = json.loads(SS_rows_raw.text)['rows']

# Strip all keys but 'id', 'rowNumber', 'cells', and 'attachments' for each row dict in list (SS_rows)

SS_rows_keys_keep = ['id', 'rowNumber', 'cells', 'attachments']
for row in range(len(SS_rows)):
    for key in list(SS_rows[row].keys()):
        if key not in SS_rows_keys_keep:
            del SS_rows[row][key]

# Strip all keys but 'columnId' and 'value' for each dict in 'cells' sublist for each row dict in list (SS_rows)
        
SS_cells_keys_keep = ['columnId', 'value']       
for row in range(len(SS_rows)):
    for cell in range(len(SS_rows[row]['cells'])):
        for key in list(SS_rows[row]['cells'][cell].keys()):
            if key not in SS_cells_keys_keep:
                del SS_rows[row]['cells'][cell][key]

# Initialize empty list for .mp3 row attachments

SS_rows_attachments_mp3s = []

# Strip all keys but 'id' and 'name' for each dict in 'attachments' sublist for each row dict in list (SS_rows)

SS_attachments_keys_keep = ['id', 'name']
for row in range(len(SS_rows)):
    if 'attachments' in SS_rows[row]:
        for attachment in range(len(SS_rows[row]['attachments'])):
            for key in list(SS_rows[row]['attachments'][attachment].keys()):
                if key not in SS_attachments_keys_keep:
                    del SS_rows[row]['attachments'][attachment][key]
                    
# Generate list of all .mp3 row attachments
                  
                elif key == 'name' and SS_rows[row]['attachments'][attachment]['name'].endswith('.mp3'):
                    SS_rows_attachments_mp3s.append(SS_rows[row]['attachments'][attachment]['name'])


# ========================================================
# Generate list of all .mp3 files in local directory
# ========================================================


# Initialize empty list for .mp3 files in directory

local_mp3s = []

# Define local mp3s directory pathname

path = r'...path to local directory where mp3s are stored...'

# Add all files in directory ending with ".mp3" to local_mp3s list

[local_mp3s.append(mp3) for mp3 in os.listdir(path) if mp3.endswith('.mp3')]


# ==========================================================
# Generate list of .mp3 files that haven't been uploaded yet
# ==========================================================


local_mp3s_not_uploaded = list(set(local_mp3s) - set(SS_rows_attachments_mp3s))

# Define function to strip '.mp3' suffix from each member of .mp3 files list

def mp3strip(ItemIDs):
    ItemIDs = [ItemIDs[ItemID].rstrip('.mp3') for ItemID in range(len(ItemIDs))]
    return ItemIDs

# Call function to strip .mp3 suffix from each member of local_mp3s_not_uploaded list

local_mp3s_not_uploaded_ItemIDs = mp3strip(local_mp3s_not_uploaded)


# =========================================================================================================
# Remove all row dicts from SS_rows list except for those whose ItemIDs match the .mp3 files to be uploaded
# =========================================================================================================


# Initialize empty list for ItemIDs that match .mp3 files to be uploaded

SS_mp3s_to_upload_ItemIDs = []

# Trim row dicts from SS_rows list that already have .mp3 file attachments

for row in reversed(range(len(SS_rows))):
    if 'attachments' in SS_rows[row]:
        for attachment in range(len(SS_rows[row]['attachments'])):
            for key in list(SS_rows[row]['attachments'][attachment].keys()):
                if key == 'name' and SS_rows[row]['attachments'][attachment]['name'].endswith('.mp3'):
                    del SS_rows[row]
                    break
            break
        
# Check row dicts in SS_rows list for ItemIDs that don't have .mp3 attachments
            
    elif 'attachments' not in SS_rows[row]:
        for cell in reversed(range(len(SS_rows[row]['cells']))):

# Generate list of all ItemIDs for rows that match .mp3 files to be uploaded
            
            if SS_rows[row]['cells'][cell]['columnId'] == SS_columnIDs['ItemID'] and SS_rows[row]['cells'][cell].get('value') in local_mp3s_not_uploaded_ItemIDs:
                SS_mp3s_to_upload_ItemIDs.append(SS_rows[row]['cells'][cell]['value'])
                
# Trim row dicts from SS_rows list whose ItemIDs don't match .mp3 files to be uploaded
                
            if SS_rows[row]['cells'][cell]['columnId'] == SS_columnIDs['ItemID'] and SS_rows[row]['cells'][cell].get('value') not in local_mp3s_not_uploaded_ItemIDs:
                del SS_rows[row]
                break

# Generate list of .mp3 files in local mp3 directory that don't match any ItemIDs in the sheet on Smartsheet

local_mp3s_diff_ItemIDs = list(set(local_mp3s_not_uploaded_ItemIDs) - set(SS_mp3s_to_upload_ItemIDs))

# Define function to add .mp3 suffix to each member of local_mp3s_diff_ItemIDs list

def mp3stripe(mp3s):
    mp3s = [mp3s[mp3]+'.mp3' for mp3 in range(len(mp3s))]
    return mp3s

# Call function to add .mp3 suffix to members of local_mp3s_diff_ItemIDs list
    
diff_mp3s = mp3stripe(local_mp3s_diff_ItemIDs)

# Remove entries from local_mp3s_not_uploaded_ItemIDs with no matching ItemIDs in the sheet on Smartsheet

if len(diff_mp3s) > 0:
    print('\nMatching item IDs could not be found in the the sheet on Smartsheet for the following .mp3s: \n', 
          *diff_mp3s, 
          sep = '\n',
          end = '\n \nThese .mp3s will not be uploaded at this time.')
    for ItemID in range(len(local_mp3s_diff_ItemIDs)):
        local_mp3s_not_uploaded_ItemIDs.remove(local_mp3s_diff_ItemIDs[ItemID])
    
# Generate list of duplicate ItemIDs found in the sheet on Smartsheet

SS_duplicate_ItemIDs = [ItemID for ItemID, instance in collections.Counter(SS_mp3s_to_upload_ItemIDs).items() if instance > 1]

# Remove entries from local_mp3s_not_uploaded_ItemIDs that correspond to duplicate ItemIDs in the sheet on Smartsheet

if len(SS_duplicate_ItemIDs) > 0:
    print('\n \nThe following item IDs are duplicates in the the sheet on Smartsheet: \n',
          *SS_duplicate_ItemIDs,
          sep = '\n',
          end = '\n \nThe corresponding .mp3s will not be uploaded at this time. \n \n \n')
    for ItemID in range(len(SS_duplicate_ItemIDs)):
        local_mp3s_not_uploaded_ItemIDs.remove(SS_duplicate_ItemIDs[ItemID])

# Trim row dicts with duplicate ItemIDs from SS_rows list

for row in reversed(range(len(SS_rows))):
    for cell in range(len(SS_rows[row]['cells'])):
        if SS_rows[row]['cells'][cell]['columnId'] == SS_columnIDs['ItemID'] and SS_rows[row]['cells'][cell].get('value') not in local_mp3s_not_uploaded_ItemIDs:
            del SS_rows[row]
            break

input('\n \nPress the Enter key to begin uploading .mp3 file attachments.')

# ===================================================================
# Upload .mp3 files to corresponding rows in the sheet on Smartsheet
# ===================================================================
        

# Set Smartsheet API credentials to update status column for rows in the sheet on Smartsheet

headers = {'',
           'Content-Type': 'application/json'} # Smartsheet API token

params = {'allowPartialSuccess': 'True'}

url = f'https://api.smartsheet.com/2.0/sheets/{SheetID}/rows/'

# Set Smartsheet API access token for SDK access to upload .mp3 file attachments to rows in the sheet on Smartsheet

access_token = '' # Smartsheet API token
SS_client = smartsheet.Smartsheet(access_token)

# Loop through SS_rows list and attach corresponding .mp3 files to rows in the sheet on Smartsheet

if len(SS_rows) > 0:
    print('\n\nAttachments uploaded:\n')
else: print('\n\nNo attachments to upload at this time.\n')

for row in range(len(SS_rows)):
    
    if (row + 1) % 10 == 0:
        time.sleep(30) # 30 second time delay after the first 9 attachments uploaded, then after every 10 attachments uploaded
        
    rowID = SS_rows[row]['id']
    rowNumber = SS_rows[row]['rowNumber']           
    for cell in range(len(SS_rows[row]['cells'])):
        if SS_rows[row]['cells'][cell]['columnId'] == SS_columnIDs['Item ID (Item 1)']:
            mp3 = SS_rows[row]["cells"][cell]["value"]+'.mp3'

            SS_mp3_attachment = SS_client.Attachments.attach_file_to_row(
                                SheetID, 
                                rowID, 
                                (f'{mp3}', open(path+f'\{mp3}', 'rb'), 'audio/mp3'))
            
            if SS_mp3_attachment.message == 'SUCCESS':
                SS_rows[row]['message'] = 'SUCCESS'
                print('row #'+f'{rowNumber}', mp3, SS_mp3_attachment.message, '\n', sep = ' ')
                
# Update status column for each .mp3 file attached to row in the sheet on Smartsheet
               
                payload = f'{{"id": {rowID}, "cells": [{{"columnId": {SS_columnIDs["Status"]}, "value": "Status...", "strict": "False"}}]}}'
                response = requests.request("PUT", url, data=payload, params=params, headers=headers)
                time.sleep(2) # Time delay for 2 seconds after every attachment is uploaded
