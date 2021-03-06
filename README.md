# Smartsheet_Attachments_Upload

This is a Python script for uploading .mp3 attachments to rows on [Smartsheet](https://smartsheet.com) via the Smartsheet RESTful API.

# Python module dependencies

- collections
- json
- os
- requests
- smartsheet
- time

# Script operation

This script takes .mp3 files from a local directory and attaches the files to designated rows in a sheet on Smartsheet. For the script to upload an .mp3, the following conditions must be met.

- The filename of the .mp3 must be an exact match with an ID#* in the designated row's "ItemID" column cell (*not to be confused with the Smartsheet API `rowID` assigned to each row).
- An .mp3 file will only be uploaded if its designated row does not already have an .mp3 attachment.
- The "ItemID" ID# for the row and the filename of the .mp3 must be a unique match with no duplicates hiding elsewhere in the sheet.

If no match is found between an .mp3 filename and an ID# for a row, then the .mp3 is assumed to be mislabeled and will not be uploaded.

For .mp3 files that do get uploaded, when an .mp3 is successfully attached to a row, the row's "Status" column cell is updated. 
