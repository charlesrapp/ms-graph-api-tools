# Microsoft Graph API tools

This repository is a set of scripts to take advantages of Microsoft Graph API in some real situations.

## 1st case: Extract all email addresses from my mailbox

The script `extractor.py` will crawl your mailbox and collect all senders available in your mailbox and its subfolders. It will also count the number of emails you received by sender and generate a csv file with the results.

This script is running with Python3 in command line : `python3 extractor.py`
It requires to get a valid token for the Graph API with Mail.Read permission and that needs to be replaced in the file `extractor.py`.
The token can be found using the Graph Explorer <https://developer.microsoft.com/fr-fr/graph/graph-explorer>
