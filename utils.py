#!/usr/bin/env python3
import random
import json
import csv
import datetime
import base64
from os import path as ospath
from os import makedirs
from sys import path
import pickle
from collections import OrderedDict
from email.mime.text import MIMEText
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow


class GoogleFriend:
    def __init__(self):
        self.test = True
        self.creds = None
        self.moduleFolder = ospath.abspath(ospath.dirname(__file__))
        self.authenticate()

    def authenticate(self):
        # Setup our scope for our request
        scopes = ['https://www.googleapis.com/auth/spreadsheets']

        print('Authenticating to google!')

        if ospath.exists(f'{self.moduleFolder}/creds/token.pickle'):
            with open(f'{self.moduleFolder}/creds/token.pickle', 'rb') as token:
                self.creds = pickle.load(token)

        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    f'{self.moduleFolder}/creds/credentials.json', scopes)
                self.creds = flow.run_local_server()

            # Save the credentials for the next run
            with open(f'{self.moduleFolder}/creds/token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('sheets', 'v4', credentials=self.creds, cache_discovery=False)

    def google_exporter(self, sheetName, sheetId, data, append=False):
        if isinstance(data, dict):
            data_dict = True
        else:
            data_dict = False

        sheet = self.service.spreadsheets()

        startOfRange = 'A'
        endOfRange = 'KZ'

        print('Sending data to google!')

        if append is True:
            keys = [list(data.keys())]
            keyRange = f'{sheetName}!{startOfRange}1:{endOfRange}1'
            keyRangeDict = {'range': keyRange,
                            'values': keys}

            currentValues = sheet.values().get(spreadsheetId=sheetId, range=f'{sheetName}!{startOfRange}:{endOfRange}').execute()

            if 'values' in currentValues.keys():
                header_set = True
                append_row = len(currentValues['values']) + 1
                valueRange = f'{sheetName}!{startOfRange}{str(append_row)}:{endOfRange}{str(append_row)}'

            else:
                headerUpdate = sheet.values().update(spreadsheetId=sheetId, range=keyRange, valueInputOption='RAW', body=keyRangeDict).execute()
                valueRange = f'{sheetName}!{startOfRange}2:{endOfRange}2'

            values = [list(data.values())]

            valueRangeDict = {'range': valueRange,
                              'values': values}

            valuesUpdate = sheet.values().update(spreadsheetId=sheetId, range=valueRange, valueInputOption='RAW', body=valueRangeDict).execute()

        else:
            if data_dict is True:
                print('data dict is true')
                for result in data:
                    keys = [list(data[result].keys())]
            else:
                keys = [list(data[0].keys())]

            keys[0].append('Last update date -> ' + str(datetime.datetime.now()))
            keyRange = f'{sheetName}!{startOfRange}1:{endOfRange}'
            keyRangeDict = {'range': keyRange,
                            'values': keys}

            valueRange = f'{sheetName}!{startOfRange}2:{endOfRange}'
            values = []
            for result in data:
                if data_dict is True:
                    # print(f'There are {len(list(data[result].values()))} Values for this item')
                    values.append(list(data[result].values()))
                else:
                    # print(f'There are {len(list(result.values()))} Values for this item')
                    values.append(list(result.values()))
            valueRangeDict = {'range': valueRange,
                              'values': values}

            try:
                clearSpreadsheetResults = sheet.values().clear(spreadsheetId=sheetId, range=keyRange).execute()
            except Exception as e:
                print('Page not existing! Creating')
                batchUpdate = {"requests": [{"addSheet": {"properties": {"title": sheetName}}}]}
                createSpreadsheetResults = sheet.batchUpdate(spreadsheetId=sheetId, body=batchUpdate).execute()

            headerUpdate = sheet.values().update(spreadsheetId=sheetId, range=keyRange, valueInputOption='RAW', body=keyRangeDict).execute()

            valuesUpdate = sheet.values().update(spreadsheetId=sheetId, range=valueRange, valueInputOption='RAW', body=valueRangeDict).execute()

    def retrieveSheet(self, id, range):
        sheet = self.service.spreadsheets()

        print(f'Retrieving results from the Google Sheet ({id})!')

        try:
            form_results = sheet.values().get(spreadsheetId=id, range=range).execute()
            return form_results
        except Exception as e:
            print(e)
            print('Issue getting that dataset!')
            return False

    def createMessage(self, sender, to, subject, message_text):
        message = MIMEText(message_text)
        message['To'] = to
        message['From'] = sender
        message['Subject'] = subject
        return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}

    def sendEmail(self, sender, recipient, subject, body):
        email = build('gmail', 'v1', credentials=self.creds)

        print(f'\n\nAttempting to send an email:\n\nFrom: {sender}\nTo: {recipient}\nSubject: {subject}\n\nBody:\n{body}')

        if email:
            print('\nAuthenticated to Gmail as you! Sending your message now.')

            ready_to_send_message = self.createMessage(sender=sender, to=recipient, subject=subject, message_text=body)

            send_email = email.users().messages().send(userId=recipient, body=ready_to_send_message).execute()

            print(send_email)


class color:
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    DARKCYAN = '\033[36m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'

    COLORS = [PURPLE, CYAN, DARKCYAN, BLUE, GREEN, YELLOW, RED, BOLD]

    def line_rainbow():
        color_blocks = '-' * 56
        color_line = ''
        for block in color_blocks:
            random_number = random.randint(0, len(color.COLORS) - 1)
            color_line += f'{color.COLORS[random_number]}{block}{color.END}'
        print(color_line)

    def mini_line_rainbow():
        color_blocks = '-' * 28
        color_line = ''
        for block in color_blocks:
            random_number = random.randint(0, len(color.COLORS) - 1)
            color_line += f'{color.COLORS[random_number]}{block}{color.END}'
        print(color_line)

    def print_big_update(text):
        # Print visual queue starting first section
        print('\n')
        color.line_rainbow()
        print(f'{color.BOLD}{text}{color.END}')
        color.line_rainbow()

    def print_mini_update(text):
        # Print visual queue starting first section
        print('\n')
        color.mini_line_rainbow()
        print(f'{color.BOLD}{text}{color.END}')
        color.mini_line_rainbow()

    def cool_text(text):
        color_line = ''
        for block in text:
            random_number = random.randint(0, len(color.COLORS) - 1)
            color_line += f'{color.COLORS[random_number]}{block}{color.END}'
        print(color_line)


# The purpose of this function is to send Messy lists of dictionaries to to have them flattened + all have matching key lists
def organize_into_clean_list(data):
    # first thing we do is loop through the list and flatten the dictionaries
    flatDictList = []

    for item in data:
        flatData = flatten_dictionary(item)
        flatDictList.append(flatData)

    all_keys = []
    ready_list = []

    # Loop through all items to identify the full list of possible keys!
    for item in flatDictList:
        for key in item.keys():
            if key not in all_keys:
                all_keys.append(key)

    for item in flatDictList:
        flat_ordered_dict = {}

        current_keys = list(item.keys())

        for all_key in all_keys:
            if all_key in current_keys:
                flat_ordered_dict[all_key] = item[all_key]
            else:
                flat_ordered_dict[all_key] = 'None'

        ready_list.append(flat_ordered_dict)

    return ready_list


# The purpose of this function is to send a messy single dictionary to to have it flattened
def flatten_dictionary(d):
    # # print(json.dumps(d, indent=3))
    result = {}
    stack = [iter(d.items())]  # Create a list of the dictionarie's keys + values in touples (k, v), (k, v) then put all that into a list
    keys = []
    list_of_dict_nums = {}
    num_dict_items_seen = 0
    while stack:
        for k, v in stack[-1]:  # Examine the LAST item in the list of touples
            keys.append(k)
            # input('\n\nReady to start next?')
            # print(f'\n\nStarting new Key/Value\nCurrent list of keys: {", ".join(keys)}')
            # print(f'Current list of dict_nums: {str(list_of_dict_nums)}')
            # print(f'Current key/value:\n{str(k)} : {str(v)}')
            if isinstance(v, list):
                # print('Value is a List')
                if len(v) > 0:
                    item_num = 0
                    for item in v:
                        if item:
                            if isinstance(item, dict):
                                # print('Item Value in the list is a dictionary')
                                if len(item.keys()) < 1:
                                    # print('Empty dictionary found')
                                    result['.'.join(keys)] = 'None'
                                else:
                                    for list_dict in v:
                                        item_num += 1
                                        list_of_dict_nums[str(item_num)] = {}
                                        list_of_dict_nums[str(item_num)]['index'] = len(keys)
                                        list_of_dict_nums[str(item_num)]['num_items'] = len(item)
                                        stack.append(iter(list_dict.items()))
                                    break
                            elif isinstance(item, list):
                                result['.'.join(keys)] = '.'.join(item)
                                keys.pop()  # This may need to be re-commented out
                            else:
                                result['.'.join(keys)] = ''.join(str(v))
                                keys.pop()
                                break
                    break
                else:
                    # print('Empty list, adding nothing! Will get filled with data later if data exists')
                    # # print(f'Key used: {".".join(keys)}')
                    # result['.'.join(keys)] = 'None'
                    keys.pop()
            elif isinstance(v, dict):
                if len(v.keys()) < 1:
                    result['.'.join(keys)] = 'None'
                    keys.pop()
                else:
                    stack.append(iter(v.items()))
                    break
            else:
                dict_nums_keys = list(list_of_dict_nums.keys())

                if len(list_of_dict_nums) > 0:
                    # print('Theres dict nums!')
                    # print(f'dict_nums_keys == {str(dict_nums_keys)}')
                    # print(f'num_dict_items_seen: {str(num_dict_items_seen)} <= list_of_dict_nums[dict_nums_keys[-1]]["num_items"]: {str(list_of_dict_nums[dict_nums_keys[-1]]["num_items"])}')
                    if num_dict_items_seen <= list_of_dict_nums[dict_nums_keys[-1]]['num_items']:
                        # print('We still are seeing items from the previous list!')
                        last_num = list_of_dict_nums[dict_nums_keys[-1]]['index']
                        # print(f'The location in the "keys" list our number should go: {str(last_num)}')
                        keys.insert(last_num, dict_nums_keys[-1])
                        # print(f'Key were going to use: {".".join(keys)}')
                        result['.'.join(keys)] = str(v)
                        # print(f'Keys before removing #{str(last_num)} and last key: {".".join(keys)}')
                        # This removes the Number identifier
                        keys.pop(last_num)
                        # This removes the last key like normal
                        keys.pop()
                        num_dict_items_seen += 1
                    if num_dict_items_seen == list_of_dict_nums[dict_nums_keys[-1]]['num_items']:
                        # print('Hit the last item of the Listed dictionary!')
                        # print(keys)
                        list_of_dict_nums.popitem()
                        dict_nums_keys.pop()
                        num_dict_items_seen = 0
                else:
                    # print('Value added directly (was str or int)')
                    result['.'.join(keys)] = str(v)
                    # print(f'Key used: {".".join(keys)}')
                    if len(dict_nums_keys) == 0:
                        keys.pop()
        else:
            # print('Hit the else that removes all keys')
            if len(dict_nums_keys) == 0:
                if keys:
                    keys.pop()
            stack.pop()
    ordered_data = OrderedDict(sorted(result.items(), key=lambda t: t[0]))
    return ordered_data


# The purpose of this function is to send a messy single dictionary to to have it flattened
def old_flatten_dictionary(d):
    result = {}
    stack = [iter(d.items())]  # Create a list of the dictionarie's keys + values in touples (k, v), (k, v) then put all that into a list
    keys = []
    while stack:
        for k, v in stack[-1]:  # Examine the LAST item in the list of touples
            keys.append(k)
            if isinstance(v, list):
                if len(v) > 0:
                    for item in v:
                        if item:
                            if isinstance(item, dict):
                                if len(item.keys()) < 1:
                                    result['.'.join(keys)] = 'None'
                                else:
                                    stack.append(iter(item.items()))
                            elif isinstance(item, list):
                                result['.'.join(keys)] = '.'.join(item)
                                keys.pop()  # This may need to be re-commented out
                            else:
                                result['.'.join(keys)] = ''.join(str(v))
                                keys.pop()
                                break
                    break
                else:
                    result['.'.join(keys)] = 'None'
                    keys.pop()
            elif isinstance(v, dict):
                if len(v.keys()) < 1:
                    result['.'.join(keys)] = 'None'
                    keys.pop()
                else:
                    stack.append(iter(v.items()))
                    break
            else:
                result['.'.join(keys)] = str(v)
                keys.pop()
        else:
            if keys:
                keys.pop()
            stack.pop()
    return result


def csv_maker(filename, data, preorganized=False):
    file_name = filename
    home_dir = f'{path[0]}/reports'
    if not ospath.exists(home_dir):
        makedirs(home_dir)

    if type(data) == dict:
        data_list = []
        for item in data:
            data_list.append(data[item])
        data = data_list

    if preorganized is False:
        ready_info = organize_into_clean_list(data)
    else:
        ready_info = data

    with open(f'{home_dir}/{file_name}.csv', 'w') as outfile:
        output = csv.writer(outfile)
        # Exporting Dictionary to csv

        all_keys = list(ready_info[0].keys())

        output.writerow(all_keys)
        for row in ready_info:
            output.writerow(row.values())
        print(f'\n\033[92mSuccessfully exported ({str(len(data))} items) to CSV -> {file_name}.csv \033[0m')
