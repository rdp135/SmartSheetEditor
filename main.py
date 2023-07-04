#!venv/bin/python3
"""

Test to show that SmartSheet API requires Authentication and Access Tokens
Token is hard coded so a password will need to be required in the compiled version

*Sheets cannot be bulk updated (import .xlsx will create a new sheet & id)
*Workaround --> use "Update Rows" from API after pulling data out of existing .xlsx into np or pd array

"""
__author__ = "Robert Palmere"
__version__ = "0.0.1"

import smartsheet
from xlsx_handler import *
from utils import *
import random
import string
import sys
import time

class SmartSheetHandler:

    def __init__(self, token):
        self.token = token
        self.user = None
        self.client = None
        self.sheets = None
        self.generated_rows = None
        self.generated_cols = None

    def __str__(self):
        return str(self.__class__)

    def connect(self):
        smartsheet_client = smartsheet.Smartsheet(self.token)
        smartsheet_client.errors_as_exceptions()
        user_info = smartsheet_client.Users.get_current_user()
        self.user = str(user_info._first_name) + " " + str(user_info._last_name)
        self.client = smartsheet_client
        print('\nConnected as %s\n' % self.user)
        return smartsheet_client

    @staticmethod
    def random_title():
        letters_ = string.ascii_letters
        title_list = []
        for i in range(50):
            choice = random.choice(letters_)
            title_list.append(choice)
        title = ''.join(title_list)
        return title


    def get_workspaces(self):
        '''
        Prints available work spaces and associated ID numbers
        :return: data attribute of client.Workspaces.list_workspaces().data
        '''
        print('--Available Workspaces--')
        ws_ = self.client.Workspaces.list_workspaces().data
        for n, i in enumerate(ws_):
            print('Name: "{}", ID: {}'.format(i.name, i.id))
            if n == len(ws_)-1:
                print('\n')
        return ws_

    def get_sheets(self, workspace):
        '''
        Prints available sheets and their associated ID numbers
        :param workspace: smartsheet_client.Workspaces.get_workspace(<workspace_ID>)
        :return: workspace.sheets as list
        '''
        print('--Available Sheets in Workspace: "{}"--'.format(workspace.name))
        sh_ = workspace.sheets
        self.sheets = sh_
        for n, i in enumerate(sh_):
            print('Name: "{}", ID: {}'.format(i.name, i.id))
            if n == len(sh_)-1:
                print('\n')
        return list(self.sheets)

    def get_row_ids(self, id):
        '''
        Gets all occupied row ids from target sheet id
        :param id: - specific sheet to get rows from
        :return: list containing row ids as ints
        '''
        result = self.client.Sheets.get_sheet(id)
        row_objs = list(result.rows)
        row_ids = [int(str(r._id_)) for r in row_objs]
        #print('--Occupied Row IDs--')
        #for n, i in enumerate(row_ids):
        #    print('Row ID: {}'.format(i))
        #    if n == len(row_ids)-1:
        #        print('\n')
        return row_ids

    def get_column_ids(self, id):
        '''
        Gets all occupied row ids from target sheet id
        :param id: - specific sheet to get columns from
        :return: list containing column ids as ints
        '''
        result = self.client.Sheets.get_sheet(id)
        col_objs = list(result.columns)
        col_ids = [int(str(c._id_)) for c in col_objs]
        #print('--Occupied Column IDs--')
        #for n, i in enumerate(col_ids):
        #    print('Column ID: {}'.format(i))
        #    if n == len(col_ids)-1:
        #        print('\n')
        return col_ids

    def generate_rows(self, col_ids, data):
        '''
        Function to generate Row() objects & subsequently append their data
        :param col_ids: column ids for which rows can append to
        :param data: np array of excel sheet
        :return: list of Row objects with data loaded
        '''
        row_count_ = 0
        rows = [self.client.models.Row() for i in range(data.shape[0])]
        
        # Populate the row objects with data
        for n, r in enumerate(rows):
            r.to_top = True # Adds rows from top to bottom of the sheet
            for k, c in enumerate(data[n]):
                r.cells.append({
                        'column_id': col_ids[k],
                        'value': str(data[n][k]), # must be string
                        'strict': False
                    })
        self.generated_rows = rows
        return self.generated_rows

    def generate_cols(self, data):
        '''
        Function to generate Col() objects & subsequently append their data
        Should only need if columns arent already generated!
        :param data: np array of excel sheet
        :return: list of Column objects with data loaded
        '''
        count = 0
        col_dicts = []
        cols = []
        # Generate list of dicts()
        while data.shape[1]:
            d = {
                'title' : self.random_title(),
                'type' : 'TEXT_NUMBER',
                'index' : 4
            }
            col_dicts.append(d)
            count += 1
            if count == data.shape[1]:
                break
        for d in col_dicts:
            col_ = self.client.models.Column(d)
            cols.append(col_)
        self.generated_cols = cols
        return self.generated_cols

    def clear_data(self, sheet_id, row_ids):
        '''
        Function to clear all data from the smart sheet with id "sheet_id"
        Caveat: must be done in groups of 200 otherwise "Request-URI Too Large"
        '''
        print('Deleting sheet contents by row - this may take a while...')
        row_list_del = [] # list of row_ids for deletion
        del_count = 0
        for r in row_ids:
            del_count += 1
            row_list_del.append(r)
            if len(row_list_del) > 199:
                self.client.Sheets.delete_rows(sheet_id, row_list_del)
                row_list_del = []
        if len(row_list_del) > 0:
            self.client.Sheets.delete_rows(sheet_id, row_list_del)
        print('Finished deleting {} rows.'.format(del_count))

    @staticmethod
    def look_up(target, sheets):
        '''
        Function to check if .xlsx being uploaded matches name of existing sheet name
        :param target: .xlsx file supplied from std.in
        :param sheets: Sheet objects from SmartSheet API
        :return: Sheet object from SmartSheet API that matches the name of supplied .xlsx file
        '''
        t_ = target.split('.xlsx')
        target_name = t_[0]
        sheet_names = [s.name for s in sheets]
        if target_name in sheet_names:
            for s_obj in sheets:
                if s_obj.name == target_name:
                    print('{} found in SmartSheet Workspace as {}'.format(target_name, s_obj.name))
                    return s_obj
        else:
            print('{} not found in SmartSheet Workspace.')
            exit(1)

    def push_rows(self, sheet_id, rows):
        '''
        Essentially the same function as self.clear_data() but for adding rows
        :param sheet_id:
        :param rows:
        :return:
        '''
        print('Adding sheet contents by row - this may take a while...')
        row_list_add = [] # list of row_ids for deletion
        add_count = 0
        for r in rows:
            add_count += 1
            row_list_add.append(r)
            if len(row_list_add) > 200:
                self.client.Sheets.add_rows(sheet_id, row_list_add)
                row_list_add = []
        if len(row_list_add) > 0:
            self.client.Sheets.add_rows(sheet_id, row_list_add)
        print('Finished adding {} rows.'.format(add_count))



if __name__ == '__main__':

    if len(sys.argv) > 2 or len(sys.argv) == 1:
        print('Usage: main.py <path_to_excel_sheet>')
        exit()
    target_xlsx = sys.argv[1]

    utils = Utility()
    utils.password_check()

    start_time = time.time()
    # Get SmartSheet Handle (** IMPORTANT: Password protect this using if statement **)
    SmartSheet = SmartSheetHandler(utils.access_key)

    # Connect to SmartSheet
    smartsheet_client = SmartSheet.connect()

    # Display available SmartSheet workspaces
    SmartSheet.get_workspaces()

    # Load Workspace for ID associated with Palmere - Test 2
    test_workspace = smartsheet_client.Workspaces.get_workspace(6533867277444996)

    # Display available sheets in Palmere - Test 2
    sheets = SmartSheet.get_sheets(test_workspace)

    # Check if supplied .xlsx exists within workspace already
    target_sheet = SmartSheet.look_up(target_xlsx, sheets)

    sheet_id = target_sheet._id_

    # Display Row IDs in current sheet
    row_ids = SmartSheet.get_row_ids(sheet_id)
    col_ids = SmartSheet.get_column_ids(sheet_id)

    # Get data from existing .xlsx sheet (converts to np array)
    data = xlsx_handler.get_xlsx(target_xlsx) # "test.xlsx"


    # Clear existing sheet data by row
    SmartSheet.clear_data(sheet_id, row_ids)

    # Generate new row objects w/ data

    rows = SmartSheet.generate_rows(col_ids, data)

    # Push rows with data from .xlsx to SmartSheets
    SmartSheet.push_rows(sheet_id, rows)

    print('\n---Finished---')

    end_time = time.time()
    time_taken = (end_time - start_time)/60
    print("Took {} minutes.".format(time_taken))










