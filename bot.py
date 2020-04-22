# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from botbuilder.core import ActivityHandler, TurnContext, MessageFactory
from botbuilder.schema import ChannelAccount, Mention

# import library
import datetime
import pytz


class MyBot(ActivityHandler):
    # See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.

    # async def on_message_activity(
    #         self,
    #         turn_context: TurnContext
    # ):
    #     if turn_context.activity.text.lower() == "punch in" or turn_context.activity.text.lower() == "punch in.":
    #         update_check_in(turn_context.activity.from_property.name)
    #         await turn_context.send_activity(
    #             f"Hello {turn_context.activity.from_property.name}, your {turn_context.activity.text} has been registered.")
    #
    #     if turn_context.activity.text.lower() == "punch out" or turn_context.activity.text.lower() == "punch out.":
    #         update_check_out(turn_context.activity.from_property.name)
    #         await turn_context.send_activity(
    #             f"Hello {turn_context.activity.from_property.name}, your {turn_context.activity.text} has been registered.")

    async def on_turn(
            self,
            turn_context: TurnContext
    ):
        if turn_context.activity.text is not None:
            if turn_context.activity.text.lower()[-8:] == "punch in" or turn_context.activity.text.lower()[-9:] == "punch in.":
                update_check_in(turn_context.activity.from_property.name)
                await turn_context.send_activity(
                    f"Hello {turn_context.activity.from_property.name}, your punch in has been registered.")

            if turn_context.activity.text.lower()[-9:] == "punch out" or turn_context.activity.text.lower()[-10:] == "punch out.":
                working_hours = update_check_out(turn_context.activity.from_property.name)
                await turn_context.send_activity(
                    f"Hello {turn_context.activity.from_property.name}, your punch out has been "
                    f"registered.\n\n\u200CYour effective working hours are " + working_hours + ".")

            if turn_context.activity.text.lower()[-17:] == "faberworkbot help" or turn_context.activity.text.lower()[-18:] == "faberworkbot help.":
                await turn_context.send_activity("Hello! I'm Faberwork Bot. I'll take care of Faberwork "
                                                 "attendance. \n\n\u200CTo use me, you have to mention me and "
                                                 "use following commands"
                                                 "\n\n\u200C @FaberworkBot punch in \n\n\u200C @FaberworkBot punch out "
                                                 "\n\n\u200C @FaberworkBot status \n\n\u200C @FaberworkBot status "
                                                 "date (in dd/mm/yyyy format)")

            if turn_context.activity.text.lower() == "help" or turn_context.activity.text.lower() == "help.":
                await turn_context.send_activity("Hello! I'm Faberwork Bot. I'll take care of Faberwork "
                                                 "attendance. \n\n\u200CTo use me, "
                                                 "use following commands"
                                                 "\n\n\u200Cpunch in \n\n\u200Cpunch out "
                                                 "\n\n\u200Cstatus \n\n\u200Cstatus "
                                                 "date (in dd/mm/yyyy format)")

            if "status" in turn_context.activity.text.lower():
                if turn_context.activity.text.lower()[-6:] == "status":
                    currentDT = datetime.datetime.now(pytz.timezone("Asia/Calcutta"))
                    date = currentDT.strftime("%d/%m/%Y")
                    status = attendance_status(date)
                    await turn_context.send_activity(f"Here is the attendance of " + date + " \n\n\u200C" + status)

                elif turn_context.activity.text.lower().endswith("2020"):
                    date = turn_context.activity.text.lower()[-10:]
                    status = attendance_status(date)
                    await turn_context.send_activity(f"Here is the attendance of " + date + " \n\n\u200C" + status)


    async def on_members_added_activity(
            self,
            members_added: ChannelAccount,
            turn_context: TurnContext
    ):
        for member_added in members_added:
            if member_added.id != turn_context.activity.recipient.id:
                await turn_context.send_activity("Hello and welcome! I'm Faberwork Bot. I'll take care of Faberwork "
                                                 "attendance. \n\n\u200CTo use me, "
                                                 "use following commands"
                                                 "\n\n\u200Cpunch in \n\n\u200Cpunch out "
                                                 "\n\n\u200Cstatus \n\n\u200Cstatus "
                                                 "date (in dd/mm/yyyy format)")


import gspread
# Service client credential from oauth2client
from oauth2client.service_account import ServiceAccountCredentials
# Print nicely
import pprint

# Create scope
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
# create some credential using that scope and content of startup_funding.json
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
# create gspread authorize using that credential
client = gspread.authorize(creds)
# Now will can access our google sheets we call client.open on

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("Attendance").sheet1

# Extract and print all of the values
list_of_hashes = sheet.get_all_records()
print(list_of_hashes)


def get_last_row_number(sheet, first_column_index):
    """Returns last empty row number in the sheet"""
    values_list = sheet.col_values(first_column_index)
    last_index = len(values_list)
    if values_list[last_index - 1] == 'Sr. No.':
        return 2
    return last_index + 1


def get_last_serial_number(sheet, first_column_index):
    """Returns last filled serial number in the sheet"""
    values_list = sheet.col_values(first_column_index)
    last_index = len(values_list)
    if values_list[last_index - 1] == 'Sr. No.':
        return 1
    return last_index


def update_sheet(sheet, row, column, data):
    """Updates data in the given sheet, row and column """
    sheet.update_cell(row, column, data)


def update_check_in(employee_name):
    currentDT = datetime.datetime.now(pytz.timezone("Asia/Calcutta"))
    last_row = get_last_row_number(sheet, 1)
    row = get_checkinout_row_number(sheet, employee_name)
    if sheet.cell(row, 3).value != currentDT.strftime("%d/%m/%Y"):
        sheet.update_cell(last_row, 1, get_last_serial_number(sheet, 1))
        sheet.update_cell(last_row, 2, employee_name)
        sheet.update_cell(last_row, 3, currentDT.strftime("%d/%m/%Y"))
        sheet.update_cell(last_row, 4, currentDT.strftime("%I:%M:%S %p"))
    print(sheet.get_all_records())


def get_checkinout_row_number(sheet, employee_name):
    """Returns row number in which the given data cell is present of given cloumn"""
    # cell = sheet.find(data)
    employee_checkin_list = sheet.findall(employee_name)
    if len(employee_checkin_list) == 0:
        return get_last_row_number(sheet, 1)
    else:
        return employee_checkin_list[len(employee_checkin_list) - 1].row


def update_check_out(employee_name):
    row = get_checkinout_row_number(sheet, employee_name)
    currentDT = datetime.datetime.now(pytz.timezone("Asia/Calcutta"))
    if sheet.cell(row, 3).value == currentDT.strftime("%d/%m/%Y"):
        sheet.update_cell(row, 5, currentDT.strftime("%I:%M:%S %p"))
    return sheet.cell(row, 6).value




def attendance_status(date):
    status = sheet.get_all_records()
    day_status = "*Sr. No.*  |  *Name*  |  *Date*  |  *Check-in Time*  |  *Check-out Time*  |  *Working Hours* \n\n\u200C"
    for employee in status:
        if employee["Date"] == date:
            day_status = day_status + str(employee["Sr. No."]) + " | " + str(
                employee["Name"]) + " | " + str(employee["Date"]) + " | " + str(
                employee["Check-in Time"]) + " | " + str(employee["Check-out Time"]) + " | " + str(employee["Working Hours"]) + "\n\n\u200C"
    return day_status
