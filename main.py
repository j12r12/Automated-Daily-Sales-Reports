import pandas as pd
from data_2 import Daily_Data
from graph_2 import Graph
from excel_2 import Excel
import os
import sys
from datetime import datetime, date
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from win32com.client import Dispatch
import win32ui


def outlook_running():

    """This function checks whether or not Microsoft Outlook is open. If not, then it opens it."""

    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return True
    except win32ui.error:
        return False


def save_attachment():
    
    """This function searches through Microsoft Outlook inbox for emails received today with set subject"""

    if not outlook_running():
        os.startfile("outlook")

    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder("6")
    all_emails = inbox.items
    subject = "Daily Report Data"

    input_path = os.path.join(application_path, input_format)

    for i in all_emails:
        if i.Subject == subject and i.ReceivedTime.date() == today:
            for att in i.Attachments:
                att.SaveAsFile(input_path)


def send_email():

    """This function sends the completed Excel file to all attended recipients."""

    # Create email
    outlook_send = Dispatch("Outlook.Application")
    mail = outlook_send.CreateItem(0)

    # Attach completed file
    att_out = "{}/{}".format(application_path, output_filename)
    mail.Attachments.Add(att_out)
    mail.To = "EMAIL ADDRESS"
    mail.Subject = "Test"
    mail.Body = "Good Afternoon,\n\nPlease see attached today's Daily Report.\n\nKind Regards."

    mail.send


if __name__ == "__main__":

    # Customisable sections
    today = datetime.today().date()
    not_required = ["Sales Type 1", "Sales Type 2"]
    salesperson = ["Salesperson 1", "Salesperson 2", "Salesperson 3"]
    stage = ["Won", "Lost"]

    # It can be useful to setup an executable programme so this can be shared with others who
    # do not have python installed on their machine. As part of this, it is necessary to
    # change the file path depending on whether or not the programme is being run from
    # the same directory as the code or from another directory (as with a shared executable).

    if getattr(sys, "frozen", False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    else:
        application_path = ""
        print("Error with the application path.")

    output_filename = "Daily Report - {}.xlsx".format(today)
    input_format = "InputData - {}.csv".format(today)

    # Run save_attachment function to download attachment from today's email.
    save_attachment()

    # Read in data
    df = pd.read_csv(os.path.join(application_path, input_format))

    # Create Excel workbook
    write1 = pd.ExcelWriter(os.path.join(application_path, output_filename), engine="xlsxwriter")

    # Create pivoted dataframes
    qt = Daily_Data(df, not_required, salesperson, stage).pivot_quantity("Quantity",
                                                                         "Salesperson")
    rev = Daily_Data(df, not_required, salesperson, stage).pivot_revenue("Total Price",
                                                                         "Salesperson")
    prod_rev = Daily_Data(df, not_required, salesperson, stage).pivot_revenue("Total Price",
                                                                              "Product")

    # Create excel sheets
    quant_sheet = Excel(write1, qt, "WonLost for day QTY", application_path).create_workbook()
    revenue_sheet = Excel(write1, rev, "WonLost for day (£)", application_path, True).create_workbook()
    products_sheet = Excel(write1, prod_rev, "WonLost for day Products (£)", application_path,
                           True).create_workbook()

    # Save file in working directory
    write1.save()

    send_email()
