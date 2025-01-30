from sys import path_hooks
import win32com.client as win32
import glob
import os

def mail_to_sign(self, path, listpath):

    path_asset = listpath +"\*(Asset List).pdf"
    list_of_files = glob.glob(path_asset) # * means all if need specific format then *.csv
    latest_file_list = max(list_of_files, key=os.path.getctime)
    print(latest_file_list)

    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')

    # construct email item object
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = 'Please Sign'
    mailItem.BodyFormat = 1
    
    mailItem.HTMLBody = "<p>Mr.Rivera <br><br> Please sign and save the inspection letter using the following link: <br></p><a href='{0}'>{1}</a><br><p>Asset List:</p><a href='{2}'>{3}</a>".format(path, path,latest_file_list,latest_file_list)
    mailItem.To = '<receipent email>'
    mailItem.Sensitivity  = 2
    # optional (account you want to use to send the email)
    # mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('<email@gmail.com')))
    mailItem.Display()
    # mailItem.Save()
    # mailItem.Send()

def mail_signed(self, path):
    path_list = path + "\*(Letter).pdf"
    list_of_files = glob.glob(path_list)
    latest_file_letter = max(list_of_files, key=os.path.getctime)
    print(latest_file_letter)

    path_asset = path +"\*(Asset List).pdf"
    list_of_files = glob.glob(path_asset) 
    latest_file_list = max(list_of_files, key=os.path.getctime)
    print(latest_file_list)

    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')

    # construct email item object
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = 'Please Sign'
    mailItem.BodyFormat = 1
    
    mailItem.HTMLBody = "<a href='{0}'>{1}</a> <br/><a href='{2}'>{3}</a>  ".format(latest_file_letter, latest_file_letter, latest_file_list,latest_file_list)
    mailItem.To = '<receipent email>'
    mailItem.Sensitivity  = 2
    # optional (account you want to use to send the email)
    # mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('<email@gmail.com')))
    mailItem.Display()
    # mailItem.Save()
    # mailItem.Send()