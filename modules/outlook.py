import re
import pandas as pd
import win32com.client as cl
import datetime as dt
from pathlib import Path

class Outlook:
    ## Mutable initial vars
    output_folder: Path = None    # Downloads here
    language: str = "pt-BR"
    
    ## Immutable initial vars
    def __init__(self, look_account: str):
        """ This class brings tools to work with the Outlook desktop app.
        
        Args:
            look_account (str): The e-mail account that you'll be using.
            
        Params:
            output_folder (Path): Sets a folder to store downloaded files.
            language (str): Sets the language of your outlook app, influences in default params.
        """
        
        self.languages = {
            'pt-BR': {'inbox': "Caixa de Entrada"},
            'en-US': {'inbox': "Inbox"}
        }
        
        try:
            self.outapp = cl.Dispatch('Outlook.Application')    # Open Outlook
            self.nmspce = self.outapp.GetNamespace('MAPI')      # Get Namespace
            self.account = self.nmspce.Folders[look_account]    # Get account
        except:
            raise ValueError(f"Is this: '{look_account}' right?")
    
    def search_emails(self, look_folder: str = None, look_sender: str = None, look_subject: str = None, date_interval: list = []):
        """Args:
            look_folder (str, optional): Target folder to look on. Defaults to standard inbox based on the language.
            look_sender (str, optional): Target sender to look on. Defaults to None.
            look_subject (str, optional): Target Phrase, Word or Charachter to look on. Defaults to None.
            date_interval (list, optional): Target interval of dates to look on. Defaults to [].

        Returns:
            list: All of the emails based on the arguments.
        """
        
        ## Searches the subfolders of a folder
        def sub_folders(folder: cl.CDispatch):
            folders = []
            for sub_folder in folder:
                folders.append(sub_folder)
            
            return folders
        
        ## Searches for the folder by path
        def find_folder(folder_path: str):
            
            cur_folder = self.account.folders
            target = folder_path.split('/')
            
            for folder in target:
                for sub_folder in sub_folders(cur_folder):
                    if sub_folder.Name == folder:
                        if folder == target[len(target)-1]:
                            cur_folder = sub_folder
                        else:
                            cur_folder = sub_folder.folders
            
            return cur_folder
        
        ## returns all the items in a folder
        def folder_items(folder: cl.CDispatch):
            emails = []
            for email in folder.items:
                emails.append(email)
            
            return emails
        
        ## Checks the e-mail for list creation
        def check_sender(mail_sender: str, look_sender: str):
            if look_sender == None:
                email = True
            else:
                email = mail_sender == look_sender
            
            return email
        
        ## Checks the e-mail receivement date
        def check_date(mail_date: cl.CDispatch, look_date_range: list):
            if look_date_range == []:
                date = True
            else:
                mail_date = pd.to_datetime(mail_date, utc=True)
                target_date = dt.date(mail_date.year, mail_date.month, mail_date.day)
                min_date = dt.date(look_date_range[0].year, look_date_range[0].month, look_date_range[0].day)
                max_date = dt.date(look_date_range[1].year, look_date_range[1].month, look_date_range[1].day)
                date = min_date <= target_date <= max_date
                
        
            return date
        
        ## Checks if the e-mail has the str
        def check_subject(mail_subject: cl.CDispatch, look_subject: str):
            
            if look_subject == None:
                subject = True
            elif re.search(look_subject, mail_subject):
                subject = True
            else:
                subject = False
            
            return subject
        
        if look_folder == None:
            look_folder = self.languages[self.language]["inbox"]
        
        folder = find_folder(look_folder)
        emails = [
            email
                for email in folder_items(folder)
                    if
                        check_sender(email.SenderEmailAddress, look_sender) and
                        check_date(email.ReceivedTime, date_interval) and
                        check_subject(email.Subject, look_subject)
        ]
        
        return emails
    
    # TODO: Add exceptions
    def read_email(self, email: cl.CDispatch):
        """Reads the text content of a e-mail

        Args:
            email (cl.CDispatch): A unique e-mail returned from search emails

        Returns:
            List: Content of the body
        """
        content = []
        body = email.Body.split('\n')
        for line in body:
            line = line.replace('\t', "")
            line = line.replace('\r', "")
            if line != '' and line != ' ':
                content.append(line)
            else:
                pass
        return content
    
    def download_attachments(self, emails: list, file_search: str):
        
        ## Exceptions
        if emails == []:
            raise ValueError("list of emails is blank")
        
        ## download file, and return it's name
        file_name = []
        for email in emails:
            attachments = email.Attachments
            for file in attachments:
                if re.search(file_search, file.FileName) != None:
                    file.SaveAsFile(self.output_folder / file.FileName)
                    file_name.append(file.FileName)
                    
        return file_name

    def send_email(self, to: str, subject: str, attachments: list, html_body: str):
        outlook = self.outapp
        email = outlook.CreateItem(0)
        email.To = to
        email.Subject = subject
        for file in range(len(attachments)):
            email.Attachments.Add(str(attachments[file]))
        email.HTMLBody = html_body
        email.Send()
        
        return subject