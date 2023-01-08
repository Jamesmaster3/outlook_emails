import win32com.client
from pathlib import Path
from string import Template
import pandas as pd
import pathlib

outlook = win32com.client.Dispatch('outlook.application')

file = pd.read_csv('bijlagen\\name_list.csv')
file = file.fillna(0)

# Create email object and add email, CC emails, subject, message and (multiple) attachments [input path to directory]
def create_mail(email_to, subject, message, attachment_dir=None, email_cc=None):
    mail = outlook.CreateItem(0)
    mail.To = email_to
    mail.Subject = subject
    mail.HTMLBody = message

    if attachment_dir: #check if attachment is added
        # create absolute paths for files and loops trough them to add to mail object
        p = pathlib.Path(attachment_dir).absolute() 
        attachments = list(p.iterdir())
        for child in attachments:
            print(child)
        mail.Attachments.Add(str(child))     

    if email_cc:
        if str(email_cc): #check if value in CC is not empty
            mail.CC = email_cc

    return mail

# creates message by substituting required fields
def subsitute_message(pahttohtml, name):
    html = Template(Path(pahttohtml).read_text())
    html = html.substitute({'name': name})
    return html

if __name__ == '__main__':
    for i in file.values:
        try:
            html = subsitute_message('formatted_body.html', i[0])
            mail = create_mail(email_to=i[1], subject='Sample', message=html, attachment_dir='bijlagen', email_cc=i[2])
            mail.Display()

        except ValueError as err:
            print(err)
        except TypeError as err:
            print(err)
        except IndexError as err:
            print(err)