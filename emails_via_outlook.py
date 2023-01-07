import win32com.client
from pathlib import Path
from string import Template
import pandas as pd
import pathlib

path = pathlib.Path('name_list.csv')

outlook = win32com.client.Dispatch('outlook.application')

file = pd.read_csv('name_list.csv')

# input paths to attachment_files as args, use path from source
def create_mail(emailadress, subject, message, *args):
    mail = outlook.CreateItem(0)
    mail.To = emailadress
    mail.Subject = subject
    mail.HTMLBody = message
    if args:
        attachment_paths = list(args)
        for i in attachment_paths:
            mail.Attachments.Add(i)
    return mail

# creates message by substituting required fields
def subsitute_message(pahttohtml, name):
    html = Template(Path(pahttohtml).read_text())
    html = html.substitute({'name': name})
    return html

if __name__ == '__main__':
    for i in file.values:
        try:
            html = subsitute_message('outlook_emails\\formatted_body.html', i[0])
            mail = create_mail(i[1], 'Sample', html)
            mail.Display()
        except ValueError as err:
            print(err)
        except TypeError as err:
            print(err)
        except IndexError as err:
            print(err)