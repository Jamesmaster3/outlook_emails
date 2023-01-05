import win32com.client
from pathlib import Path
from string import Template
import csv

outlook = win32com.client.Dispatch('outlook.application')

with open('name_list.csv', mode='r', newline='') as f:
        file = f.read()
file = file.splitlines()

# use the code below to change certain elements in email message, use a for loop





def create_mail(emailadress, subject, message):
    mail = outlook.CreateItem(0)
    mail.To = emailadress
    mail.Subject = subject
    mail.HTMLBody = message
    return mail

def subsitute_message(pahttohtml, name):
    html = Template(Path(pahttohtml).read_text())
    html = html.substitute({'name': name})
    return html


for i in file:
    print(i)
    if i != 'Name':
        html = subsitute_message('outlook_emails\\body_with_placeholder.html', i)
        mail = create_mail('email@gmail.com', 'Sample', html)
        mail.Display()






# if __name__ == '__main__':
#     print(html)
#     try:
#         mail.Display()
#         # or mail.Send()
#         print('All good!')
#     except:
#         print('error!')
