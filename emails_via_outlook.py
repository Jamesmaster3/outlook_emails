import win32com.client
from pathlib import Path
from string import Template


outlook = win32com.client.Dispatch('outlook.application')

html = open('outlook_emails\\test_body.html').read()

# use the code below to change certain elements in email message, use a for loop
# html = Template(Path('outlook_emails\\body_with_placeholder.html').read_text())
# html = html.substitute({'name': 'Jim'})

mail = outlook.CreateItem(0)

mail.To = 'jimbentem@gmail.com'
mail.Subject = 'Sample Email'
mail.HTMLBody = html


if __name__ == '__main__':
    print(html)
    try:
        mail.Display()
        # or mail.Send()
        print('All good!')
    except:
        print('error!')
