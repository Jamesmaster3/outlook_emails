# outlook_emails
Script using the win32 library to send out multiple emails automatically via outlook locally. This version supports substiution of the recipients name or names. This program is especially usefull for corporate networks for example, where you cannot send emails via SMTP.


HOW TO USE
1. Create a template email in outlook and save it as html in this directory.
2. Input the name(s), emails, and optional CC emails in the csv file. An email can be sent to multiple emailadresses by seperating them with an ';'.
3. Attachments can be added to an email by adding to the directory 'bijlagen'.
4. Run the script and press the send button on the emails or save them as drafts