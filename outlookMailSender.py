import win32com.client

# Handling Outlook
ol = win32com.client.Dispatch('Outlook.Application')

olmailitem = 0x0

newmail = ol.CreateItem(olmailitem)

# Email receiver and content
newmail.Subject = 'Test automatized email'
newmail.To = 'claudio.torres.burgos@gmail.com'
newmail.Body = 'Automatized email content body'

# Attached file
attach = 'C:\\Users\\me\\Desktop\\Test desarrollo script\\node_status.txt'

newmail.Attachments.Add(attach)

# Uncomment to visually display sending email
# newmail.Display()

newmail.Send()