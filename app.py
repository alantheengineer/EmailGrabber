import win32com.client
import os
import json
from datetime import datetime, timedelta

# this is the code to get the 'inbox', ['GetDefaultFolder(6)', but other folders are possible]
# filters it by sender, time and subject... depends on the info you have...
# [would also be possible to filter with a subject wildcard like this
# messages = messages.Restrict("@SQL=(urn:schemas:httpmail:subject LIKE '%Sample Report%')")]
# then saves it to the current folder as a JSON file...
def save_report_email_to_JSON():
	outlook = win32com.client.Dispatch('outlook.application')
	mapi = outlook.GetNamespace("MAPI")
	inbox = mapi.GetDefaultFolder(6) 
	messages = inbox.Items
	received_dt = datetime.now() - timedelta(days=1)
	received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
	messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
	messages = messages.Restrict("[SenderEmailAddress] = 'guy@placeThatSendsInfo.com'")
	messages = messages.Restrict("[Subject] = 'Sample Report'")
	outputDir = os.getcwd()
	try:
		for i, message in enumerate(list(messages)):
			try:
				s = message.sender
				with open(os.path.join(outputDir, 'data_{}.json'.format(i)), 'w') as fp:
					json.dump(message.HTMLBody, fp)
					print(f"Email recieved at {message.ReceivedTime}")
			except Exception as e:
				print("error when saving the HTML body:" + str(e))
	except Exception as e:
		print("error when processing email messages:" + str(e))

if __name__ == "__main__":
    save_report_email_to_JSON()