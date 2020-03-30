# [+]google api notes [+]
# criteria.from='sender@example.com'	All emails from sender@example.com
# criteria.size=10485760
# criteria.sizeComparison='larger'	All emails larger than 10MB
# criteria.hasAttachment=true	All emails with an attachment
# criteria.subject='[People with Pets]'	All emails with the string [People with Pets] in the subject
# criteria.query='"my important project"'	All emails containing the string my important project
# criteria.negatedQuery='"secret knock"'	All emails that do not contain the string secret knock

import win32com.client, csv, pickle, os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/gmail.settings.basic']


class OutlookRule:
	def __init__(self, rule):
		self.name = rule.Name
		self.ruleEnabled = rule.Enabled
		self.subjectEnabled = rule.Conditions.Subject.Enabled
		self.subjectText = rule.Conditions.Subject.Text[0] if self.subjectEnabled else 'None'
		self.bodyEnabled = rule.Conditions.Body.Enabled
		self.fromAddress = rule.Conditions.SenderAddress.Address
		self.bodyText = rule.Conditions.Body.Text if self.bodyEnabled else 'None'
		self.moveToFolderEnabled = rule.Actions.MoveToFolder.Enabled
		self.moveToFolderPath = rule.Actions.MoveToFolder.Folder.FolderPath if self.moveToFolderEnabled else 'None'
		self.copyToFolderEnabled = rule.Actions.CopyToFolder.Enabled
		self.copyToFolderPath = rule.Actions.CopyToFolder.Folder.FolderPath if self.copyToFolderEnabled else 'None'
		self.size = '0'
		self.sizeCompare = '='
		self.hasAttachment = 'None'
		self.query = rule.Conditions.Body.Text
		self.negatedQuery = 'None'

def get_rules(outlook):
	ruleList = []
	rules = outlook.Session.DefaultStore.GetRules() 
	ruleList = []
	i = 1
	while i < rules.Count+1:
		ruleList.append(OutlookRule(rules.Item(i)))
		i += 1
	return ruleList

def show_rules(rules):
	print(f"\nNumber of rules: {len(rules)}")
	for r in rules:
		print('\n')
		print(f"Rule: {rules.index(r)+1}")
		print(f"Name: {r.name}")
		print(f"Enabled: {r.ruleEnabled}")
		print(f"FromAddress: {r.fromAddress}")
		if r.bodyEnabled:
			print(f"BodyEnabled: {r.bodyEnabled}")
			print(f"ConditionBody: {r.bodyText}")
		if r.subjectEnabled:
			print(f"SubjectEnabled: {r.subjectEnabled}")
			print(f"SubjectText: {r.subjectText}")
		if r.moveToFolderEnabled:
			print(f"MoveToFolder: {r.moveToFolderEnabled}")
			print(f"FolderPath: {r.moveToFolderPath}")
		if r.copyToFolderEnabled:
			print(f"CopyToFolder: {r.copyToFolderEnabled}")
			print(f"Path: {r.copyToFolderPath}")
		print(f"Size: {r.size}")
		print(f"SizeOperator: {r.sizeCompare}")
		print(f"HasAttachment: {r.hasAttachment}")

def create_filter(rules):
	service = generate_token()
	for r in rules:
		print(f"Converting rule {rules.index(r)+1}")
		label_id = r.name # ID of user label to add
		label_list = []
		filter = {
		'criteria': {},
		'action': {}
		}
		if r.fromAddress != None:
			filter['criteria']['from'] = r.fromAddress[0]
		if r.subjectEnabled:
			filter['criteria']['subject'] = r.subjectText
		if r.bodyEnabled:
			filter['criteria']['query'] = r.bodyText
		if r.moveToFolderEnabled:
			filter['action']['removeLabelIds'] = ['INBOX']
			label_list.append(r.moveToFolderPath.split('\\').pop())
		if r.copyToFolderEnabled:
			label_list.append(r.copyToFolderPath.split('\\').pop())
		filter['action']['addLabelIds'] = label_list
		print(filter)
		# result = service.users().settings().filters().create(userId='me', body=filter).execute()
		# print(f"Created filter: {result.get('id')}")


def generate_csv(rules):
	with open('OutlookRules.csv','a', newline='') as f1:
		csv_writer = csv.writer(f1, delimiter=',')
		header = ['Rule Name','IsEnabled?','SubjectEnabled','BodyEnabled','MoveToFolderEnabled','CopyToFolderEnabled']
		csv_writer.writerow(header)
		print('\n')
		for r in rules:
			print('Writing row....')
			row = [r.name, r.ruleEnabled, r.subjectEnabled, r.bodyEnabled, r.moveToFolderEnabled, r.copyToFolderEnabled]
			csv_writer.writerow(row)

def generate_token():
	creds = None
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
			creds = flow.run_local_server(port=0)
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)
	service = build('gmail', 'v1', credentials=creds)
	return service

def main():
	outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
	rules = get_rules(outlook)
	show_rules(rules)
	generate_csv(rules)
	create_filter(rules)

if __name__ == '__main__':
	main()
