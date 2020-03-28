# criteria.from='sender@example.com'	All emails from sender@example.com
# criteria.size=10485760
# criteria.sizeComparison='larger'	All emails larger than 10MB
# criteria.hasAttachment=true	All emails with an attachment
# criteria.subject='[People with Pets]'	All emails with the string [People with Pets] in the subject
# criteria.query='"my important project"'	All emails containing the string my important project
# criteria.negatedQuery='"secret knock"'	All emails that do not contain the string secret knock

# ParseFromAddresses(r, mr);
# ParseLabelMove(r, mr, storeName);
# ParseLabelCopy(r, mr, storeName);
# ParseSubject(r, mr);
# ParseBody(r, mr);

import win32com.client, csv

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

class GoogleRule:
	def __init__(self):
		self.name = 'GoogleRule'

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

# convert the rule to a G Suite label
def convert_rules(rules):
	print('\n')
	print('Converting Rules...')
	for r in rules:
		print('\n')
		if r.fromAddress is not None:
			print(f"From: {r.fromAddress[0]}")
		print(f"Subject Query: {r.subjectText}")
		print(f"Body Query: {r.bodyText}")
		if r.moveToFolderEnabled:
			print(f"Move Destination: {r.moveToFolderPath}")
		if r.copyToFolderEnabled:
			print(f"Copy Destination: {r.copyToFolderPath}")

def generate_csv(rules):
	with open('OutlookRules.csv','a', newline='') as f1:
		csv_writer = csv.writer(f1, delimiter=',')
		header = ['Rule Name','IsEnabled?','SubjectEnabled','BodyEnabled','MoveToFolderEnabled','CopyToFolderEnabled']
		csv_writer.writerow(header)
		for r in rules:
			print('Writing row....')
			row = [r.name, r.ruleEnabled, r.subjectEnabled, r.bodyEnabled, r.moveToFolderEnabled, r.copyToFolderEnabled]
			csv_writer.writerow(row)


def main():
	outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
	rules = get_rules(outlook)
	show_rules(rules)
	generate_csv(rules)
	convert_rules(rules)

	

if __name__ == '__main__':
	main()


# y = outlook.GetNamespace("MAPI")

# storeList = y.GetStores()
# z = y.DefaultStore
# a = z.GetRules()
# b = a.item('TEST RULE')