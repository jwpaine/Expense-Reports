import getpass, imaplib, email, re
import numbers
import decimal




    
def connect():

	username = input("Username: ")
	password = input("Password: ")

	print("Logging into mail server...")
	M = imaplib.IMAP4_SSL('imap.gmail.com', 993)
	M.login(username, password)
	M.select('inbox')

	return M

def get_new_reports(M):
	
	print("Looking at mailbox for unseen reports...")
	(retcode, messages) = M.search(None, '(UNSEEN)', 'SUBJECT "Expense Report"')
	num_messages =  len(messages[0].split())
	print("New Reports: ", num_messages)

	if retcode == 'OK' and num_messages > 0:
		return messages
	else:
		return None


def parse_reports(messages, M):
	n = 1
	for num in messages[0].split():
		print("Fetching report: ", n)
		typ, data = M.fetch(num, '(RFC822)')
		raw_email = data[0][1];
		n = n + 1
		print("Decoding...")
		raw_email_string = raw_email.decode('utf-8')
		# converts byte literal to string removing b''
		email_message = email.message_from_string(raw_email_string)
		
		
		if (email_message.is_multipart()):
			print("multipart")
		else:
			print("not multipart")

		for part in email_message.walk():
                        if part.get_content_type() == "text/plain": # ignore attachments/html
                                body = str(part).replace("Content-Type: text/plain; charset=UTF-8", "");
                                print("Parsing...")
                                process(body);
    i = input("")
	M.close()
	M.logout()

def process(body):

	expense_report = Report();
        # Use a hashmap to hold entries to map account_num --> list [start_date, end_date, amount, vender, requester] all under the same project number
	entries = dict();
	sections = body.split("------------------------------")
	# ----variables global to each report ---
	project_num = '';
	report_name = '';
	employee_name = '';
	last_code = None;
	# ---------------------------------------
	for sub in sections:
	   # If Expense report section, grab project # ------------
		if ("Expense Report" in sub):
			expense_report.project_num = re.findall( r'\d+\.*\d*', sub )[0]
		#	print("Project #", expense_report.project_num)
		   # get Report name by finding index
			i = sub.find("Report Name :") + 14
			j = i
			while(sub[j] != '\n'):
				j = j + 1
			expense_report.report_name = sub[i:j]
		# If Exployee Name section, grab Employee name ----------
		if ("Employee Name :" in sub):
			i = sub.find("Employee Name :")+16
			#get index of 'Employee Name' substring
			j = i
			while(sub[j] != '\n'):
				#concatenate employee name
				j = j + 1
			expense_report.employee_name = sub[i:j]
			#split by transaction
		
		#divide up by expense
		for expense in sub.split("Expense Type"):
			#get account code unique to expense set
			account_code = ''
			result = (re.findall(r"\D(\d{5})\D", expense))

			if result:
				#get account code -------------------------------
				account_code = int(result[0])
				#filter out strange results
				if (account_code < 6000 or account_code > 69999):
					continue
			#	print("Account Code Group:", account_code)

				#get [Expense Type] -----------------------------
				if ("Amount" in expense):
					i = expense.find("Amount") + 20
					j = i
				while (expense[j] != "\n"):
					j = j + 1
				expense_type = expense[i:j]
			#	print("Expense_type: ", expense_type)

				#get allocations -------------------------------
				for transaction in expense.split("Allocations"):
					amount = '';
					i = -1
					j = -1
					i = transaction.find("(")
					j = transaction.find(")");	
					if (i < 0 or j < 0):
						# if amount not found, continue to next iteration
						continue
					#set amount of allocation
					amount = float(transaction[i+2:j])
					#get date
					result = re.search("([0-9]{2}\/[0-9]{2}\/[0-9]{4})", transaction)
					date = ''
					if result is not None:
						date = result.group(1)
				#	print("Transaction Date: ", date)
				#	print("Transaction Amount: ", amount)

					#make entry in dictionary
					expense_report.allocate(account_code, expense_type, date, amount)
	
	print("---------New report entry ready--------")
	print("Project Number: ", expense_report.project_num)
	print("Start Date: ", expense_report.getFirstDate())
	print("Report name: ", expense_report.report_name)
	print("Requester: ", expense_report.employee_name)
	print("Approver: VP Administration")
	for key in expense_report.entries:
		print("Account Code: ", key, ", Date: ", expense_report.entries[key][0], ", Amount: ", expense_report.entries[key][1],  ", Expense Type: ", expense_report.entries[key][2])
	print("---------------------------------------")

class Report():
	print("New Report")
	def __init__(self):
		self.project_num = None    # instance variable unique to each instance
		self.report_name = None
		self.first_date = None
		self.employee_name = None
		self.entries = dict()
	def getFirstDate(self):
		if ("-" in self.report_name): # if date duration
			i = self.report_name.find("-");
			return self.report_name[0:i];
		i = self.report_name.find(" ");
		return self.report_name[0:i];
	def allocate(self, account_code, expense_type, date, amount):
	# check if we've already made an entry under this account code:
		if (account_code in self.entries):
			# we have, so just append the allocation
		#	print("Allocation already made for account code: ", account_code, "Appending new allocation: ", amount)
			self.entries[account_code][1] += amount;
		else:
			# we have not made an entry for this account code
		#	print("No entry present under account code: ", account_code, "adding new entry!")

			self.entries[account_code] = list()
			self.entries[account_code].append(date)
			self.entries[account_code].append(amount)
			self.entries[account_code].append(expense_type)

mailbox = connect();
reports = get_new_reports(mailbox);
# if reports is not None, parse
if(reports):
	parse_reports(reports, mailbox);


			
