import win32com.client
import datetime
import time
import os
import zipfile

def check_shared(namespace,recip = None): 
	"""Function takes two arguments:
		.) Names-Space: e.g.: 
			which is set in the following way: outlook = Dispatch("Outlook.Application").GetNamespace("MAPI") and
		.) Recipient of an eventual shared account as string: e.g.: Shared e-Mail adress is "shared@shared.com"
			--> This is optional --> If not added, the standard-e-Mail is read out"""
	
	
	if recip is None:
		for i in range(1,100):                           
			try:
				inbox = namespace.GetDefaultFolder(i)     
				print ("%i %s" % (i,inbox))            
			except:
				print ("%i does not work"%i)
				continue
	else:
		print('The folders from the following shared account will be printed: '+recip)
		tmpRecipient = outlook.CreateRecipient(recip)
		for i in range(1,100):                           
			try:
				inbox = namespace.GetSharedDefaultFolder(tmpRecipient, i)     
				print ("%i %s" % (i,inbox))            
			except:
				#print ("%i does not work"%i)
				continue
	print("Done")
	Recipient = outlook.CreateRecipient(recip)
	inbox = namespace.GetSharedDefaultFolder(Recipient, 6)     
	print ("%i %s" % (i,inbox))    
	messages = inbox.Items
	#message = messages.GetFirst()
	#subject_line = message.subject
	#body_content = message.body
	i = 0
	today = datetime.date.today()
	today_formatted = format(today, '%#m/%#d/%Y')
	print ("Todays Date : ", today_formatted)       # test

	delta = datetime.timedelta(days=7)

	next_week = today + delta
	next_week_formatted = format(next_week, '%#m/%#d/%Y')
	
	print ("Next Weeks Date : ", next_week_formatted )
	
	obtained_todays = False
	obtained_nextweeks = False
	
	message = messages.GetFirst()
	while i < 100:
		to_line = message.to
		#from_line = message.from
		subject_line = message.subject
		body_content = message.body

		request_detail_csv_long = message.body[107:116]
		#print (request_detail_csv_long)
		request_detail_csv_short = message.body[107:115]
		#print (request_detail_csv_short)

		if subject_line == "eSchedule Ingestion Successful" and request_detail_csv_long == today_formatted or request_detail_csv_long == next_week_formatted or request_detail_csv_short == today_formatted or request_detail_csv_short == next_week_formatted:
			print ("---------------------") 
			print ("TO: " + to_line)
			#print ("Request Detail CSV : " + request_detail_csv)
			print ("BODY: " + body_content)
			print ("---------------------") 
			if request_detail_csv_long == today_formatted or request_detail_csv_short == today_formatted:
				try:
					attachments = message.attachments
					attachment = attachments.Item(1)
					if os.path.isfile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\" + str(attachment)):
						print("Today's file was retrieved")
						obtained_todays = True
					else:
						attachment.SaveAsFile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\" + str(attachment))
						with zipfile.ZipFile("K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\" + str(attachment),"r") as zip_ref:
							zip_ref.extractall(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files")
						obtained_todays = True
				except Exception as e: 
					print (e)
					print ("Trouble retrieving todays file")
			if request_detail_csv_long == next_week_formatted or request_detail_csv_short == next_week_formatted:
				
				try:
					attachments = message.attachments
					attachment = attachments.Item(1)
					if os.path.isfile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\7 Days Files\\" + str(attachment)):
						print("Next Week's file was retrieved")
						obtained_nextweeks = True
					else:
						attachment.SaveAsFile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\7 Days Files\\" + str(attachment))
						with zipfile.ZipFile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\7 Days Files\\" + str(attachment),"r") as zip_ref:
							zip_ref.extractall(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\7 Days Files")
						obtained_nextweeks = True
				except Exception as e: 
					print (e)
					print ("Trouble retrieving 7 day file")
			
			if obtained_todays and obtained_nextweeks:
				print ("All Files Retrieved")
				return
			else:
				print("Only Detected One File")
		message = messages.GetNext()
		i += 1

		
def check_shared_monday(namespace,recip = None): 
	"""Function takes two arguments:
		.) Names-Space: e.g.: 
			which is set in the following way: outlook = Dispatch("Outlook.Application").GetNamespace("MAPI") and
		.) Recipient of an eventual shared account as string: e.g.: Shared e-Mail adress is "shared@shared.com"
			--> This is optional --> If not added, the standard-e-Mail is read out"""
	
	
	if recip is None:
		for i in range(1,100):                           
			try:
				inbox = namespace.GetDefaultFolder(i)     
				print ("%i %s" % (i,inbox))            
			except:
				#print ("%i does not work"%i)
				continue
	else:
		print('The folders from the following shared account will be printed: '+recip)
		tmpRecipient = outlook.CreateRecipient(recip)
		for i in range(1,100):                           
			try:
				inbox = namespace.GetSharedDefaultFolder(tmpRecipient, i)     
				print ("%i %s" % (i,inbox))            
			except:
				#print ("%i does not work"%i)
				continue
	print("Done")
	Recipient = outlook.CreateRecipient(recip)
	inbox = namespace.GetSharedDefaultFolder(Recipient, 6)     
	print ("%i %s" % (i,inbox))    
	messages = inbox.Items
	#message = messages.GetFirst()
	#subject_line = message.subject
	#body_content = message.body
	i = 0
	today = datetime.date.today()
	today_formatted = format(today, '%#m/%#d/%Y')
	print ("Todays Date : ", today_formatted)       # test
	
	sunday_delta = datetime.timedelta(days=1)
	sunday = today - sunday_delta
	sunday_formatted = format(sunday, '%#m/%#d/%Y')
	print ("Sunday Date: " + sunday_formatted )
	
	saturday_delta = datetime.timedelta(days=2)
	saturday = today - saturday_delta
	saturday_formatted = format(saturday, '%#m/%#d/%Y')
	print ("Saturday Date: " + saturday_formatted )
	
	next_week_delta = datetime.timedelta(days=7)
	next_week = today + next_week_delta
	next_week_formatted = format(next_week, '%#m/%#d/%Y')
	print ("Next Weeks Date : ", next_week_formatted )
	
	obtained_todays = False
	obtained_sundays = False
	obtained_saturdays = False
	obtained_nextweeks = False
	
	message = messages.GetFirst()
	while i < 100:
		
		to_line = message.to
		#from_line = message.from
		subject_line = message.subject
		body_content = message.body

		request_detail_csv_long = message.body[107:116]
		#print (request_detail_csv_long)
		request_detail_csv_short = message.body[107:115]
		#print (request_detail_csv_short)

		if subject_line == "eSchedule Ingestion Successful" and request_detail_csv_long == today_formatted or request_detail_csv_long == next_week_formatted or request_detail_csv_short == today_formatted or request_detail_csv_short == next_week_formatted:
			print ("---------------------") 
			print ("TO: " + to_line)
			#print ("Request Detail CSV : " + request_detail_csv)
			print ("BODY: " + body_content)
			print ("---------------------") 
			if request_detail_csv_long == today_formatted or request_detail_csv_short == today_formatted:
				try:
					attachments = message.attachments
					attachment = attachments.Item(1)
					if os.path.isfile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\" + str(attachment)):
						print("Today's file was retrieved")
						obtained_todays = True
						
					else:
						attachment.SaveAsFile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\" + str(attachment))
						with zipfile.ZipFile("K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\" + str(attachment),"r") as zip_ref:
							zip_ref.extractall(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files")
						obtained_todays = True
				except Exception as e: 
					print (e)
					print ("Trouble retrieving todays file")
			if request_detail_csv_long == sunday_formatted or request_detail_csv_short == sunday_formatted:
				try:
					attachments = message.attachments
					attachment = attachments.Item(1)
					if os.path.isfile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\Files for loading\\" + str(attachment)):
						print("Sunday's file was retrieved")
						obtained_sundays = True
						
					else:
						attachment.SaveAsFile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\Files for loading\\" + str(attachment))
						obtained_sundays = True
				except:
					print ("Trouble retrieving sundays file")
			if request_detail_csv_long == saturday_formatted or request_detail_csv_short == saturday_formatted:
				try:
					attachments = message.attachments
					attachment = attachments.Item(1)
					if os.path.isfile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\Files for loading\\" + str(attachment)):
						print("Saturday's file was retrieved")
						obtained_saturdays = True
						
					else:
						attachment.SaveAsFile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\Today Files\\Files for loading\\" + str(attachment))
						obtained_sundays = True
				except:
					print ("Trouble retrieving saturdays file")
			if request_detail_csv_long == next_week_formatted or request_detail_csv_short == next_week_formatted:
				try:
					attachments = message.attachments
					attachment = attachments.Item(1)
					if os.path.isfile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\7 Days Files\\" + str(attachment)):
						print("Next Week's file was retrieved")
						obtained_nextweeks = True
						
					else:
						attachment.SaveAsFile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\7 Days Files\\" + str(attachment))
						with zipfile.ZipFile(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\7 Days Files\\" + str(attachment),"r") as zip_ref:
							zip_ref.extractall(r"K:\\OTS\\OTSPLNG\\Stores Finance\\Labor\\eSchedule Data\\eSchedule New Format\\7 Days Files")
						obtained_nextweeks = True
				except Exception as e: 
					print (e)
					print ("Trouble retrieving 7 day file")
			
			if obtained_todays and obtained_sundays and obtained_saturdays and obtained_nextweeks:
				print ("All Files Retrieved")
				return	
				
		message = messages.GetNext()	
		i += 1

account = "LCReports@luxotticaretail.com"	
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# check the day of the week. If day of the week is monday run the monday version of the script
# monday is 0, sunday is 6
weekday_number = datetime.datetime.today().weekday()

# Set up email check persistence. This will keep running for an hour (range(10) * time.sleep(360) <-- 6 minutes) before officially shutting off

for attempt in range(10):
	print ("Attempt: " + str(attempt))
	try:
		if weekday_number == 0:
			print ("It's monday!")
			check_shared_monday(outlook,account)
		else:
			print("Not monday")
			check_shared(outlook,account)
	except:
		print("Nothing detected, waiting six minutes....")
		time.sleep(360)
	else:
		# script failed/didn't detect new email
		print ("##################")
		print ("Shutting Off Script")
		break




