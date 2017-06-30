import cgitb, cgi, os
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui
import sys
import signal
import subprocess
import csv
import re
import time
from itertools import islice
import cgi
import os
import urllib
import datetime
from datetime import datetime as dt
import md5
import cgitb
cgitb.enable()

form = cgi.FieldStorage()

print "Content-type: text/html"
print

print """<HTML>
<HEAD>
<TITLE>Automatically Populate Form</TITLE>
</HEAD>
<BODY style="background-color: #ededed; font-family: tahoma;">""" 

print '<div style="margin: 0 auto; border: 1px solid #3986E2; border-radius: 10px; width: 75%; padding: 15px; background-color: #B3E1F1;">'
print '<h2 style="text-align: center;">Automatically Populate Form</h2>'
print '<h2 style="text-align: center;">How To Use</h2>'
print '1. Upload two .csv files.  The first CSV file should contain the various Jobs data.  The second contains the WebForm data.  (<a href="../samplefiles/Example1.csv" target="_blank">First Sample File</a>) (<a href="../samplefiles/Example2.csv" target="_blank">Second Sample File</a>)<br>' 
print '2. Enter your email address. A link to the output of the file will be emailed to this address when the script completes.<br>'
print '3. Click Submit. You will receive a pop-up window, then you may exit the webpage.<p>'



if form.has_key('file1') and form.has_key('file2') and form.has_key('emailTo') is True:
	fileitem1 = form['file1']
	fileitem2 = form['file2']
	fileitem1.filename = os.path.basename(fileitem1.filename)
	fileitem2.filename = os.path.basename(fileitem2.filename)

	# variables to check if user uploaded correct file types
	typeCheck1 = fileitem1.filename
	typeCheck2 = fileitem2.filename

	if not typeCheck1 and not typeCheck2:
		print '<script>alert("Please enter files of type CSV")</script>'
	elif typeCheck1.split('.')[1] != "csv" or typeCheck2.split('.')[1] != "csv":
		print '<script>alert("Please enter files of type CSV")</script>'
	else:

		# validation checks
		def checkFields(requiredHeaders,headerList,headerList2):
			reqHeadSet = set(requiredHeaders)
			headSet1 = set(headerList)
			headSet2 = set(headerList2)
			both = headSet1.union(headSet2)
			if reqHeadSet.issubset(both):
				return True
			else:
				return False

		def checkEmail(email):
			strEmail = str(email)
			if "@" not in strEmail:
				return False
			else:
				return True

		# each input area on the form is populated by the following functions
		# these functions are associated with file input #1 (multiple jobs) and listed in order that they appear in the form
		def JobType(titleList,valueList):
			index = None
			try:
				index = titleList.index("JobType")
				v = valueList[index]                                                                            

				select = Select(driver.find_element_by_name("JobType"))
				select.select_by_value(v)
			except ValueError:
				print "Job Type value not found"
		
		def workstation(titleList,valueList):
			index = None
			try:
				index = titleList.index("ExecutionServerName")
				v = valueList[index]

				tab = driver.find_element_by_name("Server").click()
				advanced_routing=ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_element_by_name("Server")
				)
				advanced_routing.click()
				element = ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_elements_by_name("Server")
				)
				element[0].send_keys(v)
			except ValueError:
				print "Workstation/Server Name value not found"

		def scriptName(titleList,valueList):
			index = None
			try:
				index = titleList.index("JobName")
				v = valueList[index]

				if v is not None:
					tab = driver.find_element_by_name("Job").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Job")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Job")
					)
					element[0].send_keys(v)
				else:
					tab = driver.find_element_by_name("Job").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Job")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Job")
					)
					element[0].send_keys("Default Script Name")
			except ValueError:
				print "Script Name value not found"

		def JobStreamName(titleList,valueList):
			index = None
			try:
				index = titleList.index("JobStreamName")
				v = valueList[index]

				tab = driver.find_element_by_name("JobStreamName").click()
				advanced_routing=ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_element_by_name("JobStreamName")
				)
				advanced_routing.click()
				element = ui.WebDriverWait(driver, 10).until(
				lambda driver : driver.find_elements_by_name("JobStreamName")
				)
				element[0].send_keys(v)
			except ValueError:
				print "Job Stream Name value not found"
		
		def JobDesc(titleList,valueList):
			index = None
			try:
				index = titleList.index("JobDesc")
				v = valueList[index]

				tab = driver.find_element_by_name("JobDescription").click()
				advanced_routing=ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_element_by_name("JobDescription")
				)
				advanced_routing.click()
				element = ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_elements_by_name("JobDescription")
				)
				element[0].send_keys(v)
			except ValueError:
				print "Job Description value not found"
		
		def scriptPath(titleList,valueList):
			folderIndex = None
			jobIndex = None
			try:
				folderIndex = titleList.index("FolderName")
				jobIndex = titleList.index("JobName")
				folderVal = valueList[folderIndex]
				jobVal = valueList[jobIndex]
				inputStr = folderVal + " " + jobVal

				if inputStr is not None:
					tab = driver.find_element_by_name("Script").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Script")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Script")
					)
					element[0].send_keys(inputStr)
			except:
				print "Folder Name or Job Name value(s) not found"

		def Predecessor(titleList,valueList):
			clear = driver.find_element_by_name("Predecessor").clear()
			regex = r"\APredecessor\d"
			testStr = "Predecessor5"
			string = ", "
			firstClick = 0
			strNone = "None"
			matchList = []

			for i in titleList:
				matches = re.findall(regex, i)
				if matches != []:
					index = titleList.index(matches[0])
					v = valueList[index]

					if not firstClick:
						tab = driver.find_element_by_name("Predecessor").click()
						advanced_routing=ui.WebDriverWait(driver, 10).until(
							lambda driver : driver.find_element_by_name("Predecessor")
						)
						advanced_routing.click()
						element = ui.WebDriverWait(driver, 10).until(
							lambda driver : driver.find_elements_by_name("Predecessor")
						)
						firstClick = 1
					element[0].send_keys(v)
					if v:
						element[0].send_keys(string)

			element = driver.find_element_by_name("Predecessor")
			strValue = element.get_attribute("value")
			if not strValue:
				tab = driver.find_element_by_name("Predecessor").click()
				advanced_routing=ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_element_by_name("Predecessor")
				)
				advanced_routing.click()
				element = ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_elements_by_name("Predecessor")
				)
				element[0].send_keys(strNone)

		def Successor(titleList,valueList):
			clear = driver.find_element_by_name("Successor").clear()
			regex = r"\ASuccessor\d"
			string = ", "
			firstClick = 0
			strNone = "None"
			matchList = []

			for i in titleList:
				matches = re.findall(regex, i)

				if matches != []:
					index = titleList.index(matches[0])
					v = valueList[index]

					if not firstClick:
						tab = driver.find_element_by_name("Successor").click()
						advanced_routing=ui.WebDriverWait(driver, 10).until(
							lambda driver : driver.find_element_by_name("Successor")
						)
						advanced_routing.click()
						element = ui.WebDriverWait(driver, 10).until(
							lambda driver : driver.find_elements_by_name("Successor")
						)
						firstClick = 1
					element[0].send_keys(v)
					if v:
						element[0].send_keys(string)

			element = driver.find_element_by_name("Successor")
			strValue = element.get_attribute("value")
			if not strValue:
				tab = driver.find_element_by_name("Successor").click()
				advanced_routing=ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_element_by_name("Successor")
				)
				advanced_routing.click()
				element = ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_elements_by_name("Successor")
				)
				element[0].send_keys(strNone)
		
		def fileDependency(titleList,valueList):
			index = None
			try:
				index = titleList.index("FileDependency")
				v = valueList[index]

				if v is not None:
					clear = driver.find_element_by_name("FileDep").clear()
					tab = driver.find_element_by_name("FileDep").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("FileDep")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("FileDep")
					)
					element[0].send_keys(v)
			except ValueError:
				print "File Dependency value not found"

		def startTime(titleList,values):
			index = None
			try:
				index = titleList.index("StartTime")
				v = str(values[index])
				temp = v.split(",")
				hr = temp[0]
				mm = temp[1]
				ampm = temp[2]
				zone = temp[3]

				selHour = Select(driver.find_element_by_name("TimeHH"))
				selHour.select_by_visible_text(hr)

				selMin = Select(driver.find_element_by_name("TimeMM"))
				selMin.select_by_visible_text(mm)

				selAMPM = Select(driver.find_element_by_name("TimeAMPM"))
				selAMPM.select_by_visible_text(ampm)

				selZone = Select(driver.find_element_by_name("TimeZone"))
				selZone.select_by_visible_text(zone)
			except ValueError:
				print "Could not find Start Time value"

		def jobInterval(titleList,valueList):
			index = None
			try:
				index = titleList.index("JobInterval")
				v = valueList[index]
				select = Select(driver.find_element_by_name("JobInterval"))
				select.select_by_visible_text(v)
			except ValueError:
				print "Job Interval value not found"

		def Frequency(titleList,valueList):
			index = None
			try:
				index = titleList.index("Frequency")
				temp = valueList[index]
				vList = temp.strip().split(",")

				nameList = ['mon','tues','wed','thurs','fri','sat','sun']
				for i, item in enumerate(vList):
					if item == 'Y':
						x = nameList[i]
						tab = driver.find_element_by_name(x).click()
			except ValueError:
				print "Could not find Frequency value"

		def otherFreq(titleList,valueList):
			index = None
			try:
				index = titleList.index("OtherFrequency")
				v = valueList[index]

				if v is not None:
					clear = driver.find_element_by_name("OtherFreq").clear()
					tab = driver.find_element_by_name("OtherFreq").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("OtherFreq")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("OtherFreq")
					)
					element[0].send_keys(v)
			except ValueError:
				print "Other Frequency value not found"
		# end functions for file input #1

		# start functions for file input #2 (repeating data) and listed in order that they appear in the form
		def requestedBy(titleList,valueList):
			index = None
			try:
				index = titleList.index("RequestedBy")
				v = valueList[index]

				tab = driver.find_element_by_name("RequestedBy").click()
				advanced_routing=ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_element_by_name("RequestedBy")
				)
				advanced_routing.click()
				element = ui.WebDriverWait(driver, 10).until(
					lambda driver : driver.find_elements_by_name("RequestedBy")
				)
				element[0].send_keys(v)
			except ValueError:
				print "Requested By value not found"

		def crq(titleList,valueList):
			index = None
			try:
				index = titleList.index("CRQ#")
				v = valueList[index]

				if v is not None:
					tab = driver.find_element_by_name("Vantive").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Vantive")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Vantive")
					)
					element[0].send_keys(v)
				else:
					tab = driver.find_element_by_name("Vantive").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Vantive")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Vantive")
					)
					element[0].send_keys("1234")
			except ValueError:
				print ("Could not find CRQ#")

		def goLive(titleList,valueList):
			# initialize variables
			index = None
			inputVal = None
			todayDate = None

			# get today's date and convert to string format for concatenation
			now = datetime.datetime.now()
			str(now)
			strMonth = str(now.month)
			strDay = str(now.day)
			strYear = str(now.year)
			tempNow = strMonth + "/" + strDay + "/" + strYear
			try:
				index = titleList.index("GoLiveDate")
				v = str(valueList[index])

				# convert input value date and today's date to datetime.datetime for easy comparison
				inputVal = dt.strptime(v, "%m/%d/%Y")
				todayDate = dt.strptime(tempNow, "%m/%d/%Y")
				
				# comparison to determine what value is inserted in the field
				if inputVal >= todayDate:
					tab = driver.find_element_by_name("LastUpdate").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("LastUpdate")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("LastUpdate")
					)
					element[0].send_keys(v)
				if todayDate > inputVal:
					l = str(todayDate)
					currDate = l.split("-")
					curYr = currDate[0]
					curMo = currDate[1]
					temp = currDate[2].split(" ")
					curDay = temp[0]
					strCurDate = curMo + "/" + curDay + "/" + curYr
					tab = driver.find_element_by_name("LastUpdate").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("LastUpdate")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("LastUpdate")
					)
					element[0].send_keys(strCurDate)
			except Exception as e:
				print e

		def workstationType(titleList,valueList):
			index = None
			try:
				index = titleList.index("WorkStationType")
				v = valueList[index]
				val = v.lower()
				prod = "production".lower()
				dev = "development".lower()
				if re.match(prod, val) is not None:
					tab = driver.find_element_by_xpath("/html/body/form/table[2]/tbody/tr[2]/td[2]/div/font/input").click()
				else:
					tab = driver.find_element_by_xpath("/html/body/form/table[2]/tbody/tr[3]/td/div/font/input").click()
			except ValueError:
				print "Workstation Type value not found"

		def decomission(titleList,valueList):
			index = None
			try:
				index = titleList.index("DecomissionDate")
				v = valueList[index]

				if v is not None:
					clear = driver.find_element_by_name("DecommissionDate").clear()
					tab = driver.find_element_by_name("DecommissionDate").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("DecommissionDate")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("DecommissionDate")
					)
					element[0].send_keys(v)
			except ValueError:
				print "Decommission Date value not found"

		def region(titleList,valueList):
			index = None
			try:
				index = titleList.index("Region")
				v = valueList[index]
				selJobReg = Select(driver.find_element_by_name("JobRegion"))
				selJobReg.select_by_visible_text(v)
			except ValueError:
				print "Region value not found"

		def instance(titleList,valueList):
			index = None
			try:
				index = titleList.index("Instance")
				v = valueList[index]
				select = Select(driver.find_element_by_name("JobInstance"))
				select.select_by_visible_text(v)
			except ValueError:
				print "Job Instance value not found"
		
		def app(titleList,valueList):
			index = None
			try:
				index = titleList.index("Application")
				v = valueList[index]
				select = Select(driver.find_element_by_name("JobApp"))
				select.select_by_visible_text(v)
			except ValueError:
				print "Job Application value not found"

		def busImpact(titleList,valueList):
			index = None
			try:
				index = titleList.index("BusImpact")
				v = valueList[index]

				if v is not None:
					tab = driver.find_element_by_name("BusinessImpact").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("BusinessImpact")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("BusinessImpact")
					)
					element[0].send_keys(v)
				else:
					tab = driver.find_element_by_name("BusinessImpact").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("BusinessImpact")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("BusinessImpact")
					)
					element[0].send_keys("None")
			except ValueError:
				print "Business Impact value not found"

		def patSafety(titleList,valueList):
			index = None
			try:
				index = titleList.index("PatSafetyInfo")
				v = valueList[index]

				if v is not None:
					tab = driver.find_element_by_name("PatientSafety").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("PatientSafety")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("PatientSafety")
					)
					element[0].send_keys(v)
				else:
					tab = driver.find_element_by_name("PatientSafety").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("PatientSafety")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("PatientSafety")
					)
					element[0].send_keys("None")
			except ValueError:
				print "Patient Safety value not found"

		def login(titleList,valueList):
			index = None
			try:
				index = titleList.index("Login")
				v = valueList[index]

				if v is not None:
					tab = driver.find_element_by_name("Login").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Login")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Login")
					)
					element[0].send_keys(v)
				else:
					tab = driver.find_element_by_name("Login").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Login")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Login")
					)
					element[0].send_keys("Default Login")
			except ValueError:
				print "Login value not found"

		def jobRecoveryOpt(titleList,valueList):
			index = None
			try:
				index = titleList.index("RecoveryOpt")
				v = valueList[index].lower()

				select = Select(driver.find_element_by_name("Restart"))
				if v == 'rerun':
					select.select_by_value("Rerun")
				elif v == 'continue':
					select.select_by_value("Continue")
				else:
					select.select_by_value("Stop")
			except ValueError:
				print "Job Recovery Option value not found"

		def acctRetCodes(titleList,valueList):
			index = None
			try:
				index = titleList.index("AcctReturnCodes")
				v = valueList[index]
				
				if v is not None:
					clear = driver.find_element_by_name("RetCode").clear()
					tab = driver.find_element_by_name("RetCode").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("RetCode")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("RetCode")
					)
					element[0].send_keys(v)
			except ValueError:
				print "Acceptable Return Code value not found"

		def escalation(titleList,valueList):
			index = None
			try:
				index = titleList.index("Escalation")
				v = valueList[index].lower()
				v = v.replace(',', '')
				temp = v.split(" ")

				select = Select(driver.find_element_by_name("Escalation"))

				if temp[0] == 'immediately':
					select.select_by_value("Immediately")
				elif temp[1] == 'day':
					select.select_by_visible_text("Next Day, after 8:30 AM")
				else:
					select.select_by_visible_text("Next Business Day, after 8:30 AM")
			except ValueError:
				print "Could not find escalation value"

		def marr(titleList,valueList):
			index = None
			try:
				index = titleList.index("MARRInst")
				v = valueList[index]

				if v is not None:
					tab = driver.find_element_by_name("RRE").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("RRE")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("RRE")
					)
					element[0].send_keys(v)
				else:
					tab = driver.find_element_by_name("RRE").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("RRE")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("RRE")
					)
					element[0].send_keys("Contact ON CALL person")
			except ValueError:
				print "Monitoring/Abend/Recovery/Rerun Instructions value not found"

		def priority(titleList,valueList):
			index = None
			try:
				index = titleList.index("Priority")
				v = valueList[index]
				select = Select(driver.find_element_by_name("Priority"))
				select.select_by_visible_text(v)
			except ValueError:
				print "Priority value not found"

		def primSupport(titleList,valueList):
			index = None
			try:
				index = titleList.index("PrimSupGroup")
				v = valueList[index]

				if v is not None:
					clear = driver.find_element_by_name("Support2").clear()
					tab = driver.find_element_by_name("Support2").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Support2")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Support2")
					)
					element[0].send_keys(v)
				else:
					tab = driver.find_element_by_name("Support2").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("Support2")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("Support2")
					)
					element[0].send_keys("Default Support Group")
			except ValueError:
				print "Primary Support Group value not found"

		def timeDependency(titleList,valueList):
			index = None
			try:
				index = titleList.index("TimeDependencyLevel")
				v = valueList[index]
				val = v.lower()

				if val == "job":
					select = Select(driver.find_element_by_name("TimeDependencyLevel"))
					select.select_by_visible_text("Job")
			except ValueError:
				print "Time Dependency value not found"

		def holiday(titleList,valueList):
			index = None
			try:
				index = titleList.index("HolidayProc")
				v = valueList[index]
				select = Select(driver.find_element_by_name("Holiday"))
				select.select_by_visible_text(v)
			except ValueError:
				print "Holiday value not found"
				
		def specialInstr(titleList,valueList):
			index = None
			strTemp = "None"
			try:
				index = titleList.index("SpecialInstr")
				v = valueList[index]

				if v is not None:
					clear = driver.find_element_by_name("SpecialInstructions").clear()
					tab = driver.find_element_by_name("SpecialInstructions").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("SpecialInstructions")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("SpecialInstructions")
					)
					element[0].send_keys(v)
				else:
					tab = driver.find_element_by_name("SpecialInstructions").click()
					advanced_routing=ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_element_by_name("SpecialInstructions")
					)
					advanced_routing.click()
					element = ui.WebDriverWait(driver, 10).until(
						lambda driver : driver.find_elements_by_name("SpecialInstructions")
					)
					element[0].send_keys(strTemp)
			except ValueError:
				print "Special Instructions value not found"

		def save(titleList,valueList):
			index = None
			try:
				index = titleList.index("SaveOption")
				v = valueList[index].lower()
				val = v.partition(' ')[0]
				if val == 'assigned':
					tab = driver.find_element_by_xpath("/html/body/form/table[6]/tbody/tr[2]/td/font/input[2]").click()
				else:
					tab = driver.find_element_by_xpath("/html/body/form/table[6]/tbody/tr[2]/td/font/input[1]").click()
			except ValueError:
				print "Save Option value not found"

		def getJobName(titleList,valueList):
			index = titleList.index("JobName")
			v = valueList[index]
			return str(v)

		def getReqNum():
			content = driver.find_element_by_xpath("/html/body/form/table/tbody/tr[1]/td/div/p/font[1]/font").get_attribute('innerHTML')
			tmp = content.split("<")
			return str(tmp[0])

		# calls all functions associated with file #2
		def getList2(headerlist,numrows2):
			for j, item in enumerate(contentList2[1:numrows2]):
				requestedBy(headerList2,item)
				crq(headerList2,item)
				goLive(headerList2,item)
				workstationType(headerList2,item)
				decomission(headerList2,item)
				region(headerList2,item)
				instance(headerList2,item)
				app(headerList2,item)
				busImpact(headerList2,item)
				patSafety(headerList2,item)
				login(headerList2,item)
				jobRecoveryOpt(headerList2,item)
				acctRetCodes(headerList2,item)
				escalation(headerList2,item)
				marr(headerList2,item)
				priority(headerList2,item)
				primSupport(headerList2,item)
				timeDependency(headerList2,item)
				holiday(headerList2,item)
				save(headerList2,item)
			driver.find_element_by_name("Submit").click()

		requiredHeaders = ["RequestedBy","CRQ#","GoLiveDate","JobType","WorkStationType","ExecutionServerName", "Region","Instance","Application","JobName","JobStreamName","JobDesc","BusImpact","PatSafetyInfo","Login","RecoveryOpt","Escalation","MARRInst","Priority","PrimSupGroup","FolderName","JobName","Frequency","HolidayProc","SaveOption"]
		
		email = form["emailTo"].value

		fn1 = os.path.basename(fileitem1.filename)
		open(fn1, 'wb').write(fileitem1.file.read())
		fn2 = os.path.basename(fileitem2.filename)
		open(fn2, 'wb').write(fileitem2.file.read())

		tempFileName = email + "_" + str(datetime.datetime.now().date()) + '.csv'

		rowNum = 1

		# write the output file to the correct directory with the name created above
		pathName = os.path.join('R:/FileDownload/',tempFileName)

		# open file input #1
		# puts first row (headers) into a list to be referenced later
		# puts all following data into separate list
		with open(cgi.escape(fileitem1.filename,"r")) as fileToRead:
			reader = csv.reader(fileToRead)
			contentList = list(reader)
			headerList = contentList[0]
			numrows = len(contentList)

		# open and read file #2 and puts data into lists
		# second file contains data that repeats for all jobs
		with open(cgi.escape(fileitem2.filename,"r")) as file2Read:
			reader2 = csv.reader(file2Read)
			contentList2 = list(reader2)
			headerList2 = contentList2[0]
			numrows2 = len(contentList2)

		 # write the header to the output file
		if checkFields(requiredHeaders,headerList,headerList2) is False:
			print '<script>alert("Please ensure you have all the required header titles in your CSV files")</script>'
		elif checkFields(requiredHeaders,headerList,headerList2) is True:
			if checkEmail(email) is True:	
				with open(pathName, 'wb') as csvfile:
					
					#writer = csv.writer(csvfile, delimiter=',')
					writer = csv.writer(csvfile)
					writer.writerow(['Job Name','Request Number'])
					for i, item in enumerate(contentList[1:numrows]):
						tempList = []
						content = None
						driver = webdriver.Chrome()
						driver.get("link name removed for privacy reasons")
						JobType(headerList,item)
						workstation(headerList,item)
						scriptName(headerList,item)
						JobStreamName(headerList,item)
						JobDesc(headerList,item)
						scriptPath(headerList,item)
						Predecessor(headerList,item)
						Successor(headerList,item)
						fileDependency(headerList,item)
						startTime(headerList,item)
						jobInterval(headerList,item)
						Frequency(headerList,item)
						otherFreq(headerList,item)
						specialInstr(headerList,item)
						getList2(contentList2,numrows2)
						reqNum = getReqNum()
						v = getJobName(headerList,item)
						tempList.append(v)
						tempList.append(reqNum)
						writer.writerow(tempList)

				plainFilename = "filename"
				emailLink = "path name removed for security reasons" + tempFileName
				subprocess.call("cmd /c mail.vbs " + email + " " + plainFilename + " " + emailLink)

				driver.service.process.send_signal(signal.SIGTERM)

			elif checkEmail(email) is False:
				print '<script>alert("Your Email Address is invalid")</script>'

print """<FORM METHOD="POST" ACTION="%s" enctype="multipart/form-data" onsubmit="confirmation()">
File 1: <INPUT TYPE="file" NAME="file1" required>
File 2: <INPUT TYPE="file" NAME="file2" required>
<br/><br/>
Email: <input type="text" name="emailTo" required>
""" % os.environ['SCRIPT_NAME']
print '<br/><br/><br/>'
print '<INPUT TYPE="submit" NAME="submit" VALUE="Submit" style="background-color: #EC7D4F; border-radius: 10px; padding: 10px 15px 10px 15px; border: 0px; color: white;">'
print '</FORM>'
print '<script type="text/javascript">'
print 'function confirmation(){'
print 'alert("Thank you for submitting a file.  Please press OK and you will be notified via email with the results.")'
print '}'
print '</script>'
print """<h4>What does this program do?</h4><p>The "Populate" script reads in two CSV files and automatically populates the form online.
This is done by reading each CSV line by line, and putting each line into a list.  The program finds the 
data based on the title of the row, then calls the associated function to populate the appropriate text
area.  When a new list is read, a new chrome window is opened to be populated.</p></body></html>"""