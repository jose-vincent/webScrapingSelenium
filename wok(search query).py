'''Please make sure that 
the version of your Chrome browser should be 
compatible with the chromedriver that 
you are going to download.'''

from environs import Env
import os, datetime
import time
import sys 
from selenium import webdriver

env = Env()
env.read_env()

email = env("USER_NAME")
password = env("PASSWORD")
path = os.path.abspath("chromedriver")

# Enter your query
query = input("--> Enter your query : ")

print("--> Loading...")

# Set download directory with date & time as folder name
results_dir = (os.getcwd()) + '/Results'
down_dir = os.path.join(results_dir, datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S'))
os.makedirs(down_dir)
options = webdriver.ChromeOptions()
prefs = {"download.default_directory":down_dir}
options.add_experimental_option("prefs", prefs)

# Opening the wed_driver
pilot = webdriver.Chrome(executable_path=path, options = options)
pilot.get("https://www.webofknowledge.com/")

# Sign in
try:
	pilot.find_element_by_name("username").send_keys(email)
	pilot.find_element_by_name("password").send_keys(password)
	pilot.find_element_by_class_name("no-focus-style").click()
	pilot.find_element_by_xpath('//*[@id="WoKerror"]/div/table[2]/tbody/tr/td[2]/p/a[1]').click()
	pilot.find_element_by_id("value(input1)").send_keys(query)
	pilot.find_element_by_xpath('//*[@id="searchCell1"]/span[1]/button').click()
except:
	pilot.find_element_by_id("value(input1)").send_keys(query)
	pilot.find_element_by_xpath('//*[@id="searchCell1"]/span[1]/button').click()

errMsg = ""

try:
	errMsg = pilot.find_element_by_class_name("newErrorHead").text
except:
	pass

if errMsg == "Your search found no records.":
	print("--> ",errMsg,"\n","--> Exiting...")
	pilot.quit()
	sys.exit()

# Number of results for your query
result = pilot.find_element_by_xpath('//*[@id="hitCount.top"]')
result_count = int(result.text.replace(',', ''))
print(f"--> {result_count} results found")
print(f"--> Files are being downloaded, please wait...")

time.sleep(3)

if result_count <= 500:
	# Filing the box with required values & options
	pilot.find_element_by_name("export").click()
	pilot.find_element_by_name("Export to Excel").click()
	pilot.find_element_by_id("numberOfRecordsRange").click()
	pilot.find_element_by_id('markFrom').clear()
	pilot.find_element_by_id("markFrom").send_keys(1)
	pilot.find_element_by_id('markTo').clear()
	pilot.find_element_by_id("markTo").send_keys(result_count)
	box = pilot.find_element_by_id("bib_fields")

	# select 'Full Record option'
	for option in box.find_elements_by_tag_name('option'):
	    if option.text == 'Full Record     ':
	    	option.click()

	pilot.find_element_by_id("excelButton").click()
else:
	if 500 < result_count <=1000:
		pilot.find_element_by_name("export").click()
		pilot.find_element_by_name("Export to Excel").click()
		pilot.find_element_by_id("numberOfRecordsRange").click()
		pilot.find_element_by_id('markFrom').clear()
		pilot.find_element_by_id("markFrom").send_keys(1)
		pilot.find_element_by_id('markTo').clear()
		pilot.find_element_by_id("markTo").send_keys(500)
		box = pilot.find_element_by_id("bib_fields")

		for option in box.find_elements_by_tag_name('option'):
		    if option.text == 'Full Record     ':
		    	option.click()

		pilot.find_element_by_id("excelButton").click()

		time.sleep(6)

		pilot.find_element_by_xpath('//*[@id="page"]/div[1]/div[26]/div[2]/div/div/div/div[2]/div[3]/div[3]/div[2]/div[1]/ul/li/button').click()
		pilot.find_element_by_id("numberOfRecordsRange").click()
		pilot.find_element_by_id('markFrom').clear()
		pilot.find_element_by_id("markFrom").send_keys(501)
		pilot.find_element_by_id('markTo').clear()
		pilot.find_element_by_id("markTo").send_keys(result_count)
		box = pilot.find_element_by_id("bib_fields")

		for option in box.find_elements_by_tag_name('option'):
		    if option.text == 'Full Record     ':
		    	option.click()

		pilot.find_element_by_id("excelButton").click()

		time.sleep(6)	

	else:			
		pilot.find_element_by_name("export").click()
		pilot.find_element_by_name("Export to Excel").click()
		pilot.find_element_by_id("numberOfRecordsRange").click()
		pilot.find_element_by_id('markFrom').clear()
		pilot.find_element_by_id("markFrom").send_keys(1)
		pilot.find_element_by_id('markTo').clear()
		pilot.find_element_by_id("markTo").send_keys(500)
		box = pilot.find_element_by_id("bib_fields")

		for option in box.find_elements_by_tag_name('option'):
		    if option.text == 'Full Record     ':
		    	option.click()

		pilot.find_element_by_id("excelButton").click()

		time.sleep(6)

		# Lists of ranges to download files
		if (result_count%500) == 0:
			no_iter = (int(result_count/500))
		else:
			no_iter = (int(result_count/500))+1

		FROM_start = 501
		FROM_interval = 500
		FROM = list(range(FROM_start, FROM_interval*(no_iter), FROM_interval))
		TO_start = 1000
		TO_interval = 500
		TO = list(range(TO_start, TO_interval*(no_iter), TO_interval))
		TO.append(result_count)

		for start, end in zip(FROM, TO):
			pilot.find_element_by_xpath('//*[@id="page"]/div[1]/div[26]/div[2]/div/div/div/div[2]/div[3]/div[3]/div[2]/div[1]/ul/li/button').click()
			pilot.find_element_by_id("numberOfRecordsRange").click()
			pilot.find_element_by_id('markFrom').clear()
			pilot.find_element_by_id("markFrom").send_keys(start)
			pilot.find_element_by_id('markTo').clear()
			pilot.find_element_by_id("markTo").send_keys(end)
			box = pilot.find_element_by_id("bib_fields")

			for option in box.find_elements_by_tag_name('option'):
			    if option.text == 'Full Record     ':
			        option.click()

			pilot.find_element_by_id("excelButton").click()
			print(f"* Downloaded {end} of {result_count} results")
			time.sleep(6)

print("--> Download complete, Exiting...")
time.sleep(1)
pilot.quit()