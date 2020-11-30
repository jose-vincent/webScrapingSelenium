from environs import Env
import os
import time
from selenium import webdriver

env = Env()
env.read_env()

email = env("USER_NAME")
password = env("PASSWORD")
url = env("URL")
path = os.path.abspath("chromedriver")

# Opening the wed_driver
pilot = webdriver.Chrome(executable_path=path)
pilot.get(url)

# Sign in
pilot.find_element_by_name("username").send_keys(email)
pilot.find_element_by_name("password").send_keys(password)
pilot.find_element_by_class_name("no-focus-style").click()
pilot.find_element_by_xpath('//*[@id="WoKerror"]/div/table[2]/tbody/tr/td[2]/p/a[1]').click()

#query
# query = input("Enter you query...")
query = 'agriculture'

pilot.find_element_by_id("value(input1)").send_keys(query)
pilot.find_element_by_xpath('//*[@id="searchCell1"]/span[1]/button').click()

result = pilot.find_element_by_xpath('//*[@id="hitCount.top"]')
result_count = result.text

print(result_count)

# time.sleep(5)

# pilot.quit()
