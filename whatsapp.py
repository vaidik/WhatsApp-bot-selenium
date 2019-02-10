# Note: For proper working of this Script Good and Uninterepted Internet Connection is Required
# Keep all contacts unique
# Can save contact with their phone Number

# Import required packages
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.firefox.options import Options

options = Options()
options.headless = True


import datetime
import time
import openpyxl as excel

# function to read contacts from a text file
def readContacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    secondCol = sheet['B']
    for cell in range(len(firstCol)):
        if firstCol[cell].value is None:
            continue
        contact = str(firstCol[cell].value)
        contact = "\"" + contact + "\""
        detail = str(secondCol[cell].value)
        lst.append((contact, detail))
    return lst

# Target Contacts, keep them in double colons
# Not tested on Broadcast
targets = readContacts("contacts.xlsx")

# can comment out below line
#targets = [
#    '"Vaibhav Kapoor"',
#    '"Isha Bhabhi"'
#]

print(targets)

# Driver to open a browser
driver = webdriver.Firefox(options=options)

#link to open a site
driver.get("https://web.whatsapp.com/")

# 10 sec wait time to load, if good internet connection is not good then increase the time
# units in seconds
# note this time is being used below also
wait = WebDriverWait(driver, 10)
wait5 = WebDriverWait(driver, 5)
wait70 = WebDriverWait(driver, 70)

driver.save_screenshot('qr_code.png')

# send screenshot by message
import os
PB_ACCESS_TOKEN = os.environ['PB_ACCESS_TOKEN']
from pushbullet import Pushbullet

pb = Pushbullet(PB_ACCESS_TOKEN)

with open("qr_code.png", "rb") as pic:
    file_data = pb.upload_file(pic, "watsapp_web_qr_code.png")

push = pb.push_file(**file_data)

pb.push_note('You have 60 seconds to login to whatsapp web', 'Find bhaia quickly')

print('waiting')

# wait for login
wait70.until(EC.presence_of_element_located((
    By.CSS_SELECTOR, 'input.copyable-text.selectable-text'
)))
print('Logged In')

# input("Scan the QR code and then press Enter")

# Message to send list
# 1st Parameter: Hours in 0-23
# 2nd Parameter: Minutes
# 3rd Parameter: Seconds (Keep it Zero)
# 4th Parameter: Message to send at a particular time
# Put '\n' at the end of the message, it is identified as Enter Key
# Else uncomment Keys.Enter in the last step if you dont want to use '\n'
# Keep a nice gap between successive messages
# Use Keys.SHIFT + Keys.ENTER to give a new line effect in your Message

new_line = lambda x: (Keys.SHIFT + Keys.ENTER) * x

MAIN_MESSAGE = '''Good Morning!

To avoid any hassles at the station and for a comfortable train journey from Varanasi to Delhi please see below your seat details:

%s

We will try our best to make the train journey comfortable for everyone but are restricted by seats allotted to us by Indian Railways. Wherever possible elders and women have been prioritised for easy access seats.

Look forward to seeing you at Maduahdih Railway Station (Platform No. 8) at 6:00 PM.

Thanks,
Kapoor & Sons.
(Isha, Vibhu & Rinku)'''

msgToSend = [
    [1, 1, 0, MAIN_MESSAGE]
]
                #[12, 32, 0, "Hello! This is test Msg. Please Ignore." + Keys.SHIFT + Keys.ENTER + "http://bit.ly/mogjm05"]
            #]

# Count variable to identify the number of messages to be sent
count = 0
while count<len(msgToSend):
    # Identify time
    curTime = datetime.datetime.now()
    curHour = curTime.time().hour
    curMin = curTime.time().minute
    curSec = curTime.time().second

    # if time matches then move further
    #if msgToSend[count][0]==curHour and msgToSend[count][1]==curMin and msgToSend[count][2]==curSec:
    if True:
        # utility variables to tract count of success and fails
        success = 0
        sNo = 1
        failList = []

        # Iterate over selected contacts
        for target in targets:
            print(sNo, ". Target is: " + target[0])
            print (target)
            sNo+=1
            try:
                # Select the target
                x_arg = '//span[contains(@title,' + target[0] + ')]'
                try:
                    wait5.until(EC.presence_of_element_located((
                        By.XPATH, x_arg
                    )))
                except Exception as e:
                    print(e)
                    # If contact not found, then search for it
                    searBoxPath = '//*[@id="input-chatlist-search"]'
                    #wait5.until(EC.presence_of_element_located((
                    #    By.ID, "input-chatlist-search"
                    #)))
                    print("Searching selector")
                    inputSearchBox = driver.find_element_by_css_selector(".copyable-text.selectable-text")
                    #inputSearchBox = driver.find_element_by_css_selector(
                    #    "span[title=%s]" % target[0])
                    time.sleep(0.5)
                    # click the search button
                    #driver.find_element_by_xpath('/html/body/div/div/div/div[2]/div/div[2]/div/button').click()
                    time.sleep(1)
                    inputSearchBox.clear()
                    inputSearchBox.send_keys(target[0][1:len(target[0]) - 1])
                    print('Target Searched')
                    # Increase the time if searching a contact is taking a long time
                    time.sleep(4)

                # Select the target
                inputSearchBox = driver.find_element_by_css_selector(
                    "span[title=%s i]" % target[0]).click()
                #driver.find_element_by_xpath(x_arg).click()
                print("Target Successfully Selected")
                time.sleep(2)

                # Select the Input Box
                inp_xpath = "//div[@contenteditable='true']"
                input_box = wait.until(EC.presence_of_element_located((
                    By.XPATH, inp_xpath)))
                time.sleep(1)

                # Send message
                # taeget is your target Name and msgToSend is you message
                #input_box.send_keys("Hello, " + target[0] + "."+ Keys.SHIFT + Keys.ENTER + msgToSend[count][3] + Keys.SPACE)
                for part in (msgToSend[count][3] % target[1]).split('\n'):
                    input_box.send_keys(part)
                    ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.ENTER).perform()

                #input_box.send_keys(msgToSend[count][3].lower() % target[1])
                # + Keys.ENTER (Uncomment it if your msg doesnt contain '\n')
                # Link Preview Time, Reduce this time, if internet connection is Good
                time.sleep(10)
                input_box.send_keys(Keys.ENTER)
                print("Successfully Send Message to : "+ target[0] + '\n')
                success+=1
                time.sleep(0.5)

            except Exception as e:
                print("Exception", e)
                # If target Not found Add it to the failed List
                print("Cannot find Target: " + target[0])
                failList.append(target[0])
                pass

        print("\nSuccessfully Sent to: ", success)
        print("Failed to Sent to: ", len(failList))
        print(failList)
        print('\n\n')

        title = 'Successfully sent all messages'
        if len(failList) > 0:
            title = 'Failed to send all messages'

        body = '''
Successfully sent to: %s
Failed to send to: %s
%s
''' % (success, len(failList), failList)
        pb.push_note(title, body)
        count+=1
driver.quit()
