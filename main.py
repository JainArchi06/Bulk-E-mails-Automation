from selenium import webdriver
import time
from config import EMAIL, PASSWORD
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.common.keys import Keys

dataframe = pd.read_excel('Data.xlsx')
print(dataframe.index)

browser = webdriver.Chrome()


browser.maximize_window()

browser.get('https://outlook.live.com/owa/?nlp=1')


email_input = browser.find_element(By.CSS_SELECTOR, ('input[type="email"]'))
email_input.send_keys(EMAIL)
time.sleep(1)

next_button = browser.find_element(By.CSS_SELECTOR, ('input[type="submit"]'))
next_button.click()

pass_input = browser.find_element(By.CSS_SELECTOR, ('input[type="password"]'))
pass_input.send_keys(PASSWORD)
time.sleep(3)

sign_in_button = browser.find_element(
    By.CSS_SELECTOR, ('input[type="submit"]'))
sign_in_button.click()
time.sleep(1)

yes_button = browser.find_element(By.CSS_SELECTOR, ('input[type="submit"]'))
yes_button.click()

time.sleep(3)


for i in dataframe.index:
    print(dataframe.loc[i], end='\n\n')
    time.sleep(5)

    Newmail = browser.find_element(
        By.CSS_SELECTOR, ('button.splitPrimaryButton.root-185'))
    Newmail.click()
    time.sleep(10)

    To = browser.find_element(
        By.CSS_SELECTOR, ('.VbY1P.T6Va1.Z4n09.EditorClass'))
    To.send_keys(dataframe.loc[i]['EMAIL'])
    To.send_keys(Keys.RETURN)
    time.sleep(5)

    Subject = browser.find_element(
        By.CSS_SELECTOR,('#docking_InitVisiblePart_0 > div > div.QFieH.J8sYv > div.f1MMn > div.P6mmz.qeaxG > div > div > div > input'))
    Subject.send_keys(f"Hello {dataframe.loc[i]['NAME']}")
    time.sleep(10)

    body = browser.find_element(By.CLASS_NAME, ('elementToProof'))
    body_content = f"""Hello {dataframe.loc[i]['NAME']},
    Greeting!
    Congratulations, You are selected For Next Round.
    That is HR Round
    """
    body.send_keys(body_content)
    time.sleep(3)
    # body.send_keys(f"""hello {dataframe.loc[i]['NAME']}

    # """)
    # send_btn = browser.find_element(
    #     "xpath", '#docking_InitVisiblePart_2 > div > div.vBoqL.iLc1q.cc0pa.tblbU.SVWa1.dP5Z2.cF0pa > div.quEtc > div > span > button.ms-Button.ms-Button--primary.ms-Button--hasMenu.root-370')
    body.send_keys(Keys.CONTROL + Keys.RETURN)
    time.sleep(5)
