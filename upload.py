import os
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from functions.methods import open_chrome_webdriver, returnFile


# read excel sheet
# open workbook
wb = openpyxl.load_workbook('SCOPE-HCPs-BD.xlsx')
# extract sheet
sheet = wb['scope']
sheet = sheet['A50': 'D56']
# browser.switch_to_frame(0)
# change firectory to that with certificate names

os.chdir(os.getcwd() + '/certificates')

# start filling the form
i = 1
for row in sheet:
    # open chrome browser
    browser = open_chrome_webdriver(
        "https://www.worldobesity.org/training-and-events/training/scope/certification/apply-for-scope-certification")
    browser.implicitly_wait(20)
    # get the hcp name
    hcp_name = row[0].value
    hcp_email = row[1].value
    dof = row[2].value
    # get th name input
    fullname = browser.find_element_by_name("full_name")
    fullname.send_keys(hcp_name)
    # get the job title
    job_title = browser.find_element_by_name("job_title")
    job_title.send_keys("Health Consultant Pharmacist")
    # get country
    country = browser.find_element_by_name("country")
    country.send_keys("Saudi Arabia")
    # get date of birth
    date_of_birth = browser.find_element_by_name(
        "date_of_birth_ddmmyy")
    date_of_birth.send_keys(dof)
    # get email
    email = browser.find_element_by_name(
        "email")
    email.send_keys(hcp_email)
    # check box 1
    checkbox_1 = browser.find_element_by_id(
        "form-input-have_you_accrued_at_least_12_scope_points")
    browser.execute_script("arguments[0].click();", checkbox_1)
    # check box 2
    checkbox_2 = browser.find_element_by_name(
        "have_you_completed_the_core_learning_path")
    browser.execute_script("arguments[0].click();", checkbox_2)

    # upload document
    upload = browser.find_element_by_name(
        "please_upload_your_supporting_evidence[]")
    for file in list(os.listdir(os.getcwd())):
        name = returnFile(file)
        if name == hcp_name:
            upload.send_keys(os.getcwd() + '/' + file)
            # right status to for this hcp
            row[3].value = "O.K"
            wb.save("SCOPE-HCPs-BD-2.xlsx")
            print(
                f'({i})\nHCP Name:\t{hcp_name}\nEmail:\t{hcp_email}\nFile:\t{file}\nDate Of Birth: {dof}')
            print('=======================================================')
            i += 1

    # get textarea
    text_area = browser.find_element_by_name("referee")
    text_area.send_keys(''' 
PO Box 17129 Jeddah 21484 Saudi Arabia CR 4030053868 CO C 25292
T +966 12 6535353 -  F +966 12 6074399
www.nahdi.sa
    ''')
    # get checkbox 3
    checkbox_3 = browser.find_element_by_name(
        "do_you_consent_for_your_name_job_title_and_country_to_be_included_on_our_online_list_of_scope_certified_professionals")
    browser.execute_script("arguments[0].click();", checkbox_3)
    # get checkbox 4
    checkbox_4 = browser.find_element_by_name(
        "check_this_box_to_confirm_that_you_consent_to_receive_future_email_communications_from_the_world_obesity_federation_regarding_scope_and_other_initiati")
    browser.execute_script("arguments[0].click();", checkbox_4)
    # get submit button
    # submit = browser.find_element_by_name("form_page_submit")
    # submit.click()

    browser.close()

print("------------------------- Program Finished -------------------------")
