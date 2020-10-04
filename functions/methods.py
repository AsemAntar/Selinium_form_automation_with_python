import os

from selenium import webdriver
from selenium.webdriver.common.keys import Keys


def open_chrome_webdriver(url):
    browser = webdriver.Chrome()
    # maximize the browser window
    browser.maximize_window()
    # open the url
    browser.get(url)

    return browser


def returnFile(cell_value):
    current_path = os.path.abspath('.') + '/hcp_certificates'
    base = os.path.basename(current_path + '/' + cell_value)
    name = os.path.splitext(base)[0]

    return name
