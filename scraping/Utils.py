from datetime import datetime
import json
import time
import re

from bs4 import BeautifulSoup
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def get_zip_codes():
    zip_codes = []
    for z in range(0, 99):
        zip_code_url = 'https://www.kff.org/wp-content/themes/kaiser-foundation-2016/interactives/subsidy-calculator/2023/json/zips/{}.json?_date={}'

        timestamp = str(round(time.time(), 3)).replace(".", "")

        if z < 10:
            zip_code_url = zip_code_url.format('0' + str(z), timestamp)
        else:
            zip_code_url = zip_code_url.format(str(z), timestamp)

        response = requests.get(zip_code_url)
        html_content = response.text

        soup = BeautifulSoup(html_content, 'html.parser')
        better_soup = re.findall(r'"[^"]*"', soup.text.replace("merge_zip_data(", "")
                                 .replace(");", "")
                                 .replace("null", "'None'")
                                 .replace("\n", "")
                                 .replace("\t", "")
                                 .replace("\r", ""))

        for v in better_soup:
            q = v.replace('"', '').replace("'", "")
            try:
                if q.startswith('00'):
                    z = int(q[-3:])
                    zip_codes.append(str(z).zfill(5))
                elif q.startswith('0'):
                    z = int(q[-4:])
                    zip_codes.append(str(z).zfill(5))
                elif (q.startswith('1')
                      or q.startswith('2')
                      or q.startswith('3')
                      or q.startswith('4')
                      or q.startswith('5')
                      or q.startswith('6')
                      or q.startswith('7')
                      or q.startswith('8')
                      or q.startswith('9')):
                    z = int(q)
                    zip_codes.append(str(z).zfill(5))
            except:
                pass
    return zip_codes


def print_line_separator():
    print("-------------------------------")


def get_age(dob, tax_year):
    if dob is None or dob == "" or dob == "None" or dob == "N/A" or dob == "N/A":
        return None
    if dob.__class__ == str:
        try:
            dob = datetime.strptime(dob, "%Y-%m-%d")
        except:
            dob = datetime.strptime(dob, "%m/%d/%Y")

    reference_date = datetime.strptime(str(tax_year) + "-01-01", "%Y-%m-%d")

    # Calculate initial age
    curr_age = reference_date.year - dob.year

    # Adjust age if the reference date is before the birthday in the reference year
    if (reference_date.month, reference_date.day) < (dob.month, dob.day):
        curr_age -= 1

    return curr_age
