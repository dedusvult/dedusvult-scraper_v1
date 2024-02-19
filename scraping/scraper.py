# pip install selenium
# pip install webdriver_manager
# pip install requests
# pip install pandas
from package_installer import PackageInstaller

import json
import time
import re
import requests
import ExcelProcessor
import global_vars
import Utils

from bs4 import BeautifulSoup
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# input_file = "C:\\Users\\dmitry.ershov\\OneDrive - Benefitscape, Inc\\Desktop\\work\\Dev\\Projects\\Ubers\\SingleLineUber_FarHillsDevelopment_Update_LSCPadded_AP - Copy.xlsx"
# use_home_zip = False


class WebScraper:
    def __init__(self, input_file, zip_to_use, tax_year):
        self.input_file = input_file
        self.zip_to_use = zip_to_use
        self.tax_year = tax_year
        self.driver = get_driver()

    def run(self):
        package_installer = PackageInstaller(["selenium", "webdriver_manager", "requests", "pandas", "openpyxl", "bs4"])
        package_installer.install_packages()

        employees = ExcelProcessor.get_employee_map_from_file(self.input_file, self.zip_to_use)
        try:
            Utils.print_line_separator()
            print("Employees to process: " + str(len(employees)))
            Utils.print_line_separator()

            all_ages_in_the_employee_map_for_zips = get_all_age_zip_combinations(employees, self.zip_to_use,
                                                                                 self.tax_year)

            age_zip_combinations_to_process = get_age_zip_combinations_number_to_process(
                all_ages_in_the_employee_map_for_zips)
            print("Age-ZipCode combinations to process: " + str(age_zip_combinations_to_process))
            Utils.print_line_separator()

            find_lcsp_for_all_ages_and_zip_combinations(all_ages_in_the_employee_map_for_zips,
                                                        employees,
                                                        self.zip_to_use,
                                                        self.driver,
                                                        self.tax_year)
        except:
            print(
                "Error occurred while processing employees. Writing the file with the processed employees."
                "\nRerun the program to process the remaining employees.")

        ExcelProcessor.write_employee_map_to_file(self.input_file, employees, self.zip_to_use)
        self.driver.quit()


def get_age_zip_combinations_number_to_process(all_ages_in_the_employee_map_for_zips):
    counter = 0

    for zip_code_age in all_ages_in_the_employee_map_for_zips:
        if zip_code_age.lcsp is None:
            counter += 1

    return counter


def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument("--incognito")
    return webdriver.Chrome(service=ChromeService(
        ChromeDriverManager().install()), options=options)


def get_lcsp(zip_code_to_search, age_to_search, driver):
    lcsp_dict = {}
    print("Zip code: " + zip_code_to_search + ", Age: " + age_to_search)
    curr_url = global_vars.base_url.format(zip_code_to_search, age_to_search)

    driver.get(curr_url)
    driver.refresh()
    time.sleep(0.5)
    print(curr_url)

    elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,
                                                                          "//*[contains(text(), 'Without financial help, your silver plan would cost:')]")))
    try:
        lcsp = elem.find_element(By.XPATH, "..").text.split('\n')[1]
    except:
        print("LCSP not found, trying again...")
        time.sleep(1)
        driver.close()
        driver = get_driver()
        updated_lcsp = get_lcsp(zip_code_to_search, age_to_search, driver)
        return updated_lcsp
    # driver.close()

    print("LCSP: " + lcsp)
    Utils.print_line_separator()

    # driver.refresh()
    body_element = driver.find_element(By.TAG_NAME, "body")
    # delete the element
    driver.execute_script(
        """ var element = arguments[0]; element.parentNode.removeChild(element); """
        , body_element)
    lcsp_dict[zip_code_to_search] = lcsp
    return lcsp_dict


class ZipCodeAge:
    def __init__(self, zip_code, zip_age, lcsp=None):
        self.zip_code = zip_code
        self.zip_age = zip_age
        self.lcsp = lcsp

    def __eq__(self, other):
        if isinstance(other, ZipCodeAge):
            return self.zip_code == other.zip_code and self.zip_age == other.zip_age and self.lcsp == other.lcsp
        return False

    def __hash__(self):
        return hash((self.zip_code, self.zip_age, self.lcsp))


def is_employee_record_invalid(employee, zip_to_use):
    flag = False
    if employee.dob is None or employee.dob == "":
        flag = True
    if zip_to_use == "Home Zip" and (employee.home_zip is None or employee.home_zip == ""):
        flag = True
    if zip_to_use == "Work Zip" and (employee.work_zip is None or employee.work_zip == ""):
        flag = True
    return flag


def get_all_age_zip_combinations(employees, zip_to_use, tax_year):
    all_ages_in_the_employee_map_for_zips = set()
    for employee in employees.values():
        if is_employee_record_invalid(employee, zip_to_use):
            continue

        # if all_ages_in_the_employee_map_for_zips contains the ZipCodeAge object with the same zip_code and zip_age,
        # and it has a lcsp value, then skip the current iteration
        # if it has not a lcsp value but the current employee has a lcsp value,
        # then add the current employee's lcsp and skip the current iteration
        already_has_lcsp = False

        for zip_code_age in all_ages_in_the_employee_map_for_zips:
            if zip_to_use == "Home Zip":
                if zip_code_age.zip_age == Utils.get_age(employee.dob,
                                                         tax_year) and zip_code_age.zip_code == employee.home_zip:
                    if zip_code_age.lcsp is not None:
                        already_has_lcsp = True
                        break
                    if employee.lcsp is not None:
                        zip_code_age.lcsp = employee.lcsp
                        already_has_lcsp = True
                        break
            else:
                if zip_code_age.zip_age == Utils.get_age(employee.dob,
                                                         tax_year) and zip_code_age.zip_code == employee.work_zip:
                    if zip_code_age.lcsp is not None:
                        already_has_lcsp = True
                        break
                    if employee.lcsp is not None:
                        zip_code_age.lcsp = employee.lcsp
                        already_has_lcsp = True
                        break

        if already_has_lcsp:
            continue

        if zip_to_use == "Home Zip":
            zip_code_age = ZipCodeAge(employee.home_zip, Utils.get_age(employee.dob, tax_year))
            if zip_code_age.zip_age is not None:
                all_ages_in_the_employee_map_for_zips.add(zip_code_age)
        else:
            zip_code_age = ZipCodeAge(employee.work_zip, Utils.get_age(employee.dob, tax_year))
            if zip_code_age.zip_age is not None:
                all_ages_in_the_employee_map_for_zips.add(zip_code_age)

    return all_ages_in_the_employee_map_for_zips


def find_lcsp_for_all_ages_and_zip_combinations(all_ages_in_the_employee_map_for_zips, employees, zip_to_use, driver,
                                                tax_year):
    for zip_code_age in all_ages_in_the_employee_map_for_zips:
        if zip_code_age.lcsp is not None:
            print("LCSP already found in the file for zip code: " + str(zip_code_age.zip_code) + ", age: " + str(
                zip_code_age.zip_age) + "\nLCSP: " + str(zip_code_age.lcsp))
            specify_lcsp_for_all_employees_with_current_zip_and_age(employees,
                                                                    {zip_code_age.zip_code: zip_code_age.lcsp},
                                                                    zip_to_use,
                                                                    zip_code_age.zip_code,
                                                                    zip_code_age,
                                                                    tax_year)
            continue

        age = zip_code_age.zip_age
        zip_code = str(zip_code_age.zip_code).zfill(5)
        if age < 21:
            age = 21
        elif age > 64:
            age = 64
        age = str(age)
        lcsp_map = get_lcsp(zip_code, age, driver)
        if lcsp_map.__class__ == dict and len(lcsp_map) > 0:
            specify_lcsp_for_all_employees_with_current_zip_and_age(employees, lcsp_map, zip_to_use, zip_code,
                                                                    zip_code_age,
                                                                    tax_year)


def specify_lcsp_for_all_employees_with_current_zip_and_age(employees, lcsp_map, zip_to_use, zip_code, zip_code_age,
                                                            tax_year):
    for employee in employees.values():
        if zip_to_use == "Home Zip":
            employee_zip = str(employee.home_zip).zfill(5)
        else:
            employee_zip = str(employee.work_zip).zfill(5)
        employee_age = Utils.get_age(employee.dob, tax_year)
        if (employee_zip == zip_code and employee_age == zip_code_age.zip_age and
                (employee.lcsp is None or employee.lcsp == "")):
            employee.lcsp = lcsp_map[zip_code]
