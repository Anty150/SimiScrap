import os
import time

import multiprocessing
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from selenium.common import NoSuchElementException, TimeoutException
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import re


class APIKeySelector:
    key = os.getenv("GOOGLE_API_KEY")  # Insert Your API Key here


class Driver:
    driver = None

    def __init__(self):  # Disable for Debug
        chrome_options = Options()
        # chrome_options.add_argument("--headless")  # Run in headless mode
        # chrome_options.add_argument("--disable-javascript")  # Disable JavaScript
        # chrome_options.add_argument("--disable-gpu")  # Disable GPU rendering
        self.driver = webdriver.Chrome(options=chrome_options)


class CompanyNamesSelector:
    _companyNames = None

    def __init__(self):
        file_name = CompanyNamesFileSelector.file_name
        contents = ""
        with open(file_name, encoding='utf8') as f:
            for line in f:
                contents += line
                print(f"Loaded names: {contents}")  # Only for debug
        self._companyNames = contents.split()

    def get_company_name_at_index(self, index):
        try:
            company_name = self._companyNames[index]
            print(company_name)
            return company_name
        except IndexError as e:
            print(f"{e}, No name at given index: {index}")
            return ""

    def get_company_names(self):
        return self._companyNames

    def __iter__(self):
        self.index = 0
        return self

    def __next__(self):
        if self.index < len(self._companyNames):
            company_name = self._companyNames[self.index]
            self.index += 1
            return company_name
        else:
            raise StopIteration


class KeywordsSelector:
    _keywords = None

    def set_keywords(self):  # ToDo Fix error when words in list.txt are in two or more separate lines
        file_name = TextFileSelector.file_name
        contents = ""
        with open(file_name, encoding='utf8') as f:
            for line in f:
                contents += line
                # print("Loaded keywords: {contents}")  # Only for debug
        self._keywords = contents.split()

    def get_keywords(self):
        return self._keywords


class CompanyNameSelector:
    _company_name = ""

    def set_company_name(self, company_name):
        # self.company_name = input("Set company name: ")
        # self._company_name = "Wemeco Poland Sp. z o.o."  # Used only for Debug
        # self._company_name = "SABNER"  # Insert Company name here
        self._company_name = company_name

    def get_company_name(self):
        return self._company_name


class CompanyNamesFileSelector:
    file_name = 'name_list.txt'


class TextFileSelector:
    file_name = 'list.txt'


class CookieSelector:
    locator = "L2AGLb"
    finder = By.ID


class SearchBarSelector:
    locator = 'q'
    finder = By.NAME


class CompanyCardSelector:
    locator = "sATSHe"
    finder = By.CLASS_NAME


class BasePageSelector:
    locator = "https://www.google.com/"


class CorrectPageSelector:
    locator = "div.liYKde.g.VjDLd"
    finder = By.CSS_SELECTOR


class LinkTagSelector:
    locator = 'a'
    finder = By.TAG_NAME


class LinkAttributeSelector:
    locator = "href"


class MapSiteButtonSelector:
    locator = "a.yYlJEf.Q7PwXb.L48Cpd.brKmxb"
    finder = By.CSS_SELECTOR


class TimeoutSelector:
    timeout_time = 8
    wait_time = 10


class BasePage:
    def __init__(self, driver_initializer):
        self.driver = driver_initializer.driver

    def search_target(self, company_name_initializer):
        self.open_base_page()
        self.wait_for_search_bar()
        self.operate_search_bar(company_name_initializer.get_company_name())

    def find_search_bar(self):
        search = self.driver.find_element(SearchBarSelector.finder, SearchBarSelector.locator)
        return search

    def open_base_page(self):
        self.driver.get(BasePageSelector.locator)
        try:
            search_bar = self.driver.find_element(CookieSelector.finder, CookieSelector.locator)  # Accept cookies
            search_bar.send_keys(Keys.ENTER)
        except NoSuchElementException:
            print("Invalid website provided! Make sure website is https://www.google.com/")
            self.driver.quit()
            exit()

    def wait_for_search_bar(self):
        try:
            WebDriverWait(self.driver, TimeoutSelector.wait_time).until(
                EC.presence_of_element_located((SearchBarSelector.finder, SearchBarSelector.locator))
            )
        except NoSuchElementException:
            print("Search bar not found! Make sure website is https://www.google.com/")
            self.driver.quit()
            exit()

    def operate_search_bar(self, company_name):
        search = self.find_search_bar()
        search.send_keys(company_name)

        # Wait for the search button to become clickable
        WebDriverWait(self.driver, TimeoutSelector.wait_time).until(
            EC.element_to_be_clickable((SearchBarSelector.finder, SearchBarSelector.locator))
        )

        # Click the search button
        search.submit()


class SearchPage:
    def __init__(self, driver_initializer):
        self.driver = driver_initializer.driver

    def check_if_valid(self):
        try:
            WebDriverWait(self.driver, TimeoutSelector.wait_time).until(
                EC.presence_of_element_located((CompanyCardSelector.finder, CompanyCardSelector.locator))
            )
        except NoSuchElementException:
            print("Unable to find company card")
            return 0
        return 1

    def open_similar_search(self):
        search = self.driver.find_elements(CompanyCardSelector.finder, CompanyCardSelector.locator)
        search = search[-1]

        link = search.find_element(LinkTagSelector.finder, LinkTagSelector.locator)
        link_value = link.get_attribute(LinkAttributeSelector.locator)

        print(link_value)  # Only Debug
        return link_value


class MapPage:
    _hit_urls = []
    _hit_urls_bools = []
    _hit_urls_nips = []

    def __init__(self, driver_initializer):
        self.driver = driver_initializer.driver

    def check_if_on_correct_page(self):
        is_page_found = 1
        try:
            if WebDriverWait(self.driver, 3).until(  # DOPOPRAWYYYYY!!!
                    EC.presence_of_element_located((CorrectPageSelector.finder, CorrectPageSelector.locator))
            ):
                print("Unable to locate corresponding map page")
                is_page_found = 0
        finally:
            if not is_page_found:
                print("No map page!")
                return 0
            else:
                print("Proceeding")
                return 1

    def operate_map_search(self, link_value):
        self.driver.get(link_value)
        if self.check_if_on_correct_page():
            iterator = 0
            x = 0

            potential_hits = self.driver.find_elements(MapSiteButtonSelector.finder, MapSiteButtonSelector.locator)
            print(f"Results found: {len(potential_hits)}")
            # time.sleep(5) # Use only to debug
            while iterator < len(potential_hits):
                self._hit_urls.append(potential_hits[iterator].get_attribute(LinkAttributeSelector.locator))
                iterator += 1

            while x < len(self._hit_urls):  # Used only for debug
                print(self._hit_urls[x])
                print("\n")
                x += 1

            self.run_parallel()

    def run_parallel(self):
        function_pool = multiprocessing.Pool(processes=len(self._hit_urls))

        for result in function_pool.map(self.check_hits, self._hit_urls):
            if result[0] == 0:
                self._hit_urls_bools.append(0)
                self._hit_urls_nips.append(result[1])
            else:
                self._hit_urls_bools.append(1)
                self._hit_urls_nips.append(result[1])

    @staticmethod
    def check_hits(hit_urls):
        # ToDo implement starting x number of process here
        result = MapPage.open_hit_page(hit_urls)

        if result[0]:
            print(f"Hit {hit_urls} is good \n")
            return 1, result[1]
        else:
            print(f"Hit {hit_urls} is not good \n")
            return 0, result[1]

    @staticmethod
    def open_hit_page(hit):

        chrome_options = Options()
        chrome_options.add_argument("--headless=new")  # Run in headless mode
        chrome_options.add_argument("--disable-javascript")  # Disable JavaScript
        chrome_options.add_argument("--disable-gpu")  # Disable GPU rendering

        driver = webdriver.Chrome(options=chrome_options)

        print(f"Opening URL: {hit}")
        driver.set_page_load_timeout(TimeoutSelector.timeout_time)  # Close driver if loading takes too long
        try:
            driver.get(hit)
            source = driver.page_source.lower()
            function_keyword = KeywordsSelector()
            function_keyword.set_keywords()
            function_keywords = function_keyword.get_keywords()

            iterator = 0

            for _ in function_keywords:
                if function_keywords[iterator] in source:
                    for _ in source:
                        pattern = r'nip[:\-\s]*([0-9\s-]+)'
                        match = re.search(pattern, source)

                        if match:
                            value = match.group(1).replace("-", "").replace(" ", "")
                            if len(value) == 10:
                                print(f"NIP found:", value, f" on {hit}")
                                return 1, value
                            break
                    return 1, 'None'
                iterator += 1
            # print(MapPage._keyword_list[iterator])
            return 0, 'None'
        except TimeoutException:
            print(f"Timeout while opening {hit}")
            driver.close()
            return 0, 'None'

        except Exception as e:
            print(f"{e}, Error while opening {hit}")
            driver.close()
            return 0, 'None'

    def get_hit_urls(self):
        return self._hit_urls

    def get_hit_urls_bools(self):
        return self._hit_urls_bools

    def get_hit_urls_nips(self):
        return self._hit_urls_nips


class ExportManagerGSheet:
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        "C:\\Users\\Igor\\PycharmProjects\\SimiScrap\\key2.json", scope)  # Insert your Google API credentials here

    client = gspread.authorize(credentials)

    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1_BTRg9Deova90IlbCqcha6kaKOUZhVQyA3YcrNys6Bk/edit?usp" \
                      "=sharing"  # Replace with your Google Sheets link
    current_date = datetime.now().strftime("%Y-%m-%d")

    @staticmethod
    def create_workbook(hit_urls, hit_urls_bools, hit_urls_nips):
        try:
            # Open the Google Sheets document by URL
            spreadsheet = ExportManagerGSheet.client.open_by_url(ExportManagerGSheet.spreadsheet_url)

            # Select the worksheet by title (create it if it doesn't exist)
            worksheet_title = "Scraping Result - {}".format(ExportManagerGSheet.current_date)
            try:
                existing_worksheet = spreadsheet.worksheet(worksheet_title)
                spreadsheet.del_worksheet(existing_worksheet)
            except gspread.exceptions.WorksheetNotFound:
                pass  # Worksheet doesn't exist, so no need to delete

            try:
                worksheet = spreadsheet.worksheet(worksheet_title)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=worksheet_title, rows="100", cols="10")

            # Update the worksheet with data
            header_row = [['Company Domain', 'Is hit good?', 'NIP']]
            worksheet.insert_rows(header_row, 1)

            data_to_insert = []
            for i in range(len(hit_urls)):
                data_to_insert.append([hit_urls[i], hit_urls_bools[i], hit_urls_nips[i]])

            worksheet.insert_rows(data_to_insert, 2)

            print("Data successfully exported to Google Sheets.")

        except Exception as e:
            print("Error while exporting to Google Sheets: {}".format(str(e)))


def main():
    start_time = time.time()

    company_names_initializer = CompanyNamesSelector()

    for company_name in company_names_initializer:
        driver_initializer = Driver()
        company_name_initializer = CompanyNameSelector()
        company_name_initializer.set_company_name(company_name)

        print(company_name_initializer.get_company_name())
        try:
            base_search = BasePage(driver_initializer)
            company_search = SearchPage(driver_initializer)
            map_search = MapPage(driver_initializer)

            base_search.search_target(company_name_initializer)

            if company_search.check_if_valid():
                map_search.operate_map_search(company_search.open_similar_search())
                try:
                    ExportManagerGSheet.create_workbook(map_search.get_hit_urls(), map_search.get_hit_urls_bools(),
                                                        map_search.get_hit_urls_nips())
                except PermissionError:
                    print("Error while exporting to Google Sheets.")
            driver_initializer.driver.close()
        except Exception as e:
            print(f"Error: {e}")
            driver_initializer.driver.close()

    print("--- %s seconds ---" % (time.time() - start_time))


if __name__ == '__main__':
    main()
