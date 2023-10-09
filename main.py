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


# Note, the code core functionality currently (searching keywords), does not work, it is due to nip and email scraping

# ToDo Fix a rare bug

# ToDo Add log what word was found in hit

# ToDO Add separate iterations based on what we are looking

# ToDo Jedna tabela z branżą (bazowane na liscie slow kluczowych)
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


class KeywordsSelector:
    _keywords = None

    def set_keywords(self):  # ToDo Fix error when words in list.txt are in two or more separate lines
        file_name = TextFileSelector.file_name
        contents = ""
        with open(file_name, encoding='utf8') as f:
            for line in f:
                contents += line
                # print("Loaded keywords: {contents}") # Only for debug
        self._keywords = contents.split()

    def get_keywords(self):
        return self._keywords


class CompanyNameSelector:
    _company_name = ""

    def __init__(self, company_name):
        self._company_name = company_name

    def get_company_name(self):
        return self._company_name


class WorksheetTitleSelector:
    worksheet_title = "Podobne"


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
    _hit_urls_emails = []
    _hit_urls_nips = []

    def __init__(self, driver_initializer):
        self.driver = driver_initializer.driver

    def check_if_on_correct_page(self):
        is_page_found = 1
        try:
            if WebDriverWait(self.driver, 3).until(
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
            print("Working")
            iterator = 0
            x = 0

            potential_hits = self.driver.find_elements(MapSiteButtonSelector.finder, MapSiteButtonSelector.locator)
            print(f"Results found: {len(potential_hits)}")

            while iterator < len(potential_hits):
                self._hit_urls.append(potential_hits[iterator].get_attribute(LinkAttributeSelector.locator))
                iterator += 1

            while x < len(self._hit_urls):  # Used only for debug
                print(self._hit_urls[x])
                print("\n")
                x += 1

            self.run_parallel()
            return 1
        else:
            print("Not working")
            return

    def run_parallel(self):
        function_pool = multiprocessing.Pool(processes=4)

        for result in function_pool.map(self.check_hits, self._hit_urls):
            if result[0] == 0:
                self._hit_urls_bools.append(0)
                self._hit_urls_emails.append(result[1])
                self._hit_urls_nips.append(result[2])
            else:
                self._hit_urls_bools.append(1)
                self._hit_urls_emails.append(result[1])
                self._hit_urls_nips.append(result[2])

    def get_hit_urls(self):  # Added method to access hit URLs
        return self._hit_urls

    def get_hit_urls_bools(self):
        return self._hit_urls_bools

    def get_hit_urls_emails(self):
        return self._hit_urls_emails

    def get_hit_urls_nips(self):
        return self._hit_urls_nips

    def get_cleaned_hit_urls(self):
        cleaned_urls = []

        for hit_url in self._hit_urls:
            # Split the URL based on "://"
            parts = hit_url.split("://", 1)
            if len(parts) > 1:
                domain_parts = parts[1].split("/", 1)
                cleaned_domain = parts[0] + "://" + domain_parts[0]
                cleaned_urls.append(cleaned_domain)
            else:
                cleaned_urls.append(hit_url)  # If "://" is not found, retain the original URL

        return cleaned_urls

    def get_scraped_company_names_from_url(self):
        scraped_company_names = []

        for hit_url in self._hit_urls:
            parts = hit_url.split("://", 1)
            if len(parts) > 1:
                if len(parts[1].split("www.")) < 2:
                    domain_parts = parts[1].split(".", 1)
                    cleaned_domain = domain_parts[0]
                    scraped_company_names.append(cleaned_domain)
                else:
                    domain_parts = parts[1].split(".")
                    cleaned_domain = domain_parts[1]
                    scraped_company_names.append(cleaned_domain)
            else:
                scraped_company_names.append(hit_url)  # If "://" is not found, retain the original URL

        return scraped_company_names

    @staticmethod
    def check_hits(hit_urls):
        result = MapPage.open_hit_page(hit_urls)

        if result[0]:
            print(f"Hit {hit_urls} is good \n")
            return 1, result[1], result[2]
        else:
            print(f"Hit {hit_urls} is not good \n")
            return 0, result[1], result[2]

    @staticmethod
    def open_hit_page(hit):
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")  # Run in headless mode
        chrome_options.add_argument("--disable-javascript")  # Disable JavaScript
        chrome_options.add_argument("--disable-gpu")  # Disable GPU rendering

        driver = webdriver.Chrome(options=chrome_options)

        print("Opening URL:", hit)
        driver.set_page_load_timeout(TimeoutSelector.timeout_time)  # Close driver if loading takes too long
        try:
            driver.get(hit)
            source = driver.page_source.lower()

            email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}'  # ToDo implement function doing that
            email_match = re.search(email_pattern, source)

            nip_pattern = r'nip[:\-\s]*([0-9\s-]+)'
            nip_match = re.search(nip_pattern, source)

            if email_match and nip_match:
                email = email_match.group()
                nip = nip_match.group()
                print("Email address found:", email, " on", hit)
                print("NIP found:", MapPage.clear_nip(nip), " on", hit)

                return 1, email, MapPage.clear_nip(nip)
            if email_match:
                email = email_match.group()
                print("Email address found:", email, " on", hit)
                print("No NIP found on", hit)

                return 1, email, 'None'
            if nip_match:
                nip = nip_match.group()
                print("NIP found:", MapPage.clear_nip(nip), " on", hit)
                print("No email address found on", hit)

                return 1, 'None', MapPage.clear_nip(nip)
            else:
                print("No email address found on", hit)
                print("No NIP found on", hit)

                return 1, 'None', 'None'

        except TimeoutException:
            print("Timeout while opening", hit)
            driver.close()
            return 0, 'None', 'None'

        except Exception as e:
            print("Error while opening", hit, ":", str(e))
            driver.close()
            return 0, 'None', 'None'

    @staticmethod
    def clear_nip(nip):
        pattern = r'nip[:\-\s]*([0-9\s-]+)'

        matches = re.findall(pattern, nip)

        cleared_nip = ''.join(matches).replace('-', '').replace(' ', '')

        return cleared_nip


class ExportManagerGSheet:
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        "C:\\Users\\Igor\\PycharmProjects\\SimiScrap\\key2.json",
        scope)  # Insert your Google API credentials here

    client = gspread.authorize(credentials)

    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1_BTRg9Deova90IlbCqcha6kaKOUZhVQyA3YcrNys6Bk/edit?usp" \
                      "=sharing"  # Replace with your Google Sheets link
    current_date = datetime.now().strftime("%Y-%m-%d")

    @staticmethod
    def change_sheet_name(worksheet_title):
        try:
            spreadsheet = ExportManagerGSheet.client.open_by_url(ExportManagerGSheet.spreadsheet_url)
            worksheet = None
            current_datetime = datetime.now()

            sheet_name = f"Result {current_datetime.strftime('%d:%m:%y %H:%M:%S')}"
            try:
                worksheet = spreadsheet.worksheet(worksheet_title)
            except gspread.exceptions.WorksheetNotFound:
                pass  # Worksheet doesn't exist, so no need to clear

            if worksheet:
                worksheet.update_title(sheet_name)
                print("Worksheet '{}' name updated.".format(worksheet_title))
            else:
                print("Worksheet '{}' not found, can't change name.".format(worksheet_title))

        except Exception as e:
            print("Error while changing worksheet name'{}': {}".format(worksheet_title, str(e)))

    @staticmethod
    def clear_worksheet(worksheet_title):
        try:
            spreadsheet = ExportManagerGSheet.client.open_by_url(ExportManagerGSheet.spreadsheet_url)
            worksheet = None
            try:
                worksheet = spreadsheet.worksheet(worksheet_title)
            except gspread.exceptions.WorksheetNotFound:
                pass  # Worksheet doesn't exist, so no need to clear

            if worksheet:
                worksheet.clear()
                print("Worksheet '{}' cleared successfully.".format(worksheet_title))
            else:
                print("Worksheet '{}' not found, no need to clear.".format(worksheet_title))

        except Exception as e:
            print("Error while clearing worksheet '{}': {}".format(worksheet_title, str(e)))

    @staticmethod
    def insert_data(map_search, worksheet):
        header_row = [['Company Name', 'Company Domain', 'Is hit good?', 'Email', 'NIP', 'Source']]
        if worksheet.cell(1, 1).value != 'Company Name':
            worksheet.insert_rows(header_row, 1)

        clean_hit_urls = map_search.get_cleaned_hit_urls()

        data_to_insert = []
        for i in range(len(clean_hit_urls)):
            data_to_insert.append(
                [map_search.get_scraped_company_names_from_url()[i], clean_hit_urls[i],
                 map_search.get_hit_urls_bools()[i],
                 map_search.get_hit_urls_emails()[i],
                 map_search.get_hit_urls_nips()[i],
                 'WebScrapping'])

        worksheet.insert_rows(data_to_insert, 2)

    @staticmethod
    def append_workbook(map_search):
        try:
            worksheet_title = WorksheetTitleSelector.worksheet_title

            spreadsheet = ExportManagerGSheet.client.open_by_url(ExportManagerGSheet.spreadsheet_url)
            try:
                worksheet = spreadsheet.worksheet(worksheet_title)
            except gspread.exceptions.WorksheetNotFound:
                print("Did not found GSheet")
                return

            ExportManagerGSheet.insert_data(map_search, worksheet)

            print("Data successfully exported to the '{}' worksheet.".format(worksheet_title))

        except Exception as e:
            print("Error while exporting to Google Sheets: {}".format(str(e)))

    @staticmethod
    def create_workbook(map_search):
        try:
            worksheet_title = WorksheetTitleSelector.worksheet_title
            ExportManagerGSheet.change_sheet_name(worksheet_title)

            spreadsheet = ExportManagerGSheet.client.open_by_url(ExportManagerGSheet.spreadsheet_url)
            try:
                worksheet = spreadsheet.worksheet(worksheet_title)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=worksheet_title, rows="50",
                                                      cols="6")

            ExportManagerGSheet.insert_data(map_search, worksheet)

            print("Data successfully exported to the '{}' worksheet.".format(worksheet_title))

        except Exception as e:
            print("Error while exporting to Google Sheets: {}".format(str(e)))


def main():
    start_time = time.time()
    driver_initializer = Driver()

    import sys
    if len(sys.argv) > 1:
        company_name = sys.argv[1]
        append_mode = sys.argv[2]

        print(company_name)
    else:
        company_name = "Sabner"
        append_mode = 0

    company_name_initializer = CompanyNameSelector(company_name)

    try:
        base_search = BasePage(driver_initializer)
        company_search = SearchPage(driver_initializer)
        map_search = MapPage(driver_initializer)

        base_search.search_target(company_name_initializer)

        if company_search.check_if_valid():
            if map_search.operate_map_search(company_search.open_similar_search()):
                try:
                    if append_mode:
                        ExportManagerGSheet.append_workbook(
                            map_search
                        )
                    else:
                        ExportManagerGSheet.create_workbook(
                            map_search
                        )
                except PermissionError:
                    print("Error while exporting to Google Sheets.")

        driver_initializer.driver.close()
        print("--- %s seconds ---" % (time.time() - start_time))

    except Exception as e:
        print(f"Error: as {e}")
        driver_initializer.driver.close()
        print("--- %s seconds ---" % (time.time() - start_time))


if __name__ == '__main__':
    main()
