import xlwt
import multiprocessing
from selenium.common import NoSuchElementException, TimeoutException
from xlwt import Workbook
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import re


# ToDo Fix a rare bug

# ToDo Add log what word was found in hit

# ToDO Add separate iterations based on what we are looking

# ToDo Jedna tabela z branżą (bazowane na liscie slow kluczowych)
class Driver:
    driver = None

    def __init__(self):
        self.driver = webdriver.Chrome()


class KeywordsSelector:
    _keywords = None

    def set_keywords(self):  # ToDo Fix error when words in list.txt are in two or more separate lines
        file_name = TextFileSelector.file_name
        contents = ""
        with open(file_name, encoding='utf8') as f:
            for line in f:
                contents += line
                # print(f"Loaded keywords: {contents}") # Only for debug
        self._keywords = contents.split()

    def get_keywords(self):
        return self._keywords


class CompanyNameSelector:
    _company_name = ""

    def set_company_name(self):
        # self.company_name = input("Set company name: ")
        # self._company_name = "Wemeco Poland Sp. z o.o."  # Used only for Debug
        self._company_name = "mzlaser"

    def get_company_name(self):
        return self._company_name


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
        search.send_keys(Keys.ENTER)
        search.submit()  # ToDo Figure out why it sometimes work diffrently


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
            print("Working")
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
        else:
            print("Not working")

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
        print("Running")
        result = MapPage.open_hit_page(hit_urls)

        if result[0]:
            print(f"Hit {hit_urls} is good \n")
            return 1, result[1]
        else:
            print(f"Hit {hit_urls} is not good \n")
            return 0, result[1]

    @staticmethod
    def open_hit_page(hit):

        driver = webdriver.Chrome()
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
                                print("NIP found:", value, f" on {hit}")
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

        except:
            print(f"Error while opening {hit}")
            driver.close()
            return 0, 'None'

    def get_hit_urls(self):
        return self._hit_urls

    def get_hit_urls_bools(self):
        return self._hit_urls_bools

    def get_hit_urls_nips(self):
        return self._hit_urls_nips


class ExportManager:  # ToDo Add option to toggle export on/off
    @staticmethod
    def create_workbook(hit_urls, hit_urls_bools, hit_urls_nips):
        wb = Workbook()
        iterator = 1

        sheet1 = wb.add_sheet('Scraping Result')

        sheet1.col(0).width = 12000
        sheet1.col(1).width = 20000
        sheet1.col(2).width = 20000
        header_font = xlwt.Font()
        header_font.name = 'Arial'
        header_font.bold = True

        header_style = xlwt.XFStyle()
        header_style.font = header_font

        sheet1.write(0, 0, 'Company Domain', header_style)
        sheet1.write(0, 1, 'Is hit good?', header_style)
        sheet1.write(0, 2, 'NIP', header_style)
        # sheet1.write(0, 2, 'Company Name', header_style)

        while iterator < len(hit_urls) + 1:  # Hardcoded due to indexing reason
            sheet1.write(iterator, 0, hit_urls[iterator - 1])
            sheet1.write(iterator, 1, hit_urls_bools[iterator - 1])
            sheet1.write(iterator, 2, hit_urls_nips[iterator - 1])
            # sheet1.write(iterator, 2, companyDomains[iterator - 1])
            iterator += 1

        wb.save('company_domains.xls')


# class ImportManager: #ToDo Implement Function

def main():
    driver_initializer = Driver()

    company_name_initializer = CompanyNameSelector()
    company_name_initializer.set_company_name()

    base_search = BasePage(driver_initializer)
    company_search = SearchPage(driver_initializer)
    map_search = MapPage(driver_initializer)

    base_search.search_target(company_name_initializer)

    if company_search.check_if_valid():
        map_search.operate_map_search(company_search.open_similar_search())
        ExportManager.create_workbook(map_search.get_hit_urls(), map_search.get_hit_urls_bools(), map_search.get_hit_urls_nips())

    driver_initializer.driver.close()


if __name__ == '__main__':
    main()
