import xlwt
from xlwt import Workbook
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By


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
                print(f"Loaded keywords: {contents}")
        self._keywords = contents.split()

    def get_keywords(self):
        return self._keywords


class CompanyNameSelector:
    _company_name = ""

    def set_company_name(self):
        # self.company_name = input("Set company name: ")
        self._company_name = "NSI Glogow"  # Used only for Debug

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
        search_bar = self.driver.find_element(CookieSelector.finder, CookieSelector.locator)  # Accept cookies
        search_bar.send_keys(Keys.ENTER)

    def wait_for_search_bar(self):
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((SearchBarSelector.finder, SearchBarSelector.locator))
            )
        except:
            print("Search bar not found, quiting")
            self.driver.quit()

    def operate_search_bar(self, company_name):
        search = self.find_search_bar()
        search.send_keys(company_name)
        search.send_keys(Keys.ENTER)


class SearchPage:
    def __init__(self, driver_initializer):
        self.driver = driver_initializer.driver

    def check_if_valid(self):
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((CompanyCardSelector.finder, CompanyCardSelector.locator))
            )
        except:
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

    def __init__(self, driver_initializer, keyword_list):
        self.driver = driver_initializer.driver
        self.keyword_list = keyword_list

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
            # hit_urls = []
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

            self.check_hits()
        else:
            print("Not working")

    def check_hits(self):
        iterator = 0
        for i in self._hit_urls:
            if self.open_hit_page(self._hit_urls[iterator]):
                print(f"Hit {self._hit_urls[iterator]} is good \n")
                self._hit_urls_bools.append(1)
            else:
                print(f"Hit {self._hit_urls[iterator]} is not good \n")
                self._hit_urls_bools.append(0)
            iterator += 1

    def open_hit_page(self, hit):
        self.driver.get(hit)
        try:
            source = self.driver.page_source.lower()
            iterator = 0

            for x in self.keyword_list:
                if self.keyword_list[iterator] in source:
                    return 1
                iterator += 1
            return 0
        except:
            print(f"Error while opening {hit}")
            self.driver.close()

    def get_hit_urls(self):
        return self._hit_urls

    def get_hit_urls_bools(self):
        return self._hit_urls_bools


class ExportManager:  # ToDo Add option to toggle export on/off
    @staticmethod
    def create_workbook(hit_urls, hit_urls_bools):
        wb = Workbook()
        iterator = 1

        sheet1 = wb.add_sheet('Scraping Result')

        sheet1.col(0).width = 12000
        sheet1.col(1).width = 20000
        # sheet1.col(2).width = 20000
        header_font = xlwt.Font()
        header_font.name = 'Arial'
        header_font.bold = True

        header_style = xlwt.XFStyle()
        header_style.font = header_font

        sheet1.write(0, 0, 'Company Domain', header_style)
        sheet1.write(0, 1, 'Is hit good?', header_style)
        # sheet1.write(0, 2, 'Company Name', header_style)

        while iterator < len(hit_urls):
            sheet1.write(iterator, 0, hit_urls[iterator - 1])
            sheet1.write(iterator, 1, hit_urls_bools[iterator - 1])
            # sheet1.write(iterator, 2, companyDomains[iterator - 1])
            iterator += 1

        wb.save('company_domains.xls')


def main():
    driver_initializer = Driver()

    company_name_initializer = CompanyNameSelector()
    company_name_initializer.set_company_name()

    keyword_initializer = KeywordsSelector()
    keyword_initializer.set_keywords()

    google_search = BasePage(driver_initializer)
    company_search = SearchPage(driver_initializer)
    map_search = MapPage(driver_initializer, keyword_initializer.get_keywords())

    google_search.search_target(company_name_initializer)

    if company_search.check_if_valid():
        # company_search.open_similar_search()
        map_search.operate_map_search(company_search.open_similar_search())
        ExportManager.create_workbook(map_search.get_hit_urls(), map_search.get_hit_urls_bools())
    driver_initializer.driver.close()


if __name__ == '__main__':
    main()
