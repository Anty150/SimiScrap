from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By


def search_target():
    open_google()
    wait_for_search_bar()
    operate_search_bar()


def initialize_webdriver():  # ToDo Make it work
    driver = webdriver.Chrome()
    driver.maximize_window()
    # return initialized_driver


def wait_for_search_bar():
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
    except:
        print("Search bar not found, quiting")
        driver.quit()


def input_company_name():
    name = "Premium Solutions"  # Debug
    # name = input("Input company name: ")
    return name


def find_search_bar():
    search = driver.find_element(By.NAME, "q")
    return search


def open_google():
    driver.get("https://www.google.com/")
    search_bar = driver.find_element(By.ID, "L2AGLb")  # Accept cookies
    search_bar.send_keys(Keys.ENTER)
    return search_bar


def operate_search_bar():
    search = find_search_bar()
    search.send_keys(company_name)
    search.send_keys(Keys.ENTER)


def check_if_search_correct(search):
    try:
        assert company_name in search
    except:
        return 0
    return 1


def check_if_valid():
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "sATSHe"))
        )
        # driver.find_elements(By.CLASS_NAME, "liYKde g VjDLd") # ToBe Deleted
    except:
        print("Unable to find company card")
        return 0
    return 1


def open_similar_search():  # ToDo Separate to smaller functions and rename
    search = driver.find_elements(By.CLASS_NAME, 'sATSHe')
    search = search[-1]

    link = search.find_element(By.TAG_NAME, "a")
    link_value = link.get_attribute("href")

    print(link_value)  # Only Debug
    operate_map_search(link_value)


def check_if_on_correct_page():
    try:
        is_page_found = 1
        if WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.liYKde.g.VjDLd"))
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


def operate_map_search(link_value):
    driver.get(link_value)
    if check_if_on_correct_page():
        print("Working")
        hit_urls = []
        iterator = 0
        x = 0

        potential_hits = driver.find_elements(By.CSS_SELECTOR, "a.yYlJEf.Q7PwXb.L48Cpd.brKmxb")
        print(f"Results found: {len(potential_hits)}")
        # time.sleep(5) # Use only to debug
        while iterator < len(potential_hits):
            hit_urls.append(potential_hits[iterator].get_attribute("href"))
            iterator += 1

        while x < len(hit_urls):  # Used only for debug
            print(hit_urls[x])
            print("\n")
            x += 1

        # print(hit_urls[-1])
        # open_hit_page(hit_urls[-1])
        check_hits(hit_urls)
    else:
        print("Not working")


def check_hits(hit_list):
    iterator = 0
    for i in hit_list:
        if open_hit_page(hit_list[iterator]):
            print(f"Hit {hit_list[iterator]} is good \n")
        else:
            print(f"Hit {hit_list[iterator]} is not good \n")
        iterator += 1


def open_hit_page(hit):  # ToDo Implement function
    driver.get(hit)
    try:
        source = driver.page_source.lower()
        iterator = 0

        for x in keyword_list:
            if keyword_list[iterator] in source:
                return 1
            iterator += 1
        return 0
    except:
        print(f"Error while opening {hit}")
        driver.close()


def main():
    global driver, company_name, keyword_list

    driver = webdriver.Chrome()
    # initialize_webdriver() # ToDo make it work
    company_name = input_company_name()
    keyword_list = ["cnc", "solidcam", "solidworks"]  # ToDo Do integration with external text file containing keywords
    search_target()

    if check_if_valid():
        open_similar_search()
    driver.close()


if __name__ == '__main__':
    main()
