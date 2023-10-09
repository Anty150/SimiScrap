import time
import subprocess
import schedule
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def run_scraping_code_normal_mode():
    sheet_url = "https://docs.google.com/spreadsheets/d/1_BTRg9Deova90IlbCqcha6kaKOUZhVQyA3YcrNys6Bk/edit?usp=sharing"
    cell_address = "A2"

    command = f"python main.py {get_string_from_google_sheet(sheet_url, 'Podobne', cell_address), 0}"
    subprocess.call(command, shell=True)


def run_scraping_code_append_mode():
    sheet_url = "https://docs.google.com/spreadsheets/d/1_BTRg9Deova90IlbCqcha6kaKOUZhVQyA3YcrNys6Bk/edit?usp=sharing"
    cell_address = "A2"

    command = f"python main.py {get_string_from_google_sheet(sheet_url, 'Podobne', cell_address), 1}"
    subprocess.call(command, shell=True)


def get_string_from_google_sheet(url, sheet_name, adress):
    # Define the scope and credentials
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\Igor\\PycharmProjects\\SimiScrap\\key2.json",
                                                             scope)
    client = gspread.authorize(creds)

    try:
        spreadsheet = client.open_by_url(url)
        worksheet = spreadsheet.worksheet(sheet_name)

        cell_value = worksheet.acell(adress).value
        print(cell_value)
        return cell_value

    except Exception as e:
        print("An error occurred:", e)
        return "Sabner"  # Default value


def get_strings_from_txt_file(txt_file_path):
    with open(txt_file_path, "r") as file:
        lines = file.readlines()
    strings = [line.strip() for line in lines[0].split(",")]

    return strings


# google_sheet_string = get_string_from_google_sheet(sheet_url, sheet_name, cell_address)

# txt_file_path = "your_file.txt"
# txt_file_strings = get_strings_from_txt_file(txt_file_path)

if __name__ == '__main__':
    # Run scheduled tasks
    schedule.every(1).minutes.do(run_scraping_code_append_mode)  # Remove the parentheses here
    while True:
        schedule.run_pending()
        time.sleep(1)
