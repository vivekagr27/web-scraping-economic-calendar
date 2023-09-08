from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup
import json
import csv
import time as t
import datetime as dt
import pandas as pd


# Function for scrapping data from the website and then saving it into csv file
def scraping_data():
    service = Service()
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service = service, options = options)
    # print(t.time())     # debugging
    driver.get(url)
    # print(t.time())     # debugging
    wait = WebDriverWait(driver, 600)

    # checking for pop-up of sign-in. If yes, click it else ignore
    try:
        wait.until(EC.element_to_be_clickable(('xpath', "//i[@class='popupCloseIcon largeBannerCloser']")))
        driver.find_element('xpath', "//i[@class='popupCloseIcon largeBannerCloser']").click()
    except:
        print("No pop-up encountered")

    # iterating over dates in intervals of 1 week starting from Jan 1, 2001
    start_date = dt.date(2001, 1, 1)
    end_date = dt.date.today()
    while(start_date <= end_date):

        # start date of interval
        curr_start_date = pd.to_datetime(start_date).strftime("%m-%d-%Y")

        # end date of interval
        if (end_date < start_date + dt.timedelta(days = 6)):
            curr_end_date = pd.to_datetime(end_date).strftime("%m-%d-%Y")
        else:
            curr_end_date = pd.to_datetime(start_date + dt.timedelta(days = 6)).strftime("%m-%d-%Y")

        while True:
            done = scrap_data_for_date_selection([driver, wait], curr_start_date, curr_end_date)
            if (done == 1):
                break
            else:
                driver.refresh()

        start_date += dt.timedelta(days = 7)

    # print(t.time())     # debugging
 
    driver.stop_client()
    driver.close()
    driver.quit()
    service.stop()

    # print(t.time())    # debugging

# exctracting events info for a particular date range
def scrap_data_for_date_selection(paras, start_date, end_date):

    driver = paras[0]
    wait = paras[1]

    try:
        # clicking calender icon for choosing customized date range
        wait.until(EC.element_to_be_clickable(('xpath', "//a[@id='datePickerToggleBtn']")))
        element = driver.find_element('xpath', "//a[@id='datePickerToggleBtn']")
        driver.execute_script("arguments[0].click();", element)

        # clearing already filled input field for start date
        wait.until(EC.element_to_be_clickable(('xpath', "//input[@id='startDate']")))
        element = driver.find_element('xpath', "//input[@id='startDate']")
        element.clear()

        # filling the new start date in the input field
        wait.until(EC.element_to_be_clickable(('xpath', "//input[@id='startDate']")))
        element = driver.find_element('xpath', "//input[@id='startDate']")
        element.send_keys('%s' % start_date)

        # clearing already filled input field for end date
        wait.until(EC.element_to_be_clickable(('xpath', "//input[@id='endDate']")))
        element = driver.find_element('xpath', "//input[@id='endDate']")
        element.clear()

        # filling the new end date in the input field
        wait.until(EC.element_to_be_clickable(('xpath', "//input[@id='endDate']")))
        element = driver.find_element('xpath', "//input[@id='endDate']")
        element.send_keys('%s' % end_date)

        # Clicking Apply button for date changes
        wait.until(EC.element_to_be_clickable(('xpath', "//a[@id='applyBtn']")))
        element = driver.find_element('xpath', "//a[@id='applyBtn']")
        driver.execute_script("arguments[0].click();", element)

        # Waiting for the economic calendar data to load
        wait.until(EC.presence_of_element_located((By.ID, "economicCalendarData")))
    except:
        return 0

    # scroll to bottom of webpage, wait until all event entries in the table have loaded and then scroll back to the top
    t.sleep(1)
    driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
    t.sleep(2)
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='economicCalendarLoading'][contains(@style, 'display: none;')]")))
        driver.execute_script('window.scrollTo(document.body.scrollHeight, 0);')
    except Exception as e:
        print("Error: ", e)
        return 0

    # getting economic calender table from the data
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, "html.parser")
    table = soup.find_all(class_="js-event-item")

    # creating json object for all events in the table
    json_data = []
    for row in table:

        datetime = row.get('data-event-datetime').split(' ')
        date = datetime[0]
        eventRowId = row.get('id').split('_')[1]
        time = row.find(class_="first left time js-time").text
        currency = row.find(class_="left flagCur noWrap").text
        imp = row.find(class_="left textNum sentiment noWrap")
        imp_score = len(imp.find_all(class_='grayFullBullishIcon'))
        event = row.find(class_="left event").text.strip('\n')
        print(datetime)
        actual = soup.find("td", {"id": "eventActual_%s" % eventRowId}).text
        forecast = soup.find("td", {"id": "eventForecast_%s" % eventRowId}).text
        previous = soup.find("td", {"id": "eventPrevious_%s" % eventRowId}).text

        # creating json object for each event in the table
        json_data.append({
                'date': date,
                'time': time,
                'currency': currency,
                'imp': imp_score,
                'event': event,
                'actual': actual,
                'forecast': forecast,
                'previous': previous,
        })

    # Aappending the data received for this date range to the csv file
    extract_data(json_data)

    return 1   # no error encountered


# For adding the events data from 01/01/2001 - till this date to the csv file named data.csv
def extract_data(datas):

    try:
        with open(csv_file_name, 'a', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
            # writer.writeheader()
            for data in datas:
                writer.writerow(data)
    except IOError:
        print("I/O error")


if __name__ == '__main__':

    url = "https://www.investing.com/economic-calendar/"
    csv_columns = ['date', 'time', 'currency', 'imp', 'event', 'actual', 'forecast', 'previous']
    csv_file_name = "data.csv"

    with open(csv_file_name, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
        writer.writeheader()

    scraping_data()
