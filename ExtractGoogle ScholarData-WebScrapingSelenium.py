# Import ChromeOptions to set ChromeDriver options
from selenium.webdriver import Chrome, ChromeOptions
# Import By to determine the method of finding an element
from selenium.webdriver.common.by import By
# Import WebDriverWait to wait for a page to load
from selenium.webdriver.support.ui import WebDriverWait
# Import expected_conditions to determine if a page has loaded
from selenium.webdriver.support import expected_conditions as EC

import requests  # Import requests to get the HTML code of a page
import openpyxl  # Import openpyxl to create an Excel workbook
import time     # Import time to add delays to the program

# Import Crossref to get the DOI of a publication
from habanero import Crossref

###############################################################################

# Create an Excel workbook
wb = openpyxl.Workbook()  # Create a workbook
sheet = wb.active  # Get the active sheet

# Write the header row
headers = ['Title', 'Authors', 'Year', 'Journal', 'Volume', 'Issue', 'Pages',
           'Publisher', 'Citations', 'Document Type', 'DOI', 'Link']  # Create a list of headers
sheet.append(headers)  # Write the headers to the sheet

###############################################################################

# Get the information about Literature from the Google Scholar page of each Literature


def get_literature_data(url):
    lit_data_dic = {}  # Create a dictionary to store the literature data

    browser_options = ChromeOptions()  # Create a ChromeOptions object
    # Set the option headless to True to avoid opening a browser window or False to open a browser window
    # Comment this line to open a browser window
    browser_options.add_argument('--headless')

    driver = Chrome(options=browser_options)  # Create a ChromeDriver object

    driver.get(url)  # Open the URL in the browser window

    title = driver.find_element(
        By.XPATH, "//div[@id='gsc_oci_title']").text.strip()  # Get the title
    # print(title)

    # Search for the DOI of the publication
    # try:
    cr = Crossref()  # Create a Crossref object
    doi_search = cr.works(query=title)
    # Get the DOI of the publication
    doi = doi_search['message']['items'][0]['DOI']
    print(doi)
    # except:
    # doi = 'NA'

    # Get the publication information
    i = 1
    last_field = ''
    while (last_field != 'Total citations'):  # Loop until the last field is 'Total citations'
        try:
            field = driver.find_element(
                By.XPATH, "//div[@id='gsc_oci_table']//div[@class='gs_scl'][" + str(i) + "]//div[@class='gsc_oci_field']").text.strip()  # Get the field
            value = driver.find_element(
                By.XPATH, "//div[@id='gsc_oci_table']//div[@class='gs_scl'][" + str(i) + "]//div[@class='gsc_oci_value']").text.strip()  # Get the value
            # If the field is 'Total citations', get the value from the third word
            if (field == 'Total citations'):
                value = value.split()[2]
                lit_data_dic[field] = value
                last_field = field
                # print(field + ': ' + value)
                break
            else:
                lit_data_dic[field] = value
                # print(field + ': ' + value)
            i = i + 1
        except:
            break
            pass

    # Get the publication type
    if (lit_data_dic.get("Journal", "") != ""):
        doc_type = "Article"
    elif (lit_data_dic.get("Conference", "") != ""):
        doc_type = "Conference"
    elif (lit_data_dic.get("Book", "") != ""):
        doc_type = "Book"
    else:
        doc_type = "Other"
    # print(doc_type)

    # Write the publication information to the Excel sheet
    sheet.append([title, lit_data_dic.get("Authors", ""), lit_data_dic.get("Publication date", ""), lit_data_dic.get("Journal", (lit_data_dic.get("Conference", (lit_data_dic.get("Book", ""))))), lit_data_dic.get("Volume", ""),
                 lit_data_dic.get("Issue", ""), lit_data_dic.get("Pages", ""), lit_data_dic.get("Publisher", ""), lit_data_dic.get("Total citations", ""), doc_type, doi, url])

    lit_data_dic.clear()  # Clear the dictionary
    driver.quit()  # Close the browser window

###############################################################################

# Get the Google Scholar data of a scholar


def get_google_scholar_data(url):
    browser_options = ChromeOptions()  # Create a ChromeOptions object
    # Set the option headless to True to avoid opening a browser window or False to open a browser window
    # Comment this line to open a browser window
    browser_options.add_argument('--headless')

    driver = Chrome(options=browser_options)  # Create a ChromeDriver object
    driver.maximize_window()  # Maximize the browser window
    driver.get(url)  # Open the URL in the browser window

    # Wait for the page to load
    # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//html[@class='gs_el_ta gs_el_sm']")))

    scholar_name = driver.find_element(
        By.ID, 'gsc_prf_in').text.strip()  # Get the scholar name
    print("Scholar Name: " + scholar_name)  # Print the scholar name

    count = 0  # Initialize the count of publications to 0
    num_of_literatures = driver.find_elements(
        By.XPATH, "//table[@id='gsc_a_t']//tbody[@id='gsc_a_b']//tr[@class='gsc_a_tr']")  # Get the number of publications in the initially loaded page
    for row in num_of_literatures:  # Loop through the publications
        count = count + 1  # Increment the count of publications
    # print(count)

    # keep scrolling down until all publications are loaded
    while True:
        n_count = 0
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")  # Scroll down to the bottom of the page
        time.sleep(5)  # Wait for 5 seconds
        show_more_button = driver.find_element(
            By.XPATH, "//button[@id='gsc_bpf_more']").click()  # Click the 'Show more' button
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")  # Scroll down to the bottom of the page
        time.sleep(5)  # Wait for 5 seconds
        num_of_literatures = driver.find_elements(
            By.XPATH, "//table[@id='gsc_a_t']//tbody[@id='gsc_a_b']//tr[@class='gsc_a_tr']")  # Get the number of publications in the loaded page
        for row in num_of_literatures:  # Loop through the publications
            n_count = n_count + 1  # Increment the count of publications
        if (n_count == count):  # If the count of publications in the loaded page is equal to the count of publications in the previous page,
            break
        else:
            count = n_count  # Update the count of publications
    print("Number of Publications: " + str(count))

    # Get Publications Web address
    for i in range(1, count + 1):
        publications_url = driver.find_element(
            By.XPATH, "//tbody[@id='gsc_a_b']//tr[@class='gsc_a_tr'][" + str(i) + "]//td[@class='gsc_a_t']//a").get_attribute('href')  # Get the URL of each publication
        # print(str(i) + " out of " + str(count) + " : " + publications_url)
        # Print the number of publications being executed
        print("Executing: " + str(i) + " out of " + str(count))
        # Get the information about Literature from the Google Scholar page of each Literature
        get_literature_data(publications_url)

    driver.quit()  # Close the browser window

    return scholar_name

###############################################################################

# Define the main() function


def main(url):  # Define the main() function

    # Get the Google Scholar data of a scholar
    scholar_data = get_google_scholar_data(url)

    # Save the Excel workbook
    file_name = scholar_data + '_GoogleScholar.xlsx'  # Create the file name
    wb.save(file_name)  # Save the Excel workbook
    print('Google Scholar Data of ' + scholar_data +
          ' saved to Excel file successfully.')  # Print a success message

###############################################################################


# Execute the main() function
# Replace the URL with the URL of the Google Scholar page of the scholar
# main('https://scholar.google.com/citations?user=63PvvA4AAAAJ&hl=en')
# Execute the main() function
main('https://scholar.google.com/citations?hl=en&user=B7vSqZsAAAAJ&view_op=list_works&sortby=pubdate')
