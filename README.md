# Extract-Google-Scholar-Data
This is a Python Program implemented using Selenium Library for Web Scraping the Literature Data of a Scholar from a Google Scholar page by inputting the URL of the google scholar, and output will be exported in Excel.

Packages Required:
1. Selenium: pip install selenium
2. habanero: pip install habanero
3. openpyxl
4. requests

Download Chrome Driver from Selenium Webpage and extract it to a location of your choice from https://chromedriver.chromium.org/downloads matching your Chrome browser version.
Add Environment Path Variable to the Chrome driver.

Input: Google Scholar URL Page (Eg: https://scholar.google.com/citations?hl=en&user=B7vSqZsAAAAJ&view_op=list_works&sortby=pubdate) - Google Scholar Page of Richard Feynmann

Output: Excel with the following Columns of Data
1. Title
2. Authors
3. Date
4. Year
5. Source
6. Title
7. Volume
8. Issue
9. Pages
10. Publisher
11. Citations
12. Document Type
13. DOI
14. Google Scholar Link of Literature
