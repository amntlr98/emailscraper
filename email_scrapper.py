import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
import re

def extract_emails(text):
    emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    return emails

def scrape_emails_from_url(url, driver):
    try:
        if not url.startswith('http'):
            url = 'https://' + url
        if 'www.' not in url:
            url = url.replace('https://', 'https://www.')
        driver.get(url)
        text = driver.page_source
        emails = set(extract_emails(text))
        return emails
    except WebDriverException as e:
        print(f"Error scraping {url}: {e}")
        return set()

input_file = 'H2Sscrapeinput.xlsx' #use your input file
output_file = 'scraped_emails.xlsx' #name of the output file

df = pd.read_excel(input_file)

urls = df.iloc[:, 4].tolist()

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(options=options)

unique_emails = set()

url_emails_dict = {}

for url in urls:
    if isinstance(url, str):
        emails = scrape_emails_from_url(url, driver)
        unique_emails.update(emails)
        url_emails_dict[url] = emails


driver.quit()


data = {'URL': [], 'Emails': []}
for url, emails in url_emails_dict.items():
    data['URL'].append(url)
    data['Emails'].append(','.join(emails))


df_output = pd.DataFrame(data)
df_output.to_excel(output_file, index=False) #files with section of URL and email


unique_emails_df = pd.DataFrame({'Emails': list(unique_emails)})
unique_emails_df.to_excel('unique_emails.xlsx', index=False) # seperate file for emails only

print("URLs and corresponding emails scraped and saved successfully.")
