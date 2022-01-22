import pandas as pd
import requests
from bs4 import BeautifulSoup

# Extracting the .xlsx file
dataFrame = pd.read_excel('C:/Users/TANUL/PycharmProjects/blackcofferInternshipAssessment/AssessmentDetails/Input.xlsx',
                          sheet_name='Sheet1')

# Creating a dictionary for 'URL_ID' and 'URL', where 'URL_ID' is the key and 'URL' is the value
url_dict = dict(zip(dataFrame['URL_ID'], dataFrame['URL']))

# Scrapping the text data from the url and storing it into a text file
headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0',
}
for key in url_dict:
    print('URL --', key)
    filename = "%s.txt" % key
    html_text = requests.get(url_dict[key], headers=headers).text
    soup = BeautifulSoup(html_text, 'lxml')
    heading = soup.find('h1', class_='entry-title').text
    content = soup.find('div', class_='td-post-content').text
    unwanted = soup.find('pre', class_='wp-block-preformatted')
    if unwanted != None:
        unwanted = unwanted.text
        content = "".join(content.rsplit(unwanted))
    textFileAddress = 'C:/Users/TANUL/PycharmProjects/blackcofferInternshipAssessment/textData/%s' % filename
    file = open(textFileAddress, 'w', encoding="utf-8")
    file.write(heading)
    file.write('\n')
    file.write(content)
    file.close()
