import os
import re
import RPA.Browser.Selenium
from robocorp.tasks import task
from RPA.Excel.Files import Files
from robocorp.tasks import get_output_dir
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
#How do i search for the search phrases in the output data?

class Robot_exe:
    def __init__(self, website  = "https://www.enca.com/"):

        #Initialize browser url and selenium browser
        self.browser = None
        self.website = website

        #Money formats
        self.money_formats = [
        r'^\$\d+\.\d{2}$',                               # $11.11
        r'^\$\d{1,3}(,\d{3})*\.\d{2}$',                  # $111,111.11 
        r'\b\d+\s+dollars\b',                            # 11 dollars
        r'\b\d+\s+USD\b'                                 # 11 USD
        ]

        #Dictionary as a container to temporarily store all scraped data
        self.news = {'Title:':[],
                'Description:':[],
                'Date:':[],
                'No of search phrases in title:':[],
                'No of search phrases in Description:':[],
                'Money:':[]
                }
        

    def open_browser(self, search_phrase):
        
        self.search_phrase = search_phrase

        self.browser = RPA.Browser.Selenium.Selenium()

        #Open news website: eNCA
        try:

            self.browser.open_available_browser(self.website)

        except TimeoutException as e:

            self.browser.close_all_browsers()

            print(f"Browser Timed Out: {e}")

        #wait for dynamic content to load:
        self.browser.wait_until_element_is_visible('name:search', timeout=4)

        #Input search phrase in search bar
        self.browser.input_text('name:search', self.search_phrase + Keys.ENTER)

        #Wait for search results to load
        #Check if search available
        
        try:
            if(self.browser.find_elements('css:view-unformatted') != None):

                self.browser.wait_until_element_is_visible(
                    'css:.view-unformatted', timeout=2
                    )  
            else:

                print(f"Swarch does not exits: {self.search_phrase}")

        except TimeoutException as e:

            self.browser.close_all_browsers()
            
            print(f"Browser timed out. Did not find: {self.search_phrase}\n ERROR: {e}")


    def extract_web_data(self):
        
        #Find titles
        titles_url = self.browser.find_elements('css:.card_heading')

        for titles in titles_url:

            #Append TEXT >> to List news[]
            self.news['Title:'].append(
                self.browser.get_text(titles)
                )

        #Find descriptions
        description_url = self.browser.find_elements('css:.blurb')

        for description in description_url:

            #Append TEXT >> to List news[]
            self.news['Description:'].append(
                self.browser.get_text(description)
                )

    
        #Find dates
        dates_url = self.browser.find_elements('css:.published-date')

        for dates in dates_url:

            #Append TEXT >> to List news[]
            self.news['Date:'].append(
                self.browser.get_text(dates)
                )


        return self.news

    def clean_data(self):

        #Cleans dictionary from blank spaces
        for keys, values in self.news.items():

            if isinstance(values, list):

                self.news[keys] = [item for item in values if item]

        return self.news


    #Search for instances of search phrase in dictionary 'self.news'
    def look_up_phrase_and_money(self):
        
        #Search for occurences in Titles
        for t in range(len(self.news['Title:'])):

            title = self.news['Title:'][t]
            
            title_count = title.lower().count(
                self.search_phrase.lower()
                )

            self.news['No of search phrases in title:'].append(
                title_count
                )

        #Search for occurences in Description
        for n in range(len(self.news['Description:'])):

            description = self.news['Description:'][n]
        
            description_count = description.lower().count(
                self.search_phrase.lower()
                )

            self.news['No of search phrases in Description:'].append(
                description_count
                )

        #Look for money formats in Dictionary:
        #self.money_formats: $11.1 | $111,111.11 | 11 dollars | 11 USD
        for check_money in self.news['Title:']:

            found_match = any(
                re.search(pattern, check_money)
                    for pattern in self.money_formats
                )
            
            self.news['Money:'].append(found_match)

        return self.news

    #FUNCTION THAT SAVES DATA TO EXCEL
    def save_excel(self):
        excel = Files()
        wb = excel.create_workbook()
        wb.create_worksheet('Results')
        excel.append_rows_to_worksheet(self.news,header=True,name='Results')
        wb.save(os.path.join(get_output_dir(),'results.xlsx'))

@task
def main():
    run = Robot_exe()
    
    search_phrase = 'Injured'

    run.open_browser(search_phrase)

    run.extract_web_data()

    run.clean_data()

    run.look_up_phrase_and_money()

    run.save_excel()

if __name__ == "__main__":
    main()