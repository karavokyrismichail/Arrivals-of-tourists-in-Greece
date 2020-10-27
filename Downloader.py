from bs4 import BeautifulSoup as bs
import requests
import webbrowser #for 2nd way


URL = 'https://www.statistics.gr/el/statistics/-/publication/STO04/'
MONTH = '-Q4'
YEAR = 2011
FILETYPE = 'VBZOni0vs5VJ' #Filter
List = []

def get_soup(url):
    return bs(requests.get(url).text, 'html.parser')

#finds all XLS links between 2011-2014 (4th quarter)

while YEAR < 2015:
    for link in get_soup(URL + str(YEAR) + MONTH).find_all('a'):
        file_link = link.get('href')
        if FILETYPE in file_link: #use the filter to find links of a specific type
            List.append(file_link) #adds every link in the list
    YEAR = YEAR + 1  

YEAR = 2011 #i use it again to name the xls of each year

#downloads only the xls I want to use in the same folder that the script was ran from

for x in range(2, 12, 3): 
    dls = List[x]
    responce = requests.get(dls)

    output = open(str(YEAR)+'.xls', 'wb')
    output.write(responce.content)
    output.close()
    YEAR = YEAR + 1


'''
#downloads only the links I want to use 2nd way
for x in range(2, 12, 3):
    print(List[x])
    webbrowser.open(List[x]) 
'''

