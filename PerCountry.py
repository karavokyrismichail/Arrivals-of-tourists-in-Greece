'''
In this programm I take the number of tourists that came in Greece of each Country between years 2011-2014, 
by reading Excel files from each year into pandas DataFrames.
Then I combine the DataFrames of each year into one DataFrame and group by country.
Finally for each Country and its Number of Tourists, I use the final DataFrame to extract csv file, 
insert the new table in a database and show a graph. 
'''


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sqlite3

excel_path = 'C:\\Users\\Micha\\Desktop\\Project\\'

#creation of list to put all the dataframes

dfs = []

#map for the rows i need to skip for each year

skiprows = {2011: (76, 110, 120, 124, 131), 
            2012: (78, 112, 122, 126, 132), 
            2013: (78, 113, 123, 127, 133), 
            2014: (78, 113, 123, 127, 133)} 

#tuple for the rows I want to read

nrows = (32, 7, 2, 5, 1) 

#for loop that reads the xls files into pandas dataframes for years 2011-2014 by using skiprows map and nrows tumple

for year in range(2011, 2015): 
    for skiprow, nrow in zip(skiprows[year], nrows):
        df1 = pd.read_excel(excel_path + str(year) + '.xls', sheet_name= 'ΔΕΚ',  skiprows= skiprow, nrows = nrow, usecols= "B,G", header=None, converters={6:int})
        dfs.append(df1) #add dataframe in the list 
        #print(df1)

#combine the values of all dataframes in the list and groups the values of same countries

data = pd.concat(dfs).groupby([1]).sum().reset_index() 
#print(data)

#name each column of dataframe

data.columns = ['Countries', 'Arrivals'] 
print(data)

#export dataframe to csv file

data.to_csv(excel_path + 'Arrivals_per_Country.csv', index=False, encoding="utf-8") 

#create new database if it doesn't exists, if database exists uses the current database

db = sqlite3.connect(excel_path + 'db.sqlite') 

#insert new table with the dataframe in database 

data.to_sql('Arrivals_per_Country', db, if_exists='replace', index=False) 

#sort 10 countries with most arrivals 

data = data.sort_values('Arrivals', ascending=False).head(10) 

#creation of graph

x = data['Countries']
y = data['Arrivals']
plt.bar(x, y)
plt.title('Countries with most Arrivals of tourists between 2011-2014')
plt.ylabel('Arrivals')
plt.show()

