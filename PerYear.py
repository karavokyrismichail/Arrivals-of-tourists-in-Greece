'''
In this programm I take the number of tourists that came in Greece for each year between years 2011-2014,
by reading Excel files into pandas DataFrames for each year.
Then I combine the DataFrames of each year into one DataFrame
Finally for each year and its Number of Tourists,
I use the final DataFrame to extract csv file, insert the new table in a database and show a graph. 
'''


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sqlite3

excel_path = 'C:\\Users\\Micha\\Desktop\\Project\\' 

#creation of list to put all the dataframes

dfs = [] 

#read xls file of year 2011 into pandas dataframe

df1 = pd.read_excel(excel_path + '2011.xls', sheet_name= 'ΔΕΚ',  skiprows= 134, nrows= 1, usecols= "G", header=None, converters={6:int}) 

#insert a column that shows the year of dataframe 

df1.insert(1, "Year", '2011') 

#add dataframe in the list

dfs.append(df1) 
#print(df1)

#for loop that reads the xls files for years 2012-2014 into pandas dataframes,
#inserts a new column that shows the year of each dataframe and adds the dataframe in the list

for year in range(2012, 2015): 
    df2 = pd.read_excel(excel_path + str(year) + '.xls', sheet_name= 'ΔΕΚ',  skiprows= 136, nrows= 1, usecols= "G", header=None, converters={6:int}) 
    df2.insert(1, "Year", str(year)) 
    dfs.append(df2)
    #print(df2)

#combine the values of all dataframes in the list

data = pd.concat(dfs) 
#print(data)

#names each column of dataframe

data.columns = ['Arrivals', 'Year'] 
print(data)

#export dataframe to csv file

data.to_csv(excel_path + 'Arrivals_per_Year.csv', index=False, encoding="utf-8") 

#create new database if it doesn't exists, if database exists uses the current database

db = sqlite3.connect(excel_path + 'db.sqlite') 

#insert new table with the dataframe in database 
  
data.to_sql('Arrivals_per_year', db, if_exists='replace', index=False) 

#create graph

x = data['Year']
y = data['Arrivals']
plt.bar(x, y)
plt.title('Arrivals of Tourists per Year between 2011-2014')
plt.ylabel('Arrivals')
plt.show()

