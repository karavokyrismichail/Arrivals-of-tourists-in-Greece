# Arrivals-of-tourists-in-Greece

Python Scripts that download data from the Hellenic Statistical Authority and plot four charts about tourism in greece during the period 2011-2014

## Tools
- Beautiful Soup
- pandas
- Matplotlib

## Scripts

#### Downloader.py
- Downloads XLS files about arrivals of tourists from the Hellenic Statistical Authority (page scraping - using Beautiful Soup)


#### PerCountry.py
- Processes unstructured data from xls format files (using pandas)
- Extracts a Graph that shows the number of tourists that came in Greece during 2011-2014 from the top 10 countries with most arrvals (using Matplotlib)
- Exports the structured data to csv file and store them to an sqlite3 database

#### PerQuarter.py
- Processes unstructured data from xls format files (using pandas)
- Extracts a Graph that shows for each quarter of the year during 2011-2014, the number of arrivals in Greece (using Matplotlib)
- Exports the structured data to csv file and store them to an sqlite3 database

#### PerTransport.py
- Processes unstructured data from xls format files (using pandas)
- Extracts a Graph that shows the number of arrivals in Greece for each transport during 2011-2014 (using Matplotlib)
- Exports the structured data to csv file and store them to an sqlite3 database


#### PerYear.py
- Processes unstructured data from xls format files (using pandas)
- Extracts a Graph that shows the number of arrivals in Greece for each year during 2011-2014 (using Matplotlib)
- Exports the structured data to csv file and store them to an sqlite3 database


## Charts

![Charts](https://github.com/karavokyrismichail/Arrivals-of-tourists-in-Greece/blob/main/Charts/Screenshot_1.png)

## Notes
- Run Downloader.py First
- Before running each of the rest scripts, change the excel_path 



