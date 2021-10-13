# Voting Data Parser: Perdue vs. Ossoff
# This program parses the data from the November 3rd General Election 
# from the New York Times website and compiles it into an Excel file.

# Import data parsing library, cleaning library, and Microsoft Excel library

from urllib.request import urlopen
import re
import xlsxwriter
from xlsxwriter import Workbook

# Start creating new workbook

perdueVsOssoff = Workbook("PerdueVsOssoff.xlsx", {'strings_to_numbers': True})

# Only need to import data first time; webpage changes so I use less resources when I do it once and save

# Save URL

# url = "https://www.nytimes.com/interactive/2020/11/03/us/elections/results-georgia-senate.html"

# Open and read in the bytes from the webpage

# page = urlopen(url)
# html_bytes = page.read()
# raw_html = html_bytes.decode("utf-8")

nyt_og_html = open("Perdue vs. Ossoff_HTML.txt", "r")
raw_html = nyt_og_html.read()

# Save data separated by character to these two files

raw_html = raw_html[143679:267211]
nyt_html = open("Perdue vs. Ossoff_HTML.html", "w")
nyt_txt = open("Perdue vs. Ossoff_txt.txt", "w")

nyt_html.write(raw_html)
nyt_txt.write(raw_html)

nyt_html.close()
nyt_txt.close()

# Isolate county data and separate by county

county_data = re.findall("\"name\".*?\"leader_party_id\":\".*?\"", raw_html)

# For each county, clean the data

for i in range(0, len(county_data)):
    county_data[i] = county_data[i].replace('"', '').replace("{", "").replace("}", "").split(",")
    
    # For each datapoint in the county, split it up by colon
    
    for j in range(0, len(county_data[i])):
        county_data[i][j] = county_data[i][j].split(":")

# Keep an array of essential categories

keep = ["name", "votes", "absentee_votes", "eevp", "eevp_value", "eevp_source", "absentee_max_ballots", "purdued", "perdued", "ossoffj", "hazels", "leader_margin_value", "leader_margin_display", "leader_margin_name_display", "leader_party_id", "results", "results_absentee", "write-ins"]

# For each county

for i in range(len(county_data)):
    
    # For each data point, if it isn't a necessary datapoint, remove it
    
    for j in range(len(county_data[i]) - 1, -1, -1):
        if county_data[i][j][0] not in keep:
            county_data[i].pop(j)

# Time to write the spreadsheet!

# For each county

for i in range(len(county_data)):
    
    # Create a new sheet with the county name
    
    sheet = perdueVsOssoff.add_worksheet(county_data[i][0][1])
    
    # Start the bar graph of in-person results with isolated data from each candidate
    
    results_chart = perdueVsOssoff.add_chart({"type": "bar"})
    results_chart.add_series({"categories": [county_data[i][0][1], 0, 7, 0, 10], "values": [county_data[i][0][1], 1,  7, 1, 10]})
    results_chart.set_title({"name": "Results"})
    results_chart.set_x_axis({"name": "Number of (In-Person) Votes"})
    results_chart.set_y_axis({"name": "Candidate"})
    sheet.insert_chart("D5", results_chart)
    
    # Start the bar graph of absentee results with isolated data from each candidate
    
    absentee_chart = perdueVsOssoff.add_chart({"type": "bar"})
    absentee_chart.add_series({"categories": [county_data[i][0][1], 0, 11, 0, 14], "values": [county_data[i][0][1], 1, 11, 1, 14]})
    absentee_chart.set_title({"name": "Absentee Results"})
    absentee_chart.set_x_axis({"name": "Number of Absentee Votes"})
    absentee_chart.set_y_axis({"name": "Candidate"})
    sheet.insert_chart("L5", absentee_chart)
    
    # For each datapoint
    
    for j in range(len(county_data[i])):
        
        # For each "category:data" pair
        
        for k in range(len(county_data[i][j])):
            
            # If the row is 3 long, keep the two important data points at the top and move the title to the bottom
           
            if len(county_data[i][j]) == 3:
                if k == 0:
                    sheet.write(2, j, county_data[i][j][k])
                else:
                    sheet.write(k - 1, j, county_data[i][j][k])
                continue
            
            # Otherwise, just write it to the sheet
            
            sheet.write(k, j, county_data[i][j][k])

# Close and save the sheet

perdueVsOssoff.close()

# Print diagnostic "DONE" message

print("DONE")