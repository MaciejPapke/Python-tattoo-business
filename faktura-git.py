import os
import glob
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

# Create a webdriver for Firefox and open (get) the website.
driver = webdriver.Firefox()
driver.get("https://x")

# Find and fill the email and password bracket
email = driver.find_element(By.ID, "emailAddress")
email.send_keys("x")

password = driver.find_element(By.ID, "password")
password.send_keys("x")
# Click the logg inn button
submitbutton = driver.find_element(By.XPATH, "/html/body/x").click()

# Wait for the page to load
XPATHdownload = "/html/body/x"
element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATHdownload)))

# Click the download button and wait
downloadbutton = driver.find_element(By.XPATH, XPATHdownload).click()
time.sleep(10)

# Find the latest *.csv file in a downloads folder
# Close the firefox window if the file is succesfully downloaded
list_of_files = glob.glob('C:/x') # * means all if need specific format then *.csv
tmp = []
for i in range(len(list_of_files)):
    tmp.append(os.path.getctime(list_of_files[i]))
max = tmp.index(max(tmp))
latest_file = list_of_files[max]
# Open the file to write
file = open(latest_file)
# Close the website
driver.close()

# Change into commas
tmpstr = file.read().replace(";", ",")
file.close()
file = open(latest_file, "w")
file.write(tmpstr)
file.close()

# Get current date as a string and add to the new file at the start
from time import gmtime, strftime
date = strftime("%Y-%m-%d", gmtime())
# Copy and overwrite the newest .csv file to the archive folder
file = "C:/x"+date+"_x.csv"
shutil.copy(latest_file, file)
# Delete the old file
os.remove(latest_file)
# Move copy and move
maindir = "C:/x.xlsx"
maindirarchive = "C:/x.xlsx"
shutil.copy(maindir, maindirarchive)

# Read the .csv and create a dataframe
df = pd.read_csv(file)
col = 'Fakturanummer'

# Transition Fakturanummer to int64 and sort by values ascending - ready to pluck into the main excel
pd.to_numeric(df[col])
df.sort_values(by=col, ascending=True, inplace=True)
df[["Day", "Month", "Year"]] = df["Dato"].str.split('.',expand=True)
# Append the sorted dataframe to main .xlsx file
maindir = "C:/x.xlsx"
with pd.ExcelWriter(maindir, engine='openpyxl', mode="a", if_sheet_exists="overlay") as writer:
    df.to_excel(writer, sheet_name='Sheet1', header=True, index=False, index_label=None, startrow=0, startcol=8)
    writer._save()


