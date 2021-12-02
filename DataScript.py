import pandas as pd
import PyPDF2
from pathlib import Path
import shutil, os
import os.path
import urllib.request
import urllib.error
import glob
import os
import openpyxl
import xlsxwriter


# !!NB!! column with URL's should be called: "Pdf_URL" and the year should be in column named: "Pub_Year"

# File names will be the ID from the ID column (e.g. BR2005.pdf)

# EDIT HERE:
    
# specify path to file containing the URLs
data = pd.read_excel("C:/Users/KOM/01 Scripts input/GRI_2017_2020.xlsx")
os.chdir("C:/Users/KOM/PycharmProjects")
# specify Output folder (in this case it moves one folder up and saves in the script output folder)
pth = "C:/Users/KOM/PycharmProjects/"

# Specify path for existing downloads
dwn_pth = "C:/Users/KOM/PycharmProjects/dwn/"

# check for files already downloaded
dwn_files = glob.glob(os.path.join(dwn_pth, "*.pdf")) 
exist = [os.path.basename(f)[:-4] for f in dwn_files]

# specify the ID column name
ID = "BRnum"


##########
c = 0
returned_URLs_lists = []
returned_URLs_strings = []


# read in file
df = pd.DataFrame(data, columns=['Pdf_URL'])


# filter out rows with no URL
non_empty = df.notnull()
print(os.getcwd())

for i in range(20):
    if non_empty.values[i] == True:
        returned_URLs_lists.append(df.to_numpy().tolist()[i])
        returned_URLs_strings += (df.to_numpy().tolist()[i])

writer = pd.ExcelWriter(pth+"check_3.xlsx", engine="openpyxl", options={'strings_to_urls': False})
df_URLs = pd.DataFrame(returned_URLs_lists)

# loop through dataset, try to download file.
for URL in returned_URLs_strings:
    print(c)
    savefile = pth + "dwn/"
    try:
        URLfile = urllib.request.urlretrieve(URL, savefile)
        if os.path.isfile(savefile):
            try:
                with open(savefile, "wb") as df_URLs:
                    pdfReader = PyPDF2.PdfFileReader(df_URLs.write(URL))
                    if pdfReader.numPages > 0:
                        df_URLs.at[URL, 'pdf_downloaded'] = "yes"
                    else:
                        df_URLs.at[URL, 'pdf_downloaded'] = "file_error"

            except Exception as e:
                df_URLs.at[URL, 'pdf_downloaded'] = str(e)
                print(str(str(URL)+" " + str(e)))
                print("First exception happened!")
        else:
            df_URLs.at[URL, 'pdf_downloaded'] = "404"
            
    except (urllib.error.HTTPError, urllib.error.URLError, ConnectionResetError, Exception) as e:
        df_URLs[str(URL), "error"] = str(e)
        print("Second exception happened!")
    c += 1


df_URLs.to_excel(writer, sheet_name="dwn")
writer.save()
writer.close()
