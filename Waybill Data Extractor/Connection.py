
# PLEASE DON'T EDIT ANY SYNTAX INSIDE THIS CODE. 
# THE DEVELOPER FOR THIS PROGRAM USED SCRIPTING METHOD INSTEAD OBJECT ORIENTED PROGRAMMING. THEREFORE, ANY CHANGES MADE WILL CAUSE CHANGES ALSO FOR OVERALL PROGRAM. 

from tkinter import Tk, filedialog
import requests
import pdfplumber
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from googletrans import Translator
import datetime
import pyttsx3
import tkinter as tk
from tkinter import filedialog
from glob import glob
import os
import fnmatch
import sys
from datetime import datetime
import json



# For text to speech engine
engine = pyttsx3.init ()
engine.setProperty ('rate', 160)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
translater = Translator ()

""" pyttsx3 imports drivers module based on specific platform. Fount at https://github.com/nateshmbhat/pyttsx3/issues/6 """

# System configuration
sys.stdin.reconfigure(encoding='utf-8')
sys.stdout.reconfigure(encoding='utf-8')

hiddenimports = [
    'pyttsx3.drivers',
    'pyttsx3.drivers.dummy',
    'pyttsx3.drivers.espeak',
    'pyttsx3.drivers.nsss',
    'pyttsx3.drivers.sapi5', ]

    
# For google sheet credentials
scope = ['https://www.googleapis.com/auth/spreadsheets',
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"]

cred = ServiceAccountCredentials.from_json_keyfile_name('Credentials.json', scope)
client = gspread.authorize(cred)
python_write = client.open("Receiving Waybill").worksheet("Under develop")

# For open file dialog box to search for a folder
root = Tk() # pointing root to Tk() to use it as Tk() in program.
root.withdraw() # Hides small tkinter window.
root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.
open_file = filedialog.askdirectory() # Returns opened path as str
print(open_file) 

#Checker if folder is selected or not
if open_file !="":
    engine.say ("Folder selected. The program will now start to run. Just relax and seatback!")
    engine.runAndWait ()
else:
    engine.say ("No folder selected. Program will not run.")
    engine.runAndWait ()
    

#Declare file path and number of files to iterate
multiple_files = glob (open_file+"/*pdf")
no_of_files = len(fnmatch.filter(os.listdir(open_file), '*.pdf'))
engine.say ("NUMBER OF PDF TO EXTRACT IS: {}".format(no_of_files))
engine.runAndWait ()
print ("NUMBER OF PDF TO EXTRACT: {}".format(no_of_files))

print ("**************************************************************************************************")
# Open each invoice in pdf format and print

#list_no_files = [*range(1,no_of_files,1)]

for i in multiple_files:
    with pdfplumber.open (i)as pdf:
        page = pdf.pages [0]
        waybill = page.extract_text()
        #print (waybill)
        print(os.path.splitext("C:/Users/63927/Desktop/New folder/Waybill/*pdf")[0])

          
# Iterate through each invoice and do the following data cleaning and sorting      
        
        for row in waybill.split ("\n"):
            if row.startswith ("箱上编号"):
                stra_code = row.split ()[1]

        for row in waybill.split ("\n"):
            if row.endswith ("WAYBILL DETAILS"):
                waybill_number = row.split ()[0]
                int_waybill_number = int(waybill_number)
                

        for row in waybill.split ("\n"):
            if row.startswith ("收货日期:"):
                date= row.split ()[1]
               
        #Create an empty list that will hold the informations for following variables.
        product_names = []
        cargo_specifications = []
        total_packages = []
        boxes_bags = []
        total_weights = []
        cubic_numbers = []
        remarks = []
        values = ("1 ","2 ","3 ", "4 ", "5 ", "6 ","7 ","8 ","9 ","10 ","11 ","12 ","13 ","14 ","15 ","16 ","17 ","18 ","19 ","20 "  )

        #Data cleaning
        for row in waybill.split ("\n"):       
            if row.startswith (values):
                try:
                    name = row.split ()[1]
                    specification = row.split ()[2]
                    packages = row.split ()[3]
                    box_bag = row.split ()[4]
                    total_weight = row.split ()[5]
                    cubic_number = row.split ()[6]
                    remark = row.split ()[-1]
                    total_packages.append (int(packages))
                    boxes_bags.append (float(box_bag))
                    cargo_specifications.append (specification)
                    total_weights.append (float(total_weight))
                    cubic_numbers.append (float(cubic_number))
                    product_names.append (name)
                    remarks.append (remark) 
                # This will hold the value error. 
                except:
                    name = row.split ()[1]
                    specification = row.split ()[3]
                    packages = row.split ()[4]
                    box_bag = row.split ()[5]
                    total_weight = row.split ()[6]
                    cubic_number = row.split ()[7]
                    remark = row.split ()[-1]

                    product_names.append (name)
                    total_packages.append (int(packages))
                    boxes_bags.append (float(box_bag))
                    cargo_specifications.append (specification)
                    total_weights.append (float(total_weight))
                    cubic_numbers.append (float(cubic_number))
                    remarks.append (remark) 
        
        product_names_translated = []
        for i in product_names:
            new = translater.translate((i), dest = "en")
            product_names_translated.append (new.text)

        product_names_translated
        
        print ("Stra code:",stra_code)
        print ("Waybill # :",int_waybill_number)
        print ("Waybill date:",date)
        print ("Product Name:", product_names_translated)
        print ("Cargo Specification:", cargo_specifications)
        print ("Total Cases/Packages:", total_packages)
        print ("Box/Bag KG:", boxes_bags)
        print ("Total Weight:", total_weights)
        print ("Cubic Number:", cubic_numbers)
       
        
        # Translate product names from chineese to english
        

        # Do final operations and append all informations from this console to google sheets
        stra_codes = len(product_names)*(stra_code,)
        waybill_numbers=len(product_names)*(int_waybill_number,)
        dates = len(product_names)*(date,)
        warehouse_number= waybill_number [0]

        if warehouse_number== "1":
            warehouse_text = "Shishi"
        elif warehouse_number == "2":
            warehouse_text = "Guangzhou"
        else:
            warehouse_text = "Yiwu"

        warehouse_texts = len (product_names)*(warehouse_text,)

        current_date = datetime.now()
        current_date_str = current_date.strftime("%m/%d/%Y  %H:%M:%S")
        current_date_strs = len (product_names)*(current_date_str,)

        #engine.say ("Appending Waybill {}.". format (i))
        #engine.runAndWait ()
        
        for i in range (0,len(product_names)):
            lst = [current_date_strs,stra_codes, waybill_numbers,warehouse_texts,dates,product_names,product_names_translated ,cargo_specifications, total_packages, boxes_bags, total_weights,cubic_numbers]
            lst_for_console_print = [current_date_strs,stra_codes, waybill_numbers,warehouse_texts,dates,product_names_translated ,cargo_specifications, total_packages, boxes_bags, total_weights,cubic_numbers]
            lst2 = []
            lst2.append([x[i] for x in lst])
        
            
            #print (lst2[0]) # comment out since it cause character error since it is in chineese language
            python_write.append_row (lst2[0])
        print ("**************************************************************************************************")
        # Checking for a potential error upon running the progrma
        if i == len (product_names)-1:
            engine.say ("New waybill appended".format (i))
            engine.runAndWait ()
        else:
            engine.say ("There's an error upon running this program. Please check. Thank you!")
            engine.runAndWait ()
            print("**************************************************************************************************")


