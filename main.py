import datetime
import subprocess

import PyPDF2
import os
from os.path import exists
from datetime import date, timedelta

####################################################################################################
#### Account Cash Projection Automation                                      #######################
#### Reads PDF, Writes Text to txt file, parses txt file, extracts text,     #######################
#### Writes it to a import file, then executes bat file to                   #######################
#### import into accounting software                                         #######################
####################################################################################################

class ACCOUNTCashProjections:

    def __init__(self, dailyCashFile = None, pdfReader = None, pageObj = None, filepath = None, filename = None):
        self.dailyCashFile = dailyCashFile
        self.pdfReader = pdfReader
        self.pageObj = pageObj
        self.filepath = filepath
        self.filename = filename

    def file_path(self):
        self.filepath = "REPORTS DIRECTORY"
        return self.filepath

    def file_name(self):
        yesterday = (date.today() - timedelta(days=1)).strftime('%Y%m%d')
        self.filename = f"REPORT_CODE_{yesterday}.PDF"
        return self.filename

    def file_exists(self):
        file = (self.filepath+self.filename)
        return exists(file)

    def read_file(self):
        self.dailyCashFile = open(self.filepath+self.filename, 'rb')
        self.pdfReader = PyPDF2.PdfFileReader(self.dailyCashFile)
        return self.pdfReader

    def page_to_read(self):
        # will always be first page
        self.pageObj = self.pdfReader.getPage(0)
        return self.pageObj

    def create_txt(self):
        textFile = open("REPORTS DIRECTORY/test_files/text.txt", "w")
        textFile.write(self.pageObj.extractText())
        textFile.close()

    def find_sub_red_amounts(self):
        # reads txt file
        readText = open("REPORTS DIRECTORY/test_files/text.txt", "r")
        # defines the lines in string
        lines = readText.readlines()
        # loops through lines searching for keyword and returns number based on sub/red
        for i in range(0,len(lines)):
            if 'Subscription' in lines[i]:
                sub = lines[i]
                print('subscription')
                print(sub[43:58].replace("$", "").replace(",", "").strip())

            elif 'Redemption' in lines[i]:
                red = lines[i]
                print('redemption')
                print(red[44:56].replace("$","").replace(",","").strip())

    def create_import_file(self):
        readText = open("REPORTS DIRECTORY/test_files/text.txt","r")
        lines = readText.readlines()
        trade_date = []
        settle_date = []
        sub_amount = []
        red_amount = []

        for i in range(0, len(lines)):
            if 'Subscription' in lines[i]:
                sub = lines[i]
                sub_amount = sub[43:58].replace("$", "").replace(",", "").strip()

                settle_date_unordered = lines[i - 2][31:42].replace("/", "").strip()
                settle_date_year = settle_date_unordered[0:4]
                settle_date_month = settle_date_unordered[4:6]
                settle_date_day = settle_date_unordered[6:8]
                settle_date = f"{settle_date_month}{settle_date_day}{settle_date_year}"

                date = datetime.datetime(int(lines[i-2][31:35]), int(lines[i-2][36:38]), int(lines[i-2][39:41])).weekday()

                if date == 0:
                    # zero == Monday
                    trade_date_day = (int(settle_date_day) - 4)
                    trade_date = (f"{settle_date_month}{trade_date_day}{settle_date_year}")

                else:
                    trade_date_day = (int(settle_date_day) - 2)
                    trade_date = (f"{settle_date_month}{trade_date_day}{settle_date_year}")

            elif 'Redemption' in lines[i]:
                red = lines[i]
                red_amount = red[44:56].replace("$","").replace(",","").strip()
                date = datetime.datetime(int(lines[6][31:35]), int(lines[6][36:38]), int(lines[6][39:41])).weekday()
                if date == 0:
                    trade_date_day = (int(settle_date_day) - 4)
                    trade_date = (f"{settle_date_month}{trade_date_day}{settle_date_year}")
                else:
                    trade_date_day = (int(settle_date_day) - 2)
                    trade_date = (f"{settle_date_month}{trade_date_day}{settle_date_year}")

        os.chdir("C:/Users/bpaur/Desktop/test2/")

        if sub_amount != None:
            sub_line1 = f"ACCOUNT_NUMBER,dp,,caca,pendcash,{trade_date},{settle_date},,,,,awca,none,,,,y,{sub_amount}"
            sub_line2 = f"ACCOUNT_NUMBER,dp,,caca,cash,{settle_date},{settle_date},,,,,caca,pendcash,,,,n,{sub_amount}"
        else:
            pass

        if red_amount != None:
            red_line1 = f"ACCOUNT_NUMBER,wd,,caca,pendcash,{trade_date},{settle_date},,,,,awca,none,,,,y,{red_amount}"
            red_line2 = f"ACCOUNT_NUMBER,wd,,caca,cash,{settle_date},{settle_date},,,,,caca,pendcash,,,,n,{red_amount}"
        else:
            pass

        import_lines = [f"{sub_line1}\n", f"{sub_line2}\n", f"{red_line1}\n", f"{red_line2}\n"]

        file_object = open(r"IMPORT FILE DIRECTORY", "w")
        file_object.writelines(import_lines)

    def run_bat_file(self):
        filepath = "BAT FILE DIRECTORY"
        londonOps = subprocess.Popen(filepath, shell=True, stdout=subprocess.PIPE)
        stdout, stderr = londonOps.communicate()
        print(londonOps.returncode)

    def app_run(self):
        os.chdir("FILE ROOT DIRECTORY")
        self.file_name()
        self.file_path()

        if self.file_exists() == False:
            print('File does not exist. Please research')
        else:
            self.read_file()
            self.page_to_read()
            self.create_txt()
            self.create_import_file()
            self.run_bat_file()

if __name__ == "__main__":
    ACCOUNTCashProjections().app_run()
