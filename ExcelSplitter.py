import pandas
import os
from datetime import datetime

"""
ExcelSplitter is designed to take Excel files as input as split according to user requirements. 
It was designed to be used in finance roles that deal with large Excel files on devices that struggle to load them in 
Microsoft Excel. As such, this tool allows users to split them down into smaller files by month or by quarter.
For more generic use, files can be split according to the number of rows per file.
"""

class XLSplitter:
    def __init__(self,directory,file,sheet,val=0):

        self.file = file
        self.sheet = sheet
        self.val=int(val)

        # When object is created, determine extension
        if self.file.lower().endswith(".xlsx"):
            fbase = self.file[:-5]
            self.ext = ".xlsx"
        elif self.file.lower().endswith(".xls"):
            fbase = self.file[:-4]
            self.ext = ".xls"
        else:
            print("XLSplitter currently only supports .xls and .xlsx files.")
            return 0

        # Create a new directory for the output
        path = directory + "/" + fbase
        self.path = path + "/" + fbase + "-"
        if not os.path.exists(path):
            os.mkdir(path)

    def byrows(self):

        # Checking number is valid
        if self.val <= 0 or not str(self.val).isnumeric():
            print("Please provide 'val' with a whole number above 0")
            return "ValError"

        df = pandas.read_excel(self.file, self.sheet)

        rows = df.shape[0]
        i = 0
        j = 1

        # While i is less than the number of rows, create a new dataframe and copy rows from the input file.
        # Once the number of rows in the new dataframe match the number of rows specified by the user, or once all
        # of the rows have been copied, output to Excel file.
        while i < rows:
            newdf = pandas.DataFrame(columns=df.columns)
            while len(newdf) <= self.val - 1:
                newdf = newdf.append(df.loc[i].copy())
                i += 1
                if i == rows:
                    break

            newdf.to_excel(self.path + str(j) + self.ext)
            j += 1


        return "Success"


    def bymonth(self, col):

        df = pandas.read_excel(self.file, self.sheet)

        dates = list(df[col])
        dates = sorted(set(dates))

        # If a value in the selected column is not in datetime format, return as an error
        for d in dates:
            if not isinstance(d, datetime):
                return "DateError"

        i = 0
        j = dates[0].month

        # While i is less than the number of unique dates, copy rows to a new dataframe.
        # When the next loop's date no longer matches the previous month (j), or all rows have been copied,
        # output to Excel file.
        while i < len(dates):
            y = str(dates[i].year)
            newdf = pandas.DataFrame(columns=df.columns)
            while j == dates[i].month:
                for d in df.index[df[col] == dates[i]].tolist():
                    newdf = newdf.append(df.loc[d].copy())

                i += 1
                if i == len(dates):
                    break

            if j < 10:
                newdf.to_excel(self.path + y + "-0" + str(j) + self.ext)
            else:
                newdf.to_excel(self.path + y + "-" + str(j) + self.ext)

            j += 1
            if j == 13:
                j = 1


        return "Success"


    def byquarter(self, col):

        df = pandas.read_excel(self.file, self.sheet)

        dates = list(df[col])
        dates = sorted(set(dates))

        for d in dates:
            if not isinstance(d, datetime):
                return "DateError"

        i = 0
        quarters = [4, 7, 10, 13]
        for q in quarters:
            if dates[0].month < q:
                j = q

        # Similar to bymonth above, when the next date reaches a new quarter, measured against j, output to Excel file.
        while i < len(dates):
            y = str(dates[i].year)
            newdf = pandas.DataFrame(columns=df.columns)
            while j > dates[i].month and j - dates[i].month != 12:
                for d in df.index[df[col] == dates[i]].tolist():
                    newdf = newdf.append(df.loc[d].copy())

                i += 1
                if i == len(dates):
                    break

            newdf.to_excel(self.path + y + "-Q" + str(int((j - 1)/3)) + self.ext)

            j += 3
            if j == 16:
                j = 4


        return "Success"