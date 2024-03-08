import os
import pandas
import PySimpleGUI as psg
import ExcelSplitter as xls

# Function for byte measurement conversion after passing the respective threshold
def filesizeCalc(size):
    KB = 1024
    MB = KB * 1024
    GB = MB * 1024
    TB = GB * 1024

    if size < KB:
        return str(size) + " Bytes"

    # If updated to Python 3.10 would use match and case keywords
    def switch(size):
        if KB <= size < MB:
            return str(round(size / KB, 2)) + " KB"
        elif MB <= size < GB:
            return str(round(size / MB, 2)) + " MB"
        elif GB <= size < TB:
            return str(round(size / GB, 2)) + " GB"
        elif size >= TB:
            return str(round(size / TB, 2)) + " TB"

    return switch(size)

# Function to create and display another window - mainly used for error messages
def errWin(m):
    warn = [
        [
            psg.Text(text=m)
        ],
        [
            psg.Button(button_text="OK", enable_events=True, size=(7, 2), key="-OK-")
        ]
    ]

    error = psg.Window("Alert!", warn, size=(300,80), element_justification='c')

    while True:
        e, v = error.read()
        if e == "-OK-" or e == psg.WIN_CLOSED:
            break

    error.close()

# Function to sheet details
def sheetUpdate(f, s):
    df = pandas.read_excel(f, s)
    window["-ROW-"].update(df.shape[0])
    window["-COL-"].update(df.shape[1])
    window["-COLS-"].update(list(df.columns))


# Lists of what will appear in the window
# File selection
file_select = [
    [
        psg.Text("Select folder:")
    ],
    [
        # Creates a Browse button which allows the user to browse folders
        psg.FolderBrowse(),
        psg.In(size=(50, 10), enable_events=True, key="-FOLDER-")
    ],
    [
        # Listbox to show files found
        psg.Listbox(
            values=[], enable_events=True, size=(70, 20), key="-FILELIST-"
        )
    ]
]

# File details + sheet and split method selection
details = [
    [
        psg.Text("File size: "),
        psg.Text(text="", key="-SIZE-")
    ],
    [
        psg.Text("Sheets:"),
        psg.Listbox(values=[], enable_events=True, size=(20, 2), key="-SHEETS-")
    ],
    [
        psg.Text("Rows: "),
        psg.Text(text="", key="-ROW-")
    ],
    [
        psg.Text("Columns: "),
        psg.Text(text="", key="-COL-")
    ],
    [
        psg.HSeparator(pad=(0,20))
    ],
    [
        # Split method selection
        psg.Radio(text="Rows", default=True, enable_events=True, group_id="rinput", key="-ROWS-"),
        psg.Radio(text="Monthly", enable_events=True, group_id="rinput", key="-MON-"),
        psg.Radio(text="Quarterly", enable_events=True, group_id="rinput", key="-QTR-")
    ],
    [
        psg.Text("Rows per file:", pad=(40,0), key="-RTX-"),
        psg.Text("File will be split per month from column:", pad=(0,10), visible=False, key="-MTX-"),
        psg.Text("File will be split per quarter from column:", pad=(0,10), visible=False, key="-QTX-")
    ],
    [
        psg.Input(default_text=1000, size=(20, 1), pad=(40,10), key="-IN-"),
        psg.Listbox(values=[], enable_events=True, pad=(34,0), size=(20, 4), visible=False, key="-COLS-")
    ],
    [
        psg.Button(button_text="Exit", enable_events=True, size=(7,2), pad=(30,30), key="-EXIT-"),
        psg.Button(button_text="Split", enable_events=True, size=(7,2), key="-SPLIT-")
    ]
]

# Final layout
layout = [
    [
        psg.Column(file_select),
        psg.VSeparator(),
        psg.Column(details)
    ]
]

# Creating instance of window using the layout
window = psg.Window("SplitSheet", layout)

# Default split method
splitstyle = "r"

while True:

    # Read the window and display
    event, values = window.read()

    # If Exit button is pressed or window is closed, end the loop and finish the script
    if event == "-EXIT-" or event == psg.WIN_CLOSED:
        break

    # If folder is selected, make a list of files in the folder
    if event == "-FOLDER-":
        folder = values["-FOLDER-"]

        # Change directory to folder selected - if no folder, unpopulate all folder/file details
        try:
            os.chdir(folder)
        except:
            window["-FILELIST-"].update("")
            window["-SIZE-"].update("")
            window["-SHEETS-"].update([])
            window["-ROW-"].update("")
            window["-COL-"].update("")
            window["-COLS-"].update([])
            continue

        # Get list of files in folder
        try:
            file_list = os.listdir(folder)
        except:
            file_list = []

        # Only show Excel files with extensions xlsx and xls
        fnames = [
            f for f in file_list if os.path.isfile(os.path.join(folder, f)) and f.lower().endswith((".xlsx", ".xls"))
        ]

        # Update the window to show file names
        window["-FILELIST-"].update(fnames)

    # If a file is selected from the Listbox
    elif event == "-FILELIST-":
        try:
            filename = os.path.join(folder, values["-FILELIST-"][0])
        except:
            continue

        # Error handling - If a sheet is chosen for one file, but then the user changes to another file, that sheet
        # is still selected. If sheet exists, the sheet will be deleted when changing to a different file.
        try:
            del sheet
        except:
            pass

        # Populating details of file selected
        window["-SIZE-"].update(filesizeCalc(os.stat(filename).st_size))
        exfile = pandas.ExcelFile(filename)
        window["-SHEETS-"].update(exfile.sheet_names)

        window["-ROW-"].update("")
        window["-COL-"].update("")
        window["-COLS-"].update([])

        # If only one sheet, store to sheet variable
        if len(exfile.sheet_names) == 1:
            sheet = exfile.sheet_names[0]
            sheetUpdate(exfile, sheet)

    # If a sheet is selected
    elif event == "-SHEETS-":
        try:
            sheet = values["-SHEETS-"][0]
        except:
            continue
        sheetUpdate(exfile, sheet)

    # User selecting split method - changed visibility of options according to choice
    elif event == "-ROWS-":
        window["-MTX-"].update(visible=False)
        window["-QTX-"].update(visible=False)
        window["-COLS-"].update(visible=False)
        window["-RTX-"].update(visible=True)
        window["-IN-"].update(visible=True)
        splitstyle = "r"

    elif event == "-MON-":
        window["-QTX-"].update(visible=False)
        window["-RTX-"].update(visible=False)
        window["-IN-"].update(visible=False)
        window["-MTX-"].update(visible=True)
        window["-COLS-"].update(visible=True)
        splitstyle = "m"

    elif event == "-QTR-":
        window["-RTX-"].update(visible=False)
        window["-IN-"].update(visible=False)
        window["-MTX-"].update(visible=False)
        window["-QTX-"].update(visible=True)
        window["-COLS-"].update(visible=True)
        splitstyle = "q"

    # Split button
    elif event == "-SPLIT-":
        # Checking the user has selected everything required
        try:
            if not os.path.exists(folder + "/SplitSheet"):
                os.mkdir("SplitSheet")
            filename = filename
            sheet = sheet
        except:
            errWin("Please select a file (and sheet).")
            continue

        if not values["-IN-"].isnumeric() and splitstyle == "r":
            window["-IN-"].update("Numbers only")
            continue
        if values["-COLS-"] == [] and splitstyle != "r":
            errWin("Please select the column with dates.")
            continue

        # Creating object using directory, file, sheet, (and number of rows to split by) selected
        toSplit = xls.XLSplitter(folder + "/SplitSheet", values["-FILELIST-"][0], sheet, values["-IN-"])

        # Call method based on split method selected
        if splitstyle == "r":
            errCheck = toSplit.byrows()
            if errCheck == "ValError":
                errWin("Please enter a whole number above 0.")
        elif splitstyle == "m":
            errCheck = toSplit.bymonth(values["-COLS-"][0])
            if errCheck == "DateError":
                errWin("There is data in this column that aren't dates.\nCheck the data and try again.")
        elif splitstyle == "q":
            errCheck = toSplit.byquarter(values["-COLS-"][0])
            if errCheck == "DateError":
                errWin("There is data in this column that aren't dates.\nCheck the data and try again.")

        if errCheck == "Success":
            errWin("Split complete!")


# Close window when user closes window!
window.close()