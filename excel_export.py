from appJar import gui
import openpyxl
import re
import os
import datetime
buttons = [
    'Go!',
    'Cancel'
]
# TODO create regex to check file names per report
fileRegex = [
    ""
]


def launch(button):
    if button == "Cancel":
        quit()
    elif not app.getEntry(name="find_dir"):
        msg = "You need to pick a directory to start the app."
        app.warningBox(title="ERROR: Missing Information", message=msg)
    else:
        direc = app.getEntry(name="find_dir")
        # check the chosen directory for # of files
        entry = os.scandir(direc)
        file_cnt = 0
        filelist = []
        for i in entry:
            if not i.name.startswith('.') and i.is_file():
                # if file is csv, warn user the utility doesn't accept CSV
                # TODO - add handling that CSV files are converted instead of errored out
                if os.path.splitext(i)[1] == '.csv':
                    msg = "Utility can only take Excel files. This file is a .csv: "
                    app.infoBox(title="Critical Error", message=msg + i.name)
                    quit()
                file_cnt += 1
                filelist.append(i.name)
        # create workbook
        wb = openpyxl.Workbook()
        # count each loop for file naming
        counter = 1
        # copy/paste from other xlsx into new file
        for excel in filelist:
            # create a new sheet per file (uses default naming)
            wb_sheet = wb.create_sheet(title="HG Report " + str(counter))
            # load saved off workbook based on selected directory
            load_wb = openpyxl.load_workbook(os.path.join(direc, excel))
            # pull the active sheet
            copy_sheet = load_wb.active
            rows = []
            # pull in each row
            for row in copy_sheet.iter_rows():
                # create list to save row data
                row_data = []
                # go through each cell in the row and save data
                for cell in row:
                    row_data.append(cell.value)
                # add full row of data and save in master 'rows' data; loop to next row
                rows.append(row_data)
            # send all data to newly created sheet
            for source_rows in rows:
                wb_sheet.append(source_rows)
            counter += 1
        wb.save(filename=(os.path.join(direc, "CS_Compiled_Reports" + str(datetime.date.today()) + ".xlsx")))
        finish_up(file_cnt=file_cnt)
        quit()


def run_app():
    msg = "Runs calculations over multiple spreadsheets. Select where all the source spreadsheets are stored.\n"
    msg2 = "A spreadsheet will be created in that folder combining this data."
    app.addLabel(title="header", text=msg + msg2)
    app.addHorizontalSeparator()
    app.addDirectoryEntry(title="find_dir", )
    app.addButtons(names=buttons, funcs=launch)
    app.go()


def calculate():
    quit()


def finish_up(file_cnt):
    comp_msg = "Your workbook is complete.\n"
    total_msg = "Spreadsheets combined: "
    app.infoBox(title="Workbook Complete!", message=comp_msg + total_msg + str(file_cnt))

app = gui(title="CS Report Utility")
run_app()
