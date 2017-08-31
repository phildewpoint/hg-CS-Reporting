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
                file_cnt += 1
                filelist.append(i.name)
        app.setLabel(name="file_count", text=("Files Located: " + str(file_cnt)))
        # create workbook
        wb = openpyxl.Workbook()
        # copy/paste from other xlsx into new file
        for excel in filelist:
            load_wb = openpyxl.load_workbook(os.path.join(direc, excel))
            copy_sheet = load_wb.active
            rows = []
            for row in copy_sheet.iter_rows():
                row_data = []
                for cell in row:
                    row_data.append(cell.value)

        wb.save(filename=(os.path.join(direc, "CS_Compiled_Reports" + str(datetime.date.today()) + ".xlsx")))
        app.addMessage(title="Workbook Complete!", text="Your workbook is complete.")
        quit()


def run_app():
    msg = "Runs calculations over multiple spreadsheets. Select where all the source spreadsheets are stored.\n"
    msg2 = "A spreadsheet will be created in that folder combining this data."
    app.addLabel(title="header", text=msg + msg2)
    app.addHorizontalSeparator()
    app.addDirectoryEntry(title="find_dir", )
    app.addLabel(title="file_count", text="Files Located: 0")
    app.addButtons(names=buttons, funcs=launch)
    app.go()


app = gui(title="CS Report Utility")
run_app()
