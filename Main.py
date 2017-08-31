from appJar import gui
import openpyxl
import re
import os
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
    elif app.getEntry(name="find_dir") is not True:
        msg = "You need to pick a directory to start the app."
        app.warningBox(title="ERROR: Missing Information", message=msg)
    else:
        direc = app.getEntry(name="find_dir")
        entry = os.scandir(direc)
        app.setLabel(name="file_count", text=("Files Located: " + str(len(entry))))
        wb = openpyxl.Workbook()
        dest_wb = "PJD_Test_Workbook.xlsx"
        wb2 = openpyxl.load_workbook(direc.join('test.xlsx'))
        print(wb2.sheetnames)


def run_app():
    msg = "Runs calculations over multiple spreadsheets. Select where all the source spreadsheets are stored and GO!"
    app.addLabel(title="header", text=msg)
    app.addHorizontalSeparator()
    app.addDirectoryEntry(title="find_dir", )
    app.addLabel(title="file_count", text="Files Located: 0")
    app.addButtons(names=buttons, funcs=launch)
    app.go()


app = gui(title="CS Report Utility")
run_app()
