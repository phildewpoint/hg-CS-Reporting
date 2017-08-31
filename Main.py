from appJar import gui
import openpyxl
import re
buttons = [
    'Go!',
    'Cancel'
]
# TODO create regex to check file names per report
fileRegex = [
    ""
]


def run_app():
    msg = "Runs calculations over multiple spreadsheets. Select where all the source spreadsheets are stored and GO!"
    file_count = 0
    app.addDirectoryEntry(title="find_dir")
    app.addLabel(title="file_count", text=("Files Located: " + str(file_count)))
    app.go()


app = gui(title="CS Report Utility")
run_app()
