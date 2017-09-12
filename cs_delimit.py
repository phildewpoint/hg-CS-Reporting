import appJar
import openpyxl
import os
import datetime
import csv
goButtons = [
    'Submit',
    'Cancel'
]
app = appJar.gui(title='CS File Delimiter')


def create_file(file_name, file_dir):
    """This function will take a name and directory and return a file object.


    Keyword arguments:
    file_name -- name of a file to save (str)
    file_dir -- directory path where the file is saved (str)
    """

    if file_name is None or file_dir is None:
        msg = "The file you're trying to save is missing a directory or name. Please try the program again"
        app = appJar.gui()
        app.errorBox(title="Missing Key Information", message=msg)
        app.go()
        app.stop()
        quit()
    else:
        file_name = file_name + datetime.date.today().strftime("%d%m%Y")
        file_path = os.path.join(file_dir, file_name)
        file = open(file_path, mode='w')
        return file


def excel_pull(file):
    # get active worksheet
    ws = openpyxl.load_workbook(filename=app.getEntry(name='excelFile')).active
    ws_values = ws.values
    dl = app.getEntry(name='Delimiter')
    for i in ws_values:
        mycsv = csv.writer(file, delimiter=dl)  # type: csv.writer()
        mycsv.writerow(i)


def submit(button):
    if button == 'Cancel':
        quit()
    elif button == 'Submit':
        if not app.getEntry(name='Delimiter'):
            app.errorBox(title='EXCEPTION TO ADDRESS', message='You need a delimiter entered to run this utility')
        elif not app.getEntry(name='excelFile'):
            app.errorBox(title='EXCEPTION TO ADDRESS', message='You need an excel file to run this utility')
        else:
            # create file
            file = create_file(file_name='my_delimited_file', file_dir=os.path.dirname(app.getEntry(name='excelFile')))
            # pull xlsx file
            excel_pull(file=file)
            app.infoBox(title='Utility Complete!', message='Your file was created @ ' + file.name)
            quit()


def main():
    # header label
    app.addLabel(title='appMsg', text='Select a file and a delimiter. A new .txt file will be created.')
    app.addHorizontalSeparator()
    # pick the excel file to delimit
    app.addLabel(title='excelBox', text='Select an Excel file to delimit. Only the first tab is used.')
    app.addFileEntry(title='excelFile')
    # pick the delimiter
    app.addLabelEntry(title='Delimiter')
    # add buttons to cancel/submit
    app.addButtons(names=goButtons, funcs=submit)
    app.setEntryDefault(name='Delimiter', text='Character to add between values')
    app.go()


if __name__ == '__main__':
    main()
