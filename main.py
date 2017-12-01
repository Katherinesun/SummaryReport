from appJar import gui
from timesheet import *


dcw_file = ""
paysht_file = ""
output_file = ""


def generate_timesheet(dcw_file, paysht_file, output_file):
    employees = get_employee_list(dcw_file)
    timesheet = TimeSheet(employees)
    timesheet.load_from_paysht_inp(paysht_file)
    timesheet.save_to_excel(output_file)


# handle button events
def press(button):
    if button == "Cancel":
        app.stop()
    else:
        global output_file
        output_file = app.saveBox(title='Save to Excel file',
                                  fileName='Timesheet & Summary Report',
                                  fileExt=".xlsx",
                                  fileTypes=[('Excel', '*.xlsx')])
        try:
            generate_timesheet(dcw_file, paysht_file, output_file)
            app.infoBox('Success!', '%s is generated!' % output_file)
        except Exception as e:
            app.errorBox('Failed!',
                         'Failed to generate report.\n' +
                         'Error:\n' +
                         '%s' % e)
        finally:
            app.stop()

def press_b1(button):
    global dcw_file
    dcw_file = app.openBox(title='Select DCW file',
                           fileTypes=[('Excel', '*.xlsx')])
    app.setEntry('DCW file', dcw_file)


def press_b2(button):
    global paysht_file
    paysht_file = app.openBox(title='Select PAYSHT file',
                              fileTypes=[('csv', '*.csv'), ('INP', '*.INP')])
    app.setEntry('PAYSHT file', paysht_file)


if __name__ == "__main__":

    app = gui("Generate timesheet & summary report", "600x200")
    #app.setSticky("news")
    #app.setExpand("both")

    app.addLabelEntry("DCW file", 0, 0)
    app.addNamedButton('Browse', 'b1', press_b1, 0, 1)
    app.addLabelEntry("PAYSHT file", 1, 0)
    app.addNamedButton('Browse', 'b2', press_b2, 1, 1)
    app.addButtons(["Generate", "Cancel"], press)

    # start the GUI
    app.go()
