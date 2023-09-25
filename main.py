import time
from HelperFunctions import CheckFile
from YesNoPopUp import tkinter_main
if __name__ == '__main__':
    while True:
        meeting_report_file_path = CheckFile()
        print(meeting_report_file_path)
        if meeting_report_file_path is None:
            pass
        else:
            tkinter_main(file_path=meeting_report_file_path)
        time.sleep(15)
