import shutil
from pathlib import Path
import time

import pandas as pd  # pip install pandas
import xlwings as xw  # pip install xlwings
from watchdog.events import FileSystemEventHandler  # pip install watchdog
from watchdog.observers import Observer

MASTER_FILE = './workarea/master.xlsx'


class MyMonitor(FileSystemEventHandler):
    def on_created(self, event):
        """Method that executes each time a new file is detected in the specified directory.

        Args:
            event (dict): directory and event details
        """

        # Excluding temporary files
        if '~$' in str(event.src_path):
            return

        print(f"A new file has been detected in --> {event.src_path}")

        self.path_f = Path(event.src_path)

        if event.src_path.endswith(".xlsx") and self.path_f.is_file():
            self.status = 1
            self.__copy_file_data(event)
            self.__move_file(event)

        if not event.src_path.endswith(".xlsx") and self.path_f.is_file():
            self.status = 0
            self.__move_file(event)

        return

    def __copy_file_data(self, event):
        """Method to extract data from each detected excel sheet and copy it to the master excel.

        Args:
            event (dict): directory and event details
        """
        wb = xw.Book(event.src_path)

        # For each excel sheet
        for sheet in wb.sheets:
            wb_temp = xw.Book(MASTER_FILE)

            # Each processed sheet will be copied to the master excel.
            sheet.copy(after=wb_temp.sheets[0])

            wb_temp.save()
            wb_temp.close()

        # Close the processed excel file
        if len(wb.app.books) == 1:
            wb.app.quit()
        else:
            wb.close()

    def __move_file(self, event):
        """Method to move files to specific folders.
        0 If the file worked on is not valid,
        1 If the file worked on is valid.

        Args:
            event (dict): directory and event details
        """
        try:
            if self.status == 0:
                res = shutil.move(event.src_path, "./workarea/Not Applicable/")
                print("Moved To --> Not Applicable", res)

            if self.status == 1:
                res = shutil.move(event.src_path, "./workarea/Processed/")
                print("Moved To --> Processed", res)

        except Exception as e:
            print(e)


def folder_observer(path):
    """Main function to monitor X folder for .xlsx files to be consolidated into a master file.

    Args:
        path (str): _description_
    """
    # Validaciones iniciales
    if not Path(MASTER_FILE).is_file():
        print(f"The field {MASTER_FILE} does not exist, do you want to create it?")
        res = input("(y/n): ")

        if res.lower() == 'y':
            pd.DataFrame([]).to_excel(MASTER_FILE, index=False)
            print(f"File {MASTER_FILE} created.\n")
        else:
            return

    # Inicializacion de observador
    observer = Observer()
    event_handler = MyMonitor()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()

    # Observador
    try:
        print('Detecting changes... (Press CTRL + C to stop the program)')
        while True:
            time.sleep(2)

    except KeyboardInterrupt:
        observer.stop()

    observer.join()
