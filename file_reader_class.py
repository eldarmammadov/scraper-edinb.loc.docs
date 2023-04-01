import sys
import os
import pandas as pd
import re


class CustomFileReader:
    def __init__(self, filename='./output.xlsx'):
        self._filename = filename
        self._path = None
        self._df = None
        self._pattern = None

    @property
    def filename(self):
        return self._filename

    @property
    def path(self):
        if self._path is None:
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                # the following code looks highly dubious
                path_actual = os.getcwd()
                path_main_folder = path_actual[:-4]
                self._path = path_main_folder + self.filename
                print('frozen path', os.path.normpath(self._path))
            else:
                self._path = self.filename
        return self._path
