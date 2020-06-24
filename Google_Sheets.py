import gspread
from oauth2client.service_account import ServiceAccountCredentials

file_path = r"D:\project\BD_ODBC_laba"


class Google_Sheets:

    def __init__(self, name_Google_Sheet):
        self.name_Sheet = name_Google_Sheet
        self.Sheet = self.authorization()

    def authorization(self):
        # this method is used to connect to the table
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(r'googleSheet.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open(self.name_Sheet).sheet1
        return sheet

    def __get_col_index(self, str_name):
        # method finds the column index by name its name
        if str_name in self.Sheet.row_values(1):
            return self.Sheet.row_values(1).index(str_name) + 1
        else:
            #  print("column not found")
            return -1

    def get_data(self, str_col):
        # the method returns a list of values of a str_col column
        index = self.__get_col_index(str_col)
        if index != -1:
            lst = self.Sheet.col_values(index)
            lst.remove(str_col)
            return lst
        else:
            return []

    def add_data(self, str_col, lst):
        # the method is used to add lst in str_col column
        # (the data is presented in the form of a list and is pushed by the column to the str_column)
        index = self.__get_col_index(str_col)
        if index == -1:
            self.add_col(str_col)
        index = len(self.Sheet.row_values(1))
        cell_range = chr(64 + index) + "2" + ":" + chr(64 + index) + str(len(lst) + 1)
        cell_list = self.Sheet.range(cell_range)  # pushing range
        for i in range(0, len(cell_list)):
            cell_list[i].value = lst[i]
        self.Sheet.update_cells(cell_list)
        self.res_sum()

    def res_sum(self):
        # the method calculates the sum of values
        for i in range(7, len(self.Sheet.row_values(1)) + 1):
            sum_range = chr(64 + i) + "2:" + chr(64 + i) + str(len(self.Sheet.col_values(1)) - 1)
            sum = "=СУММ(" + sum_range + " )"
            self.Sheet.update_cell(len(self.Sheet.col_values(1)), i, sum)

    def get_dates(self):
        # method returns a list of dates
        buf_lst = self.Sheet.row_values(1)
        buf = []
        for i in range(6, len(buf_lst)):
            buf.append(buf_lst[i])
        return buf

    def clear(self):
        self.Sheet.clear()

    def add_col(self, col):
        lst_buf = self.Sheet.row_values(1)
        for i in lst_buf:
            if str(i) == str(col):
                return

        self.Sheet.update_cell(1, len(lst_buf) + 1, col)

    def add_start(self, list_start):
        self.clear()
        self.Sheet.append_row(list_start)

    def test(self, lst):
        self.Sheet.append_row(lst)

    def add_info_bd(self, list_inf):
        index = len(self.Sheet.row_values(1))
        cell_range = "A2" + ":" + chr(64 + index) + str(len(list_inf) + 1)
        self.Sheet.append_rows(list_inf, table_range = cell_range)

