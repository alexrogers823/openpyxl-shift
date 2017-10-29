import openpyxl
from openpyxl.cell import get_column_letter

# This class works with openpyxl to add/delete rows and columns in Excel
# Openpyxl didn't have this feature to my knowledge, so I coded a way to do it
class Shift():
    def __init__(self, workbook, max_column, max_row, change_row_col):
        self.wb = workbook
        self.y = max_column
        self.z = max_row
        self.x = change_row_col

    def down(self): #Inserts row
        k = self.z
        for i in range(self.x, self.z+1):
            for j in range(1, self.y + 1):
                if self.wb.cell(row=k, column=j).value == None:
                    self.wb[get_column_letter(j)+str(k+1)] = ''
                else:
                    shift = self.wb[get_column_letter(j)+str(k)].value
                    self.wb[get_column_letter(j)+str(k+1)] = str(shift)
            k -= 1

    def right(self): #Inserts column
        k = self.y
        for i in range(self.x, self.y+1):
            for j in range(1, self.z+1):
                if self.wb.cell(row=j, column=self.y).value == None:
                    self.wb[get_column_letter(self.y+1)+str(j)] = ''
                else:
                    shift = self.wb.cell(row=j, column=self.y).value
                    self.wb[get_column_letter(self.y+1)+str(j)] = str(shift)
            k -= 1

    def up(self): #Deletes row
        for i in range(self.x, self.z+1):
            for j in range(1, self.y+1):
                if self.wb.cell(row=i+1, column=j).value == None:
                    self.wb[get_column_letter(j)+str(i)] = ''
                else:
                    shift = self.wb.cell(row=i+1, column=j).value
                    self.wb[get_column_letter(j)+str(i)] = str(shift)


    def left(self): #Deletes column
        for i in range(self.x, self.z+1):
            for j in range(1, self.y+1):
                if self.wb.cell(row=j, column=i+1).value == None:
                    self.wb[get_column_letter(i)+str(j)] = ''
                else:
                    shift = self.wb.cell(row=j, column=i+1).value
                    self.wb[get_column_letter(i)+str(j)] = str(shift)
