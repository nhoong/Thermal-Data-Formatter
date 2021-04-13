import tkinter as tk
from tkinter import filedialog
import pandas as pd

#init TK Window
root= tk.Tk()
root.title('CSV Thermal Data Formatter')

#create Window
canvas1 = tk.Canvas(root, width = 300, height = 400, bg = 'White', relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, bg= 'White', fg= 'Black', text= 'Conditional Fomatting for \nThermal Data')
canvas1.create_window(150, 50, window=label1)

#main function to get CSV/xlsx file
def getCSV():
    global df

    #import file
    import_file_path= filedialog.askopenfilename()
    if import_file_path.endswith('.csv'):
        df = pd.read_csv(import_file_path)
    if import_file_path.endswith('.xlsx'):
        df = pd.read_excel(import_file_path)    

#apply conditional formatting
def applyFormat():
    sheet_name = 'Sheet 1'

    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    writer = pd.ExcelWriter(export_file_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=sheet_name, index = False, header = False)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    worksheet.conditional_format('A1:FD120', {'type': '3_color_scale', 'min_color': "#60c460", 'mid_color': "#fffd94", 'max_color': "#f896b"})
    writer.save()

browseButton_CSV = tk.Button(text= 'Import CSV File', command= getCSV, bg= 'LightGrey', fg= 'Black', font= ('helevtica', 12, 'bold'))
applyFormat = tk.Button(text= 'Apply Format', command= applyFormat, bg= 'LightGrey', fg= 'Black', font= ('helevetica', 12, 'bold'))
exitButton = tk.Button(root, text= 'Exit', command= root.destroy, bg= 'LightGrey', fg= 'Black', font= ('Arial', 12, 'bold'))

canvas1.create_window(150, 150, window=browseButton_CSV)
canvas1.create_window(150, 200, window=applyFormat)
canvas1.create_window(150, 300, window=exitButton)

root.mainloop()