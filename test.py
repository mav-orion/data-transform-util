from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox

root = Tk()
root.title('Maveric Transformation Validator')
# root.geometry('850x650')
root.resizable(0, 0)



# SOURCE FILE DIRECTORY
sourceFileFrame = LabelFrame(root, text='Source File Directory', padx=20, pady=20)
sourceFileFrame.grid(row=0, column=0, padx=50, pady=20)

# DATE FORMAT FRAME
dateFormatFrame = LabelFrame(root, text='Date Format Standardization', padx=20, pady=20)
dateFormatFrame.grid(row=1, column=0, padx=50, pady=20)

# SPLIT COLUMN FRAME
splitColumnFrame = LabelFrame(root, text='Split Columns', padx=20, pady=20)
splitColumnFrame.grid(row=2, column=0, padx=50, pady=20)

# MERGE COLUMN FRAME
mergeColumnFrame = LabelFrame(root, text='Merge Columns', padx=20, pady=20)
mergeColumnFrame.grid(row=3, column=0, padx=50, pady=20)

# HARD CODE VALUE FRAME
hcodeColumnFrame = LabelFrame(root, text='Hard Code Values', padx=20, pady=20)
hcodeColumnFrame.grid(row=4, column=0, padx=50, pady=20)

# SOURCE FILE DIRECTORY
Label(sourceFileFrame, text='Source File').grid(row=0, column=0)
srcFile = Entry(sourceFileFrame, width=50)
srcFile.grid(row=0, column=1, padx=5, pady=5)


def browseFiles():
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                          filetypes=(("Excel files", "*.xlsx*"), ("all files", "*.*")))
    srcFile.insert(0, filename)


Button(sourceFileFrame, text='Browse', command=browseFiles).grid(row=0, column=2, padx=20, pady=20, sticky='')

# DATE FORMAT STANDARDIZATION
Label(dateFormatFrame, text="Select Date Format").grid(row=0, column=0, padx=10, pady=25)
n = StringVar()
formatChoosen = Combobox(dateFormatFrame, width=27, textvariable=n)
formatChoosen['values'] = (
    'dd-mmm-yy', 'dd-mmm-yyyy', 'mm/dd/yy', 'mm/dd/yyyy', 'dd.mm.yy', 'dd.mm.yyyy', 'yyddd', 'yyyyddd', 'yy/mm/dd',
    'yyyy/mm/dd', 'mmm yy', 'mmm yyyy', 'dd-mmm-yyyy hh:mm:ss.s')
formatChoosen.grid(row=0, column=1)
formatChoosen.current()

# Source and Target columns
Label(dateFormatFrame, text='Source Column').grid(row=1, column=0)
Label(dateFormatFrame, text='Target Column').grid(row=1, column=2)
# Source column input
Entry(dateFormatFrame).grid(row=1, column=1, padx=5, pady=5)
# Target column input
Entry(dateFormatFrame).grid(row=1, column=3, padx=5, pady=5)

# SPLIT COLUMNS
Label(splitColumnFrame, text='Source Column').grid(row=0, column=0)
Label(splitColumnFrame, text='Target Column 1').grid(row=1, column=0)
Label(splitColumnFrame, text='Target Column 2').grid(row=1, column=2)
Label(splitColumnFrame, text='Target Column 3').grid(row=1, column=4)
# Source column input
Entry(splitColumnFrame).grid(row=0, column=1, padx=5, pady=5)
# Target column 1 input
Entry(splitColumnFrame).grid(row=1, column=1, padx=5, pady=5)
# Target column 2 input
Entry(splitColumnFrame).grid(row=1, column=3, padx=5, pady=5)
# Target column 3 input
Entry(splitColumnFrame).grid(row=1, column=5, padx=5, pady=5)

# MERGE COLUMNS
Label(mergeColumnFrame, text='Source Column 1').grid(row=0, column=0)
Label(mergeColumnFrame, text='Source Column 2').grid(row=0, column=2)
Label(mergeColumnFrame, text='Target Column').grid(row=1, column=0)
# Target column 1 input
Entry(mergeColumnFrame).grid(row=0, column=1, padx=5, pady=5)
# Target column 2 input
Entry(mergeColumnFrame).grid(row=0, column=3, padx=5, pady=5)
# Source column input
Entry(mergeColumnFrame).grid(row=1, column=1, padx=5, pady=5)

# HARD CODE VALUES
Label(hcodeColumnFrame, text="Select Category").grid(row=0, column=0, padx=10, pady=25)
n = StringVar()
formatChoosen = Combobox(hcodeColumnFrame, width=27, textvariable=n)
formatChoosen['values'] = (
    'Countries', 'Others')
formatChoosen.grid(row=0, column=1)
formatChoosen.current()

# Source and Target columns
Label(hcodeColumnFrame, text='Source Column').grid(row=1, column=0)
Label(hcodeColumnFrame, text='Target Column').grid(row=1, column=2)
# Source column input
Entry(hcodeColumnFrame).grid(row=1, column=1, padx=5, pady=5)
# Target column input
Entry(hcodeColumnFrame).grid(row=1, column=3, padx=5, pady=5)

# Button to run transformation
Button(root, text='Run Transform').grid(padx=20, pady=20, sticky='')

root.mainloop()
