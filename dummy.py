import warnings
from pathlib import Path
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
import argparse
import pandas as pd
import numpy as np

fileSource2 = None
sourceData2 = None
targetData2 = None
targetPath2 = None
excelLabel3 = None
exportLabel3 = None
initiateLabel2 = None
columnHeaders2 = None
sheetNameList2 = None
sheetNameValues2 = None
sheetNamesDropDown2 = None
dtsrc2Value2 = None
dtsrc2 = None
trgFile2 = None
skipRow2 = None
diff2 = None
expFile2 = None
warnings.simplefilter(action='ignore', category=FutureWarning)


def getsheetnames(srcpath):
    global excelLabel3
    global sheetNamesDropDown2
    global fileSource2
    global sheetNameList2
    global sheetNameValues2
    fileSource2 = pd.ExcelFile(srcpath)
    sheetNameList2 = fileSource2.sheet_names
    Label(sourceFileFrame, text='Select Sheet').grid(row=1, column=0, padx=20, pady=20, sticky='')
    sheetNamesDropDown2 = Combobox(sourceFileFrame, width=27)
    sheetNamesDropDown2.grid(row=1, column=1, padx=5, pady=5)
    sheetNamesDropDown2['values'] = sheetNameList2
    sheetNamesDropDown2.current()

    def sheetNameEnter(event):
        global sheetNameValues2
        sheetNameValues2 = event.widget.get()

    sheetNamesDropDown2.bind('<<ComboboxSelected>>', sheetNameEnter)
    Button(sourceFileFrame, text='Initiate', command=lambda: initiate(srcFile.get())).grid(row=1, column=2,
                                                                                           padx=20,
                                                                                           pady=20,
                                                                                           sticky='')
    excelLabel3 = Label(sourceFileFrame, text="Imported", bg='green', fg='white')
    excelLabel3.grid(row=0, column=3, padx=5, pady=5, sticky='')


def initiate(srcpath):
    global initiateLabel2
    global targetData2
    global targetPath2
    global excelLabel3
    global columnHeaders2
    global sourceData2

    sourceData2 = pd.read_excel(srcpath, sheet_name=sheetNameValues2, dtype='object')  # sheet name - 'Sheet1'
    targetData2 = pd.DataFrame(sourceData2)
    columnHeaders2 = list(targetData2.columns)
    print('Excel imported')
    initiateLabel2 = Label(sourceFileFrame, text="Initiated", bg='green', fg='white')
    initiateLabel2.grid(row=1, column=2, padx=5, pady=5, sticky='')
    target()
    # keyColumn()
    exportExcel()
    reconcile()


#######################################################################################################################
root = Tk()

root.title('Maveric Transformation Validator')
root.geometry('850x650')
# root.resizable(0, 0)

mainFrame = Frame(root)
mainFrame.pack(fill=BOTH, expand=1)

mainCanvas = Canvas(mainFrame)
mainCanvas.pack(side=LEFT, fill=BOTH, expand=1)

mainScrollBar = Scrollbar(mainFrame, orient=VERTICAL, command=mainCanvas.yview)
mainScrollBar.pack(side=RIGHT, fill=Y)

innerFrame = Frame(mainCanvas, padx=20, pady=20)
innerFrame.grid(row=0, column=0, padx=50, pady=20)
innerFrame.bind("<Configure>", lambda e: mainCanvas.configure(scrollregion=mainCanvas.bbox('all')))
mainCanvas.configure(yscrollcommand=mainScrollBar.set)
mainCanvas.bind('<Configure>', lambda e: mainCanvas.configure(scrollregion=mainCanvas.bbox('all')))
mainCanvas.bind_all('<MouseWheel>', lambda e: mainCanvas.yview_scroll(int(-1 * (e.delta / 120)), 'units'))
mainCanvas.create_window((0, 0), window=innerFrame, anchor='nw')

#######################################################################################################################

# SOURCE FILE DIRECTORY
sourceFileFrame = LabelFrame(innerFrame, text='Source File Directory', padx=20, pady=20)
sourceFileFrame.grid(row=0, pady=20, sticky='')

# TARGET FILE DIRECTORY
targetFileFrame = LabelFrame(innerFrame, text='Target File Directory', padx=20, pady=20)
targetFileFrame.grid(row=1, pady=20, sticky='')

# INPUT FRAME
inputFrame = LabelFrame(innerFrame, text='Options', padx=20, pady=20)
inputFrame.grid(row=2, pady=20, sticky='')

# EXPORT FRAME
exportFrame = LabelFrame(innerFrame, padx=20, pady=20)
exportFrame.grid(row=3, pady=20, sticky='')

# OUTPUT FRAME
outputFrame = LabelFrame(innerFrame, text='Comparison', padx=20, pady=20)
outputFrame.grid(row=4, pady=20, sticky='')

# SOURCE FILE DIRECTORY
Label(sourceFileFrame, text='Source File Path').grid(row=0, column=0)
srcFile = Entry(sourceFileFrame, width=50)
srcFile.grid(row=0, column=1, padx=5, pady=5)


def browseSrcFiles():
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                          filetypes=(("Excel files", "*.xlsx*"), ("all files", "*.*")))
    srcFile.insert(0, filename)


Button(sourceFileFrame, text='Browse', command=browseSrcFiles).grid(row=0, column=2, padx=20, pady=20, sticky='')
Button(sourceFileFrame, text='Import',
       command=lambda: getsheetnames(srcFile.get())).grid(row=0, column=3, padx=20,

                                                          pady=20)


# TARGET FILE DIRECTORY

def target():
    global trgFile2
    Label(targetFileFrame, text='Target File Path').grid(row=0, column=0)
    trgFile2 = Entry(targetFileFrame, width=50)
    trgFile2.grid(row=0, column=1, padx=5, pady=5)

    def browseTrgFiles():
        filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                              filetypes=(("Excel files", "*.xlsx*"), ("all files", "*.*")))
        trgFile2.insert(0, filename)

    Button(targetFileFrame, text='Browse', command=browseTrgFiles).grid(row=0, column=2, padx=20, pady=20, sticky='')

def exportExcel():
    global expFile2
    Label(inputFrame, text='Export File Path').grid(row=0, column=0)
    expFile2 = Entry(inputFrame, width=50)
    expFile2.grid(row=0, column=1, padx=5, pady=5)

    def exportDiff():
        filedirectory = filedialog.askdirectory()
        expFile2.insert(0, filedirectory + '/Excel_Diff.xlsx')

    Button(inputFrame, text='Browse', command=exportDiff).grid(row=0, column=2, padx=20, pady=20, sticky='')

# Key Index

def keyColumn():
    global dtsrc2
    global skipRow2
    Label(inputFrame, text='Name of the column with unique row identifier').grid(row=0, column=0)
    dtsrc2 = Combobox(inputFrame, width=27)
    dtsrc2['values'] = columnHeaders2
    dtsrc2.grid(row=0, column=1, padx=5, pady=5)
    dtsrc2.current()

    def dtsrc2Enter(event):
        global dtsrc2Value2
        dtsrc2Value2 = event.widget.get()

    dtsrc2.bind('<<ComboboxSelected>>', dtsrc2Enter)
    Label(inputFrame, text='Skip Rows').grid(row=1, column=0)
    skipRow2 = Entry(inputFrame, width=27)
    skipRow2.grid(row=1, column=1, padx=5, pady=5)

def exportToExcel():
    global exportLabel3
    targetData2.to_excel(expFile2, index=False)
    print('Export to Excel Complete')
    exportLabel3 = Label(exportFrame, text="Export Success", bg='green', fg='white')
    exportLabel3.grid(row=1, column=2, padx=10, sticky='')

def report_diff2(x):
    """Function to use with groupby.apply to highlihgt value changes."""
    return x[0] if x[0] == x[1] or pd.isna(x).all() else f'{x[0]} ---> {x[1]}'


def strip(x):
    """Function to use with applymap to strip whitespaces in a dataframe."""
    return x.strip() if isinstance(x, str) else x


def diff2_pd(old_df, new_df, idx_col):
    """Identify diff2erences between two pandas DataFrames using a key column.
    Key column is assumed to have only unique data
    (like a database unique id index)
    Args:
        old_df (pd.DataFrame): first dataframe
        new_df (pd.DataFrame): second dataframe
        idx_col (str|list(str)): column name(s) of the index,
          needs to be present in both DataFrames
    """
    # setting the column name as index for fast operations
    old_df = old_df.set_index(idx_col)
    new_df = new_df.set_index(idx_col)
    # get the added and removed rows
    old_keys = old_df.index
    new_keys = new_df.index
    if isinstance(old_keys, pd.MultiIndex):
        removed_keys = old_keys.diff2erence(new_keys)
        added_keys = new_keys.diff2erence(old_keys)
    else:
        removed_keys = np.setdiff21d(old_keys, new_keys)
        added_keys = np.setdiff21d(new_keys, old_keys)
    out_data = {
        'removed': old_df.loc[removed_keys],
        'added': new_df.loc[added_keys]
    }
    # focusing on common data of both dataframes
    common_keys = np.intersect1d(old_keys, new_keys, assume_unique=True)
    common_columns = np.intersect1d(
        old_df.columns, new_df.columns, assume_unique=True
    )
    new_common = new_df.loc[common_keys, common_columns].applymap(strip)
    old_common = old_df.loc[common_keys, common_columns].applymap(strip)
    # get the changed rows keys by dropping identical rows
    # (indexes are ignored, so we'll reset them)
    common_data = pd.concat(
        [old_common.reset_index(), new_common.reset_index()], sort=True
    )
    changed_keys = common_data.drop_duplicates(keep=False)[idx_col]
    if isinstance(changed_keys, pd.Series):
        changed_keys = changed_keys.unique()
    else:
        changed_keys = changed_keys.drop_duplicates().set_index(idx_col).index
    # combining the changed rows via multi level columns
    df_all_changes = pd.concat(
        [old_common.loc[changed_keys], new_common.loc[changed_keys]],
        axis='columns',
        keys=['old', 'new']
    ).swaplevel(axis='columns')
    # using report_diff2 to merge the changes in a single cell with "-->"
    df_changed = df_all_changes.groupby(level=0, axis=1).apply(
        lambda frame: frame.apply(report_diff2, axis=1))
    out_data['changed'] = df_changed

    return out_data


def comparison(path1, path2, sheetname, exp_path):
    global diff2
    df1 = pd.read_excel(path1, sheet_name=sheetname)
    df2 = pd.read_excel(path2, sheet_name=sheetname)
    df1.equals(df2)
    comparison_values = df1.values == df2.values
    print(comparison_values)
    rows, cols = np.where(comparison_values == False)
    for item in zip(rows, cols):
        df1.iloc[item[0], item[1]] = '{} --> {}'.format(df1.iloc[item[0], item[1]], df2.iloc[item[0], item[1]])
    df1.to_excel(exp_path, index=False, header=True)


def compare_excel(
        path1, path2, sheet_name, index_col_name, **kwargs
):
    global diff2
    old_df = pd.read_excel(path1, sheet_name=sheet_name, **kwargs)
    new_df = pd.read_excel(path2, sheet_name=sheet_name, **kwargs)
    diff2 = diff2_pd(old_df, new_df, index_col_name)
    # with pd.ExcelWriter(out_path) as writer:
    #     for sname, data in diff2.items():
    #         data.to_excel(writer, sheet_name=sname)
    # print(f"Differences saved in {out_path}")


def build_parser():
    cfg = argparse.ArgumentParser(
        description="Compares two excel files and outputs the diff2erences "
                    "in another excel file."
    )
    cfg.add_argument("path1", help="Fist excel file")
    cfg.add_argument("path2", help="Second excel file")
    cfg.add_argument("sheetname", help="Name of the sheet to compare.")
    cfg.add_argument(
        "key_column",
        help="Name of the column(s) with unique row identifier. It has to be "
             "the actual text of the first row, not the excel notation."
             "Use multiple times to create a composite index.",
        nargs="+",
    )
    cfg.add_argument("-o", "--output-path", default="compared.xlsx",
                     help="Path of the comparison results")
    cfg.add_argument("--skiprows", help='number of rows to skip', type=int,
                     action='append', default=None)
    return cfg


# def main():
#     cfg = build_parser()
#     opt = cfg.parse_args()
#     compare_excel(opt.path1, opt.path2, opt.output_path, opt.sheetname,
#                   opt.key_column, skiprows=opt.skiprows)
#
#
# if __name__ == '__main__':
#     main()
def compare():
    comparison(srcFile.get(), trgFile2.get(), sheetNameValues2,expFile2.get())
    # T = Text(outputFrame, width=52).grid(row=0, column=0, padx=10, pady=10)
    # T.insert('END', diff2)


def reconcile():
    Button(exportFrame, text='Reconcile', command=compare).grid(row=0, column=0, padx=20, pady=10)


root.mainloop()
