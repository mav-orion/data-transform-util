import warnings
import pandas as dw
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
import argparse
import numpy as np

mainFrame = None
innerFrame = None
sourceFileFrame = None
targetFileFrame = None
inputFrame = None
exportFrame = None
outputFrame = None
sourceFileFrame2 = None
dateFormatFrame = None
splitColumnFrame = None
mergeColumnFrame = None
hcodeColumnFrame = None
exportFrame = None
messageFrame = None
fileSource = None
sourceData = None
targetData = None
targetPath = None
excelLabel = None
dateLabel = None
splitLabel = None
mergeLabel = None
exportLabel = None
initiateLabel = None
harcodeLabel = None
transformLabel = None
columnHeaders = None
sheetNameList = None
sheetNameValues = None
sheetNamesDropDown = None
dttarget = None
formatChoosen = None
splitSrcValue = None
splitTrgt1 = None
splitTrgt2 = None
splitTrgt3 = None
mergeTrgt = None
mergeSrcValue1 = None
mergeSrcValue2 = None
cntryTrgt = None
comboValue = None
dtsrc = None
splitSrc = None
mergeSrc1 = None
mergeSrc2 = None
chooseCountryFormatFrom = None
chooseCountryFormatTo = None
hcodeSource = None
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
dtsrc999992Value2 = None
dtsrc2 = None
trgFile2 = None
skipRow2 = None
diff2 = None
expFile2 = None
warnings.simplefilter(action='ignore', category=FutureWarning)


def TRANSFORM(root):
    global mainFrame
    global innerFrame
    global sourceFileFrame
    global sourceFileFrame2
    global dateFormatFrame
    global splitColumnFrame
    global mergeColumnFrame
    global hcodeColumnFrame
    global exportFrame
    global messageFrame
    global targetFileFrame
    global inputFrame
    global outputFrame

    try:
        sourceFileFrame.destroy()
    except:
        print('sourceFileFrame does not exist')
    try:
        dateFormatFrame.destroy()
    except:
        print('dateFormatFrame does not exist')
    try:
        splitColumnFrame.destroy()
    except:
        print('splitColumnFrame does not exist')
    try:
        mergeColumnFrame.destroy()
    except:
        print('mergeColumnFrame does not exist')
    try:
        hcodeColumnFrame.destroy()
    except:
        print('hcodeColumnFrame does not exist')
    try:
        exportFrame.destroy()
    except:
        print('exportFrame does not exist')
    try:
        messageFrame.destroy()
    except:
        print('messageFrame does not exist')
    try:
        mainFrame.destroy()
    except:
        print('mainFrame does not exist')
    try:
        innerFrame.destroy()
    except:
        print('innerFrame does not exist')
    try:
        sourceFileFrame2.destroy()
    except:
        print('sourceFileFrame2 does not exist')
    try:
        targetFileFrame.destroy()
    except:
        print('targetFileFrame does not exist')
    try:
        inputFrame.destroy()
    except:
        print('inputFrame does not exist')
    try:
        exportFrame.destroy()
    except:
        print('exportFrame does not exist')
    try:
        outputFrame.destroy()
    except:
        print('outputFrame does not exist')

    def getsheetnames(srcpath):
        global excelLabel
        global sheetNamesDropDown
        global fileSource
        global sheetNameList
        global sheetNameValues
        fileSource = dw.ExcelFile(srcpath)
        sheetNameList = fileSource.sheet_names
        Label(sourceFileFrame, text='Select Sheet').grid(row=1, column=0, padx=20, pady=20, sticky='')
        sheetNamesDropDown = Combobox(sourceFileFrame, width=27)
        sheetNamesDropDown.grid(row=1, column=1, padx=5, pady=5)
        sheetNamesDropDown['values'] = sheetNameList
        sheetNamesDropDown.current()

        def sheetNameEnter(event):
            global sheetNameValues
            sheetNameValues = event.widget.get()

        sheetNamesDropDown.bind('<<ComboboxSelected>>', sheetNameEnter)
        Button(sourceFileFrame, text='Initiate', command=lambda: initiate(srcFile.get())).grid(row=1, column=2,
                                                                                               padx=20,
                                                                                               pady=20,
                                                                                               sticky='')
        excelLabel = Label(sourceFileFrame, text="Imported", bg='green', fg='white')
        excelLabel.grid(row=0, column=3, padx=5, pady=5, sticky='')

    def initiate(srcpath):
        global initiateLabel
        global targetData
        global targetPath
        global excelLabel
        global columnHeaders
        global sourceData

        sourceData = dw.read_excel(srcpath, sheet_name=sheetNameValues, dtype='object')  # sheet name - 'Sheet1'
        targetData = dw.DataFrame(sourceData)
        columnHeaders = list(targetData.columns)
        print('Excel imported')
        initiateLabel = Label(sourceFileFrame, text="Initiated", bg='green', fg='white')
        initiateLabel.grid(row=1, column=2, padx=5, pady=5, sticky='')
        STANDARD_DATE()
        SPLIT_SINGLE_COLUMN()
        MERGE_COLUMNS()
        HARDCODE_VALUES()
        BUTTONS_EXECUTE()

    #######################################################################################################################

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

    # DATE FORMAT FRAME
    dateFormatFrame = LabelFrame(innerFrame, text='Date Format Standardization', padx=20, pady=20)
    dateFormatFrame.grid(row=1, pady=20, sticky='')

    # SPLIT COLUMN FRAME
    splitColumnFrame = LabelFrame(innerFrame, text='Split Columns', padx=20, pady=20)
    splitColumnFrame.grid(row=2, pady=20, sticky='')

    # MERGE COLUMN FRAME
    mergeColumnFrame = LabelFrame(innerFrame, text='Merge Columns', padx=20, pady=20)
    mergeColumnFrame.grid(row=3, pady=20, sticky='')

    # HARD CODE VALUE FRAME
    hcodeColumnFrame = LabelFrame(innerFrame, text='Country Hard Code', padx=20, pady=20)
    hcodeColumnFrame.grid(row=4, pady=10, sticky='')

    # EXPORT FRAME
    exportFrame = LabelFrame(innerFrame, padx=20, pady=20)
    exportFrame.grid(row=5, pady=20, sticky='')

    # Empty FRAME
    messageFrame = Frame(innerFrame, padx=20, pady=20)
    messageFrame.grid(row=6, pady=20, sticky='')

    # SOURCE FILE DIRECTORY
    Label(sourceFileFrame, text='Source File Path').grid(row=0, column=0)
    srcFile = Entry(sourceFileFrame, width=50)
    srcFile.grid(row=0, column=1, padx=5, pady=5)

    def browseSrcFiles():
        filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                              filetypes=(("Excel files", "*.xlsx*"), ("all files", "*.*")))
        srcFile.insert(0, filename)

    Button(sourceFileFrame, text='Browse', command=browseSrcFiles).grid(row=0, column=2, padx=20, pady=20, sticky='')

    # Label(sourceFileFrame, text='Select Sheet').grid(row=1, column=0, padx=20, pady=20, sticky='')
    # sheetName = Entry(sourceFileFrame)
    # sheetName.grid(row=0, column=4, padx=5, pady=5)

    # TARGET FILE DIRECTORY
    # Label(sourceFileFrame, text='Target File Path').grid(row=2, column=0)
    # trgtFile = Entry(sourceFileFrame, width=50)
    # trgtFile.grid(row=2, column=1, padx=5, pady=5)
    #
    #
    # def browseTargetFiles():
    #     filedirectory = filedialog.askdirectory()
    #     trgtFile.insert(0, filedirectory)
    #
    #
    # Button(sourceFileFrame, text='Browse', command=browseTargetFiles).grid(row=2, column=2, padx=20, pady=20, sticky='')
    Button(sourceFileFrame, text='Import',
           command=lambda: getsheetnames(srcFile.get())).grid(row=0, column=3, padx=20,
                                                              pady=20)

    #######################################################################################################################
    # PANDAS DATE FORMAT
    def date_format_standardization(sourceColumn, targetColumn, formatdate):
        # global dateLabel
        targetData[targetColumn] = dw.to_datetime(sourceData[sourceColumn], errors='coerce')
        if formatdate == 'dd/mm/yyyy':
            targetData[targetColumn] = targetData[targetColumn].dt.strftime('%d/%m/%Y')  # dd/mm/yyyy
        elif formatdate == 'mm/dd/yyyy':
            targetData[targetColumn] = targetData[targetColumn].dt.strftime('%m/%d/%Y')
        elif formatdate == 'dd/mm/yy':
            targetData[targetColumn] = targetData[targetColumn].dt.strftime('%d/%m/%y')
        elif formatdate == 'mm/dd/yy':
            targetData[targetColumn] = targetData[targetColumn].dt.strftime('%m/%d/%y')
        else:
            targetData[targetColumn] = targetData[targetColumn].dt.strftime('%d-%m-%Y')  # dd-mm-yyyy
        if targetColumn != sourceColumn:
            targetData.drop([sourceColumn], axis=1, inplace=True)
        print('Date Format Standardization Complete')
        # dateLabel = Label(dateFormatFrame, text="Transform Complete", bg='green', fg='white')
        # dateLabel.grid(row=2, padx=5, pady=5, sticky='')

    def STANDARD_DATE():
        # DATE FORMAT STANDARDIZATION
        global dtsrc
        global dttarget
        global formatChoosen
        Label(dateFormatFrame, text="Select Date Format").grid(row=0, column=0, padx=10, pady=25)
        n = StringVar()
        formatChoosen = Combobox(dateFormatFrame, width=27, textvariable=n)
        formatChoosen['values'] = ('dd/mm/yyyy', 'mm/dd/yyyy', 'dd/mm/yy', 'mm/dd/yy')
        formatChoosen.grid(row=0, column=1)
        formatChoosen.current()

        def enter(event):
            global comboValue
            comboValue = event.widget.get()

        formatChoosen.bind('<<ComboboxSelected>>', enter)

        # Source and Target columns
        Label(dateFormatFrame, text='Source Column').grid(row=1, column=0)
        Label(dateFormatFrame, text='Target Column').grid(row=1, column=2)
        # Source column input
        # dtsrc = Entry(dateFormatFrame)
        # dtsrc.grid(row=1, column=1, padx=5, pady=5)
        dtsrcValue = None
        dtsrc = Combobox(dateFormatFrame, width=27)
        dtsrc['values'] = columnHeaders
        dtsrc.grid(row=1, column=1, padx=5, pady=5)
        dtsrc.current()

        def dtsrcEnter(event):
            global dtsrcValue
            dtsrcValue = event.widget.get()

        dtsrc.bind('<<ComboboxSelected>>', dtsrcEnter)

        # Target column input
        dttarget = Entry(dateFormatFrame)
        dttarget.grid(row=1, column=3, padx=5, pady=5)

        # Button(dateFormatFrame, text='Transform',
        #        command=lambda: date_format_standardization(dtsrc.get(), dttarget.get(), comboValue)).grid(row=2, column=0,
        #                                                                                                   padx=20,
        #                                                                                                   pady=20)

    #######################################################################################################################
    def split_single_attribute(sourceColumn, targetColumn1, targetColumn2, targetColumn3):
        # global splitLabel
        new = targetData[sourceColumn].str.split("[ |; |, ||]", n=1, expand=True)
        targetData[targetColumn1] = new[0]
        targetData[targetColumn2] = new[1]
        new_1 = targetData[targetColumn2].str.split("[ |; |, ||]", n=1, expand=True)
        targetData[targetColumn2] = new_1[0]
        targetData[targetColumn3] = new_1[1]
        targetData.drop(columns=[sourceColumn], inplace=True)
        print('Column Split Complete')
        # splitLabel = Label(splitColumnFrame, text="Transform Complete", bg='green', fg='white')
        # splitLabel.grid(row=2, padx=5, pady=5, sticky='')

    def SPLIT_SINGLE_COLUMN():
        # SPLIT COLUMNS
        global splitSrc
        global splitTrgt1
        global splitTrgt2
        global splitTrgt3
        Label(splitColumnFrame, text='Source Column').grid(row=0, column=0)
        Label(splitColumnFrame, text='Target Column 1').grid(row=1, column=0)
        Label(splitColumnFrame, text='Target Column 2').grid(row=1, column=2)
        Label(splitColumnFrame, text='Target Column 3').grid(row=1, column=4)
        # Source column input
        # splitSrc = Entry(splitColumnFrame)
        # splitSrc.grid(row=0, column=1, padx=5, pady=5)
        splitSrc = Combobox(splitColumnFrame, width=27)
        splitSrc['values'] = columnHeaders
        splitSrc.grid(row=0, column=1, padx=5, pady=5)
        splitSrc.current()

        def dtsrcEnter(event):
            global splitSrcValue
            splitSrcValue = event.widget.get()

        splitSrc.bind('<<ComboboxSelected>>', dtsrcEnter)

        # Target column 1 input
        splitTrgt1 = Entry(splitColumnFrame)
        splitTrgt1.grid(row=1, column=1, padx=2, pady=2)
        # Target column 2 input
        splitTrgt2 = Entry(splitColumnFrame)
        splitTrgt2.grid(row=1, column=3, padx=2, pady=2)
        # Target column 3 input
        splitTrgt3 = Entry(splitColumnFrame)
        splitTrgt3.grid(row=1, column=5, padx=2, pady=2)

        # Button(splitColumnFrame, text='Transform',
        #        command=lambda: split_single_attribute(splitSrc.get(), splitTrgt1.get(), splitTrgt2.get(),
        #                                               splitTrgt3.get())).grid(row=2, column=0,
        #                                                                       padx=20,
        #                                                                       pady=20)

    #######################################################################################################################

    def merge_name_field(srcColumn1, srcColumn2, trgtColumn):
        # global mergeLabel
        targetData[trgtColumn] = (targetData[srcColumn1] + " " + targetData[srcColumn2])
        targetData.drop([srcColumn1], axis=1, inplace=True)
        targetData.drop([srcColumn2], axis=1, inplace=True)
        print('Column Merge Complete')
        # mergeLabel = Label(mergeColumnFrame, text="Transform Complete", bg='green', fg='white')
        # mergeLabel.grid(row=2, padx=5, pady=5, sticky='')

    def MERGE_COLUMNS():
        # MERGE COLUMNS
        global mergeSrc1
        global mergeSrc2
        global mergeTrgt
        Label(mergeColumnFrame, text='Source Column 1').grid(row=0, column=0)
        Label(mergeColumnFrame, text='Source Column 2').grid(row=0, column=2)
        Label(mergeColumnFrame, text='Target Column').grid(row=1, column=0)
        # Source column 1 input
        # mergeSrc1 = Entry(mergeColumnFrame)
        # mergeSrc1.grid(row=0, column=1, padx=5, pady=5)
        mergeSrc1 = Combobox(mergeColumnFrame, width=27)
        mergeSrc1['values'] = columnHeaders
        mergeSrc1.grid(row=0, column=1, padx=5, pady=5)
        mergeSrc1.current()

        def mergeSrcSelect1(event):
            global mergeSrcValue1
            mergeSrcValue1 = event.widget.get()

        mergeSrc1.bind('<<ComboboxSelected>>', mergeSrcSelect1)

        # Source column 2 input
        # mergeSrc2 = Entry(mergeColumnFrame)
        # mergeSrc2.grid(row=0, column=3, padx=5, pady=5)
        mergeSrc2 = Combobox(mergeColumnFrame, width=27)
        mergeSrc2['values'] = columnHeaders
        mergeSrc2.grid(row=0, column=3, padx=5, pady=5)
        mergeSrc2.current()

        def mergeSrcSelect2(event):
            global mergeSrcValue2
            mergeSrcValue2 = event.widget.get()

        mergeSrc2.bind('<<ComboboxSelected>>', mergeSrcSelect2)

        # Target column input
        mergeTrgt = Entry(mergeColumnFrame)
        mergeTrgt.grid(row=1, column=1, padx=5, pady=5)
        # Button(mergeColumnFrame, text='Transform',
        #        command=lambda: merge_name_field(mergeSrc1.get(), mergeSrc2.get(), mergeTrgt.get())).grid(row=2, column=0,
        #                                                                                                  padx=20,
        #                                                                                                  pady=20)

    #######################################################################################################################

    def decode_countries(sourceColumn, targetColumn, countryFormatFrom, countryFormatTo):
        # global harcodeLabel
        countries = ["Afghanistan", "Albania", "Algeria", "American Samoa", "Andorra", "Angola", "Anguilla",
                     "Antarctica",
                     "Antigua and Barbuda", "Argentina", "Armenia", "Aruba", "Australia", "Austria", "Azerbaijan",
                     "Bahamas (the)", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin",
                     "Bermuda", "Bhutan", "Bolivia (Plurinational State of)", "Bonaire, Sint Eustatius and Saba",
                     "Bosnia and Herzegovina", "Botswana", "Bouvet Island", "Brazil",
                     "British Indian Ocean Territory (the)", "Brunei Darussalam", "Bulgaria", "Burkina Faso", "Burundi",
                     "Cabo Verde", "Cambodia", "Cameroon", "Canada", "Cayman Islands (the)",
                     "Central African Republic (the)", "Chad", "Chile", "China", "Christmas Island",
                     "Cocos (Keeling) Islands (the)", "Colombia", "Comoros (the)",
                     "Congo (the Democratic Republic of the)",
                     "Congo (the)", "Cook Islands (the)", "Costa Rica", "Croatia", "Cuba", "Curaçao", "Cyprus",
                     "Czechia",
                     "Côte d'Ivoire", "Denmark", "Djibouti", "Dominica", "Dominican Republic (the)", "Ecuador", "Egypt",
                     "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia",
                     "Falkland Islands (the) [Malvinas]", "Faroe Islands (the)", "Fiji", "Finland", "France",
                     "French Guiana", "French Polynesia", "French Southern Territories (the)", "Gabon", "Gambia (the)",
                     "Georgia", "Germany", "Ghana", "Gibraltar", "Greece", "Greenland", "Grenada", "Guadeloupe", "Guam",
                     "Guatemala", "Guernsey", "Guinea", "Guinea-Bissau", "Guyana", "Haiti",
                     "Heard Island and McDonald Islands", "Holy See (the)", "Honduras", "Hong Kong", "Hungary",
                     "Iceland",
                     "India", "Indonesia", "Iran (Islamic Republic of)", "Iraq", "Ireland", "Isle of Man", "Israel",
                     "Italy", "Jamaica", "Japan", "Jersey", "Jordan", "Kazakhstan", "Kenya", "Kiribati",
                     "Korea (the Democratic People's Republic of)", "Korea (the Republic of)", "Kuwait", "Kyrgyzstan",
                     "Lao People's Democratic Republic (the)", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya",
                     "Liechtenstein", "Lithuania", "Luxembourg", "Macao", "Madagascar", "Malawi", "Malaysia",
                     "Maldives",
                     "Mali", "Malta", "Marshall Islands (the)", "Martinique", "Mauritania", "Mauritius", "Mayotte",
                     "Mexico", "Micronesia (Federated States of)", "Moldova (the Republic of)", "Monaco", "Mongolia",
                     "Montenegro", "Montserrat", "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal",
                     "Netherlands (the)", "New Caledonia", "New Zealand", "Nicaragua", "Niger (the)", "Nigeria", "Niue",
                     "Norfolk Island", "Northern Mariana Islands (the)", "Norway", "Oman", "Pakistan", "Palau",
                     "Palestine, State of", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines (the)",
                     "Pitcairn", "Poland", "Portugal", "Puerto Rico", "Qatar", "Republic of North Macedonia", "Romania",
                     "Russian Federation (the)", "Rwanda", "Réunion", "Saint Barthélemy",
                     "Saint Helena, Ascension and Tristan da Cunha", "Saint Kitts and Nevis", "Saint Lucia",
                     "Saint Martin (French part)", "Saint Pierre and Miquelon", "Saint Vincent and the Grenadines",
                     "Samoa",
                     "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles",
                     "Sierra Leone", "Singapore", "Sint Maarten (Dutch part)", "Slovakia", "Slovenia",
                     "Solomon Islands",
                     "Somalia", "South Africa", "South Georgia and the South Sandwich Islands", "South Sudan", "Spain",
                     "Sri Lanka", "Sudan (the)", "Suriname", "Svalbard and Jan Mayen", "Sweden", "Switzerland",
                     "Syrian Arab Republic", "Taiwan (Province of China)", "Tajikistan", "Tanzania, United Republic of",
                     "Thailand", "Timor-Leste", "Togo", "Tokelau", "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey",
                     "Turkmenistan", "Turks and Caicos Islands (the)", "Tuvalu", "Uganda", "Ukraine",
                     "United Arab Emirates (the)", "United Kingdom of Great Britain and Northern Ireland (the)",
                     "United States Minor Outlying Islands (the)", "United States of America (the)", "Uruguay",
                     "Uzbekistan", "Vanuatu", "Venezuela (Bolivarian Republic of)", "Viet Nam",
                     "Virgin Islands (British)",
                     "Virgin Islands (U.S.)", "Wallis and Futuna", "Western Sahara", "Yemen", "Zambia", "Zimbabwe",
                     "Åland Islands"]
        alpha2Code = ["AF", "AL", "DZ", "AS", "AD", "AO", "AI", "AQ", "AG", "AR", "AM", "AW", "AU", "AT", "AZ", "BS",
                      "BH",
                      "BD", "BB", "BY", "BE", "BZ", "BJ", "BM", "BT", "BO", "BQ", "BA", "BW", "BV", "BR", "IO", "BN",
                      "BG",
                      "BF", "BI", "CV", "KH", "CM", "CA", "KY", "CF", "TD", "CL", "CN", "CX", "CC", "CO", "KM", "CD",
                      "CG",
                      "CK", "CR", "HR", "CU", "CW", "CY", "CZ", "CI", "DK", "DJ", "DM", "DO", "EC", "EG", "SV", "GQ",
                      "ER",
                      "EE", "SZ", "ET", "FK", "FO", "FJ", "FI", "FR", "GF", "PF", "TF", "GA", "GM", "GE", "DE", "GH",
                      "GI",
                      "GR", "GL", "GD", "GP", "GU", "GT", "GG", "GN", "GW", "GY", "HT", "HM", "VA", "HN", "HK", "HU",
                      "IS",
                      "IN", "ID", "IR", "IQ", "IE", "IM", "IL", "IT", "JM", "JP", "JE", "JO", "KZ", "KE", "KI", "KP",
                      "KR",
                      "KW", "KG", "LA", "LV", "LB", "LS", "LR", "LY", "LI", "LT", "LU", "MO", "MG", "MW", "MY", "MV",
                      "ML",
                      "MT", "MH", "MQ", "MR", "MU", "YT", "MX", "FM", "MD", "MC", "MN", "ME", "MS", "MA", "MZ", "MM",
                      "NA",
                      "NR", "NP", "NL", "NC", "NZ", "NI", "NE", "NG", "NU", "NF", "MP", "NO", "OM", "PK", "PW", "PS",
                      "PA",
                      "PG", "PY", "PE", "PH", "PN", "PL", "PT", "PR", "QA", "MK", "RO", "RU", "RW", "RE", "BL", "SH",
                      "KN",
                      "LC", "MF", "PM", "VC", "WS", "SM", "ST", "SA", "SN", "RS", "SC", "SL", "SG", "SX", "SK", "SI",
                      "SB",
                      "SO", "ZA", "GS", "SS", "ES", "LK", "SD", "SR", "SJ", "SE", "CH", "SY", "TW", "TJ", "TZ", "TH",
                      "TL",
                      "TG", "TK", "TO", "TT", "TN", "TR", "TM", "TC", "TV", "UG", "UA", "AE", "GB", "UM", "US", "UY",
                      "UZ",
                      "VU", "VE", "VN", "VG", "VI", "WF", "EH", "YE", "ZM", "ZW", "AX"]
        alpha3Code = ["AFG", "ALB", "DZA", "ASM", "AND", "AGO", "AIA", "ATA", "ATG", "ARG", "ARM", "ABW", "AUS", "AUT",
                      "AZE", "BHS", "BHR", "BGD", "BRB", "BLR", "BEL", "BLZ", "BEN", "BMU", "BTN", "BOL", "BES", "BIH",
                      "BWA", "BVT", "BRA", "IOT", "BRN", "BGR", "BFA", "BDI", "CPV", "KHM", "CMR", "CAN", "CYM", "CAF",
                      "TCD", "CHL", "CHN", "CXR", "CCK", "COL", "COM", "COD", "COG", "COK", "CRI", "HRV", "CUB", "CUW",
                      "CYP", "CZE", "CIV", "DNK", "DJI", "DMA", "DOM", "ECU", "EGY", "SLV", "GNQ", "ERI", "EST", "SWZ",
                      "ETH", "FLK", "FRO", "FJI", "FIN", "FRA", "GUF", "PYF", "ATF", "GAB", "GMB", "GEO", "DEU", "GHA",
                      "GIB", "GRC", "GRL", "GRD", "GLP", "GUM", "GTM", "GGY", "GIN", "GNB", "GUY", "HTI", "HMD", "VAT",
                      "HND", "HKG", "HUN", "ISL", "IND", "IDN", "IRN", "IRQ", "IRL", "IMN", "ISR", "ITA", "JAM", "JPN",
                      "JEY", "JOR", "KAZ", "KEN", "KIR", "PRK", "KOR", "KWT", "KGZ", "LAO", "LVA", "LBN", "LSO", "LBR",
                      "LBY", "LIE", "LTU", "LUX", "MAC", "MDG", "MWI", "MYS", "MDV", "MLI", "MLT", "MHL", "MTQ", "MRT",
                      "MUS", "MYT", "MEX", "FSM", "MDA", "MCO", "MNG", "MNE", "MSR", "MAR", "MOZ", "MMR", "NAM", "NRU",
                      "NPL", "NLD", "NCL", "NZL", "NIC", "NER", "NGA", "NIU", "NFK", "MNP", "NOR", "OMN", "PAK", "PLW",
                      "PSE", "PAN", "PNG", "PRY", "PER", "PHL", "PCN", "POL", "PRT", "PRI", "QAT", "MKD", "ROU", "RUS",
                      "RWA", "REU", "BLM", "SHN", "KNA", "LCA", "MAF", "SPM", "VCT", "WSM", "SMR", "STP", "SAU", "SEN",
                      "SRB", "SYC", "SLE", "SGP", "SXM", "SVK", "SVN", "SLB", "SOM", "ZAF", "SGS", "SSD", "ESP", "LKA",
                      "SDN", "SUR", "SJM", "SWE", "CHE", "SYR", "TWN", "TJK", "TZA", "THA", "TLS", "TGO", "TKL", "TON",
                      "TTO", "TUN", "TUR", "TKM", "TCA", "TUV", "UGA", "UKR", "ARE", "GBR", "UMI", "USA", "URY", "UZB",
                      "VUT", "VEN", "VNM", "VGB", "VIR", "WLF", "ESH", "YEM", "ZMB", "ZWE", "ALA"]
        numericCode = ["004", "008", "010", "012", "016", "020", "024", "028", "031", "032", "036", "040", "044", "048",
                       "050", "051", "052", "056", "060", "064", "068", "070", "072", "074", "076", "084", "086", "090",
                       "092", "096", "100", "104", "108", "112", "116", "120", "124", "132", "136", "140", "144", "148",
                       "152", "156", "158", "162", "166", "170", "174", "175", "178", "180", "184", "188", "191", "192",
                       "196", "203", "204", "208", "212", "214", "218", "222", "226", "231", "232", "233", "234", "238",
                       "239", "242", "246", "248", "250", "254", "258", "260", "262", "266", "268", "270", "275", "276",
                       "288", "292", "296", "300", "304", "308", "312", "316", "320", "324", "328", "332", "334", "336",
                       "340", "344", "348", "352", "356", "360", "364", "368", "372", "376", "380", "384", "388", "392",
                       "398", "400", "404", "408", "410", "414", "417", "418", "422", "426", "428", "430", "434", "438",
                       "440", "442", "446", "450", "454", "458", "462", "466", "470", "474", "478", "480", "484", "492",
                       "496", "498", "499", "500", "504", "508", "512", "516", "520", "524", "528", "531", "533", "534",
                       "535", "540", "548", "554", "558", "562", "566", "570", "574", "578", "580", "581", "583", "584",
                       "585", "586", "591", "598", "600", "604", "608", "612", "616", "620", "624", "626", "630", "634",
                       "638", "642", "643", "646", "652", "654", "659", "660", "662", "663", "666", "670", "674", "678",
                       "682", "686", "688", "690", "694", "702", "703", "704", "705", "706", "710", "716", "724", "728",
                       "729", "732", "740", "744", "748", "752", "756", "760", "762", "764", "768", "772", "776", "780",
                       "784", "788", "792", "795", "796", "798", "800", "804", "807", "818", "826", "831", "832", "833",
                       "834", "840", "850", "854", "858", "860", "862", "876", "882", "887", "894"]

        if countryFormatFrom == 'Numeric Code':
            if countryFormatTo == 'Alpha-2 Code':
                for x in numericCode:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = alpha2Code[numericCode.index(x)]
            elif countryFormatTo == 'Alpha-3 Code':
                for x in numericCode:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = alpha3Code[numericCode.index(x)]
            elif countryFormatTo == 'Country Name':
                for x in numericCode:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = countries[numericCode.index(x)]
                    # targetData.loc[(targetData[sourceColumn] in numericCode), targetColumn] = countries[numericCode.index(x)]
                    # print(targetData[sourceColumn])
            else:
                print('No To Option Selected')
            targetData.drop(['Numeric Code'], axis=1, inplace=True)
            targetData.drop(['Alpha-2 Code'], axis=1, inplace=True)
            targetData.drop(['Alpha-3 Code'], axis=1, inplace=True)
            targetData.drop(['Country Name'], axis=1, inplace=True)

        elif countryFormatFrom == 'Alpha-2 Code':
            if countryFormatTo == 'Numeric Code':
                for x in alpha2Code:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = numericCode[alpha2Code.index(x)]
            elif countryFormatTo == 'Alpha-3 Code':
                for x in alpha2Code:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = alpha3Code[alpha2Code.index(x)]
            elif countryFormatTo == 'Country Name':
                for x in alpha2Code:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = countries[alpha2Code.index(x)]
            else:
                print('No To Option Selected')
            targetData.drop(['Numeric Code'], axis=1, inplace=True)
            targetData.drop(['Alpha-2 Code'], axis=1, inplace=True)
            targetData.drop(['Alpha-3 Code'], axis=1, inplace=True)
            targetData.drop(['Country Name'], axis=1, inplace=True)

        elif countryFormatFrom == 'Alpha-3 Code':
            if countryFormatTo == 'Alpha-2 Code':
                for x in alpha3Code:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = alpha2Code[alpha3Code.index(x)]
            elif countryFormatTo == 'Numeric Code':
                for x in alpha3Code:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = numericCode[alpha3Code.index(x)]
            elif countryFormatTo == 'Country Name':
                for x in alpha3Code:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = countries[alpha3Code.index(x)]
            else:
                print('No To Option Selected')
            targetData.drop(['Numeric Code'], axis=1, inplace=True)
            targetData.drop(['Alpha-2 Code'], axis=1, inplace=True)
            targetData.drop(['Alpha-3 Code'], axis=1, inplace=True)
            targetData.drop(['Country Name'], axis=1, inplace=True)

        elif countryFormatFrom == 'Country Name':
            if countryFormatTo == 'Alpha-2 Code':
                for x in countries:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = alpha2Code[countries.index(x)]
            elif countryFormatTo == 'Alpha-3 Code':
                for x in countries:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = alpha3Code[countries.index(x)]
            elif countryFormatTo == 'Numeric Code':
                for x in countries:
                    targetData.loc[(targetData[sourceColumn] == x), targetColumn] = numericCode[countries.index(x)]
            else:
                print('No To Option Selected')
            targetData.drop(['Numeric Code'], axis=1, inplace=True)
            targetData.drop(['Alpha-2 Code'], axis=1, inplace=True)
            targetData.drop(['Alpha-3 Code'], axis=1, inplace=True)
            targetData.drop(['Country Name'], axis=1, inplace=True)

        else:
            print('No From Option Selected')
        print('Country Hardcoding Complete')
        # harcodeLabel = Label(hcodeColumnFrame, text="Transform Complete", bg='green', fg='white')
        # harcodeLabel.grid(row=2, padx=5, pady=5, sticky='')

    def HARDCODE_VALUES():
        # HARD CODE VALUES
        global hcodeSource
        global chooseCountryFormatFrom
        global chooseCountryFormatTo
        global cntryTrgt
        Label(hcodeColumnFrame, text="Select Source Format").grid(row=0, column=0, padx=10, pady=10)
        Label(hcodeColumnFrame, text="Select Target Format").grid(row=0, column=2, padx=10, pady=10)
        n = StringVar()

        chooseCountryFormatFrom = Combobox(hcodeColumnFrame, width=27)
        chooseCountryFormatFrom['values'] = (
            'Country Name', 'Alpha-2 Code', 'Alpha-3 Code', 'Numeric Code')
        chooseCountryFormatFrom.grid(row=0, column=1)
        chooseCountryFormatFrom.current()
        countryFrom = None

        def getFrom(event):
            global countryFrom
            countryFrom = event.widget.get()

        chooseCountryFormatFrom.bind('<<ComboboxSelected>>', getFrom)

        chooseCountryFormatTo = Combobox(hcodeColumnFrame, width=27)
        chooseCountryFormatTo['values'] = (
            'Numeric Code', 'Alpha-3 Code', 'Alpha-2 Code', 'Country Name')
        chooseCountryFormatTo.grid(row=0, column=3)
        chooseCountryFormatTo.current()
        countryTo = None

        def getTo(event):
            global countryTo
            countryTo = event.widget.get()

        chooseCountryFormatTo.bind('<<ComboboxSelected>>', getTo)
        # Source and Target columns
        Label(hcodeColumnFrame, text='Source Column').grid(row=1, column=0)
        Label(hcodeColumnFrame, text='Target Column').grid(row=1, column=2)
        # Source column input
        # cntrySrc = Entry(hcodeColumnFrame)
        # cntrySrc.grid(row=1, column=1, padx=5, pady=5)

        hcodeSourceValue = None
        hcodeSource = Combobox(hcodeColumnFrame, width=27)
        hcodeSource['values'] = columnHeaders
        hcodeSource.grid(row=1, column=1, padx=5, pady=5)
        hcodeSource.current()

        def hcodeSourceSelect(event):
            global hcodeSourceValue
            hcodeSourceValue = event.widget.get()

        hcodeSource.bind('<<ComboboxSelected>>', hcodeSourceSelect)

        # Target column input
        cntryTrgt = Entry(hcodeColumnFrame)
        cntryTrgt.grid(row=1, column=3, padx=5, pady=5)

    #######################################################################################################################

    # Pandas Transformation
    def exportToExcel():
        global exportLabel
        filedirectory = filedialog.askdirectory()
        targetData.to_excel(filedirectory + '/output.xlsx', index=False)
        print('Export to Excel Complete')
        exportLabel = Label(exportFrame, text="Export Success", bg='green', fg='white')
        exportLabel.grid(row=1, column=2, padx=10, sticky='')

    def runTransform():
        global transformLabel
        if dttarget.get() != '':
            date_format_standardization(dtsrcValue, dttarget.get(), comboValue)
        if splitTrgt1.get() != '':
            split_single_attribute(splitSrcValue, splitTrgt1.get(), splitTrgt2.get(), splitTrgt3.get())

        if mergeTrgt.get() != '':
            merge_name_field(mergeSrcValue1, mergeSrcValue2, mergeTrgt.get())

        if cntryTrgt.get() != '':
            decode_countries(hcodeSourceValue, cntryTrgt.get(), countryFrom, countryTo)
        transformLabel = Label(exportFrame, text="Transform Complete", bg='green', fg='white')
        transformLabel.grid(row=1, column=0, padx=10, sticky='')

    def clear():
        try:
            excelLabel.destroy()
        except:
            print('Excel label does not exist')
        try:
            initiateLabel.destroy()
        except:
            print('Initiate label does not exist')
        try:
            exportLabel.destroy()
        except:
            print('Export label does not exist')
        try:
            transformLabel.destroy()
        except:
            print('Transform label does not exist')

        srcFile.delete('0', END)
        dttarget.delete('0', END)
        splitTrgt1.delete('0', END)
        splitTrgt2.delete('0', END)
        splitTrgt3.delete('0', END)
        mergeTrgt.delete('0', END)
        cntryTrgt.delete('0', END)
        formatChoosen.set('')
        sheetNamesDropDown.set('')
        dtsrc.set('')
        splitSrc.set('')
        mergeSrc1.set('')
        mergeSrc2.set('')
        chooseCountryFormatFrom.set('')
        chooseCountryFormatTo.set('')
        hcodeSource.set('')

    def BUTTONS_EXECUTE():
        # Button to run transformation
        Button(exportFrame, text='Transform', command=runTransform).grid(row=0, column=0, padx=20, pady=10)

        # Button to export to excel
        Button(exportFrame, text='Export', command=exportToExcel).grid(row=0, column=1, padx=20, pady=10)

        # Button to clear all the fields
        Button(exportFrame, text='Clear', command=clear).grid(row=0, column=2, padx=20, pady=10)


def EXCEL_DIFFERENCE(root):
    global mainFrame
    global innerFrame
    global sourceFileFrame
    global sourceFileFrame2
    global targetFileFrame
    global inputFrame
    global exportFrame
    global outputFrame
    global dateFormatFrame
    global splitColumnFrame
    global mergeColumnFrame
    global hcodeColumnFrame
    global messageFrame

    try:
        sourceFileFrame.destroy()
    except:
        print('sourceFileFrame does not exist')
    try:
        dateFormatFrame.destroy()
    except:
        print('dateFormatFrame does not exist')
    try:
        splitColumnFrame.destroy()
    except:
        print('splitColumnFrame does not exist')
    try:
        mergeColumnFrame.destroy()
    except:
        print('mergeColumnFrame does not exist')
    try:
        hcodeColumnFrame.destroy()
    except:
        print('hcodeColumnFrame does not exist')
    try:
        exportFrame.destroy()
    except:
        print('exportFrame does not exist')
    try:
        messageFrame.destroy()
    except:
        print('messageFrame does not exist')
    try:
        mainFrame.destroy()
    except:
        print('mainFrame does not exist')
    try:
        innerFrame.destroy()
    except:
        print('innerFrame does not exist')
    try:
        sourceFileFrame2.destroy()
    except:
        print('sourceFileFrame2 does not exist')
    try:
        targetFileFrame.destroy()
    except:
        print('targetFileFrame does not exist')
    try:
        inputFrame.destroy()
    except:
        print('inputFrame does not exist')
    try:
        exportFrame.destroy()
    except:
        print('exportFrame does not exist')
    try:
        outputFrame.destroy()
    except:
        print('outputFrame does not exist')

    def getsheetnames(srcpath):
        global excelLabel3
        global sheetNamesDropDown2
        global fileSource2
        global sheetNameList2
        global sheetNameValues2
        fileSource2 = dw.ExcelFile(srcpath)
        sheetNameList2 = fileSource2.sheet_names
        Label(sourceFileFrame2, text='Select Sheet').grid(row=1, column=0, padx=20, pady=20, sticky='')
        sheetNamesDropDown2 = Combobox(sourceFileFrame2, width=27)
        sheetNamesDropDown2.grid(row=1, column=1, padx=5, pady=5)
        sheetNamesDropDown2['values'] = sheetNameList2
        sheetNamesDropDown2.current()

        def sheetNameEnter(event):
            global sheetNameValues2
            sheetNameValues2 = event.widget.get()

        sheetNamesDropDown2.bind('<<ComboboxSelected>>', sheetNameEnter)
        Button(sourceFileFrame2, text='Initiate', command=lambda: initiate(srcFile.get())).grid(row=1, column=2,
                                                                                                padx=20,
                                                                                                pady=20,
                                                                                                sticky='')
        excelLabel3 = Label(sourceFileFrame2, text="Imported", bg='green', fg='white')
        excelLabel3.grid(row=0, column=3, padx=5, pady=5, sticky='')

    def initiate(srcpath):
        global initiateLabel2
        global targetData2
        global targetPath2
        global excelLabel3
        global columnHeaders2
        global sourceData2

        sourceData2 = dw.read_excel(srcpath, sheet_name=sheetNameValues2, dtype='object')  # sheet name - 'Sheet1'
        targetData2 = dw.DataFrame(sourceData2)
        columnHeaders2 = list(targetData2.columns)
        print('Excel imported')
        initiateLabel2 = Label(sourceFileFrame2, text="Initiated", bg='green', fg='white')
        initiateLabel2.grid(row=1, column=2, padx=5, pady=5, sticky='')
        target()
        # keyColumn()
        exportExcel()
        reconcile()

    #######################################################################################################################

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
    sourceFileFrame2 = LabelFrame(innerFrame, text='Source File Directory', padx=20, pady=20)
    sourceFileFrame2.grid(row=0, pady=20, sticky='')

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
    Label(sourceFileFrame2, text='Source File Path').grid(row=0, column=0)
    srcFile = Entry(sourceFileFrame2, width=50)
    srcFile.grid(row=0, column=1, padx=5, pady=5)

    def browseSrcFiles():
        filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                              filetypes=(("Excel files", "*.xlsx*"), ("all files", "*.*")))
        srcFile.insert(0, filename)

    Button(sourceFileFrame2, text='Browse', command=browseSrcFiles).grid(row=0, column=2, padx=20, pady=20, sticky='')
    Button(sourceFileFrame2, text='Import',
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

        Button(targetFileFrame, text='Browse', command=browseTrgFiles).grid(row=0, column=2, padx=20, pady=20,
                                                                            sticky='')

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
        return x[0] if x[0] == x[1] or dw.isna(x).all() else f'{x[0]} ---> {x[1]}'

    def strip(x):
        """Function to use with applymap to strip whitespaces in a dataframe."""
        return x.strip() if isinstance(x, str) else x

    def diff2_dw(old_df, new_df, idx_col):
        """Identify diff2erences between two pandas DataFrames using a key column.
        Key column is assumed to have only unique data
        (like a database unique id index)
        Args:
            old_df (dw.DataFrame): first dataframe
            new_df (dw.DataFrame): second dataframe
            idx_col (str|list(str)): column name(s) of the index,
              needs to be present in both DataFrames
        """
        # setting the column name as index for fast operations
        old_df = old_df.set_index(idx_col)
        new_df = new_df.set_index(idx_col)
        # get the added and removed rows
        old_keys = old_df.index
        new_keys = new_df.index
        if isinstance(old_keys, dw.MultiIndex):
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
        common_data = dw.concat(
            [old_common.reset_index(), new_common.reset_index()], sort=True
        )
        changed_keys = common_data.drop_duplicates(keep=False)[idx_col]
        if isinstance(changed_keys, dw.Series):
            changed_keys = changed_keys.unique()
        else:
            changed_keys = changed_keys.drop_duplicates().set_index(idx_col).index
        # combining the changed rows via multi level columns
        df_all_changes = dw.concat(
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
        df1 = dw.read_excel(path1, sheet_name=sheetname)
        df2 = dw.read_excel(path2, sheet_name=sheetname)
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
        old_df = dw.read_excel(path1, sheet_name=sheet_name, **kwargs)
        new_df = dw.read_excel(path2, sheet_name=sheet_name, **kwargs)
        diff2 = diff2_dw(old_df, new_df, index_col_name)
        # with dw.ExcelWriter(out_path) as writer:
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
        comparison(srcFile.get(), trgFile2.get(), sheetNameValues2, expFile2.get())
        # T = Text(outputFrame, width=52).grid(row=0, column=0, padx=10, pady=10)
        # T.insert('END', diff2)
        Label(exportFrame, text='Complete').grid(row=0, column=0, padx=20, pady=20, sticky='')

    def reconcile():
        Button(exportFrame, text='Reconcile', command=compare).grid(row=0, column=0, padx=20, pady=10)


root = Tk()
root.title('Maveric Transformation Validator')
root.geometry('850x650')
menubar = Menu(root)
menubar.add_command(label="Transformation", command=lambda: TRANSFORM(root))
menubar.add_command(label="Reconciliation", command=lambda: EXCEL_DIFFERENCE(root))

root.config(menu=menubar)
root.mainloop()
