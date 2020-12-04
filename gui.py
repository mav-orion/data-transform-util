import warnings
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
import pandas as dw

fileSource = None
sourceData = None
targetData = None
targetPath = None
excelLabel = None
dateLabel = None
splitLabel = None
mergeLabel = None
exportLabel = None
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
warnings.simplefilter(action='ignore', category=FutureWarning)


def getsheetnames(srcpath):
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
root = Tk()

root.title('Maveric Transformation Validator')
root.geometry('800x650')
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
    print('Date Format Standardization Complete')
    # dateLabel = Label(dateFormatFrame, text="Transform Complete", bg='green', fg='white')
    # dateLabel.grid(row=2, padx=5, pady=5, sticky='')


def STANDARD_DATE():
    # DATE FORMAT STANDARDIZATION
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
    new = targetData[sourceColumn].str.split(",", n=1, expand=True)
    targetData[targetColumn1] = new[0]
    targetData[targetColumn2] = new[1]
    new_1 = targetData[targetColumn2].str.split(",", n=1, expand=True)
    targetData[targetColumn2] = new_1[0]
    targetData[targetColumn3] = new_1[1]
    targetData.drop(columns=[sourceColumn], inplace=True)
    print('Column Split Complete')
    # splitLabel = Label(splitColumnFrame, text="Transform Complete", bg='green', fg='white')
    # splitLabel.grid(row=2, padx=5, pady=5, sticky='')

def SPLIT_SINGLE_COLUMN():
    # SPLIT COLUMNS
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
    print(mergeSrcValue1)
    print(mergeSrcValue2)
    # Button(mergeColumnFrame, text='Transform',
    #        command=lambda: merge_name_field(mergeSrc1.get(), mergeSrc2.get(), mergeTrgt.get())).grid(row=2, column=0,
    #                                                                                                  padx=20,
    #                                                                                                  pady=20)



#######################################################################################################################

def decode_countries(sourceColumn, targetColumn, countryFormatFrom, countryFormatTo):
    # global harcodeLabel
    countries = ["Afghanistan", "Albania", "Algeria", "American Samoa", "Andorra", "Angola", "Anguilla", "Antarctica",
                 "Antigua and Barbuda", "Argentina", "Armenia", "Aruba", "Australia", "Austria", "Azerbaijan",
                 "Bahamas (the)", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin",
                 "Bermuda", "Bhutan", "Bolivia (Plurinational State of)", "Bonaire, Sint Eustatius and Saba",
                 "Bosnia and Herzegovina", "Botswana", "Bouvet Island", "Brazil",
                 "British Indian Ocean Territory (the)", "Brunei Darussalam", "Bulgaria", "Burkina Faso", "Burundi",
                 "Cabo Verde", "Cambodia", "Cameroon", "Canada", "Cayman Islands (the)",
                 "Central African Republic (the)", "Chad", "Chile", "China", "Christmas Island",
                 "Cocos (Keeling) Islands (the)", "Colombia", "Comoros (the)", "Congo (the Democratic Republic of the)",
                 "Congo (the)", "Cook Islands (the)", "Costa Rica", "Croatia", "Cuba", "Curaçao", "Cyprus", "Czechia",
                 "Côte d'Ivoire", "Denmark", "Djibouti", "Dominica", "Dominican Republic (the)", "Ecuador", "Egypt",
                 "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia",
                 "Falkland Islands (the) [Malvinas]", "Faroe Islands (the)", "Fiji", "Finland", "France",
                 "French Guiana", "French Polynesia", "French Southern Territories (the)", "Gabon", "Gambia (the)",
                 "Georgia", "Germany", "Ghana", "Gibraltar", "Greece", "Greenland", "Grenada", "Guadeloupe", "Guam",
                 "Guatemala", "Guernsey", "Guinea", "Guinea-Bissau", "Guyana", "Haiti",
                 "Heard Island and McDonald Islands", "Holy See (the)", "Honduras", "Hong Kong", "Hungary", "Iceland",
                 "India", "Indonesia", "Iran (Islamic Republic of)", "Iraq", "Ireland", "Isle of Man", "Israel",
                 "Italy", "Jamaica", "Japan", "Jersey", "Jordan", "Kazakhstan", "Kenya", "Kiribati",
                 "Korea (the Democratic People's Republic of)", "Korea (the Republic of)", "Kuwait", "Kyrgyzstan",
                 "Lao People's Democratic Republic (the)", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya",
                 "Liechtenstein", "Lithuania", "Luxembourg", "Macao", "Madagascar", "Malawi", "Malaysia", "Maldives",
                 "Mali", "Malta", "Marshall Islands (the)", "Martinique", "Mauritania", "Mauritius", "Mayotte",
                 "Mexico", "Micronesia (Federated States of)", "Moldova (the Republic of)", "Monaco", "Mongolia",
                 "Montenegro", "Montserrat", "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal",
                 "Netherlands (the)", "New Caledonia", "New Zealand", "Nicaragua", "Niger (the)", "Nigeria", "Niue",
                 "Norfolk Island", "Northern Mariana Islands (the)", "Norway", "Oman", "Pakistan", "Palau",
                 "Palestine, State of", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines (the)",
                 "Pitcairn", "Poland", "Portugal", "Puerto Rico", "Qatar", "Republic of North Macedonia", "Romania",
                 "Russian Federation (the)", "Rwanda", "Réunion", "Saint Barthélemy",
                 "Saint Helena, Ascension and Tristan da Cunha", "Saint Kitts and Nevis", "Saint Lucia",
                 "Saint Martin (French part)", "Saint Pierre and Miquelon", "Saint Vincent and the Grenadines", "Samoa",
                 "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles",
                 "Sierra Leone", "Singapore", "Sint Maarten (Dutch part)", "Slovakia", "Slovenia", "Solomon Islands",
                 "Somalia", "South Africa", "South Georgia and the South Sandwich Islands", "South Sudan", "Spain",
                 "Sri Lanka", "Sudan (the)", "Suriname", "Svalbard and Jan Mayen", "Sweden", "Switzerland",
                 "Syrian Arab Republic", "Taiwan (Province of China)", "Tajikistan", "Tanzania, United Republic of",
                 "Thailand", "Timor-Leste", "Togo", "Tokelau", "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey",
                 "Turkmenistan", "Turks and Caicos Islands (the)", "Tuvalu", "Uganda", "Ukraine",
                 "United Arab Emirates (the)", "United Kingdom of Great Britain and Northern Ireland (the)",
                 "United States Minor Outlying Islands (the)", "United States of America (the)", "Uruguay",
                 "Uzbekistan", "Vanuatu", "Venezuela (Bolivarian Republic of)", "Viet Nam", "Virgin Islands (British)",
                 "Virgin Islands (U.S.)", "Wallis and Futuna", "Western Sahara", "Yemen", "Zambia", "Zimbabwe",
                 "Åland Islands"]
    alpha2Code = ["AF", "AL", "DZ", "AS", "AD", "AO", "AI", "AQ", "AG", "AR", "AM", "AW", "AU", "AT", "AZ", "BS", "BH",
                  "BD", "BB", "BY", "BE", "BZ", "BJ", "BM", "BT", "BO", "BQ", "BA", "BW", "BV", "BR", "IO", "BN", "BG",
                  "BF", "BI", "CV", "KH", "CM", "CA", "KY", "CF", "TD", "CL", "CN", "CX", "CC", "CO", "KM", "CD", "CG",
                  "CK", "CR", "HR", "CU", "CW", "CY", "CZ", "CI", "DK", "DJ", "DM", "DO", "EC", "EG", "SV", "GQ", "ER",
                  "EE", "SZ", "ET", "FK", "FO", "FJ", "FI", "FR", "GF", "PF", "TF", "GA", "GM", "GE", "DE", "GH", "GI",
                  "GR", "GL", "GD", "GP", "GU", "GT", "GG", "GN", "GW", "GY", "HT", "HM", "VA", "HN", "HK", "HU", "IS",
                  "IN", "ID", "IR", "IQ", "IE", "IM", "IL", "IT", "JM", "JP", "JE", "JO", "KZ", "KE", "KI", "KP", "KR",
                  "KW", "KG", "LA", "LV", "LB", "LS", "LR", "LY", "LI", "LT", "LU", "MO", "MG", "MW", "MY", "MV", "ML",
                  "MT", "MH", "MQ", "MR", "MU", "YT", "MX", "FM", "MD", "MC", "MN", "ME", "MS", "MA", "MZ", "MM", "NA",
                  "NR", "NP", "NL", "NC", "NZ", "NI", "NE", "NG", "NU", "NF", "MP", "NO", "OM", "PK", "PW", "PS", "PA",
                  "PG", "PY", "PE", "PH", "PN", "PL", "PT", "PR", "QA", "MK", "RO", "RU", "RW", "RE", "BL", "SH", "KN",
                  "LC", "MF", "PM", "VC", "WS", "SM", "ST", "SA", "SN", "RS", "SC", "SL", "SG", "SX", "SK", "SI", "SB",
                  "SO", "ZA", "GS", "SS", "ES", "LK", "SD", "SR", "SJ", "SE", "CH", "SY", "TW", "TJ", "TZ", "TH", "TL",
                  "TG", "TK", "TO", "TT", "TN", "TR", "TM", "TC", "TV", "UG", "UA", "AE", "GB", "UM", "US", "UY", "UZ",
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
        print('Excel Label does not exist')

    # try:
    #     dateLabel.destroy()
    # except:
    #     print('Date Label does not exist')
    #
    # try:
    #     splitLabel.destroy()
    # except:
    #     print('Split Label does not exist')
    #
    # try:
    #     mergeLabel.destroy()
    # except:
    #     print('Merge Label does not exist')
    #
    # try:
    #     harcodeLabel.destroy()
    # except:
    #     print('Hardcode Label does not exist')
    try:
        exportLabel.destroy()
    except:
        print('Export Label does not exist')
    try:
        transformLabel.destroy()
    except:
        print('Transform Label does not exist')

    srcFile.delete('0', END)
    # trgtFile.delete('0', END)
    dttarget.delete('0', END)
    # splitSrc.delete('0', END)
    splitTrgt1.delete('0', END)
    splitTrgt2.delete('0', END)
    splitTrgt3.delete('0', END)
    # mergeSrc1.delete('0', END)
    # mergeSrc2.delete('0', END)
    mergeTrgt.delete('0', END)
    formatChoosen.set('')
    sheetNamesDropDown.set('')
    # chooseCountryFormatFrom.set('')
    # chooseCountryFormatTo.set('')
    # cntryTrgt.delete('0', END)
    # cntrySrc.delete('0', END)

def BUTTONS_EXECUTE():
    # Button to run transformation
    Button(exportFrame, text='Transform', command=runTransform).grid(row=0, column=0, padx=20, pady=10)

    # Button to export to excel
    Button(exportFrame, text='Export', command=exportToExcel).grid(row=0, column=1, padx=20, pady=10)

    # Button to clear all the fields
    Button(exportFrame, text='Clear', command=clear).grid(row=0, column=2, padx=20, pady=10)

root.mainloop()
