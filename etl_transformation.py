import pandas as dw


class Etl:

    def de_duplication(self):
        data_1 = dw.read_excel('StaggingData.xlsx', sheet_name='Sheet1')
        data_1.head()
        data_1.drop_duplicates(subset=['Fname', 'Lname', 'Gender', 'Address', 'DOB'], keep='first', inplace=True)
        data_1.to_excel('DeDuplicationData.xlsx', index=False)
        return

    def date_format_standardization(self):
        # date format as input from the user (dd-mm-yyyy, dd-mm-yy, etc)
        data_2 = dw.read_excel('DeDuplicationData.xlsx', sheet_name='Sheet1')
        dataF_2 = dw.DataFrame(data_2)
        dataF_2['DOB'] = dw.to_datetime(data_2.DOB, errors='coerce')
        dataF_2['DOB'] = dataF_2['DOB'].dt.strftime('%e %B,%Y')
        dataF_2.to_excel('DateFormatStandardization.xlsx', index=False)
        return

    def derived_attribute_age(self):
        data_3 = dw.read_excel('DateFormatStandardization.xlsx')
        dataF_3 = dw.DataFrame(data_3)
        thisyear = dw.datetime.now().year
        dataF_3a = dataF_3.copy()
        dataF_3a['DOB'] = dw.to_datetime(data_3.DOB, errors='coerce')
        dataF_3a['DOB'] = dataF_3a['DOB'].dt.strftime("%Y")
        dataF_3a.loc[(dataF_3.DOB == 'NaT'), 'DOB'] = None
        dataF_3['Age'] = abs(thisyear - dataF_3a['DOB'].astype(float))
        dataF_3.to_excel('DerivedAttributeAge.xlsx', index=False)
        return

    def split_single_attribute(self):
        data_8 = dw.read_excel('DerivedAttributeAge.xlsx')
        dataF_8 = dw.DataFrame(data_8)
        new = dataF_8['Address'].str.split(", | ", n=1, expand=True)
        dataF_8["Street"] = new[0]
        dataF_8["Colony"] = new[1]
        new_1 = dataF_8['Colony'].str.split(", | ", n=1, expand=True)
        dataF_8["Colony"] = new_1[0]
        dataF_8["City"] = new_1[1]
        dataF_8.drop(columns=["Address"], inplace=True)
        dataF_8.to_excel('SplitSingleAttribute.xlsx', index=False)
        return

    def merge_name_field(self):
        data_4 = dw.read_excel('SplitSingleAttribute.xlsx')
        dataF_4 = dw.DataFrame(data_4)
        dataF_4['Full Name'] = (dataF_4["Fname"] + " " + dataF_4["Lname"])
        dataF_4.drop(['Fname'], axis=1, inplace=True)
        dataF_4.drop(['Lname'], axis=1, inplace=True)
        dataF_4.to_excel('MergeNameField.xlsx', index=False)
        return

    def decode_gender_field(self):
        data_5 = dw.read_excel('MergeNameField.xlsx')
        dataF_5 = dw.DataFrame(data_5)
        dataF_5.loc[(dataF_5.Gender == 'male') | (dataF_5.Gender == 'Male') |
                    (dataF_5.Gender == 'm') | (dataF_5.Gender == 'M') | (dataF_5.Gender == 1)
        , 'Gender'] = 'male'
        dataF_5.loc[(dataF_5.Gender == 'female') | (dataF_5.Gender == 'Female') |
                    (dataF_5.Gender == 'f') | (dataF_5.Gender == 'F') | (dataF_5.Gender == 0)
        , 'Gender'] = 'female'
        dataF_5.to_excel('DecodeGenderFi    eld.xlsx', index=False)
        return

    def delete_inconsistent_rows(self):
        data_6 = dw.read_excel('DecodeGenderField.xlsx')
        dataF_6 = dw.DataFrame(data_6)
        dataF_6 = dataF_6.drop(dataF_6[(dataF_6.Gender != 'male') & (dataF_6.Gender != 'female')].index)
        dataF_6.to_excel('DeleteInconsistentRows.xlsx', index=False)
        return

    def summarization_1(self):
        data_7 = dw.read_excel('DeleteInconsistentRows.xlsx')
        dataF_7 = dw.DataFrame(data_7)
        dataF_7 = dataF_7.groupby('City')['Age'].sum().reset_index(name='Sum of Ages of Customer')
        # dataF_7 = dataF_7.rename(column={'Age':'Sum of Ages of Customer'}, inplace=True)
        dataF_7.to_excel('output_1.xlsx', index=False)
        return

    def summarization_2(self):
        data_0 = dw.read_excel('DeleteInconsistentRows.xlsx')
        dataF_0 = dw.DataFrame(data_0)
        dataF_0 = dataF_0.groupby('Age')['ID'].count().reset_index(name='Number of Customers')
        # dataF_0 = dataF_0.rename(column={'ID':'Number of Customers'}, inplace=True)
        dataF_0.to_excel('output_2.xlsx', index=False)


a = Etl()
a.de_duplication()
# a.date_format_standardization()
# a.derived_attribute_age()
# a.split_single_attribute()
# a.merge_name_field()
# a.decode_gender_field()
# a.delete_inconsistent_rows()
# a.summarization_1()
# a.summarization_2()
