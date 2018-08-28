# The script below syncs the supply Chain with Engineering BOM.
# Developed by Yifei.li@byton.com



import pandas as pd
import numpy as np
import xlrd, openpyxl
import csv, datetime

#load EBOM
def load():
    EBOM = pd.read_excel("/home/yifei/Documents/BOM_tool/EBOM 8.23.18.xlsx")
    EBOM.columns = EBOM.columns.str.replace("\(R\)\ ", "") # trim (R) away from the header
#load Supply Chain BOM
    SCBOM = pd.read_excel("/home/yifei/Documents/BOM_tool/Supply Chain BOM.xlsx")
    return EBOM, SCBOM


# search PN and Rev in SCBOM
def search(df, PN, Rev):
    df = df.loc[(df["Title"]==PN) & (df["Revision"]==Rev)]
    index = df.index
    if (len(index)==1):
        return index
    else:
        #return the first index if duplicate PN found
        return index[0:1]
        

def replace_bracket(data):
    value = str(data).replace('[', '').replace(']', '').replace('\'', '')
    if value == 'nan':
        return ""
    return value

def change_datetime_to_string(data):
    if data.size == 0:
        return ""
    elif (data == np.array([' '])):
        return ""
    elif data == "nan":
        return ""
    else:
    	return data[0]

# copy Shaolong's BOM info and paste into SCBOM
def copy_and_paste_row(df1, index1, df2, index2):
	df1.loc[index1,["Identifier"]] = replace_bracket(df2.loc[index2,["Identifier"]].values)
	df1.loc[index1,["Title"]] = replace_bracket(df2.loc[index2,["Title"]].values)
	df1.loc[index1,["Revision"]] = replace_bracket(df2.loc[index2,["Revision"]].values)
	df1.loc[index1,["Description"]] = replace_bracket(df2.loc[index2,["Description"]].values)
	df1.loc[index1,["QTY"]] = replace_bracket(df2.loc[index2,["QTY"]].values)
	df1.loc[index1,["UOM"]] = replace_bracket(df2.loc[index2,["UOM"]].values)
	df1.loc[index1,["Purchased Part Type"]] = replace_bracket(df2.loc[index2,["Purchased Part Type"]].values)
	df1.loc[index1,["UOM"]] = replace_bracket(df2.loc[index2,["UOM"]].values)
	df1.loc[index1,["Maturity"]] = replace_bracket(df2.loc[index2,["Maturity"]].values)
	df1.loc[index1,["Part Type"]] = replace_bracket(df2.loc[index2,["Part Type"]].values)
	df1.loc[index1,["System"]] = replace_bracket(df2.loc[index2,["System"]].values)
	df1.loc[index1,["SubSystem"]] = replace_bracket(df2.loc[index2,["SubSystem"]].values)
	df1.loc[index1,["Legacy Part Number"]] = replace_bracket(df2.loc[index2,["Legacy Part Number"]].values)
	df1.loc[index1,["Legacy Part Revision"]] = replace_bracket(df2.loc[index2,["Legacy Part Revision"]].values)
	df1.loc[index1,["Configuration"]] = replace_bracket(df2.loc[index2,["Configuration"]].values)

	return df1

# save dataframe to Excel
def save(df):
    #https://stackoverflow.com/questions/28837057/pandas-writing-an-excel-file-containing-unicode-illegalcharactererror
    df = df.applymap(lambda x: x.encode('unicode_escape').
                 decode('utf-8') if isinstance(x, str) else x)

    writer = pd.ExcelWriter('Updated Supply Chain BOM.xlsx')
    df.to_excel(writer, sheet_name="Updated Supply Chain BOM", na_rep="" )
    writer.save()

#     df.to_csv("Supply Chain BOM.csv",quoting=csv.QUOTE_NONE, escapechar="\\")

def main():
    EBOM, SCBOM = load()
for index, row in SCBOM.iterrows():
    PN = row["Title"]
    Rev = row["Revision"]
    index_EBOM = search(EBOM, PN, Rev)

    if (index_EBOM.size == 0):
        # if not found, deactivate the part
        print("looped in")
        SCBOM.loc[index, ["Part Active"]] = "Inactivate"
        SCBOM.loc[index, ["Part Status"]] = "Removed"
        SCBOM.loc[index, ["Last Modified Date"]] = datetime.datetime.now()

    else:
        # if found, copy EBOM values to SCBOM and continue
        df = copy_and_paste_row(SCBOM, index, EBOM, index_EBOM)
    
    #check if EBOM PN quantities match SCBOM activate PN quantities
    #if matched, do nothing, continue

    # print("EBOM shape: ", EBOM.shape)
    # print("SCBOM shape: ", SCBOM.shape)
    # check if EBOM PN quantities match SCBOM activate PN quantities
    # if (EBOM.shape[0] == SCBOM.shape[0]):
    #     print("SCBOM successfully updated")
    # else:
    #     #if not matched, find out new PN added to EBOM and added them to SCBOM
    #     print("failed")



    save(df)


if __name__ == "__main__":
	main()