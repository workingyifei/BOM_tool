# The script below syncs the supply Chain BOM with Engineering BOM.
# Created by Yifei.li@byton.com



import pandas as pd
import numpy as np
import xlrd, openpyxl
import csv, datetime
import time


# put up-to-date EBOM and SCBOM here for sync
EBOM_PATH = './Data/AA-ZZ000001NN-03_VP BoM_9-21-18 at 3.42 PM.xlsx'
SCBOM_PATH = './Data/Supply Chain BOM V2 (version 1).xlsx'


#load EBOM and SCBOM
def load():
    # load Engineering BOM
    EBOM = pd.read_excel(EBOM_PATH, sheet_name="BOM", skiprows=[0])
    EBOM.columns = EBOM.columns.str.replace("\(R\)\ ", "") # trim (R) away from the header
    EBOM = EBOM.reset_index(drop=True)
    # load Supply Chain BOM
    SCBOM = pd.read_excel(SCBOM_PATH, sheet_name="Supply Chain BOM")
    SCBOM = SCBOM.reset_index(drop=True)
    return EBOM, SCBOM


# search PN and Rev in SCBOM
def search(df, PN, Rev):
    df = df.loc[(df["Title"]==PN) & (df["Revision"]==Rev)]
    index = df.index.values
    
    # only one entry found
    if (len(index)==1):
        return index   
    
    # multiple entries found
    elif (len(index)>1):
        #return the first index if duplicate PN and Rev found
        return index.tolist()   # returned indexes are a numpy.ndarray, we convert to a list and get the first value
    
    # no entry found
    else:
        return None
        

# copy EBOM info and paste into SCBOM_updated
def copy_and_paste_row(df1, index1, df2, index2):  # index1 and index2 are int
    # copy df2 info into df1
    for i in range(len(df2.columns.tolist())):
        df1.loc[index1, df1.columns.tolist()[i]] = df2.loc[index2, df2.columns.tolist()[i]]
        
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
	start_time = time.time()
	EBOM, SCBOM = load()

	# create a updated SCBOM with columns from SCBOM and data from EBOM
	SCBOM_updated = EBOM.copy()
	for each in SCBOM.columns.tolist()[15:]:
		SCBOM_updated[each] = ""


	SCBOM_columns = SCBOM.columns.tolist()
	SCBOM_columns_size = len(SCBOM_columns)
	EBOM_columns = EBOM.columns.tolist()
	EBOM_columns_size = len(EBOM_columns)

	print("Before Sync")
	print("EBOM shape: ", EBOM.shape)
	print("SCBOM shape: ", SCBOM.shape)
	print("SCBOM_updated ", SCBOM_updated.shape)


	# loop through SCBOM 
	for index, row in SCBOM.iterrows():
	    PN = row["Title"]
	    Rev = row["Revision"]
	    # same part could be structured differently but the part is the same, 
	    # so if same part found at multiple places, we will just pick one
	#     Identifier = row["Identifier"] 

	    # search if this part exists in SCBOM_updated, return the index in SCBOM_updated
	    # search() returned a list or None type
	    index_SCBOM_updated = search(SCBOM_updated, PN, Rev)
	    
	    #not found, deactivate the part, then append this part to SCBOM_updated
	    if (index_SCBOM_updated == None):
	        SCBOM.loc[index, ["Part Active"]] = "Inactivate"
	        SCBOM.loc[index, ["Part Status"]] = "Removed"
	        SCBOM.loc[index, ["Last Modified Date"]] = datetime.datetime.now()
	        SCBOM_updated = SCBOM_updated.append(SCBOM.loc[index], ignore_index=True)
	        

	    # found one entry, copy the information to updated SCBOM    
	    elif (len(index_SCBOM_updated) == 1):
	        index_SCBOM_updated = index_SCBOM_updated[0]
	        SCBOM_updated = copy_and_paste_row(SCBOM_updated, index_SCBOM_updated, SCBOM, index)
	        SCBOM_updated.loc[index_SCBOM_updated, ["Part Creation Date"]] = "9/21/2018"
	    
	    # found multiple entries, copy the information to each of the entries in SCBOM_updated
	    else:
	        for each in index_SCBOM_updated:
	            SCBOM_updated = copy_and_paste_row(SCBOM_updated, each, SCBOM, index)
	            SCBOM_updated.loc[index_SCBOM_updated, ["Part Creation Date"]] = "9/21/2018"	

	save(SCBOM_updated)

	print("\nAfter Sync")
	print("EBOM shape: ", EBOM.shape)
	print("SCBOM shape: ", SCBOM.shape)
	print("SCBOM_updated ", SCBOM_updated.shape)

	execution_time = time.time() - start_time
	print("\nThis script took--- {} seconds ---".format(execution_time))


if __name__ == "__main__":
	main()