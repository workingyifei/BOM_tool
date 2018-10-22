# The script below syncs the supply Chain BOM with Manufacturing BOM.
# Created by Yifei.li@byton.com

import pandas as pd
import numpy as np
import xlrd, openpyxl
import csv, datetime
import time

#load MBOM and SCBOM
def load():
	# load Manufacturing BOM
	# MBOM = pd.read_csv(MBOM_PATH)
	MBOM = pd.read_excel(MBOM_PATH, sheet_name="BOM")
	MBOM.columns = MBOM.columns.str.replace("\(R\)\ ", "") # trim (R) away from the header

	# Rename N Intelligent Car Experience ICE
	MBOM.loc[MBOM.System=="N Intelligent Car Experience ICE", "System"]= "N ICE"

	# change Byton Part Number column to Byton PN
	MBOM = MBOM.rename(index=str, columns={"Title": "Byton PN", })

	# change identifier type to string
	MBOM["Identifier"] = MBOM["Identifier"].apply(str)

	MBOM = MBOM.reset_index(drop=True)

	#     # load Supply Chain BOM from MULTIPLE tabs in an Excel
	#     df= pd.read_excel(SCBOM_PATH, sheet_name=system_name)
	#     SCBOM = pd.DataFrame()
	#     for each in df:
	#         SCBOM = SCBOM.append(df[each])

	# load Supply Chain BOM from SINGLE TAB in an Excel
	SCBOM= pd.read_excel(SCBOM_PATH, sheet_name="Supply Chain BOM")

	# change column N Intelligent Car Experience ICE to just N ICE
	# pd.to_excel doesn't allow column name greater than 31 characters
	SCBOM.loc[SCBOM.System=="N Intelligent Car Experience ICE", "System"]= "N ICE"

	# change Byton Part Number column to Byton PN
	SCBOM = SCBOM.rename(index=str, columns={"Title": "Byton PN", })

	# reset index of SCBOM
	SCBOM.reset_index(drop=True, inplace=True)
	print("\nCompleted loading excel files...\n")
	return MBOM, SCBOM

# search PN and/or Rev in updated SCBOM
def search(df, PN, Rev):
	# df = df.loc[(df["Title"]==PN) & (df["Revision"]==Rev)]
	# CHANGE THIS IF MBOM COLUMN NAME CHANGED!!!
	df = df.loc[(df["Byton PN"]==PN)]	# search PN only 
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
        
# copy SCBOM info and paste into SCBOM_updated
def copy_and_paste_row(df1, index1, df2, index2):  # index1 and index2 are int
	# copy df2 info into df1
	# only copy columns that are not on MBOM
	df1.loc[index1, df1.columns.tolist()[SCBOM_updated_columns_start:SCBOM_updated_columns_end]] = df2.loc[index2, df2.columns.tolist()[SCBOM_columns_start:SCBOM_columns_end]]
	return df1

# save dataframe to Excel
def save(df):
	#https://stackoverflow.com/questions/28837057/pandas-writing-an-excel-file-containing-unicode-illegalcharactererror
	df = df.applymap(lambda x: x.encode('unicode_escape').
	             decode('utf-8') if isinstance(x, str) else x)
	date = str(datetime.date.today())
	name = "Supply Chain BOM_" + date + ".xlsx"
	writer = pd.ExcelWriter(name)
	print("\n\nStart saving...")

	# saving to multiple tabs
	for each in system_name:
		system = df[df["System"]==each]
		print("\rsaving: {}".format(each), end="", flush=True)
		system.to_excel(writer, sheet_name=each, na_rep="")
	
	print("Saving complete...\n")
	# # saving to ONE tab
	# df.to_excel(writer, sheet_name="Updated Supply Chain BOM", na_rep="" )

	writer.save()

#     df.to_csv("Supply Chain BOM.csv",quoting=csv.QUOTE_NONE, escapechar="\\")

def main(MBOM, SCBOM, SCBOM_updated):

	SCBOM_columns = SCBOM.columns.tolist()
	SCBOM_columns_size = len(SCBOM_columns)
	MBOM_columns = MBOM.columns.tolist()
	MBOM_columns_size = len(MBOM_columns)

	print("Before Sync")
	print("MBOM shape: ", MBOM.shape)
	print("SCBOM shape: ", SCBOM.shape)
	print("SCBOM_updated: {}\n".format(SCBOM_updated.shape))


	# count how many new parts are added
	# how many old parts are removed
	removed_parts_count = 0
	same_parts = 0


	# loop through SCBOM 
	for index, row in SCBOM.iterrows():
		print("\rUpdating row # {0} of Supply Chain BOM. Progress: {1}%".format(index, round(index*100/SCBOM.shape[0]), 4), end="", flush=True)
		PN = row["Byton PN"]
		Rev = row["Revision"]
		# same part could be structured differently but the part is the same, 
		# so if same part found at multiple places, we will just pick one
		# Identifier = row["Identifier"] 

	    # search if this part exists in SCBOM_updated, return the index in SCBOM_updated
	    # search() returned a list or None type
		index_SCBOM_updated = search(SCBOM_updated, PN, Rev)
	    
	    #not found, deactivate the part, then append this part to SCBOM_updated
		if (index_SCBOM_updated == None):
			SCBOM.loc[index, ["Part Active"]] = "Inactivate"
			SCBOM.loc[index, ["Part Status"]] = "Removed"
			# SCBOM.loc[index, ["Part Creation Date"]] = datetime.date(2019, 9, 14)
			SCBOM.loc[index, ["Last Modified Date"]] = datetime.date.today()
			# SCBOM.loc[index, ["Last Modified Date"]] = datetime.date(2019, 9, 28)
			columns_to_copy = SCBOM.columns.tolist()[SCBOM_columns_start:SCBOM_columns_end]
			other_columns = ["Byton PN", "Revision", "Description", "System", "SubSystem", "Part Type"]
			columns_to_copy = columns_to_copy + other_columns
			SCBOM_updated = SCBOM_updated.append(SCBOM.loc[index, columns_to_copy], ignore_index=True)

			removed_parts_count = removed_parts_count + 1

	   	# found one entry, copy the information to updated SCBOM    
		elif (len(index_SCBOM_updated) == 1):
			index_SCBOM_updated = index_SCBOM_updated[0]
			SCBOM_updated = copy_and_paste_row(SCBOM_updated, index_SCBOM_updated, SCBOM, index)
			SCBOM_updated.loc[index_SCBOM_updated, ["Part Creation Date"]] = datetime.date(2019, 9, 21)
			# SCBOM_updated.loc[index_SCBOM_updated, ["Last Modified Date"]] = datetime.date.today()
			SCBOM_updated.loc[index_SCBOM_updated, ["Part Status"]] = "Lateste Revision"
			SCBOM_updated.loc[index_SCBOM_updated, ["Part Active"]] = "Active"
			same_parts = same_parts + 1

		# found multiple entries, copy the information to each of the entries in SCBOM_updated
		else:
			for each in index_SCBOM_updated:
				SCBOM_updated = copy_and_paste_row(SCBOM_updated, each, SCBOM, index)
				SCBOM_updated.loc[index_SCBOM_updated, ["Part Creation Date"]] = datetime.date(2019, 9, 21)
				# SCBOM_updated.loc[index_SCBOM_updated, ["Last Modified Date"]] = datetime.date.today()
				SCBOM_updated.loc[index_SCBOM_updated, ["Part Status"]] = "Lateste Revision"
				SCBOM_updated.loc[index_SCBOM_updated, ["Part Active"]] = "Active"
			same_parts = same_parts + 1

	save(SCBOM_updated)

	print("After Sync")
	print("MBOM shape: ", MBOM.shape)
	print("SCBOM shape: ", SCBOM.shape)
	print("SCBOM_updated ", SCBOM_updated.shape)


	print("\n# of new parts added: ", MBOM.shape[0]-same_parts)
	print("# of old parts removed: ", removed_parts_count)
	print("# of additional columns added in MBOM: ", SCBOM_updated.shape[1] - SCBOM.shape[1])


if __name__ == "__main__":
	start_time = time.time()

	# put up-to-date MBOM and SCBOM here for sync
	MBOM_PATH = './Data/AAZZ000001NN03_VP_BoM_10-8-18 at 11.10 AM_MFG_Release_v1.0.xlsx'
	SCBOM_PATH = './Data/Supply Chain BOM_2018-10-03_final.xlsx'

	MBOM, SCBOM = load()

	# indexes of the columns we are copying from and pasting into
	# 0 index
	SCBOM_columns_start = 14 # SC columns start at column index 14, "Engineer"
	SCBOM_columns_end = SCBOM.shape[1] # SC columns end at column index 77
	SCBOM_updated_columns_start = MBOM.shape[1] # 21 columns
	SCBOM_updated_columns_end = SCBOM_updated_columns_start + (SCBOM_columns_end - SCBOM_columns_start) + 1 # 21 + 77-14+1 = 85


	# tab names in supply chain BOM.xlsx and MBOM.xlsx
	system_name = ["A BIW", "B Closures", "C Exterior", "D Interior", "E Chassis", "F Thermal Management", "G Drivetrain",
				   "H Power Electronics", "J HV Battery", "K Autonomy", "L Low Voltage Systems", "M Connectivity", 
	          	   "N ICE", "X Raw Materials", "Y Fasteners", "Z Vehicle Top Level Cfg"]

	

	# # create Parent and Level column in MBOM
	# MBOM.insert(0, column="Level", value="")
	# MBOM.insert(1, column="Parent", value="")

	# for index, row in MBOM.iterrows():
	# 	identifier = str(row["Identifier"]) 
	# 	level = identifier.count("|")
	# 	MBOM.loc[index, ["Level"]] = level

	# 	# find out PN of part's parent
	# 	identifier_parent = identifier[0:-2]
	# 	if identifier_parent == '':
	# 		identifier_parent = ''
	# 	elif identifier_parent[-1]=='|':
	# 		identifier_parent = identifier_parent[0:-1]
	# 	else:
	# 		identifier_parent = identifier_parent

	# 	PN_parent = MBOM[MBOM["Identifier"]==identifier_parent]["Byton PN"]
	# 	if PN_parent.empty == True:
	# 	    PN_parent = ""
	# 	else:
	# 	    PN_parent = PN_parent.values
	# 	MBOM.loc[index, ["Parent"]] = PN_parent
	# print("Parent PN and Level are generated in CAD BOM...\n")

	# create a updated SCBOM with columns from SCBOM and data from MBOM
	SCBOM_updated = MBOM.copy()
	# copy only colums that do not exist in MBOM
	for each in SCBOM.columns.tolist()[SCBOM_columns_start:]:
		SCBOM_updated[each] = ""


	main(MBOM, SCBOM, SCBOM_updated)

	execution_time = round(time.time() - start_time, 2)
	print("\nThis script took--- {} seconds ---".format(execution_time))
