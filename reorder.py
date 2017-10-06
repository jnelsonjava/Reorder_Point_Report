

import os
import pandas as pd
import sys

homey = os.path.abspath(os.path.dirname(__file__))
sqlPath = os.path.join(homey, 'SQL')

sys.path.insert(0, 'Z:\Python projects\FishbowlAPITestProject')
import connecttest


# Using the inventory query from FB_Simulator
myresults = connecttest.create_connection(sqlPath, 'INVQuery.txt')
myexcel = connecttest.makeexcelsheet(myresults)
connecttest.save_workbook(myexcel, sqlPath, 'INVs.xlsx')

# Quick query of reorder point table
myresults = connecttest.create_connection(sqlPath, 'ReorderPointQuery.txt')
myexcel = connecttest.makeexcelsheet(myresults)
connecttest.save_workbook(myexcel, sqlPath, 'ReorderPoint.xlsx')

# read the xlsx files into dataFrames
invPath = os.path.join(sqlPath, 'INVs.xlsx')
inv = pd.read_excel(invPath, header=0)

reorderPath = os.path.join(sqlPath, 'ReorderPoint.xlsx')
reorder = pd.read_excel(reorderPath, header=0)

# put the inventory on the reorder points list
reorderInv = pd.merge(reorder.copy(), inv.copy(), how='left', on='PART')
# anything with no inventory will show up as NAN, so replace it with 0
reorderInv.fillna(0, inplace=True)
# filter to only lines where the inventory is less than or equal to the reorder point
reorderInv = reorderInv[reorderInv['INV'] <= reorderInv['Reorder Point']].copy()
# clean up the column order
reorderINV = reorderInv[['PART', 'Part Description', 'INV', 'Reorder Point', 'Order Up To Level']]
# sort by part
reorderINV.sort_values(by='PART', ascending=True, inplace=True)

# choose the excel filename
reorderFilename = 'Current Reorder Parts.xlsx'

# save to excel
writer = pd.ExcelWriter(os.path.join(homey, reorderFilename))
reorderINV.to_excel(writer, 'Sheet', index=False)
writer.save()



# import email_tool

# reorderRecipientList = ['tvay.com','jnelson@commnetsystems.com']

# email_tool.send_email(reorderRecipientList, reorderFilename)
