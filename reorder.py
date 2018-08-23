

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

# Pulling PO and MO "On Order" quantities
myresults = connecttest.create_connection(sqlPath, 'OnOrderQuery.txt')
myexcel = connecttest.makeexcelsheet(myresults)
connecttest.save_workbook(myexcel, sqlPath, 'OnOrder.xlsx')

# read the xlsx files into dataFrames
invPath = os.path.join(sqlPath, 'INVs.xlsx')
inv = pd.read_excel(invPath, header=0)

reorderPath = os.path.join(sqlPath, 'ReorderPoint.xlsx')
reorder = pd.read_excel(reorderPath, header=0)

onOrderPath = os.path.join(sqlPath, 'OnOrder.xlsx')
onOrder = pd.read_excel(onOrderPath, header=0)

# put the inventory on the reorder points list
reorderInv = pd.merge(reorder.copy(), inv.copy(), how='left', on='PART')
# anything with no inventory will show up as NAN, so replace it with 0
reorderInv.fillna(0, inplace=True)
# filter to only lines where the inventory is less than or equal to the reorder point
reorderInv = reorderInv[reorderInv['INV'] <= reorderInv['Reorder Point']].copy()
# add the current amount on order between POs and MOs
reorderInv = pd.merge(reorderInv.copy(), onOrder.copy(), how='left', on='PART')
# clean up the column order
reorderInv = reorderInv[['PART', 'Part Description', 'INV', 'Reorder Point', 'Order Up To Level', 'On Order', 'Make/Buy']]
# sort by part
reorderInv.sort_values(by='PART', ascending=True, inplace=True)
# split into make/buy
reorderBuy = reorderInv[reorderInv['Make/Buy'] == 'Buy'].copy()
reorderMake = reorderInv[reorderInv['Make/Buy'] == 'Make'].copy()


# choose the excel filename
reorderFilename = 'Current Reorder Parts.xlsx'

# save to excel
writer = pd.ExcelWriter(os.path.join(homey, reorderFilename))
reorderBuy.to_excel(writer, 'Buy Parts', index=False)
reorderMake.to_excel(writer, 'Make Parts', index=False)
writer.save()



import email_tool

reorderRecipientList = ['jnelson@commnetsystems.com','vbratcher@commnetsystems.com'] # for testing purposes
# reorderRecipientList = ['tvay@commnetsystems.com','jnelson@commnetsystems.com','mfreeling@commnetsystems.com','jmayhle@commnetsystems.com']

email_tool.send_email(reorderRecipientList, reorderFilename)

print('Reorder report done!')