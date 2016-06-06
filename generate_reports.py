import os 
import price_list_generator as plg
import MySQLdb
import pyodbc

# 'HPE_price_lists' will store the name of all the PL files in the HPE input directory
HPE_price_lists = [] 
# 'HPI_price_lists' will store the name of all the PL files in the HPI input directory
HPI_price_lists = [] 

# Go to the input directory
current_dir = os.getcwd()
HPE_input_dir = current_dir + '\input\HPE' + '\\'
HPI_input_dir = current_dir + '\input\HPI' + '\\'

# iterate over all the PL files in the HPE directory and add their names to 'HPE_price_lists'
for i in os.listdir(HPE_input_dir):
    if i.endswith(".xlsx") and i.startswith('PL'): 
        HPE_price_lists += [ i ]

# iterate over all the PL files in the HPI directory and add their names to 'HPI_price_lists'
for i in os.listdir(HPI_input_dir):
    if i.endswith(".xlsx") and i.startswith('PL'): 
        HPI_price_lists += [ i ]

hp_products_purchased = plg.get_cis_products_list("Mapping_Products_2.xlsx","Remove_Dup")

print( 'WORKING ON HPE PRICE LISTS')
# Filter every HPE price list in 'HPE_price_lists'
for pl in HPE_price_lists:
	print('Working on PL : ' + pl + ' ...')
	plg.generate_filtered_product_list(pl,'HPE',"SHEET1",hp_products_purchased)
	print('Done with PL : ' + pl)
print( 'DONE WITH HPE PRICE LISTS')

print( 'WORKING ON HPI PRICE LISTS')
# Filter every HPI price list in 'HPI_price_lists'
for pl in HPI_price_lists:
	print('Working on PL : ' + pl + ' ...')
	plg.generate_filtered_product_list(pl,'HPI',"SHEET1",hp_products_purchased)
	print('Done with PL : ' + pl)
print( 'DONE WITH HPI PRICE LISTS')

# Get all products purchased
#hp_products_purchased = plg.get_cis_products_list("Mapping_Products_2.xlsx","Remove_Dup")
#plg.generate_filtered_product_list(all_price_lists[0],"SHEET1",hp_products_purchased)

#print( plg.removeNonAscii("Numero de produit") )
#print( all_price_lists[0] )

"""
# Testing SQL Locally
cnx = MySQLdb.Connect(host="127.0.0.1", port=3306, user="root", passwd="root")
cur = cnx.cursor()
cur.execute("USE pricel")
# Use all the SQL you like
#cur.execute("SELECT * FROM pricel")
cur.execute("SELECT COUNT(*) FROM pricel")

# print all the first cell of all the rows
for row in cur.fetchall():
    print row
cnx.close()
"""


