import os 
import price_list_generator as plg
import MySQLdb
import pyodbc

# 'all_price_lists' will store the name of all the PL files in the current directory
all_price_lists = [] 

# Go to the input directory
current_dir = os.getcwd()
input_dir = current_dir + '\input' + '\\'

# iterate over all the PL files in the current directory and add their names to 'all_price_lists'
for i in os.listdir(input_dir):
    if i.endswith(".xlsx") and i.startswith('PL'): 
        all_price_lists += [ i ]
        continue
    else:
        continue
"""
hp_products_purchased = plg.get_cis_products_list("Mapping_Products_2.xlsx","Remove_Dup")

# Filter every HP price list in 'all_price_lists'
for pl in all_price_lists:
	print('Working on PL : ' + pl + ' ...')
	plg.generate_filtered_product_list(pl,"SHEET1",hp_products_purchased)
	print('Done with PL : ' + pl)
"""


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


