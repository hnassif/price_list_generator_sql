# -*- coding: utf-8 -*- 
import sys
import codecs
import pandas as pd
import MySQLdb
from sqlalchemy.types import NVARCHAR
import os 

sys.stdout = codecs.getwriter('utf8')(sys.stdout)


def get_cis_products_list(xl_file,sheetname):

    # Read and Parse the 'Remove_Dup' sheet in the Mapping Produts Excel File
    xl = pd.ExcelFile(xl_file)
    df = xl.parse(sheetname)

    # Extract the 'HP_Reference_ListP' from the file
    hp_reference_list = df['HP_Reference_ListP']

    # Remove empty rows (This is to know where the data ends)
    criterion = df[ pd.notnull(df['HP_Reference_ListP']) ] 

    # Display the Size of the data 
    print( 'CIS Product List Dataframe Size : ' + str(criterion.size) )
    print( 'CIS Product List Dataframe Shape : ' + str(criterion.shape) )

    #The number of rows corresponds to the number of unique products 
    num_of_rows = criterion.shape[0] 
    print( 'Number of Products : ' + str(num_of_rows-1)) # -1 because of the first row, which has the column names

    # Index the DataFrame. ie: assign numbers to rows 
    df.set_index([range(df.shape[0])], inplace = True)

    # generate a list of unique products
    hp_products_purchased = sorted(list(set(criterion['HP_Reference_ListP'])))

    # create a new list for products 
    prefix = tuple()
    filtered_hp_products_list = [] 
    
    # 
    for i in range(len(hp_products_purchased)): 
        if not hp_products_purchased[i].startswith(prefix):
            filtered_hp_products_list.append(hp_products_purchased[i])
            prefix += ( hp_products_purchased[i] , )

    print( 'Number of Unique Prefixes : ' + str(len(filtered_hp_products_list))) 

    return filtered_hp_products_list


# Function to remove Non-ASCII characters. This function is used in the 'generate_filtered_product_list' function.
def removeNonAscii(s): 
	#return "".join(map(lambda x: x if ord(x)<128 else 'e', s))
	s = s.lower()
	s_final = ""
	for c in s:
		if ord(c)>96 and ord(c)<128: 
			s_final += c 
		elif c == ' ' :
			s_final += '_'
		elif ord(c) == 39: #apostrophe 
			continue
		else: 
			s_final += 'e'

	#s_final = "".join(map(lambda x: x if ord(x)<128 else 'e', s))
	#return s_final.replace( ' ' , '_').lower()
	return s_final

 
def generate_filtered_product_list(hp_price_list,HPE_or_HPI,sheetname,hp_products_purchased):

    try: 

        # Go to the input directory
        input_dir = os.getcwd() + '\input' + '\\' + HPE_or_HPI + '\\'

        # Read and Parse the 'sheet1' sheet in the Mapping Produts Excel File
        xl_pl = pd.ExcelFile( input_dir + hp_price_list)
        df_pl = xl_pl.parse(sheetname)

        # Converting columns to ASCII characters, because HP uses some non-ascii characters that the Panda library cannot handle
        columns = df_pl.columns
        df_pl.columns = map( removeNonAscii , list(columns) )
        print( 'Columns are : ' + str(df_pl.columns))

        #Identify the currency
        #if 'numero_de_produit' in df_pl.columns:
        #col_name = 'numero_de_produit'
        #else: 
        col_name = 'product_number'

        # Remove empty rows (This is to know where the data ends)
        criterion_pl = df_pl[ pd.notnull(df_pl[col_name]) ] 

        # Display the Size of the data 
        print( 'HP Product List Size : ' + str(criterion_pl.size) )
        print( 'HP Product List Shape : ' + str(criterion_pl.shape) )

        #The number of rows corresponds to the number of HP Prodcuts
        num_of_rows_pl = criterion_pl.shape[0]
        print('Number of Rows in HP Price List : ' + str(num_of_rows_pl) )

        # 'frames' will store the results of each product query performed
        frames = []

        # For each hp_product purchased by CIS, search for any HP product that starts with the same number 
        for hp_product in hp_products_purchased:
            query = df_pl[col_name].map(lambda x: x.startswith(hp_product) if type(x) != float else False)
            if not df_pl[query].empty: 
                frames += [ df_pl[query] ]

        # Combining all query resuts into a single DataFrame  
        if frames: 
            output_df = pd.concat(frames)
        else:
            output_df = pd.DataFrame( columns=df_pl.columns)
            output_df = output_df.fillna(0) # with 0s rather than NaNs
            print('No Products have been purchased')
            #return 

        print( 'The filtered HP products list has Size : ' + str(output_df.size) )
        print( 'The filtered HP products list has Shape : ' + str(output_df.shape) )

        # Connecting to the SQL Database
        #cnx = MySQLdb.Connect(host="127.0.0.1", port=3306, user="root", passwd="root")
        #cnx = MySQLdb.Connect(host="localhost", port=3306, user="User1", passwd="h@nassif8") # 1433

        # create cursor to perform operations/queries 
        #cur = cnx.cursor() 

        #cur.execute("CREATE DATABASE pricel") # only needed when creating the database for the first time 
        #cur.execute("USE pricel")
        #cur.execute("SET sql_mode=(SELECT REPLACE(@@sql_mode,'STRICT_TRANS_TABLES',''))")

        # Remove the 'description_longue' column because the content is too big to write to SQL
        #if col_name == 'numero_de_produit': 
        #output_df = output_df.drop(['description_longue'], axis=1)
        #else: 
        output_df = output_df.drop(['long_description_'], axis=1)

        #print( output_df )
        #print( map(lambda x: x.dtype , output_df) )

        # Write Dataframe to Database 
        #output_df.to_sql(name='pricel', con=cnx, flavor='mysql', schema=output_df.columns, if_exists='append', index=True)

        # Delete the cursor
        #del cur 

        # Close the connection
        #cnx.close()

        #Writing the Output to an excel file 
        #output_dir = os.getcwd() + '\output' + '\\'
        writer = pd.ExcelWriter('output_' + hp_price_list  + '.xlsx' ) 
        output_df.to_excel(writer,'Sheet1', index=False)
        writer.save()

    except: 

        print("Error With File : " + hp_price_list )
        print("Error is : " , sys.exc_info())


    return 










