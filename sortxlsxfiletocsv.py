
import os
import pandas as pd
import openpyxl
import sqlite3
import csv


# giving directory name
path_of_excelfile =".\\input"
intermediate_path=".\\bufferinput"
path_of_csv_folder = ".\\output"
files =os.listdir(path_of_excelfile)
#print(files)

for eachfile in files:
 if eachfile.endswith(".xlsx"):
    name_without_xlsx_extension=eachfile.replace(".xlsx","")
    #print(name_without_xlsx_extension)
    read_file=pd.read_excel(os.path.join(path_of_excelfile,eachfile))
    read_file.to_csv(os.path.join(intermediate_path,str(name_without_xlsx_extension)+""+str('.csv')), index = False, header=True)
    #print("converted   " +eachfile +"  to  " + name_without_xlsx_extension +".csv!")



#Helps to import csv file to sqlite3 database
       
#step1:load csv file 
csvfiles =os.listdir(intermediate_path)
for eachcsvfile in csvfiles :
  if eachcsvfile.endswith(".csv"):
    name_without_csv_extension=eachcsvfile.replace(".csv","")
    # reading data from the CSV file
    # print(eachcsvfile)
    f= pd.read_csv(os.path.join(intermediate_path,eachcsvfile))

    #Step2:Data cleanup
    f.columns=f.columns.str.strip()

    #step 3:to create connection with existing database file or create database file
    connection=sqlite3.connect(".\\database_for_shortingxlsxfiletocsv.db")

    #step 4:load eachcsvfile to sqlite3
    f.to_sql(name_without_csv_extension,connection,if_exists='replace',index=False)

    #step 5:to save work done to database
    #connection.commit

    #step 6:to show in output folder 
    #sql_select_command='''SELECT * FROM {}'''.format(name_without_csv_extension)
    
    sql_select_command="SELECT * FROM "+name_without_csv_extension+" ORDER BY Sheet_No,PARCELNO ASC"

    # step 7:to create dataFrame or csv file of table
    sqlquery=pd.read_sql_query(sql_select_command,connection)
    dall=pd.DataFrame(sqlquery)
    dall.to_csv(os.path.join(path_of_csv_folder,str(name_without_csv_extension) + '_' + str('.csv')),index=False,header=True)
    
    #step 8:to save work done to database
     
    connection.commit

    #step 8:close connection
    
    connection.close



