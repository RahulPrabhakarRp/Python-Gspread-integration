# Manage your spreadsheets with gspread in Python.

This project was aimed at solving the problem of the Master file getting corrupted when multiple resources directly update their logs. Also an approach to semi automate the spreadsheet management process with python-gspread integration.

# Python Gspread integration

step 1--> Obtain OAuth2 credentials from Google Developers Console (http://gspread.readthedocs.io/en/latest/oauth2.html)

Step 2--> Start using gpread

         import gspread
         from oauth2client.service_account import ServiceAccountCredentials

         #Authorization and setting up scope
         scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
         creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
         client = gspread.authorize(creds)


In this project the resources will continue to maintain their own personal sheets and the data from these individual sources are to be updated to master file as the script is run.

When the script is run it collects all the data from the individual sheets and append the same in the master file.
The name of these sheets has to be declared and please note they are case sensitive

         resource_sheetnames = ['Rp','simi']

         #Loops each sheet names
         for sheetname in resource_sheetnames:
             #Opens source sheet 
             worksheet = client.open(sheetname).sheet1    
             #Opens sheet to which data have to be updated
             final_worksheet = client.open("Consolidated Report").sheet1      

             #Rows which have to be updated marked as FALSE
             update_status = "FALSE"   
             #Returns list of column ID with value FALSE
             cell_list = worksheet.findall(update_status)    
             #To find number of rows in consolidated sheet in order to append as last row
             finalCell_list = final_worksheet.findall(update_status)   
             #Number of rows with FALSE status + header row + 1 = next row
             count = len(finalCell_list) + 2
    
    
Since this will be a daily repetitive process, there is a need to identify the recently updated data in the individual sheet to avoid duplication of data while appending the master file, hence used a flag 'FALSE' in order to identify the data to be updated from the individual sheets. As the master file is updated the corresponding rows are tagged 'TRUE' in the individual sheets and will be skipped while updating the data next time.

          #Looping source sheet with update status False
             for cell in cell_list:
        
                 #Finds data from the row using the row ID of cell
                 values_list = worksheet.row_values(cell.row)
                 row = values_list
                 index = count
                 #Inserts row in consolidated sheet according to the index specified
                 final_worksheet.insert_row(row,index)
                 count = count + 1
                 #Updates source sheet status column as True
                 worksheet.update_cell(cell.row,cell.col, 'True')
