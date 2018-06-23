import gspread
from oauth2client.service_account import ServiceAccountCredentia

print('*****Started merging*****')

#Authorization and setting up scope
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

#Name of the sheets shared with client_email (source sheets). Please note sheet names are case sensitive
resource_sheetnames = ['Rp','simi']

#Counter
count = 0

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

print('*****Finished merging*****')



    
    
    
    
      
