# UOfIExcel

This project is developed using VS2022 and .NET6. 

Excel Files Location:
By default, this project expects all the input files in the location "C:\\Projects\\UIUC\\UOfIExcel\\excels\". 
If you want to change the input location of the excel files, please update the location of the files in the code and the application can work seamless. 
Files needs to be changed: Subrecipient.cs, UOfIMain.cs

Goal of the Project:
- Reads all spreadsheets from a folder (use the 3 attached Excel spreadsheets). 
- For each file, output to the console the file name followed by each subrecipient name from that file. The subrecipient names will be under “G. Other Direct Costs” in the format
 “Subaward: {SubRecipientName}”      
- Finally, output a distinct list of all subrecipients along with the total subaward amount that subrecipient received across all files. 
