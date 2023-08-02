using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace UOfISpace
{

    public class SubrecipientData
    {
        public string ExcelFileName { get; set; }
        public string SubLocation { get; set; }
        public string SubAmount { get; set; }
    }


    public class Subrecipient
    {

        public List<SubrecipientData> getSubrecipientDataFromFile(string excelFileName)
        {

            string path = @"C:\\Projects\\UIUC\\\UOfIExcel\excels";
            string fullFileLoc = path + "\\" + excelFileName;
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(@fullFileLoc);
            Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            List<SubrecipientData> lstSubData = new List<SubrecipientData>();

            try
            {

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                object[,] valueArray = (object[,])xlRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].value2 != null)
                        {
                            var cellVal = xlRange.Cells[i, j].value2?.ToString().Trim() ?? "";

                            if (cellVal.StartsWith("Subaward:"))
                            {
                                var dLoc = xlRange.Cells[i, 3].value2?.ToString() ?? "";
                                // check if the previos cell have the location name in it
                                if (String.IsNullOrEmpty(dLoc))
                                {
                                    string[] cellValArr = cellVal.Trim().Split(":");
                                    if (cellValArr.Length > 1)
                                    {
                                        dLoc = cellValArr[1].Trim();
                                    }
                                }
                                if (String.IsNullOrEmpty(dLoc)) dLoc = "Null";

                                var dAmt = xlRange.Cells[i, 5].value2?.ToString().Trim() ?? "";

                                lstSubData.Add(new SubrecipientData
                                {
                                    ExcelFileName = excelFileName,
                                    SubLocation = dLoc,
                                    SubAmount = dAmt
                                });

                            }
                        }

                    }
                }

            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            return lstSubData;
        } // end of getSubrecipientDataFromFile

    }

}

