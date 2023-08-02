UOfISpace.UOfIMain.ProcessExcel();

Console.WriteLine("");
Console.WriteLine("");
Console.WriteLine("Press any key to exit... ");
Console.ReadLine();

namespace UOfISpace
{
    public class UOfIMain
    {
        public static void ProcessExcel()
        {
            // go over the list of the excel files and get the sub receipent data
            List<SubrecipientData> lstSubDataFull = new List<SubrecipientData>();
            string path = @"C:\\Projects\\UIUC\\\UOfIExcel\excels";
            DirectoryInfo dir = new DirectoryInfo(path);
            Console.WriteLine("");
            Console.WriteLine("Input Subward Excel Files Information:");
            Console.WriteLine("------------------------------------------------------------------");
            Console.WriteLine("File Name                      | Size       | Creation Date & Time");
            Console.WriteLine("------------------------------------------------------------------");
            foreach (FileInfo fileInfo in dir.GetFiles())
            {
                String fileName = fileInfo.Name;

                if (fileName.StartsWith("Subaward") && fileName.EndsWith(".xlsx"))
                {
                    long fileSize = fileInfo.Length;
                    DateTime creationTime = fileInfo.CreationTime;
                    Console.WriteLine("{0, -32:g} {1,-12:N0} {2} ", fileName, fileSize, creationTime);
                    Subrecipient subReceip = new Subrecipient();
                    var lstSubData = subReceip.getSubrecipientDataFromFile(fileName);

                    if (lstSubData.Count > 0) lstSubDataFull.AddRange(lstSubData);
                }

            }


            // after processing all the excel files from the directory, find unique locations and get the aggregated total
            if (lstSubDataFull.Count > 0)
            {
                Console.WriteLine("");
                Console.WriteLine("------------");
                Console.WriteLine("Full Data:");
                Console.WriteLine("------------");
                Console.WriteLine("{0,-32:g} {1,-15:g} {2} ", "Excel File Name", "Location", "Amount");
                Console.WriteLine("-------------------------------------------------------");
                foreach (SubrecipientData subData in lstSubDataFull)
                {
                    Console.WriteLine("{0,-32:g} {1,-15:g} {2} ", subData.ExcelFileName, subData.SubLocation, subData.SubAmount);
                }

                // sort the data by location 
                lstSubDataFull = lstSubDataFull.OrderBy(o => o.SubLocation).ToList();

                // find distinct locations
                var distinctSubDataLocation = lstSubDataFull.Select(std => std.SubLocation).Distinct().ToList();

                Console.WriteLine("");
                Console.WriteLine("------------");
                Console.WriteLine("Final Aggregated Data by Location:");
                Console.WriteLine("------------");
                Console.WriteLine("{0,-15:g} {1} ", "Location", "Amount");
                Console.WriteLine("--------------------------");
                foreach (string location in distinctSubDataLocation)
                {
                    var newList = lstSubDataFull.Where(c => c.SubLocation.Equals(location)).ToList();

                    if (newList.Count > 0)
                    {
                        Decimal totalAmount = 0;
                        string locationName = "";
                        foreach (SubrecipientData subData in newList)
                        {
                            totalAmount += Convert.ToDecimal(subData.SubAmount);
                            locationName = subData.SubLocation;
                        }
                        Console.WriteLine("{0,-15:g} {1} ", locationName, totalAmount);
                    }

                }

            }
        } // end of ProcessExcel

    }


}




