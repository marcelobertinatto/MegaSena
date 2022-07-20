using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MegaSena
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var random = new Random();
            StringBuilder excelResult = new StringBuilder();
            List<int> finalList = new List<int>();
            var list = ReadExcelFile(@"C:\temp\MegaSenaResultados.xlsx", "C");
            var duplicate = list.GroupBy(x => x)
              .Where(g => g.Count() > 1)
              .Select(y => y.Key)
              .ToList();

            var query = list.GroupBy(x => x)
              .Where(g => g.Count() > 1)
              .ToDictionary(x => x.Key, y => y.Count());

            var orderedDuplicateListByAsc = query.OrderBy(x => x.Key);
            var orderedDuplicateListByDesc = query.OrderByDescending(x => x.Value);
            for (int i = 0; i < 6; i++)
            {
                int index = random.Next(orderedDuplicateListByDesc.Select(x => x.Key).Count());
                var item = orderedDuplicateListByDesc.Select(x => x.Key).ToList()[index];
                if (!finalList.Contains(item))
                {
                    finalList.Add(orderedDuplicateListByDesc.Select(x => x.Key).ToList()[index]);
                }
                else
                {
                    i = i - 1;
                }
                //excelResult.Append(orderedDuplicateListByDesc.Select(x => x.Key).ToList()[index].ToString() + " ");
            }

            excelResult.AppendLine("Fézinha ");
            excelResult.AppendLine("----------------------------------------------- ");

            foreach (var item in finalList.OrderBy(x => x))
            {
                excelResult.Append(item + " ");
            }
            
            Console.WriteLine(excelResult.ToString());
            Console.ReadLine();

        }

        static List<int> ReadExcelFile(string filePath, string columnToBeSearched)
        {
            var listMegaSenaResults = new List<int>();
            try
            {
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    StringBuilder excelResult = new StringBuilder();

                    //using for each loop to get the sheet from the sheetcollection  
                    foreach (Sheet thesheet in thesheetcollection)
                    {
                        //excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                        //excelResult.AppendLine("----------------------------------------------- ");
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        foreach (Row thecurrentrow in thesheetdata)
                        {
                            foreach (Cell thecurrentcell in thecurrentrow)
                            {
                                if (thecurrentcell.CellReference.Value.StartsWith(columnToBeSearched))
                                {
                                    //statement to take the integer value  
                                    string currentcellvalue = string.Empty;
                                    if (thecurrentcell.DataType != null)
                                    {
                                        if (thecurrentcell.DataType == CellValues.SharedString)
                                        {
                                            int id;
                                            if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                            {
                                                SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                                if (item.Text != null)
                                                {
                                                    //code to take the string value  
                                                    excelResult.Append(item.Text.Text + " ");
                                                    //string[] textSplit = item.Text.Text.Trim().Split(" ");
                                                    listMegaSenaResults.AddRange(item.Text.Text.Trim().Split(" ").Select(int.Parse).ToList());
                                                }
                                                else if (item.InnerText != null)
                                                {
                                                    currentcellvalue = item.InnerText;
                                                }
                                                else if (item.InnerXml != null)
                                                {
                                                    currentcellvalue = item.InnerXml;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + " ");
                                    } 
                                }
                            }
                            //excelResult.AppendLine();
                        }
                        //excelResult.Append("");
                        //Console.WriteLine(excelResult.ToString());
                        //Console.ReadLine();
                    }
                }
            }
            catch (Exception)
            {

            }
            return listMegaSenaResults;
        }

    }
}
