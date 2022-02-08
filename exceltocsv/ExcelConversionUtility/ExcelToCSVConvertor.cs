using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;

namespace ExcelConversionUtility
{

    public static class ExcelToCSVConvertor
    {
       
        public static List<BlobInput> Convert(List<BlobOutput> inputs)
        {
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
            WorksheetCollection worksheets = workbook.Worksheets;
            workbook.Worksheets.RemoveAt("Sheet1");
          
            string b_name = "";
            
            var dataForBlobInput = new List<BlobInput>();
            try
            {
                foreach (BlobOutput item in inputs)
                {

                    b_name = item.BlobName;
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(item.BlobContent, false))
                    {
                        foreach (Sheet _Sheet in document.WorkbookPart.Workbook.Descendants<Sheet>())
                        {
                            Aspose.Cells.Worksheet worksheet = worksheets.Add(_Sheet.Name);
                            workbook.Worksheets.RemoveAt("Evaluation Warning");
                            WorksheetPart _WorksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(_Sheet.Id);
                            DocumentFormat.OpenXml.Spreadsheet.Worksheet _Worksheet = _WorksheetPart.Worksheet;

                            SharedStringTablePart _SharedStringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                            SharedStringItem[] _SharedStringItem = _SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ToArray();
                            
                            StringBuilder stringBuilder = new StringBuilder();
                            int t = 1;
                            var tempcells = worksheet.Cells;
                            tempcells["G1"].PutValue("Status");
                            tempcells["H1"].PutValue("updated");
                            foreach (var row in _Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>())
                            {
                                if (t == 15)
                                {
                                    throw new Exception();
                                }
                                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell _Cell in row)
                                {
                                    
                                    
                                    var cells = worksheet.Cells;
                                    
                                    string Value = string.Empty;
                                    if (_Cell.CellValue != null)
                                    {
                                        if (_Cell.DataType != null && _Cell.DataType.Value == CellValues.SharedString)
                                        {
                                            Value = _SharedStringItem[int.Parse(_Cell.CellValue.Text)].InnerText;

                                            cells[_Cell.CellReference].PutValue(Value);
                                            
                                            
                                        }
                                        else
                                        {
                                            Value = _Cell.CellValue.Text;
                                            cells[_Cell.CellReference].PutValue(Value);
                                        }
                                        }
                                    stringBuilder.Append(string.Format("{0},", Value.Trim()));
                                    
                                }
                                t++;
                                string s = "G" + t;
                                string s1 = "H" + t;
                                var cells1 = worksheet.Cells;
                                cells1[s].PutValue("done");
                                cells1[s1].PutValue("0");
                                
                                stringBuilder.Append("\n");
                            }

                            byte[] data = Encoding.UTF8.GetBytes(stringBuilder.ToString().Trim());
                            string fileNameWithoutExtn = item.BlobName.ToString().Substring(0, item.BlobName.ToString().IndexOf("."));
                            string newFilename = $"{fileNameWithoutExtn}_{_Sheet.Name}.csv";
                            workbook.Worksheets.ActiveSheetIndex = 1;
                            
                            workbook.Save(item.BlobName, SaveFormat.Xlsx);
                            dataForBlobInput.Add(new BlobInput { BlobName = newFilename, BlobContent = data });
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                workbook.Save(b_name, SaveFormat.Xlsx);
                throw Ex;
            }
            return dataForBlobInput;
        }
    }
}
