using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;

namespace ExcelConversionUtility
{
    /// <summary>
    /// This class is responsible for converting Excel to CSV format
    /// </summary>
    /// 
    public static class ExcelToCSVConvertor
    {
        /// <summary>
        /// Converts Excel to CSV
        /// </summary>
        /// <param name="input">Key is excel filename and value is file content.</param>
        /// <returns></returns>
        public static List<BlobInput> Convert(List<BlobOutput> inputs)
        {
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
            WorksheetCollection worksheets = workbook.Worksheets;


            var dataForBlobInput = new List<BlobInput>();
            try
            {
                foreach (BlobOutput item in inputs)
                {
                    

                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(item.BlobContent, false))
                    {
                        foreach (Sheet _Sheet in document.WorkbookPart.Workbook.Descendants<Sheet>())
                        {
                            Aspose.Cells.Worksheet worksheet = worksheets.Add(_Sheet.Name);
                            WorksheetPart _WorksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(_Sheet.Id);
                            DocumentFormat.OpenXml.Spreadsheet.Worksheet _Worksheet = _WorksheetPart.Worksheet;

                            SharedStringTablePart _SharedStringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                            SharedStringItem[] _SharedStringItem = _SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ToArray();

                            StringBuilder stringBuilder = new StringBuilder();
                            foreach (var row in _Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>())
                            {
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
                                stringBuilder.Append("\n");
                            }

                            byte[] data = Encoding.UTF8.GetBytes(stringBuilder.ToString().Trim());
                            string fileNameWithoutExtn = item.BlobName.ToString().Substring(0, item.BlobName.ToString().IndexOf("."));
                            string newFilename = $"{fileNameWithoutExtn}_{_Sheet.Name}.csv";
                            workbook.Save("ou.xlsx", SaveFormat.Xlsx);
                            dataForBlobInput.Add(new BlobInput { BlobName = newFilename, BlobContent = data });
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
            return dataForBlobInput;
        }
    }
}
