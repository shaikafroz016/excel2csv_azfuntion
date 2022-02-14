using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.Data;

namespace ExcelConversionUtility
{
    public static class ExcelToCSVConvertor
    {
        public static List<BlobInput> Convert(List<BlobOutput> inputs)
        {
            DataTable WorkSheetMaster = new DataTable();
            WorkSheetMaster.Columns.Add("WorksheetName");
            WorkSheetMaster.Columns.Add("WorksheetCode");
            WorkSheetMaster.Rows.Add(new Object[] { "Claims Data", "ClaD" });
            WorkSheetMaster.Rows.Add(new Object[] { "Policy Data", "PolD" });
            WorkSheetMaster.Rows.Add(new Object[] { "Client Data", "CliD" });
            DataTable EventsConfigList = new DataTable();
            EventsConfigList.Columns.Add("WorksheetName");
            EventsConfigList.Columns.Add("WSFields");
            EventsConfigList.Columns.Add("EventAssetIDMapping");
            EventsConfigList.Columns.Add("Region");
            EventsConfigList.Columns.Add("EventType");
            EventsConfigList.Columns.Add("ComplianceTag");
            EventsConfigList.Columns.Add("WSEventDateField");
            EventsConfigList.Columns.Add("EventDateFormula");
            EventsConfigList.Rows.Add(new Object[] { "Client Data", "CountryCode, GCID", "TCT_Code:**$”GCID”$** AND CRBCountryCode:**$”CountryCode”$**", "US, CAN, MEX, CYM, CRI", "CRB End of Client Relationship", "ACC - Client Relationship Management", "ClientEndDate", "X+5" });
            EventsConfigList.Rows.Add(new Object[] { "Client Data", "CountryCode, GCID, Claims Ref", "TCT_Code:**$”GCID”$** AND CRBCountryCode:**$”CountryCode”$**  AND CRB_ClaimId:**$”ClaimsRef”$**", "GLOBAL", "CRB Claim Servicing Record", "INS - Broking and Reinsurance – Claims", "ClaimCloseDate", "X" });
            EventsConfigList.Rows.Add(new Object[] { "Policy Data", "CountryCode, GCID,  External Reference", "TCT_Code:**$”GCID”$** AND CRBCountryCode:**$”CountryCode”$**  AND CRB_PolicyReference:**$”External Reference”$**", "US, CAN, MEX, CYM, CRI", "CRB Policy Expired", "INS - Broking and Reinsurance", "?", "?" });
            EventsConfigList.Rows.Add(new Object[] { "Client Data", "CountryCode, GCID", "TCT_Code:**$”GCID”$** AND CRBCountryCode:**$”CountryCode”$**", "GLOBAL", "CRB End of Client Relationship", "ACC - Client Relationship Management", "ClientEndDate", "?-5" });
            EventsConfigList.Rows.Add(new Object[] { "Claims Data", "CountryCode, GCID, Claims Data", "TCT_Code:**$”GCID”$** AND CRBCountryCode:**$”CountryCode”$**  AND CRB_InternalClaimNumber:**$”ClaimsRef”$**;", "US, CAN, MEX, CYM, CRI", "CRB Claim Servicing Record", "INS - Broking and Reinsurance – Claims", "ClaimClosedDate", "?+5" });



            DataSet ds = new DataSet();
            string b_name = ""; 
            var dataForBlobInput = new List<BlobInput>();
            try
            {
                foreach (BlobOutput item in inputs)
                {
                    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(item.BlobName);
                    WorksheetCollection worksheets = workbook.Worksheets;
                    workbook.Worksheets.RemoveAt("Sheet1");
                    b_name = item.BlobName;
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(item.BlobContent, false))
                    {
                        int x = 0;
                        foreach (Sheet _Sheet in document.WorkbookPart.Workbook.Descendants<Sheet>())
                              {
                            var st = Vali(WorkSheetMaster, _Sheet.Name);
                            if (st != null)
                            {
                                DataTable dt = workbook.Worksheets[0].Cells.ExportDataTable(0, 0, workbook.Worksheets[x].Cells.MaxDataRow + 1, workbook.Worksheets[x].Cells.MaxDataColumn + 1, true);
                                dt.TableName = _Sheet.Name;
                                ds.Tables.Add(dt);
                                Console.WriteLine(_Sheet.Name + " is proccessed for file "+ item.BlobName);
                                x++;
                            }
                            else
                            {
                                Console.WriteLine(_Sheet.Name+ " is ignored fo file " + item.BlobName);
                            }
                            
                            
                        }
                        x = 0;
                            }
                           
                }
            }
            catch (Exception Ex)
            {
               // workbook.Save(b_name, SaveFormat.Xlsx);
                throw Ex;
            }
            Console.WriteLine(ds.Tables.Count);
            return dataForBlobInput;
        }
        public static string Vali(DataTable wsn, string na)
        {
            
                if (wsn.Rows[0].ItemArray[0].ToString() == na  || wsn.Rows[1].ItemArray[0].ToString() == na || wsn.Rows[2].ItemArray[0].ToString() == na)
                {
                    return na;
                }
                
                
            
            return null;
        }
    }
    
    
}
