using System;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

public class ExcelHandler
{
    private Excel.Application ExcelApp = null;

    Excel.Workbook xlReferenceWorkbook = null;
    Excel.Workbook xlTargetWorkbook = null;

    public Workbook XlReferenceWorkbook { get => xlReferenceWorkbook; set => xlReferenceWorkbook = value; }
    public Workbook XlTeamMemberWorkbook { get => xlTargetWorkbook; set => xlTargetWorkbook = value; }

    public bool OpenReferenceWorkbook(string location)
    {
        try
        {
            if (XlReferenceWorkbook == null && !location.StartsWith("~"))
            {
                XlReferenceWorkbook = ExcelApp.Workbooks.Open(location);
            }
            return true;
        }
        catch (Exception e)
        {
            return false;
        }
    }

    public bool OpenTeamMemberWorkbook(string location)
    {
        try
        {
            XlTeamMemberWorkbook = ExcelApp.Workbooks.Open(location);
            return true;
        }
        catch(Exception e)
        {
            return false;
        }
    }

    public bool CloseWorkbook(Excel.Workbook workbook)
    {
        if (workbook != null)
        {
            try
            {
                workbook.Save();
                workbook.Close();
            }
            finally { }

            Marshal.ReleaseComObject(workbook);

            return true;
        }

        return false;
    }

    public bool HideSheetByName(Excel.Workbook workbook, string workSheetName)
    {
        if(FindSheetByName(workbook, workSheetName))
        {
            ((Excel.Worksheet)workbook.Worksheets[workSheetName]).Visible = XlSheetVisibility.xlSheetHidden;
        }

        return false;
    }

    private bool FindSheetByName(Excel.Workbook workbook, string workSheetName)
    {
        foreach(Excel.Worksheet sheet in workbook.Worksheets)
        {
            if(sheet.Name == workSheetName)
            {
                return true;
            }
        }

        return false;
    }

    public string GetSheetNameByIndex(Excel.Workbook workbook, int i)
    {
        if (workbook.Worksheets.Count >= i)
        {
            return ((Excel.Worksheet)workbook.Worksheets[i]).Name;
        }

        return string.Empty;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="workbook"></param>
    /// <param name="i">NON ZERO based index!</param>
    /// <returns></returns>
    public string GetVisibleSheetNameByIndex(Excel.Workbook workbook, int i)
    {
        List<Excel.Worksheet> visibleSheets = new List<Worksheet>();
        foreach(Excel.Worksheet sheet in workbook.Sheets)
        {
            if(sheet.Visible == XlSheetVisibility.xlSheetVisible)
            {
                visibleSheets.Add(sheet);
            }
        }

        if (visibleSheets.Count >= 1)
        {
            var usedIndex = 0;
            if (i > 0)
            {
                usedIndex = i - 1;
            }
            return visibleSheets[usedIndex].Name;
        }

        return string.Empty;
    }

    public bool UpdateLinks(Excel.Workbook workbook)
    {
        //Array links = workbook.LinkSources(XlLink.xlExcelLinks) as Array;

        //if (links != null && links.Length > 0)
        //{
        //    foreach (var link in links)
        //    {
        //        workbook.BreakLink(link.ToString(), XlLinkType.xlLinkTypeExcelLinks);
        //        //workbook.ChangeLink(link.ToString(), workbook.FullName, XlLinkType.xlLinkTypeExcelLinks);
        //    }
        //}

        foreach (Excel.Name name in workbook.Names)
        {
            if(name.Name == "TOWList")
            {
                name.Delete();
            }
        }

        var indexReference = GetSheetIndexByName(workbook, "Input");
        Excel.Worksheet targetSheet = workbook.Worksheets[indexReference];

        Excel.Range range = targetSheet.Columns["H:H"];
        range.Validation.Delete();
        workbook.Save();
        // .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _xlBetween, Formula1:= "=INDIRECT(""lookup_projects[Project]"")"
        range.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, @"=INDIRECT(""lookup_projects[Project]"")");

        workbook.Save();
        return true;
    }

    private int GetSheetIndexByName(Excel.Workbook workbook, string workSheetName)
    {
        for (int i = 1; i <= workbook.Worksheets.Count; i++)
        {
            if(workbook.Worksheets[i].Name == workSheetName)
            {
                return i;
            }
        }

        return int.MinValue;
    }

    public bool CopyValuesByTableName(string sheetName, string tableName)
    {
        var indexReference = GetSheetIndexByName(XlReferenceWorkbook, sheetName);
        var indexTarget = GetSheetIndexByName(XlTeamMemberWorkbook, sheetName);

        Excel.Range sourceTableRange = null;
        Excel.Range targetTableRange = null;

        foreach (Excel.ListObject table in XlReferenceWorkbook.Worksheets[indexReference].ListObjects)
        {
            if (table.Name == tableName)
            {
                sourceTableRange = table.Range;
            }
        }

        foreach (Excel.ListObject table in XlTeamMemberWorkbook.Worksheets[indexTarget].ListObjects)
        {
            if (table.Name == tableName)
            {
                targetTableRange = table.Range;
            }
        }

        if (targetTableRange != null && sourceTableRange != null)
        {
            int rowCount = sourceTableRange.Rows.Count;
            int colCount = sourceTableRange.Rows.Count;

            int targetSize = targetTableRange.Rows.Count;
            if (targetSize > rowCount)
            {
                var diff = targetSize - rowCount;

                for(int i=1; i<=diff; i++)
                {
                    targetTableRange.Rows[rowCount + i].Delete();
                }
            }

            var lastInsertRowIndex = 2; //include header

            for (int i = 2; i <= rowCount; i++) // skip header
            {
                //if(i == 3) { break; } //debug mode: Only 1 row

                for (int j = 1; j <= colCount; j++)
                {
                    if (sourceTableRange.Cells[i, j] != null && sourceTableRange.Cells[i, j].Value2 != null && sourceTableRange.Cells[i, j].Value2.ToString() != string.Empty)
                    {
                        targetTableRange.Cells[lastInsertRowIndex, j].Value2 = sourceTableRange.Cells[i, j].Value2;
                    }
                    else
                    {
                        break;
                    }
                }

                lastInsertRowIndex++;
            }

            
            XlTeamMemberWorkbook.Save();
        }


        sourceTableRange = null;
        targetTableRange = null;

        //cleanup
        GC.Collect();
        GC.WaitForPendingFinalizers();

        return true;
    }

    public bool CopyReferenceSheets(List<int> sheetIndexes)
    {
        Excel._Worksheet templateSheet = null;
        bool success = true;

        try
        {
            foreach (int i in sheetIndexes)
            {
                templateSheet = XlReferenceWorkbook.Sheets[i];

                // Delete sheets with duplicate names
                if (FindSheetByName(XlTeamMemberWorkbook, templateSheet.Name)) 
                {
                    ExcelApp.DisplayAlerts = false;
                    ((Excel.Worksheet)XlTeamMemberWorkbook.Sheets[templateSheet.Name]).Visible = XlSheetVisibility.xlSheetVisible;
                    ((Excel.Worksheet)XlTeamMemberWorkbook.Sheets[templateSheet.Name]).Delete();
                    ExcelApp.DisplayAlerts = true;

                    XlTeamMemberWorkbook.Save();
                }

                templateSheet.Copy(XlTeamMemberWorkbook.Worksheets[1]);

                XlTeamMemberWorkbook.Save();
            }
        }
        catch (Exception e)
        {
            success = false;
        }
        finally
        {
            if (XlTeamMemberWorkbook != null)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (templateSheet != null)
            {
                Marshal.ReleaseComObject(templateSheet);
                templateSheet = null;
            }
        }

        return success;
    }

    public void Dispose()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();

        if (XlReferenceWorkbook != null)
        {
            try
            {
                if (!XlReferenceWorkbook.Saved)
                {
                    XlReferenceWorkbook.Save();
                }

                XlReferenceWorkbook.Close(false, Missing.Value, Missing.Value);
            }
            catch (Exception e)
            {
                // die silently (for now :))
            }

            Marshal.ReleaseComObject(XlReferenceWorkbook);
            XlReferenceWorkbook = null;
        }

        if (XlTeamMemberWorkbook != null)
        {
            try
            {
                XlTeamMemberWorkbook.Close();
            }
            catch (Exception e)
            {
                // die silently (for now :))
            }

            Marshal.ReleaseComObject(XlTeamMemberWorkbook);
            XlTeamMemberWorkbook = null;
        }

        if(ExcelApp != null)
        {
            //quit and release
            ExcelApp.Quit();
            Marshal.ReleaseComObject(ExcelApp);
        }
    }

    public ExcelHandler()
    {
        if (ExcelApp == null)
        {
            ExcelApp = new Excel.Application();
            ExcelApp.DisplayAlerts = false;
        }
    }

    public bool CopyPasteTimeTable(string sourceWorkSheetName, string targetWorkSheetName)
    {
        /*ActiveWindow.SmallScroll Down:=-249
        Range("D2:H52").Select
        Selection.Copy
        Sheets.Add After:=ActiveSheet
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

        Windows("Urenstaten FY19 - Wouter Overmeer.xlsx").Activate
        Sheets("Input").Select
        Selection.Copy
        Windows("Digital Studio Account Overview.xlsx").Activate
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("G3922").Select
         * 
         * */
        Excel.Worksheet hourSheet = null;
        Excel.Worksheet consolidationSheet = null;
        var success = true;

        try
        {
            

            hourSheet = XlTeamMemberWorkbook.Worksheets[sourceWorkSheetName];
            Excel.Range xlRange = hourSheet.UsedRange;
            //Excel.Range range = sheet.get_Range("A1", "B4");
           
            int rowCount = xlRange.Rows.Count;
            int emptyRowCount = 0;
            var lastValidSourceRow = 0;

            consolidationSheet = XlReferenceWorkbook.Worksheets[targetWorkSheetName];
            // sheet has headers
            var lastInsertRowIndex = consolidationSheet.UsedRange.Rows.Count;
            lastInsertRowIndex += 1;

            for (int i = 2; i <= rowCount; i++)
            {
                //check if cells 4, 5, 6, 8 (date, name, hour, project) have values, these are the ones that count.
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4].Value2.ToString() != string.Empty &&
                    xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5].Value2.ToString() != string.Empty &&
                    xlRange.Cells[i, 6] != null && xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6].Value2.ToString() != string.Empty &&
                    xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8].Value2.ToString() != string.Empty)
                {
                    
                    emptyRowCount = 0;
                }
                else
                {
                    if (emptyRowCount == 4) // break after 5 consequtive empty rows
                    {
                        if (i > 5) // do we have input?
                        {
                            lastValidSourceRow = i-5;
                        }
                        break;
                    }

                    emptyRowCount++;
                }
            }

            Excel.Range actuals = hourSheet.get_Range("D2", "H" + lastValidSourceRow);
            actuals.Copy();
            consolidationSheet.Cells[lastInsertRowIndex, 4].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, true, false);

            //consolidationSheet.UsedRange.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, true, false);

            XlReferenceWorkbook.Save();
        }
        catch (Exception e)
        {
            success = false;
        }
        finally
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (hourSheet != null)
            {
                Marshal.ReleaseComObject(hourSheet);
                hourSheet = null;
            }

            if (consolidationSheet != null)
            {
                Marshal.ReleaseComObject(consolidationSheet);
                consolidationSheet = null;
            }
        }

        return success;

    }

    public bool ConsolidateTimesheet(string sourceWorkSheetName, string targetWorkSheetName)
    {
        Excel.Worksheet hourSheet = null;
        Excel.Worksheet consolidationSheet = null;
        var success = true;

        try
        {
            hourSheet = XlTeamMemberWorkbook.Worksheets[sourceWorkSheetName];
            Excel.Range xlRange = hourSheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int emptyRowCount = 0;

            consolidationSheet = XlReferenceWorkbook.Worksheets[targetWorkSheetName];
 
            // sheet has headers
            var lastInsertRowIndex = consolidationSheet.UsedRange.Rows.Count;
            lastInsertRowIndex += 1;

            // Uncomment for short run for debugging purposes
            //if (j == 6) { break; }

            for (int i = 2; i <= rowCount; i++)
            {
                //check if cells 4, 5, 6, 8 (date, name, hour, project) have values, these are the ones that count.
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4].Value2.ToString() != string.Empty &&
                    xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5].Value2.ToString() != string.Empty &&
                    xlRange.Cells[i, 6] != null && xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6].Value2.ToString() != string.Empty &&
                    xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8].Value2.ToString() != string.Empty)
                {
                    for (int j = 4; j <= 8; j++)
                    {
                        consolidationSheet.Cells[lastInsertRowIndex, j].Value2 = xlRange.Cells[i, j].Value2;
                    }

                    emptyRowCount = 0;
                    lastInsertRowIndex++;
                }
                else
                {
                    if (emptyRowCount == 4) // break after 5 consequtive empty rows
                    {
                        break;
                    }

                    emptyRowCount++;
                }

                
            }

            XlReferenceWorkbook.Save();
        }
        catch (Exception e)
        {
            success = false;
        }
        finally
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (hourSheet != null)
            {
                Marshal.ReleaseComObject(hourSheet);
                hourSheet = null;
            }

            if (consolidationSheet != null)
            {
                Marshal.ReleaseComObject(consolidationSheet);
                consolidationSheet = null;
            }
        }

        return success;
    }

    public bool ConvertTimeSheet(string sourceWorkSheetName, string targetWorkSheetName)
    {
        Excel.Worksheet hourSheet = null;
        Excel.Worksheet timeSheet = null;
        var success = true;

        try
        {
            hourSheet = XlTeamMemberWorkbook.Worksheets[sourceWorkSheetName];
            Excel.Range xlRange = hourSheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            int count = XlTeamMemberWorkbook.Worksheets.Count;
            timeSheet = XlTeamMemberWorkbook.Worksheets[targetWorkSheetName];

            // sheet has headers
            var lastInsertRowIndex = 2;

            for (int j = 5; j <= colCount; j++)
            {
                var date = DateTime.MinValue;

                // Uncomment for short run for debugging purposes
                //if(j == 6) { break;  }

                for (int i = 2; i <= rowCount; i++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value != null && xlRange.Cells[i, j].Value.ToString() != string.Empty)
                    {
                        if (i == 2)
                        {
                            date = DateTime.Parse(xlRange.Cells[i, j].Value.ToString());
                        }
                        else
                        {
                            timeSheet.Cells[lastInsertRowIndex, 4].Value = date;
                            timeSheet.Cells[lastInsertRowIndex, 5].Value = xlRange.Cells[i, 1].Value;
                            timeSheet.Cells[lastInsertRowIndex, 6].Value = xlRange.Cells[i, j].Value;
                            timeSheet.Cells[lastInsertRowIndex, 7].Value = xlRange.Cells[i, 4].Value;
                            lastInsertRowIndex++;
                        }
                    }
                }
            }

            XlTeamMemberWorkbook.Save();
        }
        catch (Exception e)
        {
            success = false;
        }
        finally
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (hourSheet != null)
            {
                Marshal.ReleaseComObject(hourSheet);
                hourSheet = null;
            }

            if (timeSheet != null)
            {
                Marshal.ReleaseComObject(timeSheet);
                timeSheet = null;
            }
        }

        return success;
    }
}
