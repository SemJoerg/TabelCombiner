using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.IO;

namespace TabelCombiner
{
    public static class ExcelLogic
    {
        public readonly static BackgroundWorker excelWorker;

        static ExcelLogic()
        {
            excelWorker = new BackgroundWorker();
            excelWorker.WorkerReportsProgress = false;
            excelWorker.WorkerSupportsCancellation = false;
            excelWorker.DoWork += ExcelWorker_DoWork;
        }

        private static void ExcelWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Excel.Application excelApp = null;
            Excel.Workbooks workbooks = null;
            Excel._Workbook mainWorkbook = null;
            try
            {
                IEnumerable<FileInfo> files = e.Argument as IEnumerable<FileInfo>;

                excelApp = new Excel.Application();
                excelApp.Visible = false;
                workbooks = excelApp.Workbooks;
                mainWorkbook = workbooks.Add();
                //Excel._Worksheet mainSheet = (Excel._Worksheet)mainWorkbook.ActiveSheet;
                Excel._Worksheet mainSheet = (Excel._Worksheet)mainWorkbook.Sheets[1];
                int mainSheetRowCounter = 0;
                bool copyTabelHeadder = true;

                //Read And Combine Files
                foreach (FileInfo file in files)
                {
                    Excel._Workbook newWorkbook = null;
                    try
                    {
                        newWorkbook = workbooks.Open(file.FullName);
                        Excel._Worksheet newSheet = null;
                        if (copyTabelHeadder)
                        {
                            newSheet = (Excel.Worksheet)newWorkbook.Sheets[1];
                            int lastRow = LastRowTotal(newSheet);
                            Excel.Range source = newSheet.Range["1:" + lastRow];
                            Excel.Range destination = mainSheet.Range["1:1"];
                            source.Copy(destination);
                            mainSheetRowCounter = lastRow + 1;
                            for (int i = 2; i <= newWorkbook.Sheets.Count; i++)
                            {
                                newSheet = (Excel._Worksheet)newWorkbook.Sheets[i];
                                newSheet.Copy(After: mainSheet);
                                mainSheet.Activate();
                            }
                            copyTabelHeadder = false;
                        }
                        else
                        {
                            newSheet = (Excel._Worksheet)newWorkbook.Sheets[1];

                            int lastRow = LastRowTotal(newSheet);
                            if(lastRow > 1)
                            {
                                Excel.Range source = newSheet.Range["2:" + lastRow];
                                Excel.Range destination = mainSheet.Range[mainSheetRowCounter + ":" + mainSheetRowCounter];
                                mainSheetRowCounter += lastRow - 1;
                                source.Copy(destination);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.ErrorMessage(ex.Message);
                    }
                    finally
                    {
                        if (newWorkbook != null)
                        {
                            newWorkbook.Close(false);
                            Marshal.FinalReleaseComObject(newWorkbook);
                            newWorkbook = null;
                        }
                    }
                }

                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                Log.ErrorMessage(ex.Message);
            }
            finally
            {
                if(mainWorkbook != null)
                {
                    Marshal.FinalReleaseComObject(mainWorkbook);
                    mainWorkbook = null;
                }
                if(workbooks != null)
                {
                    Marshal.FinalReleaseComObject(workbooks);
                    workbooks = null;
                }
                if(excelApp != null)
                {
                    Marshal.FinalReleaseComObject(excelApp);
                    excelApp = null;
                }
            }
        }

        public static int LastRowTotal(Excel._Worksheet wks)
        {
            Excel.Range lastCell = wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            return lastCell.Row;
        }
    }
}
