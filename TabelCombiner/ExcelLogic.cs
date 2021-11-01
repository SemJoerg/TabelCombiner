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
            IEnumerable<FileInfo> files = e.Argument as IEnumerable<FileInfo>;

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel._Workbook mainWorkbook = workbooks.Add();
            Excel._Worksheet mainSheet = (Excel._Worksheet)mainWorkbook.ActiveSheet;
            int mainSheetRowCounter = 2;
            bool copyTabelHeadder = true;

            //Read And Combine Files
            foreach (FileInfo file in files)
            {
                Excel._Workbook newWorkbook = null;
                try
                {
                    newWorkbook = workbooks.Open(file.FullName);
                    for (int i = 0; i < newWorkbook.Sheets.Count; i++)
                    {
                        Excel._Worksheet newSheet = (Excel._Worksheet)newWorkbook.Sheets[i + 1];
                        int lastRow = LastRowTotal(newSheet);

                        if(copyTabelHeadder)
                        {
                            Excel.Range headderSource = newSheet.Range["1:1"];
                            Excel.Range headderDestination = mainSheet.Range["1:1"];
                            headderSource.Copy(headderDestination);
                            copyTabelHeadder = false;
                        }

                        Excel.Range source = newSheet.Range["2:" + lastRow];
                        Excel.Range destination = mainSheet.Range[mainSheetRowCounter + ":" + mainSheetRowCounter];
                        mainSheetRowCounter += lastRow - 1;
                        source.Copy(destination);
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
            Marshal.FinalReleaseComObject(mainWorkbook);
            Marshal.FinalReleaseComObject(workbooks);
            Marshal.FinalReleaseComObject(excelApp);
            mainWorkbook = null;
            workbooks = null;
            excelApp = null;
        }

        public static int LastRowTotal(Excel._Worksheet wks)
        {
            Excel.Range lastCell = wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            return lastCell.Row;
        }
    }
}
