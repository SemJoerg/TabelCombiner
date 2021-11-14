using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.IO;
using Forms = System.Windows.Forms;
using System.Diagnostics;


namespace TabelCombiner
{
    public static class ExcelLogic
    {
        public readonly static BackgroundWorker excelWorker;

        static ExcelLogic()
        {
            excelWorker = new BackgroundWorker();
            excelWorker.WorkerReportsProgress = true;
            excelWorker.WorkerSupportsCancellation = true;
            excelWorker.DoWork += ExcelWorker_DoWork;
        }

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private static Process GetExcelProcess(Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }

        private static void ExcelWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Excel.Application excelApp = null;
            Process excelProcess = null;
            Excel.Workbooks workbooks = null;
            Excel._Workbook mainWorkbook = null;
            ExcelWorkerArgs args = e.Argument as ExcelWorkerArgs;

            void ReleaseEverything()
            {
                if (mainWorkbook != null)
                {
                    Marshal.FinalReleaseComObject(mainWorkbook);
                    mainWorkbook = null;
                }
                if (workbooks != null)
                {
                    Marshal.FinalReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (excelApp != null)
                {
                    Marshal.FinalReleaseComObject(excelApp);
                    excelApp = null;
                }
            }

            try
            {
                excelApp = new Excel.Application();
                excelProcess = GetExcelProcess(excelApp);
                excelApp.Visible = false;
                workbooks = excelApp.Workbooks;
                mainWorkbook = workbooks.Add();
                Excel._Worksheet mainSheet = (Excel._Worksheet)mainWorkbook.Sheets[1];
                mainSheet.Name = "DeleteSheet";
                int mainSheetRowCounter = 2;
                bool copyTabelHeadder = true;

                //Read And Combine Files
                foreach (FileInfo file in args.FileList)
                {
                    Excel._Workbook newWorkbook = null;
                    try
                    {
                        newWorkbook = workbooks.Open(file.FullName);
                        Excel._Worksheet newSheet = null;

                        if (copyTabelHeadder)
                        {
                            newSheet = (Excel.Worksheet)newWorkbook.Sheets[1];
                            newSheet.Copy(After: mainSheet);
                            mainSheet.Delete();
                            mainSheet = (Excel._Worksheet)mainWorkbook.Sheets[1];

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

                            Excel.Range source = newSheet.Range["2:2"];
                            Excel.Range destination = mainSheet.Range[++mainSheetRowCounter + ":" + mainSheetRowCounter];
                            source.Copy(destination);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.ErrorMessage(ex.Message);
                    }

                    //Close and release newWorkbook
                    if (newWorkbook != null)
                    {
                        newWorkbook.Close(false);
                        Marshal.FinalReleaseComObject(newWorkbook);
                        newWorkbook = null;
                    }

                    //Check for cancelation
                    if (excelWorker.CancellationPending)
                    {
                        return;
                    }

                    excelWorker.ReportProgress(mainSheetRowCounter - 1);
                }


                if(args.ExportAsTextFile)
                {
                    string filePath = null;

                    Thread fileSaveThread = new Thread(() =>
                    {
                        Forms.SaveFileDialog saveFileDialog = new Forms.SaveFileDialog();
                        saveFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                        saveFileDialog.AddExtension = true;
                        saveFileDialog.DefaultExt = ".txt";
                        saveFileDialog.ShowDialog();
                        filePath = saveFileDialog.FileName;
                    });
                    fileSaveThread.SetApartmentState(ApartmentState.STA);
                    fileSaveThread.Start();
                    fileSaveThread.Join();
                    
                    if(!String.IsNullOrEmpty(filePath))
                    {
                        mainSheet.SaveAs2(filePath, Excel.XlFileFormat.xlTextWindows);
                    }
                }

                if(args.ShowExcel)
                {
                    excelApp.Visible = true;
                }

            }
            catch (Exception ex)
            {
                Log.ErrorMessage(ex.Message);
            }
            finally
            {
                if(!excelApp.Visible)
                {
                    mainWorkbook.Close(false);
                    excelApp.Quit();
                    ReleaseEverything();
                    excelProcess.Kill(true);
                }
                else
                {
                    ReleaseEverything();
                }
            }
        }
    }
}
