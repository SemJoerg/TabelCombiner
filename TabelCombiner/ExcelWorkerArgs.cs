using System;
using System.Collections.Generic;
using System.IO;

namespace TabelCombiner
{
    public class ExcelWorkerArgs
    {
        public readonly FileInfo[] FileList;
        public readonly bool ExportAsTextFile;
        public readonly bool ShowExcel;

        public ExcelWorkerArgs(FileInfo[] fileList, bool? exportAsTextFile, bool? showExcel)
        {
            FileList = fileList.Clone() as FileInfo[];
            ExportAsTextFile = exportAsTextFile ?? false;
            ShowExcel = showExcel ?? false;
        }
    }
}
