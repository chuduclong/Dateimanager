using System;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Model;
using System.IO;

namespace View
{
    class RunFile
    {
        PowerPoint.Application PowerPointApp;
        Excel.Application ExcelApp;
        PowerPoint.Presentation Pres;
        Word.Application WordApp;
        Word.Document Doc;
        Excel.Workbook Workbook;
        Datenbank ds = new Datenbank();

        public void openExistingWordFile(string name)
        {
            object missingValue = Type.Missing;
            object fileName = name;
            object readOnly = false;
            object isVisible = true;
            WordApp = new Word.Application();
            WordApp.Visible = true;
            Doc = WordApp.Documents.Open(ref fileName);
            Doc.Activate();
        }

        public void openExistingPowerPointPresentation(string name)
        {
            string fileName = name;
            PowerPointApp = new PowerPoint.Application();
            PowerPointApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            Pres = PowerPointApp.Presentations.Open(fileName);
        }

        public void openExistingExcelWorkbook(string name)
        {
            object missingValue = Type.Missing;
            string fileName = name;
            object readOnly = false;
            object isVisible = true;
            ExcelApp = new Excel.Application();
            ExcelApp.Visible = true;
            Workbook = ExcelApp.Workbooks.Open(name, missingValue, false);
        }

        public void openNewWordFile()
        {
            WordApp = new Word.Application();
            WordApp.Visible = true;
            Doc = WordApp.Documents.Add();
        }

        public void openNewExcelWorkbook()
        {
            ExcelApp = new Excel.Application();
            ExcelApp.Visible = true;
            Workbook = ExcelApp.Workbooks.Add();
        }

        public void openNewPowerPointPresentation()
        {
            PowerPointApp = new PowerPoint.Application();
            PowerPointApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            Pres = PowerPointApp.Presentations.Add();
        }

        public void closeCurrentWordFile()
        {
            Doc.Close();
        }

        public void closeCurrentExcelWorkbook()
        {
            Workbook.Close();
        }
        public void closeCurrentPowerPointPresentation()
        {
            Pres.Close();
        }

        public void addFile(FileInfo fi, String dateiart)
        {
            ds.AddDokus(fi.Name, fi.Name, fi.CreationTime, dateiart);
        }
    }
}
