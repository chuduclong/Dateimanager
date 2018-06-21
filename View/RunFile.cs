using System;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Model;

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
            Object oMissing = System.Reflection.Missing.Value;
            Doc = WordApp.Documents.Open(name, oMissing, false);
        }

        public void openExistingPowerPointPresentation(string name)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Pres = PowerPointApp.Presentations.Open(name);
        }

        public void openExistingExcelWorkbook(string name)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Workbook = ExcelApp.Workbooks.Open(name, oMissing, false);
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

        public void chooseData(String application)
        {
            switch(application)
            {
                case "word":
                    ds.WordAnzeigen();
                    break;

                case "powerpoint":
                    break;

                case "excel":
                    break;

                default:
                    break;
            }
        }
    }
}
