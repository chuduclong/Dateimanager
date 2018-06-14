using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace View
{
    class RunFile
    {
        PowerPoint.Application PowerPointApp;
        PowerPoint.Presentation Pres;
        Word.Application WordApp;
        Word.Document Doc;

        public void openExistingWordFile(string name)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Doc = WordApp.Documents.Open("DATNAME", oMissing, false);
        }

        public void openExistingPowerPointPresentation(string name)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Pres = PowerPointApp.Presentations.Open("DATNAME");
        }

        public void openNewWordFile()
        {
            WordApp = new Word.Application();
            WordApp.Visible = true;
            Doc = WordApp.Documents.Add();
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

        public void closeCurrentPowerPointPresentation()
        {
            Pres.Close();
        }
    }
}
