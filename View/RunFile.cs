using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace View
{
    class RunFile
    {
        Word.Application WordApp;
        Word.Document Doc;

        public void openExistingWordFile(string name)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Doc = WordApp.Documents.Open("DATNAME", oMissing, false);
        }

        public void openNewWordFile()
        {
            WordApp = new Word.Application();
            WordApp.Visible = true;
            Doc = WordApp.Documents.Add();
        }

        public void closeCurrentWordFile()
        {
            Doc.Close();
        }
    }
}
