using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.Application.DocumentOpen +=new Word.ApplicationEvents4_DocumentOpenEventHandler(WorkWithDocument);

            //((Word.ApplicationEvents4_Event)this.Application).NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(WorkWithDocument);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        //private void WorkWithDocument(Microsoft.Office.Interop.Word.Document Doc)
        //{
        //   DefaultFormatting();
        //}
        //public void DefaultFormatting()
        //{
        //    //Word.Selection currentSelection = this.Application.Selection;
        //    //Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
        //    //document.Paragraphs.Space15();
        //    //for (int i = 1; i <= currentSelection.Paragraphs.Count; i++)
        //    //{  
        //    //    Word.Range rng = currentSelection.Paragraphs[i].Range;
        //    //    rng.Font.Name = "Arial";
        //    //    rng.Font.Size = 12;
        //    //    rng.Font.Bold = 0;
        //    //    rng.Font.Italic = 0;
        //    //    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
        //    //}  
        //    //rng.ParagraphFormat.LeftIndent = Application.InchesToPoints(1.25f);
        //    //rng.ParagraphFormat.RightIndent = Application.InchesToPoints(1.25f);
        //}

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
