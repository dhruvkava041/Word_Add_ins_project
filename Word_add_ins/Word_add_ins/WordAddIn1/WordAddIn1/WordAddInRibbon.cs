using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
namespace WordAddIn1
{
    public partial class WordAddInRibbon
    {
        private void WordAddInRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void ChNameNNum_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            Word.Range rng = currentSelection.Paragraphs[1].Range;
            rng.Font.Name = "Arial";
            rng.Font.Size = 16;
            rng.Font.AllCaps = 1;
            rng.Font.Bold = 1;
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            for (int i = 0; i < 2; i++) 
            currentSelection.TypeParagraph();
            //MessageBox.Show(currentSelection.Text);
            
            currentSelection.ClearFormatting();
            //Globals.ThisAddIn.DefaultFormatting();
            DefaultFormatting();

        }

        private void SecNameNNum_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            Word.Range rng = currentSelection.Paragraphs[1].Range;
            rng.Font.Name = "Arial";
            rng.Font.Size = 14;
            rng.Font.AllCaps = 1;
            rng.Font.Bold = 1;
            //rng.ParagraphFormat.set_Style(Word.WdArabicNumeral.wdNumeralArabic);
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            //for (int i = 0; i < 2; i++)
            currentSelection.TypeParagraph();
            
            currentSelection.ClearFormatting();
            //Globals.ThisAddIn.DefaultFormatting();
            DefaultFormatting();
        }

        private void SubSecNameNNum_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            Word.Range rng = currentSelection.Paragraphs[1].Range;
            rng.Font.Name = "Arial";
            rng.Font.Size = 12;
            rng.Case = Word.WdCharacterCase.wdTitleWord;
            rng.Font.Bold = 1;
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            //for (int i = 0; i <= 3; i++)
            currentSelection.TypeParagraph();
            currentSelection.ClearFormatting();
            //Globals.ThisAddIn.DefaultFormatting();
            DefaultFormatting();
        }

        private void GeneralText_Click(object sender, RibbonControlEventArgs e)
        {
            //Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            //currentSelection.ClearFormatting();
            //currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            //Globals.ThisAddIn.DefaultFormatting();
            DefaultFormatting();
        }

        private void SpecialText_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            for (int i = 1; i <= currentSelection.Paragraphs.Count; i++)
            {
                Word.Range rng = currentSelection.Paragraphs[i].Range;
                rng.Case = Word.WdCharacterCase.wdLowerCase;
                rng.Font.Name = "Arial";
                rng.Font.Size = 12;
                rng.Font.Italic = 1;
                rng.Font.Bold = 0;
                rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            }
            currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }
        public void DefaultFormatting()
        {
            Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            //document.Paragraphs.Space15();
            for (int i = 1; i <= currentSelection.Paragraphs.Count; i++)
            {  
                Word.Range rng = currentSelection.Paragraphs[i].Range;
                rng.Case = Word.WdCharacterCase.wdLowerCase;
                rng.Font.Name = "Arial";
                rng.Font.Size = 12;
                rng.Font.Bold = 0;
                rng.Font.Italic = 0;
                rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            }  
            //rng.ParagraphFormat.LeftIndent = Application.InchesToPoints(1.25f);
            //rng.ParagraphFormat.RightIndent = Application.InchesToPoints(1.25f);
        }

        private void SaveAsPdf_Click(object sender, RibbonControlEventArgs e)
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            string sfileName_Document = doc.Name;
            string sPath = doc.Path;
            string sFullpath_pdf = sPath + "\\" + sfileName_Document + ".pdf";
            doc.ExportAsFixedFormat(sFullpath_pdf, Word.WdExportFormat.wdExportFormatPDF, OpenAfterExport: true);
        }
    }
}
