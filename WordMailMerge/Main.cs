using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordMailMerge
{
    public class Main
    {
        public void DoMerge()
        {
            const string sourceFile = @"E:\Projects\tests\WordMailMerge\MergeForm.docx";
            const string targetFile = @"E:\Projects\tests\WordMailMerge\MergedForm.docx";

            const string mergeFieldName = "Property";
            const string replacementText = "The Property, 1 The Street";

            File.Copy(sourceFile, targetFile, true);

            using (var document = WordprocessingDocument.Open(targetFile, true))
            {
                // If your sourceFile is a different type (e.g., .DOTX), you will need to change the target type like so:
                document.ChangeDocumentType(WordprocessingDocumentType.Document);

                // Get the MainPart of the document
                var mainPart = document.MainDocumentPart;
                var mergeFields = mainPart.RootElement.Descendants<FieldCode>();

                ReplaceMergeFieldWithText(mergeFields, mergeFieldName, replacementText);

                // Save the document
                mainPart.Document.Save();

            }
        }


        private static void ReplaceMergeFieldWithText(IEnumerable<FieldCode> fields, string mergeFieldName, string replacementText)
        {
            var field = fields
                .FirstOrDefault(f => f.InnerText.Contains(mergeFieldName));

            if (field == null) return;
            // Get the Run that contains our FieldCode
            // Then get the parent container of this Run
            var rFldCode = (Run)field.Parent;

            // Get the three (3) other Runs that make up our merge field
            var rBegin = rFldCode.PreviousSibling<Run>();
            var rSep = rFldCode.NextSibling<Run>();
            var rText = rSep.NextSibling<Run>();
            var rEnd = rText.NextSibling<Run>();

            // Get the Run that holds the Text element for our merge field
            // Get the Text element and replace the text content 
            var t = rText.GetFirstChild<Text>();
            t.Text = replacementText;

            // Remove all the four (4) Runs for our merge field
            rFldCode.Remove();
            rBegin.Remove();
            rSep.Remove();
            rEnd.Remove();
        }
    }
}
