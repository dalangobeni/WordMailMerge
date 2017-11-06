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
            var infringements = new List<Infringement>()
            {
                new Infringement{InfringementText= "You have done this thing incorrectly", ActionRequired = "Do this thing"},
                new Infringement{InfringementText= "You have done this other thing incorrectly", ActionRequired = "Do this other thing"}
            }; 

            const string mergeFieldName = "Property";
            const string replacementText = "The Property, 1 The Street";

            File.Copy(sourceFile, targetFile, true);

            using (var document = WordprocessingDocument.Open(targetFile, true))
            {
                // If your sourceFile is a different type (e.g., .DOTX), you will need to change the target type like so:
                document.ChangeDocumentType(WordprocessingDocumentType.Document);

                // Get the MainPart of the document
                var mainPart = document.MainDocumentPart;
                var customComponents = mainPart.RootElement.Descendants<SdtBlock>().Where(block => block.SdtProperties.GetFirstChild<Tag>().Val != "");
                var mergeFields = mainPart.RootElement.Descendants<FieldCode>();

                ReplaceMergeFieldWithText(mergeFields, mergeFieldName, replacementText);

                foreach (var customComponent in customComponents)
                {
                    var source = customComponent.SdtProperties.GetFirstChild<Tag>().Val;
                    if (source == "Infringements")
                    {
                        var table = customComponent.Descendants<Table>().Single();
                        var rowTemplate = table.Descendants<TableRow>().Last();

                        foreach (var infringement in infringements)
                        {
                            var rowCopy = rowTemplate.CloneNode(true) as TableRow;

                            var rowMergeFields = rowCopy.Descendants<FieldCode>();
                            ReplaceMergeFieldWithText(rowMergeFields, "InfringementText", infringement.InfringementText);
                            ReplaceMergeFieldWithText(rowMergeFields, "ActionRequired", infringement.ActionRequired);

                            table.AppendChild(rowCopy);
                        }

                        table.RemoveChild(rowTemplate);
                    }
                }

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

    public class Infringement
    {
        public string InfringementText { get; set; }
        public string ActionRequired { get; set; }
    }
}
