using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace WordMailMerge
{
    public class Main
    {
        public void DoMerge()
        {
            const string sourceFile = @"E:\Projects\tests\WordMailMerge\MergeForm.docx";
            const string targetFile = @"E:\Projects\tests\WordMailMerge\MergedForm.docx";

            File.Copy(sourceFile, targetFile, true);

            using (var document = WordprocessingDocument.Open(targetFile, true))
            {
                // If your sourceFile is a different type (e.g., .DOTX), you will need to change the target type like so:
                document.ChangeDocumentType(WordprocessingDocumentType.Document);

                // Get the MainPart of the document
                var mainPart = document.MainDocumentPart;

                var forEachData = new Dictionary<string, List<Dictionary<string, string>>>
                {
                    {
                        "Infringements",
                        new List<Dictionary<string, string>> {
                            new Dictionary<string, string>
                            {
                                { "InfringementText", "Some sort of infringement E:\\Projects\\tests\\WordMailMerge\\test.jpg of the standards." },
                                { "ActionRequired", "Do something about this thing." },
                                { "Photo", "E:\\Projects\\tests\\WordMailMerge\\test.jpg" }
                            },
                            new Dictionary<string, string>
                            {
                                { "InfringementText", "Some other sort of infringement of the standards." },
                                { "ActionRequired", "Do something else about this thing." },
                                { "Photo", "E:\\Projects\\tests\\WordMailMerge\\test2.jpg" }
                                //{ "Photo", "http://intcdn.telemetry.aws/Images/Drop/header.png" }
                            },
                            new Dictionary<string, string>
                            {
                                { "InfringementText", "Some other other sort of infringement of the standards." },
                                { "ActionRequired", "You seem to enjoy infringing." }
                            }
                        }
                    },
                    {
                        "Something",
                        new List<Dictionary<string, string>> {
                            new Dictionary<string, string>
                            {
                                { "Someprop", "a document property to be replaced." }
                            },
                            new Dictionary<string, string>
                            {
                                { "Someprop", "another document property to be replaced." }
                            }
                        }
                    }
                };

                var mergeData = new Dictionary<string, string> { { "Property", "The Property, 1 The Street" }, { "Ref", "AWPREF00001" } };

                ReplaceFont(document, "Consolas", "Bauhaus 93");

                MergeForEach(document, GetForEachFields(mainPart.RootElement), forEachData);

                RemoveForEachFields(mainPart.RootElement);

                ReplaceMergeFields(document, mainPart.RootElement, mergeData);

                // Save the document
                mainPart.Document.Save();
            }
        }

        private void ReplaceFont(WordprocessingDocument document, string fontFrom, string fontTo)
        {
            var fonts = document.MainDocumentPart.RootElement.Descendants<RunFonts>().Where(runFonts => runFonts.Ascii == fontFrom);
            foreach (var font in fonts)
            {
                font.Ascii = fontTo;
                font.HighAnsi = fontTo;
                font.ComplexScript = fontTo;
            }
        }

        private void MergeForEach(WordprocessingDocument wordprocessingDocument, IEnumerable<FieldCode> repeated, Dictionary<string, List<Dictionary<string, string>>> forEachData)
        {
            foreach (var repeat in repeated)
            {
                var name = GetFieldName(repeat);

                OpenXmlElement container = repeat.Parent as Table
                    ?? repeat?.Parent?.Parent as Table
                    ?? repeat?.Parent?.Parent?.Parent as Table
                    ?? repeat?.Parent?.Parent?.Parent?.Parent as Table
                    ?? repeat?.Parent?.Parent?.Parent?.Parent?.Parent as Table
                    ?? repeat?.Parent?.Parent?.Parent?.Parent?.Parent?.Parent as Table;

                OpenXmlElement template = repeat.Parent as TableRow
                    ?? repeat?.Parent?.Parent as TableRow
                    ?? repeat?.Parent?.Parent?.Parent as TableRow
                    ?? repeat?.Parent?.Parent?.Parent?.Parent as TableRow
                    ?? repeat?.Parent?.Parent?.Parent?.Parent?.Parent as TableRow;

                if (container == null || template == null)
                {
                    container = repeat.Parent.Parent.Parent;
                    template = repeat.Parent.Parent;
                }

                foreach (var datum in forEachData[name])
                {
                    ProcessTemplateAndAppendToContainer(wordprocessingDocument, container, template, datum);
                }

                container.RemoveChild(template);
            }
        }

        private string GetFieldName(FieldCode field)
        {
            // should be one of these forms:
            // MERGEFIELD COMMAND:name
            // MERGEFIELD \"COMMAND:name\"
            // MERGEFIELD  COMMAND:name
            // MERGEFIELD name
            // MERGEFIELD \"name\"
            // MERGEFIELD  name

            var fieldText = field.InnerText.Replace("MERGEFIELD", "").Replace(" ", "");
            return fieldText.Contains(":") ? fieldText.Split(':')[1] : fieldText;
        }

        private void ProcessTemplateAndAppendToContainer(WordprocessingDocument wordprocessingDocument, OpenXmlElement container, OpenXmlElement template, Dictionary<string, string> data)
        {
            var templateClone = template.CloneNode(true);

            ReplaceMergeFields(wordprocessingDocument, templateClone, data);

            container.AppendChild(templateClone);
        }

        private IEnumerable<FieldCode> GetMergeFields(OpenXmlElement element)
        {
            return element.Descendants<FieldCode>().Where(code => code.InnerText.Contains("MERGEFIELD") && !code.InnerText.Contains("FOREACH:"));
        }

        private IEnumerable<FieldCode> GetForEachFields(OpenXmlElement element)
        {
            return element.Descendants<FieldCode>().Where(code => code.InnerText.Contains("MERGEFIELD") && code.InnerText.Contains("FOREACH:"));
        }

        /// <summary>
        /// </summary>
        /// <example> Merge Field XML
        ////<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        ////    <w:rPr>
        ////        <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" />
        ////        <w:sz w:val="20" />
        ////        <w:szCs w:val="20" />
        ////    </w:rPr>
        ////    <w:fldChar w:fldCharType="begin" />
        ////</w:r>
        ////<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        ////    <w:rPr>
        ////        <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" />
        ////        <w:sz w:val="20" />
        ////        <w:szCs w:val="20" />
        ////    </w:rPr>
        ////    <w:instrText xml:space="preserve"> MERGEFIELD FOREACH:Infringements</w:instrText>
        ////</w:r>
        ////<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        ////    <w:rPr>
        ////        <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" />
        ////        <w:sz w:val="20" />
        ////        <w:szCs w:val="20" />
        ////    </w:rPr>
        ////    <w:fldChar w:fldCharType="separate" />
        ////</w:r>
        ////<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        ////    <w:rPr>
        ////        <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" />
        ////        <w:noProof />
        ////        <w:sz w:val="20" />
        ////        <w:szCs w:val="20" />
        ////    </w:rPr>
        ////    <w:t>«FOREACH:Infringements»</w:t>
        ////</w:r>
        ////<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        ////    <w:rPr>
        ////        <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" />
        ////        <w:sz w:val="20" />
        ////        <w:szCs w:val="20" />
        ////    </w:rPr>
        ////    <w:fldChar w:fldCharType="end" />
        ////</w:r>
        /// </example>
        /// <param name="element"></param>
        /// <param name="data"></param>
        private void ReplaceMergeFields(WordprocessingDocument wordprocessingDocument, OpenXmlElement element, Dictionary<string, string> data)
        {
            var fields = GetMergeFields(element);

            foreach (var field in fields.ToList())
            {
                var fieldname = GetFieldName(field);

                if (data.ContainsKey(fieldname))
                {
                    if (field.InnerText.Contains("IMAGE:"))
                    {
                        ReplaceMergeFieldImage(wordprocessingDocument, field, data[fieldname]);
                    }
                    else
                    {
                        ReplaceMergeField(field, data[fieldname]);
                    }
                }
                else
                {
                    RemoveMergeField(field);
                }
            }
        }

        private void RemoveForEachFields(OpenXmlElement element)
        {
            var fields = GetForEachFields(element);

            foreach (var field in fields.ToList())
            {
                RemoveMergeField(field);
            }
        }

        private void ReplaceMergeField(FieldCode mergeField, string replacementText)
        {
            // Get the Run that contains our FieldCode
            // Then get the parent container of this Run
            var fieldParent = (Run)mergeField.Parent;

            // Get the three (3) other Runs that make up our merge field
            var begin = fieldParent.PreviousSibling<Run>();
            var separator = fieldParent.NextSibling<Run>();
            var text = separator.NextSibling<Run>();
            var end = text.NextSibling<Run>();

            // Get the Run that holds the Text element for our merge field
            // Get the Text element and replace the text content 
            var t = text.GetFirstChild<Text>();
            t.Text = replacementText;

            // Remove all the four (4) Runs for our merge field
            fieldParent.Remove();
            begin.Remove();
            separator.Remove();
            end.Remove();
        }

        private void ReplaceMergeFieldImage(WordprocessingDocument wordprocessingDocument, FieldCode mergeField, string imageUri)
        {
            var fieldParent = mergeField.Parent.Parent;

            ImagePart imagePart = wordprocessingDocument.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(imageUri, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            RemoveMergeField(mergeField);

            fieldParent.AppendChild(new Paragraph());
            fieldParent.AppendChild(new Paragraph(new Run(CreateImageElement(wordprocessingDocument.MainDocumentPart.GetIdOfPart(imagePart)))));
        }

        private void RemoveMergeField(FieldCode mergeField)
        {
            // Get the Run that contains our FieldCode
            var fieldParent = (Run)mergeField.Parent;

            // Get the other Runs that make up our merge field
            var begin = fieldParent.PreviousSibling<Run>();
            var separator = fieldParent.NextSibling<Run>();
            var text = separator.NextSibling<Run>();
            var end = text.NextSibling<Run>();

            // Remove all Runs for our merge field
            begin.Remove();
            fieldParent.Remove();
            separator.Remove();
            text.Remove();
            end.Remove();
        }

        private static OpenXmlElement CreateImageElement(string relationshipId)
        {
            // Define the reference of the image.
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = 990000L, Cy = 792000L },
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value)1U,
                            Name = "Picture 1"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                    new PIC.Picture(
                                        new PIC.NonVisualPictureProperties(
                                            new PIC.NonVisualDrawingProperties()
                                            {
                                                Id = (UInt32Value)0U,
                                                Name = "New Bitmap Image.jpg"
                                            },
                                            new PIC.NonVisualPictureDrawingProperties()),
                                        new PIC.BlipFill(
                                            new A.Blip(
                                                new A.BlipExtensionList(
                                                    new A.BlipExtension()
                                                    {
                                                        Uri =
                                                            "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                    })
                                            )
                                            {
                                                Embed = relationshipId,
                                                CompressionState =
                                                    A.BlipCompressionValues.Print
                                            },
                                            new A.Stretch(
                                                new A.FillRectangle())),
                                        new PIC.ShapeProperties(
                                            new A.Transform2D(
                                                new A.Offset() { X = 0L, Y = 0L },
                                                new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                            new A.PresetGeometry(
                                                    new A.AdjustValueList()
                                                )
                                                { Preset = A.ShapeTypeValues.Rectangle }))
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U,
                        EditId = "50D07946"
                    });

            // Append the reference to body, the element should be in a Run.
            return element; //wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
    }
}
