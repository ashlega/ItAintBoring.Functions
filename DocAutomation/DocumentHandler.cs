using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

//Based on the Andrew Best post here: https://www.andrew-best.com/posts/adventures-in-azure-functions/

namespace ItAintBoring.Functions
{
    public class DocumentHandler 
    {
        public string Read(Stream content)
        {
            var xmlDoc = WordprocessingDocument.Open(content, false);
            var bodyElement = xmlDoc.MainDocumentPart.Document.Body;
            var docText = GetPlainText(bodyElement);
            xmlDoc.Close();

            return docText;
        }

        public void Merge(Stream doc1, Stream doc2)
        {
            var wordDoc = WordprocessingDocument.Open(doc1, true);
            var mainPart = wordDoc.MainDocumentPart;
            var chunk = mainPart.AddAlternativeFormatImportPart(
                DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPartType.WordprocessingML);
            doc2.Position = 0;
            chunk.FeedData(doc2);
            var altChunkId = mainPart.GetIdOfPart(chunk);
            var altChunk = new DocumentFormat.OpenXml.Wordprocessing.AltChunk();
            altChunk.Id = altChunkId;
            Paragraph para = new Paragraph(new Run((new Break() { Type = BreakValues.Page })));
            mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);
            mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.LastChild);
            wordDoc.Close();
        }

        private static string GetPlainText(OpenXmlElement element)
        {
            var builder = new StringBuilder();
            foreach (var section in element.Elements())
            {
                switch (section.LocalName)
                {
                    case "t":
                        builder.Append(section.InnerText);
                        break;
                    case "cr":
                    case "br":
                        builder.Append(Environment.NewLine);
                        break;
                    case "tab":
                        builder.Append("\t");
                        break;
                    case "p":
                        builder.Append(GetPlainText(section));
                        builder.AppendLine(Environment.NewLine);
                        break;
                    default:
                        builder.Append(GetPlainText(section));
                        break;
                }
            }

            return builder.ToString();
        }
    }
}
