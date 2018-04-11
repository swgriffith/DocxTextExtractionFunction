using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using System.Collections.Generic;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using System.IO.Packaging;
using System;
using System.Xml;
using System.Linq;
using System.Text;

namespace DocxTextExtractionFunction
{
    public static class DocxToText
    {
        [FunctionName("DocxToText")]
        [return: Blob("output/{name}.txt", FileAccess.Write, Connection = "dropsite")]
        public static string Run([BlobTrigger("drop/{name}", Connection = "dropsite")]Stream myBlob, string name, TraceWriter log)
        {
            log.Info($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");

            MemoryStream ms = new MemoryStream();
            myBlob.CopyTo(ms);
            
            List<string> output = ProcessWordDoc(ms);

            //string outFileName = name.Substring(0, name.IndexOf('.'));
            return string.Join(Environment.NewLine,output);
        }
        

        private static List<string> ProcessWordDoc(MemoryStream fileBytes)
        {
            const string documentRelationshipType =
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            const string stylesRelationshipType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
            const string wordmlNamespace =
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace w = wordmlNamespace;

            XDocument xDoc = null;
            XDocument styleDoc = null;

            using (Package wdPackage = Package.Open(fileBytes, FileMode.Open, FileAccess.Read))
            {
                PackageRelationship docPackageRelationship =
                  wdPackage
                  .GetRelationshipsByType(documentRelationshipType)
                  .FirstOrDefault();
                if (docPackageRelationship != null)
                {
                    Uri documentUri =
                        PackUriHelper
                        .ResolvePartUri(
                           new Uri("/", UriKind.Relative),
                                 docPackageRelationship.TargetUri);
                    PackagePart documentPart =
                        wdPackage.GetPart(documentUri);

                    //  Load the document XML in the part into an XDocument instance.  
                    xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                    //  Find the styles part. There will only be one.  
                    PackageRelationship styleRelation =
                      documentPart.GetRelationshipsByType(stylesRelationshipType)
                      .FirstOrDefault();
                    if (styleRelation != null)
                    {
                        Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);
                        PackagePart stylePart = wdPackage.GetPart(styleUri);

                        //  Load the style XML in the part into an XDocument instance.  
                        styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));
                    }
                }
            }

            string defaultStyle =
                (string)(
                    from style in styleDoc.Root.Elements(w + "style")
                    where (string)style.Attribute(w + "type") == "paragraph" &&
                          (string)style.Attribute(w + "default") == "1"
                    select style
                ).First().Attribute(w + "styleId");

            // Find all paragraphs in the document.  
            var paragraphs =
                from para in xDoc
                             .Root
                             .Element(w + "body")
                             .Descendants(w + "p")
                let styleNode = para
                                .Elements(w + "pPr")
                                .Elements(w + "pStyle")
                                .FirstOrDefault()
                select new
                {
                    ParagraphNode = para,
                    StyleName = styleNode != null ?
                        (string)styleNode.Attribute(w + "val") :
                        defaultStyle
                };

            // Retrieve the text of each paragraph.  
            var paraWithText =
                from para in paragraphs
                select new
                {
                    ParagraphNode = para.ParagraphNode,
                    StyleName = para.StyleName,
                    Text = ParagraphText(para.ParagraphNode)
                };

            List<string> array = new List<string>();

            foreach (var para in paraWithText)
            {
                array.Add(para.Text);

            }

            return array;
        }

        private static string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            return e
                   .Elements(w + "r")
                   .Elements(w + "t")
                   .StringConcatenate(element => (string)element);
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source,
    Func<T, string> func)
        {
            StringBuilder sb = new StringBuilder();
            foreach (T item in source)
                sb.Append(func(item));
            return sb.ToString();
        }

    }
}
