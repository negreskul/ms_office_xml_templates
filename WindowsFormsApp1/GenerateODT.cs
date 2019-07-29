using System.Collections.Generic;
using System.IO;
using System.Linq;
using AODL.Document.TextDocuments;
using AODL.Document.Content;
using System.Xml;
using System.Windows.Forms;
namespace WindowsFormsApp1
{
    class GenerateODT
    {
        public static void GenerateDocument(string xmlData, string template, string output)
        {
            if (File.Exists(output))
            {
                File.Delete(output);
            }
            TextDocument doc = new TextDocument();
            doc.Load(template);
            IEnumerable<IContent> content = doc.Content.Cast<IContent>();
            XmlNode docNode = doc.XmlDoc.DocumentElement.LastChild.PreviousSibling.LastChild;
            XmlDocument newXml = new XmlDocument();
            newXml.Load(xmlData);
            XmlNode newNode = newXml.DocumentElement.LastChild.LastChild;
            XmlNode importNode = docNode.OwnerDocument.ImportNode(newNode, true);
            doc.Content.Clear();
            doc.XmlDoc.DocumentElement.LastChild.PreviousSibling.ReplaceChild(importNode, docNode);
            doc.SaveTo(output);
            MessageBox.Show("Document successfully generated.", "Complete");
        }
    }
}