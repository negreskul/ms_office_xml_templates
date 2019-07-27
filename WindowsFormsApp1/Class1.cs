using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using WindowsFormsApp1;
using System.Windows.Forms;

namespace ConsoleApp1
{
    class Class1
    {
        public static void GenerateDocument(string xmlData, string template, string output)
        {
            if (File.Exists(output))
            {
                File.Delete(output);
            }
            File.Copy(template, output);
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(output, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                mainPart.DeleteParts(mainPart.CustomXmlParts);
                CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using (FileStream stream = new FileStream(xmlData, FileMode.Open))
                {
                    myXmlPart.FeedData(stream);
                }
            }
            MessageBox.Show("Document successfully generated.", "Complete");
        }
    }
}
