using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp1
{
    class Class2
    {
        public static void GenerateXml()
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>
                <mydata xmlns = ""http://CustomDemoXML.htm"">
                    <Contacts> 
                        <Contact> 
                        <Name>NEW NAME</Name> 
                        <Phone Type=""Home"">NEW HOME NUMBER</Phone> 
                        <Phone Type=""Work"">N E W  W O R K  N U M B E R</Phone> 
                        <Phone Type=""Mobile"">mobileeeeeeee</Phone> 
                        <Address> 
                            <Street1>123 Main St</Street1> 
                            <City>Mercer Island</City> 
                            <State>WA</State> 
                            <Postal>68042</Postal> 
                        </Address> 
                        </Contact> 
                        <Contact> 
                        <Name>Dorothy Lee</Name> 
                        <Phone Type=""Home"">910-555-1212</Phone> 
                        <Phone Type=""Work"">336-555-0123</Phone> 
                        <Phone Type=""Mobile"">336-555-0005</Phone> 
                        <Address> 
                            <Street1>ленинский проспект дом 4 ска</Street1> 
                            <City>москва</City> 
                            <State>россия</State> 
                            <Postal>хз</Postal> 
                        </Address> 
                        </Contact>
                    </Contacts>
                </mydata>";
            File.WriteAllText("doc.xml", xml);
        }

        public static void GenerateEmptyDocx()
        {
            string path = "file.docx";            
            using (MemoryStream mem = new MemoryStream())
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    mainPart.Document = new Document();
                    Body docBody = new Body();

                }
                using (FileStream fs = File.Create(path))
                {
                    mem.Position = 0;
                    mem.CopyTo(fs);
                }
            }                
        }
    }
}
