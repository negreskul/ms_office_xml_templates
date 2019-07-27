using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Class2.GenerateXml();
            Class1.GenerateDocument("doc.xml", "XMLThings.docx", "new_file.docx");
        }
    }
}
