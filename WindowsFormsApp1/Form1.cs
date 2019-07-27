using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConsoleApp1;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        string xmlData;
        string template;
        public Form1()
        {
            InitializeComponent();
            Class2.GenerateXml();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "XML Files (*.xml)|*.xml";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.DefaultExt = "xml";
            bool t = true;
            while (t)
            {
                t = false;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (!String.Equals(Path.GetExtension(openFileDialog1.FileName),
                                       ".xml",
                                       StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show("The type of the selected file is not supported by this application. You must select an XML file.",
                                        "Invalid File Type",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                        t = true;
                    }
                    else
                    {
                        xmlData = openFileDialog1.FileName;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Word File (.docx)|*.docx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.DefaultExt = "docx";
            bool t = true;
            while (t)
            {
                t = false;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (!String.Equals(Path.GetExtension(openFileDialog1.FileName),
                                       ".docx",
                                       StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show("The type of the selected file is not supported by this application. You must select a .docx file.",
                                        "Invalid File Type",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                        t = true;
                    }
                    else
                    {
                        template = openFileDialog1.FileName;
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Class1.GenerateDocument(xmlData, template, "new_file.docx");
        }
    }
}
