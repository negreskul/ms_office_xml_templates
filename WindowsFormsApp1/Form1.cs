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

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        string xmlData;
        string template;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "XML Files (*.xml)|*.xml";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.FileName = null;
            bool t = true;
            while (t)
            {
                t = false;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    checkBox1.Checked = true;
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
                else
                {
                    checkBox1.Checked = false;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Word File (*.docx)|*.docx|OpenDocument Text File (*.odt)|*.odt";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.FileName = null;
            bool t = true;
            while (t)
            {
                t = false;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    checkBox2.Checked = true;
                    if ((!String.Equals(Path.GetExtension(openFileDialog1.FileName),
                                       ".docx",
                                       StringComparison.OrdinalIgnoreCase)) && (!String.Equals(Path.GetExtension(openFileDialog1.FileName),
                                       ".odt",
                                       StringComparison.OrdinalIgnoreCase)))
                    {
                        MessageBox.Show("The type of the selected file is not supported by this application. You must select either a .docx file or an .odt file.",
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
                else
                {
                    checkBox2.Checked = false;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (String.Equals(Path.GetExtension(openFileDialog1.FileName),
                                       ".docx",
                                       StringComparison.OrdinalIgnoreCase))
            {
                GenerateDOCX.GenerateDocument(xmlData, template, "new_file.docx");
            }
            else
            {
                GenerateODT.GenerateDocument(xmlData, template, "new_file.odt");
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if  ((checkBox1.Checked == true) && (checkBox2.Checked == true))
            {
                button3.Enabled = true;
            }
            else
            {
                button3.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if ((checkBox1.Checked == true) && (checkBox2.Checked == true))
            {
                button3.Enabled = true;
            }
            else
            {
                button3.Enabled = false;
            }
        }
    }
}
