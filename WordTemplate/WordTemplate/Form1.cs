using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.Serialization.Formatters;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace WordTemplate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private TemplateData tdata = new TemplateData();
        private string _contName = "";
        private string _sdtPropId = "";

        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox1.Text != "") tdata.institute.Add(textBox1.Text);
            if (textBox2.Text != "") tdata.department.Add(textBox2.Text);
            if (textBox3.Text != "") tdata.theme.Add(textBox3.Text);
            if (textBox4.Text != "") tdata.code.Add(textBox4.Text);
            if (textBox5.Text != "") tdata.specialization.Add(textBox5.Text);
            if (textBox6.Text != "") tdata.section.Add(textBox6.Text);
            if (textBox7.Text != "") tdata.student.Add(textBox7.Text);
            if (textBox8.Text != "") tdata.head_dep_name.Add(textBox8.Text);
            if (textBox9.Text != "") tdata.head_dep_degree.Add(textBox9.Text);
            if (textBox10.Text != "") tdata.adviser_name.Add(textBox10.Text);
            if (textBox11.Text != "") tdata.adviser_degree.Add(textBox11.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (string t in tdata.theme)
            {
                MessageBox.Show(t);
            }

            BinaryFormatter formatter = new BinaryFormatter();

            using (FileStream fs = new FileStream("template_data.dat", FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, tdata);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            BinaryFormatter formatter = new BinaryFormatter();
            using (FileStream fs = new FileStream("template_data.dat", FileMode.OpenOrCreate))
            {
                tdata = (TemplateData)formatter.Deserialize(fs);
            }
            foreach (string t in tdata.theme)
            {
                MessageBox.Show(t);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open("C:\\Users\\ikega\\Desktop\\ООП\\kurs1.docx", true)){ 
            foreach (var item in wordDoc.MainDocumentPart.Document.Body)
            {
                var oo = item.Descendants<SdtProperties>();
                foreach (var f1 in oo)
                {
                    Tag tag = f1.Elements<Tag>().FirstOrDefault();
                    string _tag = "";

                    if (tag != null) _tag = tag.Val;

                    if (_tag.Contains("_institute")) // #1
                    {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox1.Text;
                    }
                    if (_tag.Contains("_department")) // #2
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox2.Text;
                        }
                    if (_tag.Contains("_theme")) // #3
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox3.Text;
                        }
                    if (_tag.Contains("_code")) // #4
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox4.Text;
                        }
                    if (_tag.Contains("_specialization")) // #5
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox5.Text;
                    }
                    if (_tag.Contains("_section")) // #6
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox6.Text;
                    }
                    if (_tag.Contains("_student")) // #7
                    {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox7.Text;
                    }
                    if (_tag.Contains("_head_dep_name")) // #8
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox8.Text;
                    }
                    if (_tag.Contains("_head_dep_degree")) // #9
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox9.Text;
                    }
                    if (_tag.Contains("_adviser_name")) // #10
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox10.Text;
                    }
                    if (_tag.Contains("_adviser_degree")) // #11
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox11.Text;
                    }
                    if (_tag.Contains("_year")) // #12
                        {
                        SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == _tag);
                        element.Descendants<Text>().First().Text = textBox12.Text;
                    }
                }
            }
            wordDoc.Close();
            }
        }

        [Serializable]
        public class TemplateData
        {
            public List<string> institute = new List<string>();
            public List<string> department = new List<string>();
            public List<string> theme = new List<string>();
            public List<string> code = new List<string>();
            public List<string> specialization = new List<string>();
            public List<string> section = new List<string>();
            public List<string> student = new List<string>();
            public List<string> head_dep_name = new List<string>();
            public List<string> head_dep_degree = new List<string>();
            public List<string> adviser_name = new List<string>();
            public List<string> adviser_degree = new List<string>();
        }

    }
}


