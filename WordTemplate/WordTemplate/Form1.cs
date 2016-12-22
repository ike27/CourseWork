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
        private int[] search_index = new int[11];

        private void ComboBox_Update()
        {
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox9.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox10.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox11.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();
            comboBox6.Items.Clear();
            comboBox7.Items.Clear();
            comboBox8.Items.Clear();
            comboBox9.Items.Clear();
            comboBox10.Items.Clear();
            comboBox11.Items.Clear();
            for (int i = 0; i < tdata.institute.Count; i++) comboBox1.Items.Add(tdata.institute[i]);
            for (int i = 0; i < tdata.department.Count; i++) comboBox2.Items.Add(tdata.department[i]);
            for (int i = 0; i < tdata.theme.Count; i++) comboBox3.Items.Add(tdata.theme[i]);
            for (int i = 0; i < tdata.code.Count; i++) comboBox4.Items.Add(tdata.code[i]);
            for (int i = 0; i < tdata.specialization.Count; i++) comboBox5.Items.Add(tdata.specialization[i]);
            for (int i = 0; i < tdata.section.Count; i++) comboBox6.Items.Add(tdata.section[i]);
            for (int i = 0; i < tdata.student.Count; i++) comboBox7.Items.Add(tdata.student[i]);
            for (int i = 0; i < tdata.head_dep_name.Count; i++) comboBox8.Items.Add(tdata.head_dep_name[i]);
            for (int i = 0; i < tdata.head_dep_degree.Count; i++) comboBox9.Items.Add(tdata.head_dep_degree[i]);
            for (int i = 0; i < tdata.adviser_name.Count; i++) comboBox10.Items.Add(tdata.adviser_name[i]);
            for (int i = 0; i < tdata.adviser_degree.Count; i++) comboBox11.Items.Add(tdata.adviser_degree[i]);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text != "") && !tdata.institute.Contains(textBox1.Text)) tdata.institute.Add(textBox1.Text);
            if ((textBox2.Text != "") && !tdata.department.Contains(textBox2.Text)) tdata.department.Add(textBox2.Text);
            if ((textBox3.Text != "") && !tdata.theme.Contains(textBox3.Text)) tdata.theme.Add(textBox3.Text);
            if ((textBox4.Text != "") && !tdata.code.Contains(textBox4.Text)) tdata.code.Add(textBox4.Text);
            if ((textBox5.Text != "") && !tdata.specialization.Contains(textBox5.Text)) tdata.specialization.Add(textBox5.Text);
            if ((textBox6.Text != "") && !tdata.section.Contains(textBox6.Text)) tdata.section.Add(textBox6.Text);
            if ((textBox7.Text != "") && !tdata.student.Contains(textBox7.Text)) tdata.student.Add(textBox7.Text);
            if ((textBox8.Text != "") && !tdata.head_dep_name.Contains(textBox8.Text)) tdata.head_dep_name.Add(textBox8.Text);
            if ((textBox9.Text != "") && !tdata.head_dep_degree.Contains(textBox9.Text)) tdata.head_dep_degree.Add(textBox9.Text);
            if ((textBox10.Text != "") && !tdata.adviser_name.Contains(textBox10.Text)) tdata.adviser_name.Add(textBox10.Text);
            if ((textBox11.Text != "") && !tdata.adviser_degree.Contains(textBox11.Text)) tdata.adviser_degree.Add(textBox11.Text);
            ComboBox_Update();
        }

        private void button2_Click(object sender, EventArgs e)
        {
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
            ComboBox_Update();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            ComboBox_Update();
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;
            comboBox6.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;
            comboBox8.SelectedIndex = 0;
            comboBox9.SelectedIndex = 0;
            comboBox10.SelectedIndex = 0;
            comboBox11.SelectedIndex = 0;
            DateTime now = DateTime.Now;
            textBox13.Text = now.ToString("yyyy");
        }

        private void buttonTRY_Click_1(object sender, EventArgs e)
        {
            richTextBox1.LoadFile("template_vkr_bachelor.rtf");
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open("template_vkr_bachelor.docx", true))
            {
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
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if(comboBox1.SelectedIndex!=-1) element.Descendants<Text>().First().Text = comboBox1.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_department")) // #2
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox2.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_theme")) // #3
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox3.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_code")) // #4
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox4.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_specialization")) // #5
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox5.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_section")) // #6
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox6.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_student")) // #7
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox7.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_head_dep_name")) // #8
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox8.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_head_dep_degree")) // #9
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox9.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_adviser_name")) // #10
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox10.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_adviser_degree")) // #11
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = comboBox11.SelectedItem.ToString();
                        }
                        if (_tag.Contains("_year")) // #12
                        {
                            SdtElement element = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>().FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>() != null ? sdt.SdtProperties.GetFirstChild<Tag>().Val == _tag : false);
                            if (comboBox1.SelectedIndex != -1) element.Descendants<Text>().First().Text = textBox13.Text;
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

        private void button5_Click_1(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            List<string> result1 = new List<string>();
            result1 = tdata.institute.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result1.Count; i++)
            {
                ListViewItem item = new ListViewItem("Институт");
                item.SubItems.Add(result1[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[0] = result1.Count;

            List<string> result2 = new List<string>();
            result2 = tdata.department.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result2.Count; i++)
            {
                ListViewItem item = new ListViewItem("Кафедра");
                item.SubItems.Add(result2[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[1] = search_index[0] + result2.Count;

            List<string> result3 = new List<string>();
            result3 = tdata.theme.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result3.Count; i++)
            {
                ListViewItem item = new ListViewItem("Тема ВКР");
                item.SubItems.Add(result3[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[2] = search_index[1] + result3.Count;

            List<string> result4 = new List<string>();
            result4 = tdata.code.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result4.Count; i++)
            {
                ListViewItem item = new ListViewItem("Номер направления");
                item.SubItems.Add(result4[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[3] = search_index[2] + result4.Count;

            List<string> result5 = new List<string>();
            result5 = tdata.specialization.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result5.Count; i++)
            {
                ListViewItem item = new ListViewItem("Направление");
                item.SubItems.Add(result5[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[4] = search_index[3] + result5.Count;

            List<string> result6 = new List<string>();
            result6 = tdata.section.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result6.Count; i++)
            {
                ListViewItem item = new ListViewItem("Профиль подготовки");
                item.SubItems.Add(result6[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[5] = search_index[4] + result6.Count;

            List<string> result7 = new List<string>();
            result7 = tdata.student.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result7.Count; i++)
            {
                ListViewItem item = new ListViewItem("ФИО студента");
                item.SubItems.Add(result7[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[6] = search_index[5] + result7.Count;

            List<string> result8 = new List<string>();
            result8 = tdata.head_dep_name.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result8.Count; i++)
            {
                ListViewItem item = new ListViewItem("ФИО завед. кафедрой");
                item.SubItems.Add(result8[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[7] = search_index[6] + result8.Count;

            List<string> result9 = new List<string>();
            result9 = tdata.head_dep_degree.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result9.Count; i++)
            {
                ListViewItem item = new ListViewItem("Научная степень завед. кафедрой");
                item.SubItems.Add(result9[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[8] = search_index[7] + result9.Count;

            List<string> result10 = new List<string>();
            result10 = tdata.adviser_name.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result10.Count; i++)
            {
                ListViewItem item = new ListViewItem("ФИО научного руководителя");
                item.SubItems.Add(result10[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[9] = search_index[8] + result10.Count;

            List<string> result11 = new List<string>();
            result11 = tdata.adviser_degree.Where(t => t.Contains(textBox12.Text)).OrderBy(t => t).ToList();
            for (int i = 0; i < result11.Count; i++)
            {
                ListViewItem item = new ListViewItem("Научная степень науч. руковод.");
                item.SubItems.Add(result11[i]);
                listView1.Items.AddRange(new ListViewItem[] { item });
            }
            search_index[10] = search_index[9] + result11.Count;
        }

        private void listView1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                int line = listView1.SelectedIndices[0];
                int group = -1;
                for (int i = 0; i < 11; i++)
                {
                    if (line + 1 <= search_index[i])
                    {
                        group = i;
                        break;
                    }
                }
                switch (group)
                {
                    case 0:
                        for (int i = 0; i < comboBox1.Items.Count; i++)
                        {
                            if (comboBox1.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox1.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: Институт");
                            }
                        }
                        break;
                    case 1:
                        for (int i = 0; i < comboBox2.Items.Count; i++)
                        {
                            if (comboBox2.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox2.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: Кафедра");
                            }
                        }
                        break;
                    case 2:
                        for (int i = 0; i < comboBox3.Items.Count; i++)
                        {
                            if (comboBox3.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox3.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: Тема ВКР");
                            }
                        }
                        break;
                    case 3:
                        for (int i = 0; i < comboBox4.Items.Count; i++)
                        {
                            if (comboBox4.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox4.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: Номер направления");
                            }
                        }
                        break;
                    case 4:
                        for (int i = 0; i < comboBox5.Items.Count; i++)
                        {
                            if (comboBox5.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox5.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: Направление");
                            }
                        }
                        break;
                    case 5:
                        for (int i = 0; i < comboBox6.Items.Count; i++)
                        {
                            if (comboBox6.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox6.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: Профиль подготовки");
                            }
                        }
                        break;
                    case 6:
                        for (int i = 0; i < comboBox7.Items.Count; i++)
                        {
                            if (comboBox7.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox7.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: ФИО студента");
                            }
                        }
                        break;
                    case 7:
                        for (int i = 0; i < comboBox8.Items.Count; i++)
                        {
                            if (comboBox8.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox8.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: ФИО завед. кафедрой");
                            }
                        }
                        break;
                    case 8:
                        for (int i = 0; i < comboBox9.Items.Count; i++)
                        {
                            if (comboBox9.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox9.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: Научная степень завед. кафедрой");
                            }
                        }
                        break;
                    case 9:
                        for (int i = 0; i < comboBox10.Items.Count; i++)
                        {
                            if (comboBox10.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox10.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: ФИО научного руководителя");
                            }
                        }
                        break;
                    case 10:
                        for (int i = 0; i < comboBox11.Items.Count; i++)
                        {
                            if (comboBox11.Items[i].ToString() == listView1.Items[line].SubItems[1].Text)
                            {
                                comboBox11.SelectedIndex = i;
                                MessageBox.Show("Изменено значение для категории: Научная степень научного руководителя");
                            }
                        }
                        break;
                    default:
                        MessageBox.Show("Что-то пошло не так :(");
                        break;
                }
            }
        }
    }
}