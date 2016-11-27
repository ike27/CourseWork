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

namespace WordTemplate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        TemplateData tdata = new TemplateData();

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
   //     public List<int> year = new List<int>();
    }
}
