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
            if (textBox2.Text != "") tdata.department.Add(textBox1.Text);
            if (textBox3.Text != "") tdata.theme.Add(textBox1.Text);
            if (textBox4.Text != "") tdata.code.Add(textBox1.Text);
            if (textBox5.Text != "") tdata.specialization.Add(textBox1.Text);
            if (textBox6.Text != "") tdata.section.Add(textBox1.Text);
            if (textBox7.Text != "") tdata.student.Add(textBox1.Text);
            if (textBox8.Text != "") tdata.head_dep_name.Add(textBox1.Text);
            if (textBox9.Text != "") tdata.head_dep_degree.Add(textBox1.Text);
            if (textBox10.Text != "") tdata.adviser_name.Add(textBox1.Text);
            if (textBox11.Text != "") tdata.adviser_degree.Add(textBox1.Text);
            foreach (string t in tdata.theme)
            {
                Console.WriteLine(t);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            //    TemplateData tdata = new TemplateData();

            /*    if 

                BinaryFormatter formatter = new BinaryFormatter();

                using (FileStream fs = new FileStream("people.dat", FileMode.OpenOrCreate))
                {
                    // сериализуем весь массив people
                    formatter.Serialize(fs, people);

                    Console.WriteLine("Объект сериализован");
                }

                // десериализация
                using (FileStream fs = new FileStream("people.dat", FileMode.OpenOrCreate))
                {
                    Person[] deserilizePeople = (Person[])formatter.Deserialize(fs);

                    foreach (Person p in deserilizePeople)
                    {
                        Console.WriteLine("Имя: {0} --- Возраст: {1}", p.Name, p.Age);
                    }
                }*/
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
