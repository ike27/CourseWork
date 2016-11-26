using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordTemplate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
    }

    public class TemplateData
    {
        List<string> institute = new List<string>();
        List<string> department = new List<string>();
        List<string> theme = new List<string>();
        List<string> code = new List<string>();
        List<string> specialization = new List<string>();
        List<string> section = new List<string>();
        List<string> student = new List<string>();
        List<string> head_dep_name = new List<string>();
        List<string> head_dep_degree = new List<string>();
        List<string> adviser_name = new List<string>();
        List<string> adviser_degree = new List<string>();
        List<int> year = new List<int>();
        
         


    }
}
