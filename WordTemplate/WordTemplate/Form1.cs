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
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open("C:\\Users\\ikega\\Desktop\\ООП", false))
            {
                foreach (var item in wordDoc.MainDocumentPart.Document.Body)
                {
                    var oo = item.Descendants<SdtProperties>();
                    foreach (var f1 in oo)
                    {
                        _contName = FindPictureContainer(wordDoc, f1, ref _sdtPropId);
                    }
                }
            }
        }

   
        private string FindPictureContainer(WordprocessingDocument wdDoc, OpenXmlElement uy, ref string SdtId)
        {
            SdtAlias alias = uy.Elements<SdtAlias>().FirstOrDefault();
            SdtId sdtId = uy.Elements<SdtId>().FirstOrDefault();
            Tag tag = uy.Elements<Tag>().FirstOrDefault();

            string _tag = "";
            string _sdtId = "";
            string _alias = "";

            //Получаем тег контейнера
            if (tag != null)
                _tag = tag.Val;

            //Получаем ID контейнера
            if (sdtId != null)
                SdtId = _sdtId = sdtId.Val;

            //Получаем название контейнера
            if (alias != null)
                _alias = alias.Val;

            if (_tag.Contains("theme"))
            {
                
                    var sdtBlock = wdDoc.MainDocumentPart.Document.Descendants<SdtBlock>()
                                .Where(r => r.SdtProperties.GetFirstChild<SdtId>().Val == _sdtId);
                   
                
                
            }
            return _alias;
        }


        #region Image methods
        public static ImagePartType GetImagePartTypeFromFileName(string fileName)
        {
            ImagePartType io;
            switch (Path.GetExtension(fileName.ToLower()))
            {
                case ".bmp":
                    io = ImagePartType.Bmp;
                    break;
                case ".jpeg":
                case ".jpg":
                    io = ImagePartType.Jpeg;
                    break;
                case ".png":
                    io = ImagePartType.Png;
                    break;
                default:
                    throw new Exception("Загружен неверный формат файла!");
            }
            return io;
        }
        private static void ResizePictureContainer(Drawing d, int originalWidth, int originalHeight, ref int maxWidth, ref int maxHeight)
        {
            Extent imageSizeProps = d.Descendants<Extent>().FirstOrDefault();

            if (imageSizeProps != null)
            {
                int imageWidthOr = (int)(imageSizeProps.Cx / 9525);
                int imageHeightOr = (int)(imageSizeProps.Cy / 9525);
                maxWidth = imageWidthOr;
                maxHeight = imageHeightOr;
                //Определяем соотношение сторон
                double aspectRatio = (double)originalWidth / (double)originalHeight;

                //Проверяем, что высота изображения больше разрешенной высоты
                int newHeight = (originalHeight > imageHeightOr) ? imageHeightOr : originalHeight;
                //Проверяем, что ширина изображения больше разрешенной ширины
                int newWidth = (originalWidth > imageWidthOr) ? imageWidthOr : originalWidth;
                //Вычисляем новую высоту или ширину в зависимости от соотношения сторон (полагаясь на ориентацию изображения)
                if ((newWidth == originalWidth) && (newHeight == originalHeight))
                {
                    //Если ширина больше, то ориентация книжная
                    if (newWidth > newHeight)
                    {
                        //Вычисляем новую высоту умножением ширины на соотношение сторон
                        newHeight = (int)(imageWidthOr / aspectRatio);
                        newWidth = imageWidthOr;
                        //в некторых случаях вычисленная высота может быть больше чем разрешенная 
                        //поэтому нужно подвести высоту к разрешенной и пересчитать ширину
                        if (newHeight > imageHeightOr)
                        {
                            newHeight = imageHeightOr;
                            newWidth = (int)(aspectRatio * newHeight);
                        }
                    }
                    else //ориентация портретная
                    {
                        //Вычисляем новую ширину умножением высоты на соотношение сторон
                        newWidth = (int)(aspectRatio * imageHeightOr);
                        newHeight = imageHeightOr;
                    }
                }
                else //Если исходное изображение меньше, чем контейнер
                {
                    if (newWidth > newHeight)
                    {
                        newHeight = (int)(newWidth / aspectRatio);
                        if (newHeight > imageHeightOr)
                        {
                            newHeight = imageHeightOr;
                            newWidth = (int)(aspectRatio * newHeight);
                        }
                    }
                    else
                    {
                        newWidth = (int)(aspectRatio * newHeight);
                    }
                }
                imageSizeProps.Cx = (long)(newWidth * 9525);
                imageSizeProps.Cy = (long)(newHeight * 9525);
            }

            Extents e2 = d.Descendants<Extents>().FirstOrDefault();

            long imageWidthEmu = (long)(originalWidth * 9525);
            long imageHeightEmu = (long)(originalHeight * 9525);
            if (e2 != null)
            {
                e2.Cx = imageWidthEmu;
                e2.Cy = imageHeightEmu;
            }
        }
        #endregion

       


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

    public class WorkWithDock
    {
        public static void CreateWordprocessingDocument(string filepath)
        {
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
            }
        }

        public static void OpenAndAddTextToWordDocument(string filepath, string txt)
        {
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(filepath, true);
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));
            wordprocessingDocument.Close();
        }

        public static string ReadWordDocument(string filepath)
        {
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(filepath, true);
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            string result="";
            foreach (Text text in body.Descendants<Text>())
            {
                result += text.Text.ToString();
            }
            wordprocessingDocument.Close();
            return result;
        }
    }


}
