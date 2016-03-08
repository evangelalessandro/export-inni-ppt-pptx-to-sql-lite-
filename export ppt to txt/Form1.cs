using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace export_ppt_to_txt
{
    public partial class Form1 : Form
    {
        private SQLiteConnection sql_con;
        private SQLiteCommand sql_cmd;

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            const string dir = @"C:\Users\MioAle\Downloads\innario_completo\innario completo\";


            var fileConPunto = System.IO.Directory.EnumerateFiles(dir, "*.*.pptx").ToList()
                .Select(a => a.Replace(dir, "").Replace(".pptx", "").
                    Replace(".ppt", "")).Where(a => a.Contains(".")).ToList();
            var filesenzaPunto = System.IO.Directory.EnumerateFiles(dir, "*.pptx")
                .Select(a => a.Replace(dir, "")
             .Replace(".pptx", "").Replace(".ppt", "")).Where(a => !a.Contains(".")).ToList();

            var numeriA = fileConPunto.Select(a => int.Parse(a.Replace(".", "")
                 .Split(" ".ToArray())[0])).ToList();

            var numeriB = filesenzaPunto.Select(a => int.Parse(a.Replace(".", "").Split(" ".ToArray())[0])).ToList();

            var intersezioni = numeriA.Where(a => numeriB.Contains(a)).ToList();

            foreach (var item in intersezioni)
            {
                if (item.ToString().Length == 2)
                {
                    var file = System.IO.Directory.EnumerateFiles(dir, item.ToString() + " *.pptx").First();
                    System.IO.File.Move(file, file.Replace(dir, dir + @"todelete\"));
                }
            }

        }

 
        private void button2_Click(object sender, EventArgs e)
        {
            const string dir = @"C:\Users\MioAle\Downloads\innario_completo\innario completo\";

            foreach (var item in System.IO.Directory.EnumerateFiles(dir, "*.ppt"))
            {
                ConvertiPptToPptx(item);
                System.IO.File.Delete(item);
            }

            var innoList = new List<InnoItem>();
            foreach (var filepptx in System.IO.Directory.EnumerateFiles(dir, "*.pptx"))
            {
                InnoItem newInno = new InnoItem();
                int numberOfSlides = CountSlides(filepptx);
                //System.Console.WriteLine("Number of slides = {0}", numberOfSlides);
                string slideText;
                for (int numeroSlide = 0; numeroSlide < numberOfSlides; numeroSlide++)
                {
                    GetSlideIdAndText(out slideText, filepptx, numeroSlide);
                    slideText = slideText.Replace("  ", " ");

                    if (numeroSlide != 0)
                    {
                        if (string.IsNullOrEmpty(newInno.Testo))
                        {
                            newInno.Testo = slideText;
                        }
                        else
                        {
                            newInno.Testo += Environment.NewLine + Environment.NewLine + slideText;
                        }
                    }
                    else
                    {
                        slideText = ImpostaTitoloNumero(newInno, slideText);
                    }
                }
                innoList.Add(newInno);

                textBox1.Text = newInno.Numero.ToString();
                this.Update();
            }

            var textMess = new StringBuilder();
            foreach (var item in innoList.Select(a => a.Numero.ToString() + "||" + Environment.NewLine + a.Titolo + "||" + Environment.NewLine + a.Testo).ToList())
            {
                textMess.AppendLine(item);
            }
            textBox1.Text = textMess.ToString();
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DataSource = innoList;
            dataGridView1.Columns[2].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //dataGridView1.Columns[2] = new DataGridViewTextBoxColumn();
            dataGridView1.Refresh();
            dataGridView1.Update();

            InsertToDatabase(innoList.OrderBy(a => a.Numero).ToList());
        }
        private void InsertToDatabase(List<InnoItem> list)
        {
            sql_con = new SQLiteConnection("Data Source=InnarioAvventistaDb.db;Version=3;New=True;Compress=True;");

            ExecuteQuery("CREATE TABLE 'innarioadulti' (	'numero'	INTEGER NOT NULL," +
                    " 'titolo'	TEXT NOT NULL, " +
                    " 'testo'	BLOB NOT NULL, " +
                    " PRIMARY KEY(numero) " +
                ")");

            foreach (var item in list)
            {

                //ExecuteQuery("select numero from InnarioAdulti where Numero ='" 
                //    + item.Numero.ToString() + "'");

                ExecuteQuery("Insert into InnarioAdulti (Numero,Titolo,Testo) Values('" + item.Numero.ToString() + "',"
                   + "'" + item.Titolo + "'," + "'" + item.Testo + "')");
            }
        }

        private void ExecuteQuery(string txtQuery)
        {
            sql_con.Open();
            sql_cmd = sql_con.CreateCommand();
            sql_cmd.CommandText = txtQuery;
            sql_cmd.ExecuteNonQuery();
            sql_con.Close();
        }

        private static string ImpostaTitoloNumero(InnoItem newInno, string slideText)
        {
            //il primo contine il titolo e forse anche una strofa
            var countLine = (slideText.Length - slideText.Replace(Environment.NewLine, "").Length) / Environment.NewLine.Length;
            slideText = slideText.Trim();
            if (countLine > 1)
            {
                //se ha anche la strofa
                var titolo = slideText.Split(Environment.NewLine.ToArray())[0].Trim();
                string numero = EstrapolaNumero(newInno, titolo);

                slideText = slideText.Remove(0, titolo.Length).Trim();
                titolo = ImpostaTitolo(newInno, titolo, numero);

                newInno.Testo = slideText;
                //    slideText = slideText.Remove(0, titolo.Length).Trim();
            }
            else
            {
                var titolo = slideText;
                string numero = EstrapolaNumero(newInno, titolo);
                titolo = ImpostaTitolo(newInno, titolo, numero);

            }

            return slideText;
        }

        private static string ImpostaTitolo(InnoItem newInno, string titolo, string numero)
        {
            titolo = titolo.Remove(0, numero.Length).Trim();
            newInno.Titolo = titolo;
            return titolo;
        }

        private static string EstrapolaNumero(InnoItem newInno, string titolo)
        {
            var numero = titolo.Trim().Split(" ".ToCharArray())[0];
            newInno.Numero = int.Parse(numero.Replace(".", ""));
            if (newInno.Numero == 116)
            {
            }
            return numero;
        }

        private static string ConvertiPptToPptx(string filepptx)
        {
            using (var export = new ExportJob())
            {
                filepptx = export.convertpptTopptx(filepptx);
            }

            return filepptx;
        }

        public static int CountSlides(string presentationFile)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                // Pass the presentation to the next CountSlides method
                // and return the slide count.
                return CountSlides(presentationDocument);
            }
        }

        // Count the slides in the presentation.
        public static int CountSlides(PresentationDocument presentationDocument)
        {
            // Check for a null document object.
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;

            // Get the presentation part of document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            // Get the slide count from the SlideParts.
            if (presentationPart != null)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }
            // Return the slide count to the previous method.
            return slidesCount;
        }

        public static void GetSlideIdAndText(out string sldText, string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                string relId = (slideIds[index] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                // Build a StringBuilder object.
                StringBuilder paragraphText = new StringBuilder();

                //// Get the inner text of the slide:
                //var texts = slide.Slide.Descendants().Where(a=>a.InnerText!="").Select(a=>a.InnerText).ToList();
                //foreach (var text in texts)
                //{
                //    paragraphText.Append(text);
                //}

                var xmlText = slide.Slide.InnerXml.Replace("\t", "").Replace("\t", "");
                while (true)
                {
                    int indexTExt = xmlText.IndexOf("<a:t>");
                    if (indexTExt > -1)
                    {
                        xmlText = xmlText.Remove(0, indexTExt + "<a:t>".Length);
                        int indexEndTExt = xmlText.IndexOf("</a:t>");
                        paragraphText.Append(" " + xmlText.Substring(0, indexEndTExt));

                        indexTExt = xmlText.IndexOf("<a:t>");
                        int indexTExtbr = xmlText.IndexOf("</a:br>");
                        if (indexTExtbr > -1)
                        {

                            if (indexTExtbr < indexTExt)
                            {
                                xmlText = xmlText.Remove(0, indexTExtbr + "</a:br>".Length);
                                paragraphText.AppendLine("");
                            }
                        }
                    }
                    else
                    {
                        sldText = paragraphText.ToString();
                        return;
                    }
                }

                sldText = paragraphText.ToString();
                return;
            }
        }
    }
}
