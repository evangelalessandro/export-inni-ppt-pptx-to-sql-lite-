using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace export_ppt_to_txt
{
    public partial class Form1 : Form
    {
        public string filepptx = @"C:\Users\MioAle\Downloads\innario_completo\innario completo\298. Chi potrà dir qual sia la gloria.pptx";
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            using (var provaExport = new ExportJob())
            {
                provaExport.runInForm(filepptx);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int numberOfSlides = CountSlides(filepptx);
            System.Console.WriteLine("Number of slides = {0}", numberOfSlides);
            string slideText;
            for (int i = 0; i < numberOfSlides; i++)
            {
                GetSlideIdAndText(out slideText, filepptx, i);
                System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText);

                textBox1.Text = textBox1.Text + Environment.NewLine + Environment.NewLine + slideText;
            }
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


                        xmlText = xmlText.Remove(0,indexTExt+ "<a:t>".Length);

                        int indexEndTExt = xmlText.IndexOf("</a:t>");
                        paragraphText.Append(xmlText.Substring(0, indexEndTExt));

                        indexTExt = xmlText.IndexOf("<a:t>");
                        int indexTExtbr = xmlText.IndexOf("</a:br>");
                        if (indexTExtbr > -1)
                        {
                         
                            if (indexTExtbr < indexTExt)
                            {
                                xmlText = xmlText.Remove(0,indexTExtbr+ "</a:br>".Length);
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
