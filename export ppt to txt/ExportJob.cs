using System;
using System.IO;
using System.IO.Packaging;
using Microsoft.Office.Interop.PowerPoint;
using System.Text;

namespace export_ppt_to_txt
{

    public class ExportJob :IDisposable
    {
        public String method = "";
        public String file = "";
        public String options = "";
        //ExportingDialog ed;

        //public void run(Form f)
        //{
        //    ed = new ExportingDialog();
        //    ed.job = this;
        //    ed.Show(f);
        //}

        String getTempDir()
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }

        Microsoft.Office.Interop.PowerPoint.Application openPowerPoint()
        {
            Microsoft.Office.Interop.PowerPoint.Application powerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
            foreach (Presentation ppt in powerPoint.Presentations)
            {
                ppt.Close();
            }
            powerPoint.Presentations.Open(file);
            return powerPoint;
        }

        public void runInForm(string myFile)
        {
            file = myFile;
            System.Diagnostics.Debug.WriteLine("Exporting " + file + "...");
            String tempDir = getTempDir();
            System.Diagnostics.Debug.WriteLine("Tempoary Directory: " + tempDir);

                 Package zipPresentation = null;
                if (options.Contains("exportNotes"))
                {
                    String tmpFile = Path.Combine(tempDir, "pres.zip");
                    File.Copy(file, tmpFile);
                    zipPresentation = ZipPackage.Open(new FileStream(tmpFile, FileMode.Open));
                }

                Microsoft.Office.Interop.PowerPoint.Application powerPoint = openPowerPoint();
                System.Diagnostics.Debug.WriteLine("PowerPoint should have loaded document");
                System.Diagnostics.Debug.WriteLine("Exporting as handouts...");
                Presentation toExport = powerPoint.Presentations[1];

                //System.Diagnostics.Debug.WriteLine("Opening Publisher Template");
                //Microsoft.Office.Interop.Publisher.Application publisher = new Microsoft.Office.Interop.Publisher.Application();

                //String pubFile = Path.Combine(tempDir, "out.pub");
                //File.Copy(Path.Combine(
                //    Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath),
                //    "Templates",
                //    "Handout-3P.pub" // TODO: Add option for different templates
                //), pubFile);
                //Document publisherDoc = publisher.Open(pubFile);
                //publisher.ActiveWindow.Visible = true;


                int slideWidth = (int)toExport.PageSetup.SlideWidth;
                int slideHeight = (int)toExport.PageSetup.SlideHeight;
                int slideCount = 1, onPage = 1;
                //publisherDoc.Pages[onPage].Duplicate();

                foreach (_Slide slide in toExport.Slides)
                {
                    if (slideCount == 4)  // End of Page but only create a new one, if there is more!
                    {
                        //publisherDoc.Pages[onPage + 1].Duplicate();
                        onPage += 1;
                        slideCount = 1;
                    }
                
                    System.Diagnostics.Debug.WriteLine("Exporting slide #" + slide.SlideNumber);
                     
                    String notes = "";
                    //if (options.Contains("exportNotes"))
                    //{
                    //    // I HATE YOU POWERPOINT!
                    //    Uri partUriResource = PackUriHelper.CreatePartUri(
                    //          new Uri(@"ppt/notesSlides/notesSlide" + slide.SlideNumber + ".xml", UriKind.Relative));
                    //    if (zipPresentation.PartExists(partUriResource))
                    //    {
                    //        Stream noteStream = zipPresentation.GetPart(partUriResource).GetStream(FileMode.Open, FileAccess.Read);
                    //        StreamReader nSR = new StreamReader(noteStream);

                    //        // Strip crap
                    //        notes = nSR.ReadToEnd();
                    //        // Chop out slide numbers. I don't have a clue really :')
                    //        int s = notes.IndexOf("<p:txBody>");
                    //        if (s != -1)
                    //        {
                    //            notes = notes.Substring(s, notes.IndexOf("</p:txBody>") - s);
                    //            notes = (new Regex("<[^>]*>")).Replace(notes, "");
                    //        }

                    //        nSR.Close();
                    //        noteStream.Close();
                    //    }
                    //}


                    //foreach (Microsoft.Office.Interop.Publisher.Shape shape in publisherDoc.Pages[onPage].Shapes)
                    //{
                    //    if (shapeHasTag(shape, "Picture " + slideCount))
                    //    {
                    //        shape.Fill.UserPicture(fname);
                    //    }
                    //    else if (shapeHasTag(shape, "Text " + slideCount))
                    //    {
                    //        shape.TextFrame.AutoFitText = PbTextAutoFitType.pbTextAutoFitBestFit;
                    //        shape.TextFrame.TextRange.Text = notes;
                    //    }
                    //}

                    slideCount += 1;
                }

                //publisherDoc.Pages[publisherDoc.Pages.Count].Delete();
                toExport.Close();
                System.Diagnostics.Debug.WriteLine("Exported as handouts :)");
            
            //ed.finishJob("");
        }

        #region IDisposable Support
        private bool disposedValue = false; // Per rilevare chiamate ridondanti

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: eliminare lo stato gestito (oggetti gestiti).
                }

                // TODO: liberare risorse non gestite (oggetti non gestiti) ed eseguire sotto l'override di un finalizzatore.
                // TODO: impostare campi di grandi dimensioni su Null.

                disposedValue = true;
            }
        }

        // TODO: eseguire l'override di un finalizzatore solo se Dispose(bool disposing) include il codice per liberare risorse non gestite.
        // ~ExportJob() {
        //   // Non modificare questo codice. Inserire il codice di pulizia in Dispose(bool disposing) sopra.
        //   Dispose(false);
        // }

        // Questo codice viene aggiunto per implementare in modo corretto il criterio Disposable.
        public void Dispose()
        {
            // Non modificare questo codice. Inserire il codice di pulizia in Dispose(bool disposing) sopra.
            Dispose(true);
            // TODO: rimuovere il commento dalla riga seguente se è stato eseguito l'override del finalizzatore.
            // GC.SuppressFinalize(this);
        }
        #endregion

        //private bool shapeHasTag(Microsoft.Office.Interop.Publisher.Shape s, string p)
        //{
        //    foreach (Tag t in s.Tags)
        //    {
        //        if (t.Name == "PPE" && t.Value == p)
        //        {
        //            return true;
        //        }
        //    }
        //    return false;
        //}
     
    
}
}

