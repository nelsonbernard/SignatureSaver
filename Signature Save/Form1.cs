using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Ink;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Reflection;

namespace Signature_Save
{
    public partial class Form1 : Form
    {
        
        private InkOverlay inkOverlay;
        private InkOverlay initialsOverlay;
        

        public Form1()
        {
            InitializeComponent();
            InitializeOverlays();
            Word.Application WordApp;                     
        }

        private void btnNewCapture_Click(object sender, EventArgs e)
        {
            ClearOverlays();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            inkOverlay.Enabled = false;
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            string fileName = filePath + textName.Text;
            string initialsFileName = filePath + textName.Text + " - Initials";

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = textName.Text;
            sfd.Filter = "GIF Images (*.gif)|*.gif";
            sfd.InitialDirectory = filePath;


            if (inkOverlay.Ink.Strokes.Count > 0)
            {
                if (textName.Text != "")
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Rectangle signatureBounds = this.pictureBox1.Bounds;
                        inkOverlay.Ink.Clip(signatureBounds);
                        byte[] bytes = (byte[])inkOverlay.Ink.Save(PersistenceFormat.Gif, CompressionMode.NoCompression);

                        using (MemoryStream ms = new MemoryStream(bytes))
                        {
                            using (Bitmap gif = new Bitmap(ms))
                            {
                                gif.Save(sfd.FileName);                                
                            }
                        }
                    }
                }
                else
                    MessageBox.Show("Signature must be attached to a client name");
            }
            else
            {
                MessageBox.Show("There is no signature to save");
                this.pictureBox1.Invalidate();
                inkOverlay.Enabled = true;
            }

        }

        private void btnEditDocument_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.InitialDirectory = filePath;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string sourceFile = ofd.FileName;
                    string sourceFilePath = Path.GetDirectoryName(sourceFile);
                    UpdateDoc(sourceFile);            
                }         
                
                       
            }
            catch(Exception)
            {
                MessageBox.Show("Error in process.", "Internal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /*private void FindAndReplace(Word.Application wordApp, object findText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            byte[] signature = (byte[])inkOverlay.Ink.Save(PersistenceFormat.Gif, CompressionMode.NoCompression);
            byte[] initials = (byte[])initialsOverlay.Ink.Save(PersistenceFormat.Gif, CompressionMode.NoCompression);

            using (MemoryStream ms = new MemoryStream(signature))
            {
                Bitmap gif = new Bitmap(ms);
                gif.Save(ms, ImageFormat.Png);
                var data = new DataObject("PNG", ms);
                Clipboard.Clear();
                Clipboard.SetDataObject(data, true);

                wordApp.Selection.Find.Execute(ref findText, ref matchCase,
                ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
                ref matchAllWordForms, ref forward, ref wrap, ref format,
                Clipboard.GetDataObject(), ref replace, ref matchKashida, ref matchDiacritics,
                ref matchAlefHamza, ref matchControl);
            }
        }*/
        
        private void InitializeOverlays()
        {
            inkOverlay = new InkOverlay();
            inkOverlay.Handle = this.pictureBox1.Handle;
            inkOverlay.DefaultDrawingAttributes.AntiAliased = true;
            inkOverlay.DefaultDrawingAttributes.IgnorePressure = false;
            inkOverlay.DefaultDrawingAttributes.PenTip = PenTip.Ball;
            inkOverlay.Enabled = true;


            initialsOverlay = new InkOverlay();
            initialsOverlay.Handle = this.pictureBox2.Handle;
            initialsOverlay.DefaultDrawingAttributes.AntiAliased = true;
            initialsOverlay.DefaultDrawingAttributes.IgnorePressure = false;
            initialsOverlay.DefaultDrawingAttributes.PenTip = PenTip.Ball;
            initialsOverlay.Enabled = true;
            Rectangle initialsBounds = this.pictureBox2.Bounds;
        }

        private void ClearOverlays()
        {
            inkOverlay.Enabled = false;
            initialsOverlay.Enabled = false;
            inkOverlay.Ink.DeleteStrokes();
            initialsOverlay.Ink.DeleteStrokes();
            this.pictureBox2.Invalidate();
            this.pictureBox1.Invalidate();
            inkOverlay.Enabled = true;
            initialsOverlay.Enabled = true;
        }

        /*private void FindAndReplace()
        {
            inkOverlay.Enabled = false;
            string filePath = Path.GetTempPath();
            string fileName = filePath + textName.Text + ".gif";
            string initialsFileName = filePath + textName.Text + " - Initials.gif";

            inkOverlay.Ink.ClipboardPaste();
            byte[] signature = (byte[])inkOverlay.Ink.Save(PersistenceFormat.Gif, CompressionMode.NoCompression);
            byte[] initials = (byte[])initialsOverlay.Ink.Save(PersistenceFormat.Gif, CompressionMode.NoCompression);

            using (MemoryStream ms = new MemoryStream(signature))
            {
                /*Bitmap gif = new Bitmap(ms);
                gif.Save(ms, ImageFormat.Png);
                var data = new DataObject("PNG", ms);
                Clipboard.Clear();
                Clipboard.SetDataObject(data, true);
                
                using (Bitmap gif = new Bitmap(ms))
                {
                    gif.Save(fileName);
                }
            }

            using (MemoryStream ms = new MemoryStream(initials))
            {
                using (Bitmap gif = new Bitmap(ms))
                {
                    gif.Save(initialsFileName);
                }
            }

            string searchTerm = "<Signature>";
            string searchTerm2 = "<Initials>";

            var sel = wordApp.Selection;
            sel.Find.Text = searchTerm;
            
            sel.Find.Execute(Word.WdReplace.wdReplaceAll);
            
            sel.InlineShapes.AddPicture(fileName, false, true);

            sel.Find.Text = searchTerm2;
            sel.Range.Select();
            sel.Find.Execute(Word.WdReplace.wdReplaceAll);
            
            sel.InlineShapes.AddPicture(initialsFileName, false, true);
        }*/

        public void UpdateDoc(string source)
        {
            Word.Application WordApp = new Word.Application();
            object missing = System.Reflection.Missing.Value;
            object yes = true;
            object no = false;
            Word.Document d;
            object filename = Path.GetFullPath(source);
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            string fileName = filePath + textName.Text + ".gif";
            string initialsFileName = filePath + textName.Text + " - Initials.gif";

            d = WordApp.Documents.Open(ref filename, ref missing, ref no, ref missing,
                   ref missing, ref missing, ref missing, ref missing, ref missing,
                   ref missing, ref missing, ref yes, ref missing, ref missing, ref missing, ref missing);

            d.Activate();
            List<Word.Range> ranges = new List<Microsoft.Office.Interop.Word.Range>();

            foreach (Word.Range rngstory in d.StoryRanges)
            {
                if (rngstory.Text == "<Signature>")
                {
                    ranges.Add(rngstory);
                    rngstory.Delete();
                }
            }
            foreach (Word.Range r in ranges)
            {
                r.InlineShapes.AddPicture(fileName, ref missing, ref missing, ref missing);
            }
            //WordApp.Quit(ref yes, ref missing, ref missing);
        }

    }
}
