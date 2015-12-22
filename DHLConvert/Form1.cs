using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Spire.Pdf;

namespace DHLConvert
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Bitmap getClip(Image source, Rectangle rect)
        {
            Bitmap tempBitmap = new Bitmap(rect.Width, rect.Height);

            using (Graphics g = Graphics.FromImage(tempBitmap))
            {
                g.DrawImage(source, new Rectangle(0, 0, tempBitmap.Width, tempBitmap.Height), rect, GraphicsUnit.Pixel);
            }

            return tempBitmap;
        }

        private Bitmap convertPDF(string filename)
        {
            PdfDocument doc = new PdfDocument();
            Image source;
            Rectangle rectMain = new Rectangle(1662, 171, 1142, 685);
            Rectangle rectBarcode = new Rectangle(1803, 1240, 556, 685);
            Rectangle rectTracking = new Rectangle(1665, 857, 749, 113);
            Bitmap bitmapMain = new Bitmap(rectMain.Width, rectMain.Height);
            Bitmap bitmapTracking = new Bitmap(rectTracking.Width, rectTracking.Height);
            Bitmap bitmapBarcode = new Bitmap(rectBarcode.Width, rectBarcode.Height);
            Bitmap bitmapLabel = new Bitmap(1701, 697);

            try
            {
                doc.LoadFromFile(filename);
                source = doc.SaveAsImage(0, 273, 273);
                source.RotateFlip(RotateFlipType.Rotate90FlipNone);
                bitmapMain = getClip(source, rectMain);
                bitmapTracking = getClip(source, rectTracking);

                source = doc.SaveAsImage(0, 255, 255);
                source.RotateFlip(RotateFlipType.Rotate90FlipNone);
                bitmapBarcode = getClip(source, rectBarcode);

                using (Graphics g = Graphics.FromImage(bitmapLabel))
                {
                    g.DrawImage(bitmapMain, 0, 0);
                    g.DrawImage(bitmapTracking, 3, 570);
                    g.DrawImage(bitmapBarcode, 1142, 0);
                }

                return bitmapLabel;
            }
            catch (Spire.Pdf.Exceptions.PdfDocumentException) {
                return null;
            }
        }

        private bool savePDF(string filename)
        {
            try
            {
                PdfDocument doc = new PdfDocument();
                PdfSection section = doc.Sections.Add();
                PdfPageBase page = doc.Pages.Add(new SizeF(408.32f, 167.3f), new Spire.Pdf.Graphics.PdfMargins(0));
//            PdfPageBase page = doc.Pages.Add(new SizeF(150f, 62f), new Spire.Pdf.Graphics.PdfMargins(0));
                Spire.Pdf.Graphics.PdfImage image = Spire.Pdf.Graphics.PdfImage.FromImage(pictureBox1.Image);
                float widthFitRate = image.PhysicalDimension.Width / page.Canvas.ClientSize.Width;
                float heightFitRate = image.PhysicalDimension.Height / page.Canvas.ClientSize.Height;
                float fitRate = Math.Max(widthFitRate, heightFitRate);
                float fitWidth = image.PhysicalDimension.Width / fitRate;
                float fitHeight = image.PhysicalDimension.Height / fitRate;

                doc.DocumentInformation.Author = "(C) 2015 by Thorsten Schröpel";
                doc.DocumentInformation.Title = "DHL-Paketmarke";
                doc.DocumentInformation.Producer = "DHLConvert";
                doc.DocumentInformation.Keywords = "DHLConvert DHL Paketmarke Label Aufkleber";
                page.Canvas.DrawImage(image, 0.3f, 0.5f, fitWidth, fitHeight);
                doc.SaveToFile(filename);
//            bitmapLabel.Save("dhl.png", System.Drawing.Imaging.ImageFormat.Png);
                doc.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void enableControls()
        {
            button1.Enabled = true;
            button2.Enabled = true;
            pDFSpeichernToolStripMenuItem.Enabled = true;
            druckenToolStripMenuItem.Enabled = true;
        }

        private void printLabel(Image label)
        {
            printDocument1.DefaultPageSettings.Landscape = true;
            printDocument1.PrintPage += (sender, args) =>
            {
                Rectangle m = args.MarginBounds;

                if ((double)label.Width / (double)label.Height > (double)m.Width / (double)m.Height) 
                {
                    m.Height = (int)((double)label.Height / (double)label.Width * (double)m.Width);
                }
                else
                {
                    m.Width = (int)((double)label.Width / (double)label.Height * (double)m.Height);
                }

                args.Graphics.DrawImage(pictureBox1.Image, m);
            };
            printDocument1.Print();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printLabel(pictureBox1.Image);
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            Bitmap bitmapPDF;

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            bitmapPDF = convertPDF(files[0]);
            if (bitmapPDF == null)
            {
                MessageBox.Show("Keine gültige Paketmarke!");
            }
            else
            {
                pictureBox1.Image = bitmapPDF;
                enableControls();
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void beendenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool userClickedOK = (saveFileDialog1.ShowDialog() == DialogResult.OK);

            if (userClickedOK == true)
            {
                if (savePDF(saveFileDialog1.FileName) == false)
                {
                    MessageBox.Show("Das DHL-Label konnte nicht gespeichert werden!");
                }
            }
        }

        private void überToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 aboutBox = new AboutBox1();

            aboutBox.ShowDialog();
        }

        private void druckenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printLabel(pictureBox1.Image);
            }
        }

        private void pDFSpeichernToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool userClickedOK = (saveFileDialog1.ShowDialog() == DialogResult.OK);

            if (userClickedOK == true)
            {
                if (savePDF(saveFileDialog1.FileName) == false)
                {
                    MessageBox.Show("Das DHL-Label konnte nicht gespeichert werden!");
                }
            }
        }

        private void pDFÖffnenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Bitmap bitmapPDF;

            bool userClickedOK = (openFileDialog1.ShowDialog() == DialogResult.OK);

            if (userClickedOK == true)
            {
                bitmapPDF = convertPDF(openFileDialog1.FileName);
                if (bitmapPDF == null) {
                    MessageBox.Show("Keine gültige Paketmarke!");
                } else {
                    pictureBox1.Image = bitmapPDF;
                    enableControls();
                }
            }
        }
    }
}
