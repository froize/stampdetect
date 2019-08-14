using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Reflection;

using AForge;
using AForge.Imaging;
using AForge.Imaging.Filters;
using AForge.Math.Geometry;
using System.IO;

using Syncfusion.Pdf.Parsing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace StampDetect
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        // Exit from application
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // On loading of the form
        private void MainForm_Load(object sender, EventArgs e)
        {
            
        }
        void GetCircle(Bitmap sourceBitmap, Bitmap targetBitmap)
        {
            //int threshold = 15;
            for (int y = 0; y < sourceBitmap.Height; ++y)
            {
                for (int x = 0; x < sourceBitmap.Width; ++x)
                {
                    Color c = sourceBitmap.GetPixel(x, y);

                    if (c.B > 150 & c.R < 150 & c.G < 150)
                    {
                        targetBitmap.SetPixel(x, y, Color.Blue);
                        for (int i = -3; i < 3; i++)
                        {
                            for (int j = -3; j < 3; j++)
                            {
                                if ((x + i > 0) && (y + j > 0) && (x + i < targetBitmap.Width) && (y + j < targetBitmap.Height))
                                    targetBitmap.SetPixel(x + i, y + j, Color.Blue);
                            }

                        }
                    }
                }
            }

            targetBitmap.Save("C:/Users/Ali/Desktop/pechat9.jpg");
        }
        // Process image
        private void ProcessImage(Bitmap bitmap)
        {
            Bitmap targetBitmap = new Bitmap(bitmap.Width, bitmap.Height, bitmap.PixelFormat);
            GetCircle(bitmap, targetBitmap);

            // lock image
            BitmapData bitmapData = targetBitmap.LockBits(
                new Rectangle(0, 0, targetBitmap.Width, targetBitmap.Height),
                ImageLockMode.ReadWrite, targetBitmap.PixelFormat);

            // step 2 - locating objects
            BlobCounter blobCounter = new BlobCounter();

            blobCounter.FilterBlobs = true;
            blobCounter.MinHeight = 100;
            blobCounter.MinWidth = 100;

            blobCounter.ProcessImage(bitmapData);
            Blob[] blobs = blobCounter.GetObjectsInformation();
            targetBitmap.UnlockBits(bitmapData);

            // step 3 - check objects' type and highlight
            SimpleShapeChecker shapeChecker = new SimpleShapeChecker();
            shapeChecker.MinAcceptableDistortion = 4;
            shapeChecker.RelativeDistortionLimit = 5f;

            Graphics g = Graphics.FromImage(targetBitmap);
            Pen yellowPen = new Pen(Color.Yellow, 4); // circles
            Pen redPen = new Pen(Color.Red, 2);       // quadrilateral
            Pen brownPen = new Pen(Color.Brown, 2);   // quadrilateral with known sub-type
            Pen greenPen = new Pen(Color.Green, 2);   // known triangle
            Pen bluePen = new Pen(Color.Blue, 2);     // triangle
            double maxradius = 0;
            for (int i = 0, n = blobs.Length; i < n; i++)
            {
                List<IntPoint> edgePoints = blobCounter.GetBlobsEdgePoints(blobs[i]);

                AForge.Point center;
                float radius;

                // is circle ?
                if (shapeChecker.IsCircle(edgePoints, out center, out radius))
                {
                    g.DrawEllipse(yellowPen,
                        (float)(center.X - radius), (float)(center.Y - radius),
                        (float)(radius * 2), (float)(radius * 2));
                    if (radius > maxradius)
                    {
                        Crop cropFilter = new Crop(new Rectangle((int)(center.X - radius - 50), (int)(center.Y - radius - 50), (int)radius * 2 + 100, (int)radius * 2 + 100));
                        Bitmap croppedImage = cropFilter.Apply(bitmap);
                        if (!Directory.Exists(Environment.CurrentDirectory + "\\" + "Stamps"))
                        {
                            Directory.CreateDirectory(Environment.CurrentDirectory + "\\" + "Stamps");
                        }
                        croppedImage.Save(Environment.CurrentDirectory + "\\" + "Stamps" + "\\" + Guid.NewGuid().ToString("N") + ".jpg");
                        pictureBox.Image = croppedImage;
                        maxradius = radius;
                    }
                }

        }

    }
        void WordToJpg(string startupPath, string filename1)
        {
            //string startupPath = "C:/Users/Vitek/Desktop"; 
            //string filename1 = "1.docx"; 
            var docPath = Path.Combine(startupPath, filename1);
            var app = new Microsoft.Office.Interop.Word.Application();

            //MessageFilter.Register(); 

            app.Visible = true;

            var doc = app.Documents.Open(docPath);

            doc.ShowGrammaticalErrors = false;
            doc.ShowRevisions = false;
            doc.ShowSpellingErrors = false;

            if (!Directory.Exists(startupPath + "\\" + filename1.Split('.')[0]))
            {
                Directory.CreateDirectory(startupPath + "\\" + filename1.Split('.')[0]);
            }

            //Opens the word document and fetch each page and converts to image
            foreach (Microsoft.Office.Interop.Word.Window window in doc.Windows)
            {
                foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
                {
                    for (var i = 1; i <= pane.Pages.Count; i++)
                    {
                        var page = pane.Pages[i];
                        var bits = page.EnhMetaFileBits;
                        var target = Path.Combine(startupPath + "\\" + filename1.Split('.')[0], string.Format("{1}_page_{0}", i, filename1.Split('.')[0]));

                        try
                        {
                            using (var ms = new MemoryStream((byte[])(bits)))
                            {
                                var image = System.Drawing.Image.FromStream(ms);
                                var pngTarget = Path.ChangeExtension(target, "png");
                                //image.Save(pngTarget, ImageFormat.Png);
                                Bitmap bitmap = new Bitmap(image);
                                ProcessImage(bitmap);
                            }
                        }
                        catch (System.Exception ex)
                        { }
                    }
                }
            }
            doc.Close(Type.Missing, Type.Missing, Type.Missing);
            app.Quit(Type.Missing, Type.Missing, Type.Missing);

            //MessageFilter.Revoke(); 
        }
        private void button1_Click(object sender, EventArgs e)
        {

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = "";
                string fileName = "";

                string myFilePath = openFileDialog.FileName;
                string ext = Path.GetExtension(myFilePath);
                switch (ext)
                {
                    case ".jpg":
                        ProcessImage((Bitmap)Bitmap.FromFile(openFileDialog.FileName));
                        break;
                    case ".png":
                        ProcessImage((Bitmap)Bitmap.FromFile(openFileDialog.FileName));
                        break;
                    case ".xlsx":
                        Workbook workbook = new Workbook(openFileDialog.FileName);
                        //Get the first worksheet.
                        Worksheet sheet = workbook.Worksheets[0];

                        //Define ImageOrPrintOptions
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                        //Specify the image format
                        imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                        //Only one page for the whole sheet would be rendered
                        imgOptions.OnePagePerSheet = true;

                        //Render the sheet with respect to specified image/print options
                        SheetRender sr = new SheetRender(sheet, imgOptions);
                        //Render the image for the sheet
                        Bitmap bitmap = sr.ToImage(0);

                        //Save the image file specifying its image format.
                        //bitmap.Save("excel.jpg");
                        ProcessImage(bitmap);

                        break;
                    case ".doc":
                        filePath = Path.GetDirectoryName(openFileDialog.FileName);
                        fileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName) + ".doc";
                        WordToJpg(filePath, fileName);
                        break;
                    case ".docx":
                        filePath = Path.GetDirectoryName(openFileDialog.FileName);
                        fileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName) + ".docx";
                        WordToJpg(filePath, fileName);
                        break;
                    case ".pdf":
                        PdfLoadedDocument loadedDocument = new PdfLoadedDocument(openFileDialog.FileName);
                        Bitmap image = loadedDocument.ExportAsImage(0);
                        //image.Save("Image.jpg", ImageFormat.Jpeg);
                        ProcessImage(image);
                        loadedDocument.Close(true);
                        break;
                }
                
            }
        }
    }

}
