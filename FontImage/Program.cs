using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing.Text;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Net;
using System.Runtime.InteropServices;
using NetOffice.PowerPointApi;
using Font = System.Drawing.Font;
using System.Threading;
using NetOffice;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using Rectangle = System.Drawing.Rectangle;

namespace FontImage
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            while(true)
            {
                Console.WriteLine("1:字体截图。2:字体下载。3:图标ppt拆分。4:word加第一页图片。5:批量设置SVG颜色。");
                string strAction = Console.ReadLine();
                if (strAction == "1")
                {
                    ExportFontImage();
                }
                else if (strAction == "2")
                {
                    DownLoadFont();
                }
                else if (strAction == "3")
                {
                    ExportShapes();
                }
                else if (strAction == "4")
                {
                    AddWordFirstPageImgge();
                }
                else if (strAction == "5")
                {
                    Console.WriteLine("输入旧值:");
                    string strOld = Console.ReadLine();
                    if (!strOld.StartsWith("#") || strOld.Length != 7)
                    {
                        Console.WriteLine("输入有误");
                        continue;
                    }

                    Console.WriteLine("输入新值:");
                    string strNew = Console.ReadLine();
                    if (!strNew.StartsWith("#") || strNew.Length != 7)
                    {
                        Console.WriteLine("输入有误");
                        continue;
                    }
                    FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                    if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
                    {
                        continue;
                    }
                    SVGColor svg = new SVGColor(folderBrowserDialog.SelectedPath);
                    svg.ChangeColor(strOld, strNew);
                }
            }
        }

        static void ExportShapes()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            var application = GetApplication();
            string strFolder = folderBrowserDialog.SelectedPath;
            var files = Directory.GetFiles(strFolder).Where(d => d.Contains(".ppt")).ToArray();

            int i = 0;
            foreach(string file in files)
            {
                Presentation presSource = application.Presentations.Open(file);
                string strPath = Path.GetDirectoryName(file);
                foreach (Slide slide in presSource.Slides)
                {
                    foreach (NetOffice.PowerPointApi.Shape shape in slide.Shapes)
                    {
                        string strSaveName = GetSaveName(strPath, shape.Name);
                        Presentation presNew = application.Presentations.Add();
                        Slide slideNew = presNew.Slides.Add(1, NetOffice.PowerPointApi.Enums.PpSlideLayout.ppLayoutBlank);
                        shape.Copy();
                        NetOffice.PowerPointApi.ShapeRange shapeRange = slideNew.Shapes.Paste();
                        //shapeRange.Align(NetOffice.OfficeApi.Enums.MsoAlignCmd.msoAlignCenters | NetOffice.OfficeApi.Enums.MsoAlignCmd.msoAlignMiddles, NetOffice.OfficeApi.Enums.MsoTriState.msoTrue);
                        shapeRange.Select();
                        application.CommandBars.ExecuteMso("ObjectsAlignMiddleVerticalSmart");
                        application.CommandBars.ExecuteMso("ObjectsAlignCenterHorizontalSmart");
                        presNew.SaveAs(strSaveName);
                        Thread.Sleep(200);
                        presNew.Close();
                    }
                }
                presSource.Close();
                Console.WriteLine($"{++i}/{files.Length}");
            }
        }

        static string GetSaveName(string strPath, string strName)
        {
            string strSaveName = strPath + "/" + strName + ".pptx";
            int i = 1;
            while(true)
            {
                if(File.Exists(strSaveName))
                {
                    strSaveName = $"{strPath}/{strName}{i}.pptx";
                }
                else
                {
                    return strSaveName;
                }
                i++;
            }
        }

        static NetOffice.PowerPointApi.Application GetApplication()
        {
            object obj = null;
            try
            {
                obj = Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch
            {

            }
            
            if (obj != null)
            {
                return new NetOffice.PowerPointApi.Application(new COMObject(obj));
                //return (NetOffice.PowerPointApi.Application)obj;
            }
            return new NetOffice.PowerPointApi.Application("PowerPoint.Application");
        }

        static void DownLoadFont()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择下载列表文件";
            openFileDialog.Filter = "txt|*.txt";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string strFile = openFileDialog.FileName;
            string strFolder = folderBrowserDialog.SelectedPath;

            WebClient webClient = new WebClient();
            string[] lines = File.ReadAllLines(strFile, Encoding.GetEncoding("GB2312"));
            int i = 0;
            foreach (string line in lines)
            {
                var splits = line.Split('\t');
                if (splits.Length > 1 && splits[1].Trim().StartsWith("http"))
                {
                    try
                    {
                        webClient.DownloadFile(splits[1], $"{strFolder}/{splits[0]}");
                        Console.WriteLine($"{++i}/{lines.Length}");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }
        }

        static void ExportFontImage()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if(folderBrowserDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string strFolder = folderBrowserDialog.SelectedPath;

            var files = Directory.GetFiles(strFolder).Where(f => f.EndsWith(".ttf", true, null) || f.EndsWith(".otf", true, null));
            int count = files.Count();
            int i = 0;
            foreach (string strFile in files)
            {
                PrivateFontCollection fonts = new PrivateFontCollection();
                fonts.AddFontFile(strFile);
                Font font;
                string strName = Path.GetFileNameWithoutExtension(strFile);
                Bitmap img = new Bitmap(230, 26);
                Graphics graphics = Graphics.FromImage(img);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                graphics.Clear(Color.White);
                StringFormat format = new StringFormat();
                format.Alignment = StringAlignment.Near;
                format.LineAlignment = StringAlignment.Center;

                int size = 18;
                SizeF sizeTest;
                while (true)
                {
                    font = new Font(fonts.Families[0], size);
                    sizeTest = graphics.MeasureString(strName, font);
                    if (sizeTest.Width > img.Width)
                    {
                        size--;
                        continue;
                    }
                    break;
                }
                graphics.DrawString(strName, font, new SolidBrush(Color.Black), new Rectangle(0, 0, img.Width, (int)(sizeTest.Height > img.Height ? sizeTest.Height : img.Height)), format);
                img.Save($@"{strFolder}/1_{strName}_1.png", ImageFormat.Png);
                
                img = new Bitmap(280, 74);
                graphics = Graphics.FromImage(img);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                graphics.Clear(Color.White);
                format = new StringFormat();
                format.Alignment = StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                size = 22;
                while (true)
                {
                    font = new Font(fonts.Families[0], size);
                    sizeTest = graphics.MeasureString(strName, font);
                    if (sizeTest.Width > img.Width)
                    {
                        size--;
                        continue;
                    }
                    break;
                }

                graphics.DrawString(strName, font, new SolidBrush(Color.Black), new Rectangle(0, 0, img.Width, (int) (sizeTest.Height > img.Height ? sizeTest.Height : img.Height)), format);
                img.Save($@"{strFolder}/m_{strName}_1.png", ImageFormat.Png);


                size = 22;
                while (true)
                {
                    font = new Font(fonts.Families[0], size);
                    sizeTest = graphics.MeasureString(strName, font);
                    if (sizeTest.Width > img.Width)
                    {
                        size--;
                        continue;
                    }
                    break;
                }
                font = new Font(fonts.Families[0], size);
                img = new Bitmap(316, 74);
                graphics = Graphics.FromImage(img);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                graphics.Clear(Color.Transparent);
                format = new StringFormat();
                format.Alignment = StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                graphics.DrawString(strName, font, new SolidBrush(Color.White), new Rectangle(0, 0, img.Width, img.Height), format);
                img.Save($@"{strFolder}/{strName}_1.png", ImageFormat.Png);

                int[] iMobile = GetMobileWidth(fonts.Families[0]);
                font = new Font(fonts.Families[0], iMobile[1]);
                img = new Bitmap(iMobile[0] * font.Name.Length, 100);
                graphics = Graphics.FromImage(img);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                graphics.Clear(Color.Transparent);
                graphics.DrawString(strName, font, new SolidBrush(Color.Black), 0, 0);

                int[] iBorder = GetMobileBorder(img as Bitmap);
                Bitmap bitmapMobile = img.Clone(new System.Drawing.Rectangle(iBorder[0], iBorder[1], iBorder[2] - iBorder[0] + 1, iBorder[3] - iBorder[1] + 1), img.PixelFormat);
                Image imgSave = new Bitmap(bitmapMobile, (int)(50.0f / bitmapMobile.Height * bitmapMobile.Width), 50);
                imgSave.Save($@"{strFolder}/a_{strName}_1.png", ImageFormat.Png);

                size = 18;
                while (true)
                {
                    font = new Font(fonts.Families[0], size);
                    Size stringSize = GetStringSize(strName, font);
                    if(stringSize.Width > 290)
                    {
                        size--;
                        continue;
                    }
                    img = new Bitmap(stringSize.Width, 30);
                    graphics = Graphics.FromImage(img);
                    graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                    graphics.Clear(Color.Transparent);
                    StringFormat stringFormat = StringFormat.GenericTypographic;
                    StringFormat.GenericTypographic.LineAlignment = StringAlignment.Center;
                    StringFormat.GenericTypographic.FormatFlags = StringFormatFlags.NoWrap;
                    graphics.DrawString(strName, font, new SolidBrush(Color.Black), new Rectangle(0, 0, img.Width, img.Height + 5), stringFormat);
                    img.Save($@"{strFolder}/s_{strName}_1.png", ImageFormat.Png);
                    break;
                }

                graphics.Dispose();
                img.Dispose();
                bitmapMobile.Dispose();
                imgSave.Dispose();

                Console.WriteLine($"{++i}/{count}");
            }
        }

        static int[] GetMobileWidth(FontFamily fontFamily)
        {
            int iFontSize = 1;
            int iPreHeight = 0;
            int iPreWidth = 0;
            int iPreFontSize = 1;
            while (true)
            {
                Font font = new Font(fontFamily, iFontSize);
                Size size = TextRenderer.MeasureText("啊", font, new Size(100, 100), TextFormatFlags.SingleLine);
                if (size.Height == 100)
                {
                    return new int[] { (int)(size.Width * 0.76), iFontSize};
                }

                if(size.Height > 100)
                {
                    return new int[] { (int)(iPreWidth * 0.76), iPreFontSize };
                }

                iPreHeight = size.Height;
                iPreWidth = size.Width;
                iPreFontSize = iFontSize;
                iFontSize++;
            }

        }

        static Size GetStringSize(string strText, Font font)
        {
            TextFormatFlags flags = TextFormatFlags.Left |
        TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine;
            return TextRenderer.MeasureText(strText, font, new Size(10, 10));
        }

        static int[] GetMobileBorder(Bitmap img)
        {
            int left = -1, top = -1, right = -1, bottom = -1;
            for (int i = 0; i < img.Width; i++)
            {
                for (int j = 0; j < img.Height; j++)
                {
                    if (img.GetPixel(i, j).ToArgb() != 0)
                    {
                        left = i;
                        break;
                    }
                }
                if (left > -1)
                    break;
            }

            for (int i = 0; i < img.Height; i++)
            {
                for (int j = 0; j < img.Width; j++)
                {
                    if (img.GetPixel(j, i).ToArgb() != 0)
                    {
                        top = i;
                        break;
                    }
                }
                if (top > -1)
                    break;
            }

            for (int i = img.Width - 1; i >= 0; i--)
            {
                for (int j = 0; j < img.Height; j++)
                {
                    if (img.GetPixel(i, j).ToArgb() != 0)
                    {
                        right = i;
                        break;
                    }
                }
                if (right > -1)
                    break;
            }

            for (int i = img.Height - 1; i >= 0; i--)
            {
                for (int j = 0; j < img.Width; j++)
                {
                    if (img.GetPixel(j, i).ToArgb() != 0)
                    {
                        bottom = i;
                        break;
                    }
                }
                if (bottom > -1)
                    break;
            }

            return new int[] { left, top, right, bottom };
        }

        static void AddWordFirstPageImgge()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string strFolder = folderBrowserDialog.SelectedPath;
            var application = GetWordApplication();
            int i = 0;
            var folders = Directory.GetDirectories(strFolder);
            foreach (string folder in folders)
            {
                i++;
                string strWord = Directory.GetFiles(folder).Where(f => f.Contains(".doc")).First();
                string strImage = Directory.GetFiles(folder).Where(f => f.Contains(".jpg") || f.Contains(".png")).First();

                Document doc = application.Documents.Open(strWord);
                doc.ActiveWindow.Selection.InsertBreak(WdBreakType.wdPageBreak);
                doc.ActiveWindow.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToFirst);
                NetOffice.WordApi.Shape shape = doc.Shapes.AddPicture(strImage);
                shape.WrapFormat.Type = WdWrapType.wdWrapBehind;
                shape.Left = -91.0f;
                shape.Top = -72.75f;
                shape.Width = 596.25f;
                shape.Height = 840.75f;
                doc.Save();
                doc.Close();
                Thread.Sleep(200);
                Console.WriteLine($"{i}/{folders.Length}");
            }
        }

        static NetOffice.WordApi.Application GetWordApplication()
        {
            object obj = null;
            try
            {
                obj = Marshal.GetActiveObject("Word.Application");
            }
            catch
            {

            }

            if (obj != null)
            {
                return new NetOffice.WordApi.Application(new COMObject(obj));
                //return (NetOffice.PowerPointApi.Application)obj;
            }
            return new NetOffice.WordApi.Application("Word.Application");
        }
    }
}
