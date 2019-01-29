using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;

namespace FontImage
{
    public class FontImage
    {
        private string _strFontFile;
        public FontImage(string strFontFile)
        {
            _strFontFile = strFontFile;
        }

        public void ExportFontImage(string strFolder)
        {
            string strName = Path.GetFileNameWithoutExtension(_strFontFile);
            GetFontImage($@"{strFolder}/1_{strName}_1.png",
                Color.Black,
                Color.White,
                18,
                new Size(230, 26),
                StringAlignment.Near,
                StringAlignment.Center);

            GetFontImage($@"{strFolder}/m_{strName}_1.jpg",
                Color.Black,
                Color.White,
                22,
                new Size(280, 74),
                StringAlignment.Center,
                StringAlignment.Center);

            GetFontImage($@"{strFolder}/{strName}_1.png",
                Color.White,
                Color.Transparent,
                22,
                new Size(316, 74),
                StringAlignment.Center,
                StringAlignment.Center);

            GetFontImage($@"{strFolder}/s_{strName}_1.png",
                Color.Black,
                Color.Transparent,
                18,
                new Size(316, 74),
                StringAlignment.Center,
                StringAlignment.Center);

            GetAndroidImage($@"{strFolder}/a_{strName}_1.png",
                Color.Black,
                Color.Transparent,
                1,
                new Size(0, 50),
                StringAlignment.Center,
                StringAlignment.Center);

            GetAndroidImage($@"{strFolder}/s_{strName}_1.png",
                Color.Black,
                Color.Transparent,
                18,
                new Size(290, 30),
                StringAlignment.Center,
                StringAlignment.Center);

            GetAndroidImage($@"{strFolder}/i_{strName}_1.png",
                Color.Black,
                Color.Transparent,
                32,
                new Size(430, 44),
                StringAlignment.Center,
                StringAlignment.Center);

            //GetFontImage($@"{strFolder}/o_{strName}_1.png",
            //    Color.FromArgb(117, 185, 243),
            //    Color.Transparent,
            //    42,
            //    new Size(670, 176),
            //    StringAlignment.Near,
            //    StringAlignment.Center,
            //    "小巧、极速、全兼容，WPS Office\r\n您的办公专家。");
        }

        private void GetFontImage(string strImgName, Color colorFont, Color colorBackground, int fontSize, Size imgSize, StringAlignment alignHorizontal, StringAlignment alignVercical, string strText = "")
        {
            PrivateFontCollection fonts = new PrivateFontCollection();
            fonts.AddFontFile(_strFontFile);
            FontFamily fontFamily = fonts.Families[0];

            if (strText == "")
                strText = Path.GetFileNameWithoutExtension(_strFontFile);
            Bitmap bitmap = new Bitmap(imgSize.Width, imgSize.Height);
            Graphics graphics = Graphics.FromImage(bitmap);
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.InterpolationMode = InterpolationMode.Bicubic;
            graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
            graphics.Clear(colorBackground);

            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = alignHorizontal;
            stringFormat.LineAlignment = alignVercical;

            Font font;
            while (true)
            {
                font = new Font(fontFamily, fontSize);
                SizeF sizeTest = graphics.MeasureString(strText, font, imgSize, stringFormat);
                if (sizeTest.Width > bitmap.Width)
                {
                    fontSize--;
                    continue;
                }
                break;
            }

            graphics.DrawString(strText, font, new SolidBrush(colorFont), new Rectangle(0,0, bitmap.Width, bitmap.Height), stringFormat);
            bitmap.Save(strImgName, strImgName.EndsWith("jpg") ? ImageFormat.Jpeg : ImageFormat.Png);
            graphics.Dispose();
            bitmap.Dispose();
        }

        private void GetAndroidImage(string strImgName, Color colorFont, Color colorBackground, int fontSize, Size imgSize, StringAlignment alignHorizontal, StringAlignment alignVercical)
        {
            PrivateFontCollection fonts = new PrivateFontCollection();
            fonts.AddFontFile(_strFontFile);
            FontFamily fontFamily = fonts.Families[0];

            string strText = Path.GetFileNameWithoutExtension(_strFontFile);
            Bitmap bitmap = new Bitmap(imgSize.Height, imgSize.Height);
            Graphics graphics = Graphics.FromImage(bitmap);
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.InterpolationMode = InterpolationMode.Bicubic;
            graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
            graphics.Clear(colorBackground);

            StringFormat stringFormat = StringFormat.GenericTypographic;
            StringFormat.GenericTypographic.LineAlignment = StringAlignment.Center;
            StringFormat.GenericTypographic.FormatFlags = StringFormatFlags.NoWrap;

            bool bAdd = fontSize == 1;
            Font font;
            while (true)
            {
                font = new Font(fontFamily, fontSize);
                SizeF sizeTest = graphics.MeasureString(strText, font, Size.Add(imgSize, new Size(1000,10)), stringFormat);
                bool bContinue = true;
                if (imgSize.Width == 0)
                {
                    if (sizeTest.Height > imgSize.Height)
                        bContinue = false;
                }
                else
                {
                    if (sizeTest.Width <= imgSize.Width)
                        bContinue = false;
                }
                if (!bContinue)
                {
                    graphics.Dispose();
                    bitmap.Dispose();
                    bitmap = new Bitmap((int) sizeTest.Width, imgSize.Height);
                    graphics = Graphics.FromImage(bitmap);
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.Bicubic;
                    graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                    graphics.Clear(colorBackground);
                    break;
                }

                if (bAdd)
                    fontSize++;
                else
                    fontSize--;
            }

            font = new Font(fontFamily, fontSize);
            graphics.DrawString(strText, font, new SolidBrush(colorFont), new Rectangle(0, 0, bitmap.Width, bitmap.Height), stringFormat);
            bitmap.Save(strImgName, ImageFormat.Png);
            graphics.Dispose();
            bitmap.Dispose();
        }
    }
}
