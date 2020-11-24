using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using Newtonsoft.Json;
using Path = System.IO.Path;
using System.Drawing;
using DocumentFormat.OpenXml.Drawing;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Rectangle = DocumentFormat.OpenXml.Vml.Rectangle;
using Shape = DocumentFormat.OpenXml.Vml.Shape;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            //string strFile = @"D:\yoo\【着色算法】可根据不同着色技巧对图示进行重新染色.pdf";
            string strFile = @"D:\yoo\张甜甜-Java开发.pdf";

            PdfDocument doc = new PdfDocument();
            PdfReader reader = new PdfReader(strFile);
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                RenderFilter fontFilter = new RegionTextRenderFilter(reader.GetPageSize(i));
                ITextExtractionStrategy strategy = new TextWithFontExtractionStategy();
                string s = PdfTextExtractor.GetTextFromPage(reader, i, strategy);
            }
        }
    }

    public class TextWithFontExtractionStategy : iTextSharp.text.pdf.parser.ITextExtractionStrategy
    {
        StringBuilder result = new StringBuilder();
        private Vector lastStart;
        private Vector lastEnd;

        public void BeginTextBlock()
        {
            
        }

        public void RenderText(TextRenderInfo renderInfo)
        {
            var a = renderInfo.GetBaseline();
            var a1 = renderInfo.GetAscentLine();
            var a2 = renderInfo.GetSingleSpaceWidth();
            var a3 = renderInfo.GetDescentLine();
            var a4 = renderInfo.GetUnscaledBaseline();
            var a5 = renderInfo.GetFont();
            var a6 = renderInfo.GetRise();
            var a7 = renderInfo.GetText();

            bool flag1 = this.result.Length == 0;
            bool flag2 = false;
            LineSegment baseline = renderInfo.GetBaseline();
            Vector startPoint = baseline.GetStartPoint();
            Vector endPoint = baseline.GetEndPoint();
            if (!flag1)
            {
                Vector v = startPoint;
                Vector lastStart = this.lastStart;
                Vector lastEnd = this.lastEnd;
                if ((double)(lastEnd.Subtract(lastStart).Cross(lastStart.Subtract(v)).LengthSquared / lastEnd.Subtract(lastStart).LengthSquared) > 1.0)
                    flag2 = true;
            }
            if (flag2)
                this.AppendTextChunk('\n');
            else if (!flag1 && this.result[this.result.Length - 1] != ' ' && (renderInfo.GetText().Length > 0 && renderInfo.GetText()[0] != ' ') && (double)this.lastEnd.Subtract(startPoint).Length > (double)renderInfo.GetSingleSpaceWidth() / 2.0)
                this.AppendTextChunk(' ');
            this.AppendTextChunk(renderInfo.GetText());
            this.lastStart = startPoint;
            this.lastEnd = endPoint;
        }

        protected void AppendTextChunk(string text)
        {
            this.result.Append(text);
        }

        protected void AppendTextChunk(char text)
        {
            this.result.Append(text);
        }

        public void EndTextBlock()
        {
            
        }

        public void RenderImage(ImageRenderInfo renderInfo)
        {
            
        }

        public string GetResultantText()
        {
            return this.result.ToString();
        }
    }

    class DocumentInfo
    {
        public List<ShapeInfo> Shapes = new List<ShapeInfo>();

    }

    class ShapeInfo
    {
        public ShapeInfo(Shape shape)
        {
            string strStyle = shape.Style.Value;
            ParseStyle(strStyle);
            Text = shape.InnerText;
            Name = shape.Id.Value;
        }

        public ShapeInfo(Rectangle rectangle)
        {
            string strStyle = rectangle.Style.Value;
            ParseStyle(strStyle);
            Text = rectangle.InnerText;
            Name = rectangle.Id.Value;

            if (rectangle.Gfxdata.HasValue)
            {
                string strPicBase64 = rectangle.Gfxdata.Value.Replace("\n","");
                byte[] bytes = Convert.FromBase64String(strPicBase64);
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    Bitmap b = new Bitmap(ms);
                    Img = new Bitmap(ms);
                }
            }
        }

        static byte[] AnotherDecode64(string base64Decoded)
        {
            string temp = base64Decoded.TrimEnd('=');
            int asciiChars = temp.Length - temp.Count(c => Char.IsWhiteSpace(c));
            switch (asciiChars % 4)
            {
                case 1:
                    //This would always produce an exception!!
                    //Regardless what (or what not) you attach to your string!
                    //Better would be some kind of throw new Exception()
                    return new byte[0];
                case 0:
                    asciiChars = 0;
                    break;
                case 2:
                    asciiChars = 2;
                    break;
                case 3:
                    asciiChars = 1;
                    break;
            }
            temp += new String('=', asciiChars);

            return Convert.FromBase64String(temp);
        }

        private void ParseStyle(string strStyle)
        {
            foreach (string s in strStyle.Split(';'))
            {
                if (string.IsNullOrEmpty(s))
                    continue;

                string[] split = s.Split(':');
                Styles.Add(split[0], split[1]);
            }

            Left = GetStyleValue<double>("margin-left");
            Top = GetStyleValue<double>("margin-top");
            Width = GetStyleValue<double>("width");
            Height = GetStyleValue<double>("height");
        }

        private Dictionary<string, string> Styles = new Dictionary<string, string>();
        public double Left { get; private set; }
        public double Top { get; private set; }
        public double Width { get; private set; }
        public double Height { get; private set; }
        public string Text { get; private set; }
        public string Name { get; private set; }
        public Image Img { get; private set; }

        private T GetStyleValue<T>(string key)
        {
            bool bExist = Styles.ContainsKey(key);
            if (typeof(double) == typeof(T))
            {
                if (!bExist)
                    return (T) (object) 0;

                string val = Styles[key].Replace("pt", "");
                return (T) (object) Convert.ToDouble(val);
            }

            if (!bExist)
                return (T) (object) "";

            return (T) (object) Styles[key];
        }
    }
}
