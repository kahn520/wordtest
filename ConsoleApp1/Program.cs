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
using Rectangle = DocumentFormat.OpenXml.Vml.Rectangle;
using Shape = DocumentFormat.OpenXml.Vml.Shape;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string strFile = @"I:\yoo\Doc1.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(strFile, true))
            {
                DocumentInfo docInfo = new DocumentInfo();
                var body = doc.MainDocumentPart.Document.Body;
                var shapes = body.Descendants<Shape>();
                foreach (Shape shape in shapes)
                {
                    ShapeInfo shapeInfo = new ShapeInfo(shape);
                    if (shapeInfo.Text == "")
                        continue;

                    docInfo.Shapes.Add(shapeInfo);
                }


                var rectangles = body.Descendants<Rectangle>();
                foreach (Rectangle rect in rectangles)
                {
                    ShapeInfo shapeInfo = new ShapeInfo(rect);
                    if (shapeInfo.Text == "")
                        continue;

                    docInfo.Shapes.Add(shapeInfo);
                }

                string strJsonFile = Path.ChangeExtension(strFile, ".json");
                string strJson = JsonConvert.SerializeObject(docInfo);
                File.WriteAllText(strJsonFile, strJson, Encoding.GetEncoding("GB2312"));
            }
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
