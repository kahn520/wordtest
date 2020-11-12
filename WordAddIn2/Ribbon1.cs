using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Rectangle = System.Drawing.Rectangle;
using Shape = Microsoft.Office.Interop.Word.Shape;

namespace WordAddIn2
{

    public partial class Ribbon1
    {
        private const int MsoTrue = -1;
        private const int MsoFalse = 0;

        private Application _application;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _application = Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            _application.ScreenUpdating = false;
            Document doc = _application.ActiveDocument;

            DocumentInfo docInfo = new DocumentInfo();
            docInfo.Size.Width = doc.PageSetup.PageWidth;
            docInfo.Size.Height = doc.PageSetup.PageHeight;
            for (var i = 1; i <= doc.Paragraphs.Count; i++)
            {
                var obj = doc.Paragraphs[i].Range;
                TextRange textRange = GetTextRangInfo(obj);
                if (textRange != null)
                    docInfo.TextRanges.Add(textRange);
            }

            for (int i = 1; i <= doc.Shapes.Count; i++)
            {
                Shape s = doc.Shapes[i];
                var list = GetShapeTextRange(s);
                if(list.Any())
                    docInfo.TextRanges.AddRange(list);
            }

            _application.ScreenUpdating = true;
            File.WriteAllText(@"I:\yoo\out.json", JsonConvert.SerializeObject(docInfo), Encoding.UTF8);
        }

        private List<TextRange> GetShapeTextRange(Shape s)
        {
            List<TextRange> list = new List<TextRange>();
            if (s.Type == MsoShapeType.msoGroup)
            {
                foreach (Shape childShape in s.GroupItems)
                {
                    list.AddRange(GetShapeTextRange(childShape));
                }
            }
            else
            {
                if (s.TextFrame.HasText == MsoTrue)
                {
                    for (var i = 1; i <= s.TextFrame.TextRange.Paragraphs.Count; i++)
                    {
                        var obj = s.TextFrame.TextRange.Paragraphs[i].Range;
                        TextRange textRange = GetTextRangInfo(obj);
                        if (textRange != null)
                            list.Add(textRange);
                    }
                }
            }

            return list;
        }

        private TextRange GetTextRangInfo(Range range)
        {
            TextRange textRange = new TextRange();
            textRange.Rect = GetRangeRect(range);
            if (textRange.Rect.IsEmpty)
                return null;

            textRange.Text = TrimText(range.Text);
            textRange.Page = (int)range.Information[WdInformation.wdActiveEndPageNumber];
            textRange.TextStyle = GetTextStyle(range);
            return textRange;
        }

        private string TrimText(string text)
        {
            text = text.Replace("\r", "");
            text = text.Replace("\u0007", "");
            return text;
        }

        private Rect GetRangeRect(Range range)
        {
            try
            {
                int l, t, w, h;
                _application.ActiveDocument.ActiveWindow.GetPoint(out l, out t, out w, out h, range);
                object hPos = range.Information[WdInformation.wdHorizontalPositionRelativeToPage];
                object vPos = range.Information[WdInformation.wdVerticalPositionRelativeToPage];
                w = (int)Pixel2Pound(w);
                h = (int)Pixel2Pound(h);
                Rect r = new Rect();
                r.Left = (float)hPos;
                r.Top = (float)vPos;
                r.Width = w;
                r.Height = h;
                return r;
            }
            catch (Exception e)
            {
                return new Rect();
            }
        }

        private TextStyle GetTextStyle(Range range)
        {
            TextStyle textStyle = new TextStyle();
            textStyle.FontName = range.Font.Name;
            textStyle.FontSize = range.Font.Size;
            textStyle.Blod = range.Font.Bold == MsoTrue;
            return textStyle;
        }

        private float Pound2Pixel(float pound)
        {
            return pound / 72 * 96;
        }

        private float Pixel2Pound(float pixel)
        {
            return pixel / 96 * 72;
        }
    }

    public class DocumentInfo
    {
        public Size Size { get; set; } = new Size();
        public List<TextRange> TextRanges { get; set; } = new List<TextRange>();
    }

    public class Size
    {
        public float Width { get; set; }
        public float Height { get; set; }
    }

    public class TextRange
    {
        public string Text { get; set; }
        public int Page { get; set; }
        public TextStyle TextStyle { get; set; } = new TextStyle();
        public Rect Rect { get; set; } = new Rect();
    }

    public class TextStyle
    {
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool Blod { get; set; }
    }

    public class Rect
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }

        [JsonIgnore]
        public bool IsEmpty => Width <= 0 && Height <= 0;
    }
}
