using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Rectangle = System.Drawing.Rectangle;

namespace WordAddIn2
{
    public partial class Ribbon1
    {
        private Application _application;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _application = Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            int l, t, w, h;
            for (var i = 1; i <= _application.ActiveDocument.Paragraphs.Count; i++)
            {
                var obj = _application.ActiveDocument.Paragraphs[i].Range;
                _application.ActiveDocument.ActiveWindow.GetPoint(out l, out t, out w, out h, obj);

                object hPos = obj.Information[WdInformation.wdHorizontalPositionRelativeToPage];
                object vPos = obj.Information[WdInformation.wdVerticalPositionRelativeToPage];
                hPos = Pound2Pixel((float) hPos);
                vPos = Pound2Pixel((float) vPos);

                Rectangle rect = new Rectangle(Convert.ToInt32(hPos), Convert.ToInt32(vPos), w, h);
                string text = obj.Text;

                System.Diagnostics.Debug.WriteLine($"text:{text}");
                System.Diagnostics.Debug.WriteLine($"rect:{rect}");
            }

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
}
