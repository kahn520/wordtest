using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronPdf;

namespace IronPdfTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string strFile = @"D:\yoo\【着色算法】可根据不同着色技巧对图示进行重新染色.pdf";
            PdfDocument doc = PdfDocument.FromFile(strFile);
            for (int i = 0; i < doc.PageCount; i++)
            {
                string s = doc.ExtractTextFromPage(i);
            }
        }
    }
}
