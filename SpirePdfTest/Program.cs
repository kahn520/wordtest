using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Pdf;
using Spire.Pdf.General.Find;

namespace SpirePdfTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //string strFile = @"D:\yoo\【着色算法】可根据不同着色技巧对图示进行重新染色.pdf";
            string strFile = @"D:\yoo\张甜甜-Java开发.pdf";
            using (FileStream fs = new FileStream(strFile, FileMode.Open))
            {
                PdfDocument doc = new PdfDocument(fs);
                foreach (PdfPageWidget page in doc.Pages)
                {
                    var s = page.ExtractText();
                    var collection = page.FindAllText();
                }
            }
        }
    }
}
