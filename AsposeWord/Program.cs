using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;

namespace AsposeWord
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document(@"I:\yoo\Doc1.docx");
            foreach (Node paragraph in doc.FirstSection.Body.Paragraphs)
            {
                
            }
        }
    }
}
