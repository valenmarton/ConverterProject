using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    interface DocxToPdfRepository
    {
        int docxToPdf(string line, string station);         //konvertálás 2 paraméterrel: sor és állomás
        int docxToPdfWithOneParameter(string line);         //konvertálás 1 paraméterrel: sor (rekurzívan összes állomást)
    }
}
