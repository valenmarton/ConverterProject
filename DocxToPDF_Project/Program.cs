using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;


//FONTOS!!!
//Ha CMD-ből futtatod, vedd ki a Console.ReadLine-t a végéről, különben nem fog tovább futni a scripted !
namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        { 
            //ha nincs Main() argumentum standard inputról olvasás, 
            //egyébként a command line argumentumok olvasása (beágyazott rendszer miatt)
            //line = a sor mappája, amit rekurzívan konvertálni szeretnél
            //station = az állomás mappája, amit rekurzívan konvertálni szeretnél

            if (args.Length == 0)
            {
                Console.Write("Add meg a sort: ");
                DocxConverter.line = Console.ReadLine().ToUpper();
                Console.Write("Add meg az állomást: ");
                DocxConverter.station = Console.ReadLine().ToUpper();
                Console.WriteLine("___________________________________________________\n");

                DocxConverter.path = @"\\sgmscfs01\SGHU_TEF_Cloud$\Starter\" + DocxConverter.line + "\\" + DocxConverter.station + "\\Documentations\\Instructions\\Work_instructions\\";
                string rootPath = DocxConverter.path;
                DocxConverter docObj = new DocxConverter(rootPath);
                docObj.docxToPdf(DocxConverter.line, DocxConverter.station);
            }
            else if (args.Length == 1)
            {
                DocxConverter.staticPath = "\\Documentations\\Instructions\\Work_instructions\\";
                string rootPath = DocxConverter.path;
                DocxConverter docObj = new DocxConverter(rootPath);

                docObj.docxToPdfWithOneParameter(args[0].ToUpper());
            }
            else if (args.Length == 2)
            {
                DocxConverter.path = @"\\sgmscfs01\SGHU_TEF_Cloud$\Starter\" + DocxConverter.line + "\\" + DocxConverter.station + "\\Documentations\\Instructions\\Work_instructions\\";
                string rootPath = DocxConverter.path;
                DocxConverter docObj = new DocxConverter(rootPath);
                docObj.docxToPdf(args[0].ToUpper(), args[1].ToUpper());
            }
            else
            {
                Console.WriteLine("Túl sok argumentum! Maximum 2 argumentum lehet, az 1. a sor neve, 2. az állomás neve!");
            }
            Console.ReadLine();
        }
    }
}
