using System;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class DocxConverter : DocxToPdfRepository
    {
        public static string path;
        public DirectoryInfo rootDir;
        public DirectoryInfo[] stationDirs;
        public DirectoryInfo currentDocDir;
        public DirectoryInfo[] subDirs;
        public static bool emptyDir;
        public static Application app;
        public static Document doc;
        //függvény paraméterek
        public static string line;
        public static string station;
        public static string staticPath; //segédváltozó a path-hoz, ahhoz az esethez kell, amikor 1 paraméterünk van a függvényben 

        public DocxConverter(string path)
        {
            DocxConverter.path = path;
            this.rootDir = new DirectoryInfo(path);
            this.stationDirs = this.rootDir.GetDirectories();
            DocxConverter.emptyDir = true;
            DocxConverter.app = new Application();
            DocxConverter.doc = new Document();
        }

        //string path = @"\\sgmscfs01\SGHU_TEF_Cloud$\Starter\" + Line + "\\" + Station + "\\Documentations\\Instructions\\Work_instructions\\";
        public int docxToPdf(string line, string station)
        {
            for (int i = 0; i < stationDirs.Length; i++)
            {
                currentDocDir = new DirectoryInfo(stationDirs[i].FullName);
                for (int j = 0; j < currentDocDir.GetDirectories().Length; j++)
                {
                    subDirs = currentDocDir.GetDirectories();
                    FileInfo[] inputFiles = subDirs[j].GetFiles();
                    if (inputFiles.Length > 0)
                    {
                        emptyDir = false;
                        for (int k = 0; k < inputFiles.Length; k++)
                        {
                            if ((inputFiles[k].Extension == ".doc" || inputFiles[k].Extension == ".docx") && inputFiles[k].Attributes.ToString() == "Archive")
                            {
                                string fileName = Path.GetFileNameWithoutExtension(inputFiles[k].Name);
                                Console.WriteLine("Converting \"" + inputFiles[k] + "\" to PDF...");

                                doc = app.Documents.Open(inputFiles[k].FullName);               //open word
                                string exportPath = currentDocDir + "\\" + subDirs[j] + "\\" + fileName + ".pdf";  //ebbe a mappába rakja a konvertált fájlokat
                                doc.ExportAsFixedFormat(exportPath, WdExportFormat.wdExportFormatPDF);

                                object saveOption = WdSaveOptions.wdDoNotSaveChanges;       //docx mentés nélkül
                                object originalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
                                object routeDocument = false;

                                Console.WriteLine("\"" + inputFiles[k] + "\"" + "is READY!\n");

                                doc.Close(saveOption, originalFormat, routeDocument);
                            }
                        }
                    }
                }
            }

            if (emptyDir == true)
            {
                Console.WriteLine("There are not new documents to covert!");
                Console.WriteLine("___________________________________________________");
                Console.WriteLine("\nPress ENTER to EXIT!");
            }
            else
            {
                //Terminate background processes
                app.Quit();
                Console.WriteLine("___________________________________________________");
                Console.WriteLine("\nConverting is ready... Press ENTER to EXIT!");
            }

            return 0;
        }

        public int docxToPdfWithOneParameter(string line)
        {
            // 1-től indítom, hogy a _MainLine mappa ne, legyen benne
            for (int i = 1; i < stationDirs.Length; i++)
            {
                currentDocDir = new DirectoryInfo(stationDirs[i].FullName + staticPath);
                for (int j = 0; j < currentDocDir.GetDirectories().Length; j++)
                {
                    subDirs = currentDocDir.GetDirectories();
                    FileInfo[] inputFiles = subDirs[j].GetFiles();
                    if (inputFiles.Length > 0)
                    {
                        emptyDir = false;
                        for (int k = 0; k < inputFiles.Length; k++)
                        {
                            if ((inputFiles[k].Extension == ".doc" || inputFiles[k].Extension == ".docx") && inputFiles[k].Attributes.ToString() == "Archive")
                            {
                                string fileName = Path.GetFileNameWithoutExtension(inputFiles[k].Name);

                                Console.WriteLine("Converting \"" + inputFiles[k] + "\" to PDF...");
                                doc = app.Documents.Open(inputFiles[k].FullName);
                                string exportPath = currentDocDir + "\\" + subDirs[j] + "\\" + fileName + ".pdf";  //ebbe a mappába rakja a konvertált fájlokat
                                doc.ExportAsFixedFormat(exportPath, WdExportFormat.wdExportFormatPDF);

                                object saveOption = WdSaveOptions.wdDoNotSaveChanges;
                                object originalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
                                object routeDocument = false;

                                Console.WriteLine("\"" + inputFiles[k] + "\"" + "is READY!\n");

                                doc.Close(saveOption, originalFormat, routeDocument);
                            }
                        }
                    }
                }
            }
            if (emptyDir == true)
            {
                Console.WriteLine("There are not new documents to covert!");
                Console.WriteLine("___________________________________________________");
                Console.WriteLine("\nPress ENTER to EXIT!");
            }
            else
            {
                //Terminate background processes
                app.Quit();
                Console.WriteLine("___________________________________________________");
                Console.WriteLine("\nConverting is ready... Press ENTER to EXIT!");
            }
            return 0;
        }

    }
}






//CLOSE PROCESSES

/*
  System.Runtime.InteropServices.Marshal.FinalReleaseComObject(appWord);
    try
    {
        {
            wordDocument.Close(saveOption, originalFormat, routeDocument);
            appWord.Quit();
        }
        //Marshal.ReleaseComObject(wordDocument);

        if (wordDocument != null) Marshal.ReleaseComObject(wordDocument);
        Marshal.ReleaseComObject(wordDocument);
        Marshal.ReleaseComObject(appWord);
        wordDocument = null;
        appWord = null;
    }
    catch { }
    finally { GC.Collect(); }
    */




//NECCESSARY VARS

/*string path = @"\\sgmscfs01\SGHU_TEF_Cloud$\Starter\" + line + "\\" + station + "\\Documentations\\Instructions\\Work_instructions\\";
DirectoryInfo rootDir = new DirectoryInfo(path);    //Work_instructions
DirectoryInfo[] stationDirs = rootDir.GetDirectories();    //ATA, ATU, ELU...
DirectoryInfo currentDocDir;                       //ATA...
DirectoryInfo[] subDirs;                        //V01.00.00, V01.00.01...
bool emptyDir = true;       //Ha nincs .doc vagy .docx fájl lépjen ki a program

Application appWord = new Application();    //word applicaton object
Document wordDocument;                      //word document object
*/
