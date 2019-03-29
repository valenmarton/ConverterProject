namespace DocxToPdfCoverter
{
    interface DocxToPdfRepository
    {
        int docxToPdf(string line, string station);         //konvertálás 2 paraméterrel: sor és állomás
        int docxToPdfWithOneParameter(string line);         //konvertálás 1 paraméterrel: sor (rekurzívan összes állomást)
    }
}
