using DevExpress.Office.Utils;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Layout;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Drawing;
using System.Linq;

namespace WordProcessorLayoutAPISample
{
    class Program
    {
        //Layout API - How to perform page-based operations: get page text, page number and page count in a document, search for text on a specific page
        static void Main(string[] args)
        {
            using (RichEditDocumentServer rtfProcessor = new RichEditDocumentServer())
            {
                rtfProcessor.LoadDocument("FirstLook.docx");
                Console.WriteLine(string.Format("This document has {0} page(s)", GetPageCount(rtfProcessor)));

                DocumentRange[] textRangeCollection = FindTextOnPage(rtfProcessor, "network", 4);
                Console.WriteLine(string.Format("There are(is) {0} 'network' entries on page 4", textRangeCollection.Length));

                Bookmark b = rtfProcessor.Document.Bookmarks.Where((b) => b.Name == "Appendix").FirstOrDefault();
                if (b != null)
                {
                    int pageWithBookmark = GetPageNumberFromPosition(rtfProcessor, b.Range.Start);
                    Console.WriteLine(string.Format("The 'Appendix' bookmark is located on page {0}", pageWithBookmark));
                }

                Console.WriteLine("This text is extracted from the last page: ");
                Console.WriteLine(GetPageText(rtfProcessor, 5));
            }
            Console.ReadKey();
        }
        static int GetPageCount(RichEditDocumentServer rtfProcessor)
        {
            var documentLayout = rtfProcessor.DocumentLayout;
            int pageCount = documentLayout.GetPageCount();
            return pageCount;
        }
        static DocumentRange[] FindTextOnPage(RichEditDocumentServer rtfProcessor, string text, int pageNumber)
        {
            DocumentRange pageDocumentRange = GetPageDocumentRange(rtfProcessor, pageNumber);
            DocumentRange[] words = rtfProcessor.Document.FindAll(text, SearchOptions.WholeWord, pageDocumentRange);
            return words;
        }
        static int GetPageNumberFromPosition(RichEditDocumentServer rtfProcessor, DocumentPosition pos)
        {
            var documentLayout = rtfProcessor.DocumentLayout;
            var row = documentLayout.GetElement(pos, LayoutType.Row);
            LayoutPage page = row.GetParentByType<LayoutPage>();
            return page.Index + 1;
        }
        static string GetPageText(RichEditDocumentServer rtfProcessor, int pageNumber)
        {
            DocumentRange pageDocumentRange = GetPageDocumentRange(rtfProcessor, pageNumber);
            return rtfProcessor.Document.GetText(pageDocumentRange);
        }
       
        private static DocumentRange GetPageDocumentRange(RichEditDocumentServer rtfProcessor, int pageNumber)
        {
            var documentLayout = rtfProcessor.DocumentLayout;
            LayoutPage page = documentLayout.GetPage(pageNumber - 1);
            DocumentRange pageDocumentRange = rtfProcessor.Document.CreateRange(page.MainContentRange.Start, page.MainContentRange.Length);
            return pageDocumentRange;
        }
    }
}
