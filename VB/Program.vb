Imports System
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Layout
Imports DevExpress.XtraRichEdit.API.Native

Module Program
    Sub Main(args As String())
        Using rtfProcessor As New RichEditDocumentServer()
            rtfProcessor.LoadDocument("FirstLook.docx")
            Console.WriteLine(String.Format("This document has {0} page(s)", GetPageCount(rtfProcessor)))

            Dim textRangeCollection() As DocumentRange = FindTextOnPage(rtfProcessor, "network", 4)
            Console.WriteLine(String.Format("There are(is) {0} 'network' entries on page 4", textRangeCollection.Length))

            Dim bookmark As Bookmark = rtfProcessor.Document.Bookmarks.Where(Function(b As Bookmark) b.Name = "Appendix").FirstOrDefault()
            If bookmark IsNot Nothing Then
                Dim pageWithBookmark As Integer = GetPageNumberFromPosition(rtfProcessor, bookmark.Range.Start)
                Console.WriteLine(String.Format("The 'Appendix' bookmark is located on page {0}", pageWithBookmark))
            End If

            Console.WriteLine("This text is extracted from the last page: ")
            Console.WriteLine(GetPageText(rtfProcessor, 5))
        End Using
        Console.ReadKey()
    End Sub
    Function GetPageCount(ByVal rtfProcessor As RichEditDocumentServer) As Integer
        Dim documentLayout = rtfProcessor.DocumentLayout
        Dim pageCount As Integer = documentLayout.GetPageCount()
        Return pageCount
    End Function
    Function FindTextOnPage(ByVal rtfProcessor As RichEditDocumentServer, ByVal text As String, ByVal pageNumber As Integer) As DocumentRange()
        Dim pageDocumentRange As DocumentRange = GetPageDocumentRange(rtfProcessor, pageNumber)
        Dim words() As DocumentRange = rtfProcessor.Document.FindAll(text, SearchOptions.WholeWord, pageDocumentRange)
        Return words
    End Function
    Function GetPageNumberFromPosition(ByVal rtfProcessor As RichEditDocumentServer, ByVal pos As DocumentPosition) As Integer
        Dim documentLayout = rtfProcessor.DocumentLayout
        Dim row = documentLayout.GetElement(pos, LayoutType.Row)
        Dim page As LayoutPage = row.GetParentByType(Of LayoutPage)()
        Return page.Index + 1
    End Function
    Function GetPageText(ByVal rtfProcessor As RichEditDocumentServer, ByVal pageNumber As Integer) As String
        Dim pageDocumentRange As DocumentRange = GetPageDocumentRange(rtfProcessor, pageNumber)
        Return rtfProcessor.Document.GetText(pageDocumentRange)
    End Function

    Function GetPageDocumentRange(ByVal rtfProcessor As RichEditDocumentServer, ByVal pageNumber As Integer) As DocumentRange
        Dim documentLayout = rtfProcessor.DocumentLayout
        Dim page As LayoutPage = documentLayout.GetPage(pageNumber - 1)
        Dim pageDocumentRange As DocumentRange = rtfProcessor.Document.CreateRange(page.MainContentRange.Start, page.MainContentRange.Length)
        Return pageDocumentRange
    End Function


End Module
