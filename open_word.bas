Attribute VB_Name = "Module1"
Sub open_word_test()
Call open_word("C:\Users\user\Downloads\sample.docx")
End Sub
Sub open_word(ByVal FilePath As String)
    Dim objWord As Word.Application
    Dim objDoc As Word.Document
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    Set objDoc = objWord.Documents.Open(FilePath)
End Sub
