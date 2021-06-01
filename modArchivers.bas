Attribute VB_Name = "modArchivers"
Sub unzip(zipname As String, workfolder As String)
    vbzipnam.s(0) = vbNullString
    vbxnames.s(0) = vbNullString
    VBUnzip zipname, workfolder, 0, 1, 0, 0, 0, 0
End Sub
Sub zip(files As String)
Dim mynames As ZIPnames
    argc = Separate(files)
    For i = 2 To argc
        mynames.s(i - 1) = Items(i)
    Next i
    x = VBZip(argc - 1, Items(1), mynames, 0, 0, 0, 0, "C:\")
End Sub
