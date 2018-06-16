'version @20180306, added editable offset input.
'version @20170903, added PowerPoint icon.

Private Sub Horiz_Click()
    xoff = 1
    yoff = -1
End Sub

Private Sub Verti_Click()
    xoff = 0
    yoff = 5
End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, _
    Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim FPath, FName, FExt, FIco(2)
    
    FPath = Data.Files(1)

    If FPath <> "" Then
        FName = Mid(FPath, InStrRev(FPath, "\") + 1)
        FExt = Mid(FName, InStrRev(FName, ".") + 1)

        Select Case FExt
            Case "pdf"
                FIco(0) = "C:\Program Files (x86)\Adobe\Acrobat Reader 2015\Reader\AcroRd32Res.dll"
                FIco(1) = 0
            Case "doc", "docx"
                FIco(0) = "c:\program files (x86)\microsoft office\office16\winword.exe"
                FIco(1) = 1
            Case "xls", "xlsx"
                FIco(0) = "c:\program files (x86)\microsoft office\office16\excel.exe"
                FIco(1) = 1
            Case "ppt", "pptx"
                FIco(0) = "c:\program files (x86)\microsoft office\office16\powerpnt.exe"
                FIco(1) = 1
            Case "bmp", "jpg", "jpeg", "png", "gif"
                FIco(0) = "C:\windows\system32\SHELL32.dll"
                FIco(1) = 302
            Case Else
                FIco(0) = "C:\windows\system32\SHELL32.dll"
                FIco(1) = 0
        End Select
        
        ActiveCell = FName
        ActiveCell.Offset(1, 0).Select
    
        ActiveSheet.OLEObjects.Add(Filename:=FPath, Link:=False, _
            DisplayAsIcon:=True, IconFileName:=FIco(0), _
            IconIndex:=FIco(1), IconLabel:=FName).Select
        
        ActiveCell.Offset(yoff, xoff).Select
           
    End If

End Sub

