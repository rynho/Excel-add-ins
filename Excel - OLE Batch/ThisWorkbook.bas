Private Const TOOLBARNAME = "MyFunkyNewToolbar"

Private Sub Workbook_Open()
    On Error Resume Next
    Application.Windows(1).Activate
    ShowToolbar
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DeleteCommandBar
End Sub

Public Sub ShowToolbar()
    On Error Resume Next
    Dim iniTest
    iniTest = Application.CommandBars(TOOLBARNAME).Index
    If iniTest = 0 Then
        Application.CommandBars.Add TOOLBARNAME
        AddButton "Button caption", "This is a tooltip", 681, "ThisWorkbook.openform"
        ' call AddButton more times for more buttons '

        With Application.CommandBars(TOOLBARNAME)
            .Visible = True
            .Position = msoBarTop
        End With
    End If
End Sub

Private Sub AddButton(caption As String, tooltip As String, faceId As Long, methodName As String)
    Dim Btn As CommandBarButton
    Set Btn = Application.CommandBars(TOOLBARNAME).Controls.Add
    With Btn
        .Style = msoButtonIcon
        .faceId = faceId ' choose from a world of possible images in Excel: see http://www.ozgrid.com/forum/showthread.php?t=39992 '
        .OnAction = methodName
        .TooltipText = tooltip
    End With
End Sub

Public Sub DeleteCommandBar()
    On Error Resume Next
    Dim iniTest
    For Each wnd In Application.Windows
        wnd.Activate

        iniTest = Application.CommandBars(TOOLBARNAME).Index
        If iniTest > 0 Then _
            Application.CommandBars(TOOLBARNAME).Delete
    Next wnd
End Sub

Public Sub openform()
    UserForm1.Show
End Sub

