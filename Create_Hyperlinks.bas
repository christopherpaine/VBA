Attribute VB_Name = "Module2"
Sub create_hyperlink_record()
    Selection.Copy
    Range("G815").Select
    Application.CutCopyMode = False
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "Pdf!B2974", TextToDisplay:="Statement 69"
End Sub

