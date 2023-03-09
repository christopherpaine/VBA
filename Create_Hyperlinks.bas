Attribute VB_Name = "Module2"



'Some thoughts on creating functions\subprocedures for hyperlinks
'----------------------------------------------------------------
'----------------------------------------------------------------
'
'   user defined function
'   --------------------
'
'       a user defined function with various input parameters would be the neatest solution
'       however an initial attempt at this has not yielded results
'
'
'
'
'
'   external reading sources
'   ------------------------
'
'       https://stackoverflow.com/questions/9560742/creating-a-custom-hyperlink-function-in-excel
'


Sub create_hyperlink_record()
    Selection.Copy
    Range("G815").Select
    Application.CutCopyMode = False
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "Pdf!B2974", TextToDisplay:="Statement 69"
End Sub

