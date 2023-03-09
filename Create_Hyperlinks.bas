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
'           in combination with existing hyperlink function
'           -----------------------------------------------
'
'
'           in combination with application.caller and a manually created hyperlink
'           -----------------------------------------------------------------------
'
'           it seems that the step of creating the hyperlink manually is difficult to avoid
'           however it is possible to do this just ONCE
'           and then just copy your cell to other locations
'
'           application.caller can be used to get the cell in which the UDF is used
'           and manipulate any existing hyperlinks in that cell
'
'           see:  https://stackoverflow.com/questions/9560742/creating-a-custom-hyperlink-function-in-excel
'
'   event handling
'   --------------
'
'
'   external reading sources
'   ------------------------
'
'       https://stackoverflow.com/questions/9560742/creating-a-custom-hyperlink-function-in-excel
'
'
'       https://github.com/muyamima/VBA-Excel_LinkTableMacros/blob/master/HyperlinkMacros.bas
'
'           This code is a VBA script for auditing and testing hyperlinks in an Excel workbook. 
'           The AuditHyperlinks subroutine checks each hyperlink in the workbook and updates the formatting of the
'           cell based on whether the link is working or broken. The TestWebAddress and TestFilePath functions are 
'           helper functions that test whether a given web address or file path is valid. 
'           The code also includes subroutines for adding new hyperlinks and fixing broken relative links.
'
'


'Summary description
'   -   function returns to the cell the value put into the function
'       which ends up being the label for the hyperlink
'   -   application.caller returns the cell and allows modification
'       of any existing hyperlink
Function CustomHyperlink(Term As String,mod_choice as string,optional loc as string) As String
    Dim rng As Range

    Set rng = Application.Caller
    CustomHyperlink = Term

    If rng.Hyperlinks.Count > 0 Then
        select case mod_choice
            Case "google"
                rng.Hyperlinks(1).Address = "http://www.google.com/search?q=" & Term
            Case "internal"        
                rng.Hyperlinks(1).Address = ""
                rng.Hyperlinks(1).SubAddress = loc
        end select
    End If
End Function







'results from recording creation of a hyperlink

Sub create_hyperlink_record()
    Selection.Copy
    Range("G815").Select
    Application.CutCopyMode = False
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "Pdf!B2974", TextToDisplay:="Statement 69"
End Sub

