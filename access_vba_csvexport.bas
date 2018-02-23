
Public Sub exportToXl()

        On Error GoTo ErrorHandler 'Error handling
        
        Dim dbTable As String
        Dim db As Database
        Dim Tbl As TableDef
        Dim TblNames As String
        Dim xlworkpath As String
        
        Set db = CurrentDb 'select current access database open in access
        
        xlworkpath = "D:\TEST\" 'custom path for storing all the exported csv files
        
        For Each Tbl In db.TableDefs
                If Tbl.Attributes = 0 Then 'Ignores System Tables
                xlworkpath = xlworkpath & Tbl.Name & ".csv" 'here .csv extension can be changed to rtf
                DoCmd.TransferText TransferType:=acExportDelim, Tablename:=Tbl.Name, FileName:=xlworkpath, hasfieldnames:=True
                End If
                xlworkpath = "D:\TEST\" 'this resets the file name to avoid concatenating all the names
        Next
                db.Close
                Set db = Nothing

ErrorHandlerExit:

        Exit Sub
 
ErrorHandler:

        MsgBox "Error No: " & Err.Number & ";Description" & Err.Description
        Resume ErrorHandlerExit
End Sub
