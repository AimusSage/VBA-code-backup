Sub test()

Call StoringenDataVerwerken("S:\Stedin\Diensten\SMB\Operations KVB\Datacollectie\Uitvalbakken\05_Rapportage\StatsMFE", "MFE stats MeterInfo", "storingencount")

MsgBox "hello"

End Sub

Public Sub StoringenDataVerwerken(strpath As String, specs As String, tblName As String)
'On Error GoTo errHandler
    Dim tbl_col As Collection
    Dim dateField As String

    Dim strFilename As String 'Filename
    Dim strFileList() As String 'File  Array
    Dim intFile As Integer 'File Number
    Dim MFEtable As String

Call closeAllObjects(False)
DoCmd.SetWarnings False
Set db = CurrentDb()

'---------------------------------
'Lijst met te importeren MFE dumps genereren
Call DeleteTable("Temp")

If specs = "" Then
    SQL = "SELECT MSysIMEXSpecs.SpecName FROM MSysIMEXSpecs"
    Set rs = CurrentDb.OpenRecordset(SQL)
    
    Do While Not rs.EOF
        impspecstring = impspecstring & rs("SpecName") & "; "
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    specs = InputBox("Voor de naam van de import specificaties in, indien niets ingevoerd, wordt de standaard specificatie gebruikt" & _
    vbNewLine & "Opties: " & impspecstring)
    If specs = "" Then
        MsgBox "Geen specificatie opgegeven, probeer opnieuw en geef een importspecificatie op"
    End If
End If

Set fd = Application.FileDialog(msoFileDialogFolderPicker)
With fd
    .Title = "Selecteer de folder met de te importeren bestanden"
    .AllowMultiSelect = False
    .InitialFileName = strpath
    If .Show <> -1 Then GoTo exitVerwerken
    strpath = .SelectedItems(1) & "\"
End With
Set fd = Nothing

     'Loop through the folder & build file list
    strFilename = Dir(strpath & "*.csv")
    While strFilename <> ""
         'add files to the list
        intFile = intFile + 1
        ReDim Preserve strFileList(1 To intFile)
        strFileList(intFile) = strFilename
        strFilename = Dir()
    Wend
     'see if any files were found
    If intFile = 0 Then
        MsgBox "No files found"
        Exit Sub
    End If
     'cycle through the list of files &  import to Access
 
Call sortList(strFileList, 1, UBound(strFileList))
For intFile = 1 To UBound(strFileList)
    Call DeleteTable("temp")
    MFEDate = Format(DateSerial(Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 14, 4), Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 10, 2), Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 8, 2)), "yyyy-mm-dd")
    MFEtable = "MFE stats_" & Format(MFEDate, "dd-mm-yyyy")
    If FieldExists(tblName, CStr(Format(MFEDate, "dd-mm-yyyy"))) = False Then
        DoCmd.TransferText acImportDelim, specs, "temp", strpath & strFileList(intFile), True 'Creates table temp for importing the csv
        
        MFEDate = Format(DateSerial(Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 14, 4), Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 10, 2), Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 8, 2)), "yyyy-mm-dd")
        
        SQL = "ALTER TABLE [temp] ADD COLUMN [MFE Datum] datetime;"
        
            db.Execute SQL, dbFailOnError                                    'Adds field MFE Datum to temp table
        SQL = "UPDATE [temp] SET [temp].[MFE datum] = '" & MFEDate & "';"
            CurrentDb.Execute SQL                                               'Fills field "MFE Datum" with date in filename of CSV
                  
        SQL = "UPDATE [temp] SET [e-meternummer] = trim([e-meternummer]), [g-meternummer] = trim([g-meternummer])"
            db.Execute SQL                                               'Cleans up the meternumber field in New MFE to eliminate white spaces
        Call deleteImportErrorTables
        
        'Updating StoringenCount with new data
        Set td = db.TableDefs(tblName)
        Set rs = db.OpenRecordset("temp")
        dateField = Format(rs("MFE Datum").Value, "dd-mm-yyyy")
        rs.Close
        Set rs = Nothing
        If FieldExists(tblName, dateField) = False Then td.Fields.Append td.CreateField(dateField, dbDate)

        SQL = "UPDATE " & tblName & " as t1 INNER JOIN temp as t2 ON t1.meternummer = t2.[e-meternummer] SET t1.[" & dateField & "] = t2.[laatste comm E]"
            db.Execute SQL
            
        SQL = "UPDATE " & tblName & " as t1 INNER JOIN temp as t2 ON t1.meternummer = t2.[g-meternummer] SET t1.[" & dateField & "] = t2.[laatste comm g]"
            db.Execute SQL
    
        Call closeAllObjects(True)
    End If
Next

exitVerwerken:
    Set rs = Nothing
    Set rsSource = Nothing
    Set db = Nothing
    DoCmd.SetWarnings True
    Exit Sub

errHandler:
    Set rs = Nothing
    Set rsSource = Nothing
    Set db = Nothing
    DoCmd.SetWarnings True
    MsgBox "Er is iets fout gegaan, foutcode: " & Err.Number & " " & Err.Description

End Sub


Sub testing()

Set db = CurrentDb()
On Error Resume Next

DoCmd.DeleteObject acTable, "storingenCount"
skipDelete:
MsgBox "Hello"
Set db = Nothing

End Sub

Sub countDays()

Dim fld_col As New Collection
Dim fld As Field
Dim fldName As String
Dim fldnamePrev As String
Dim fldNameLast As String
Dim updFld As String
Call closeAllObjects(True)
Set db = CurrentDb()

Set td = db.TableDefs("storingencount")
    If FieldExists("storingencount", "commCountSeq") = False Then
        td.Fields.Append td.CreateField("commCountSeq", dbInteger)
    End If
    If FieldExists("storingencount", "commCount") = False Then
        td.Fields.Append td.CreateField("commCount", dbInteger)
    End If
    If FieldExists("storingencount", "countInstance") = False Then
        td.Fields.Append td.CreateField("countInstance", dbInteger)
    End If
Set rs = db.OpenRecordset("storingencount")

DoCmd.SetWarnings False
DoCmd.Hourglass True
Call closeAllObjects(True)
For Each fld In td.Fields
    If IsNumeric(Left(CStr(fld.Name), 1)) Then
        fld_col.Add CStr(fld.Name)
    End If
Next

'   Get last field in collection (last day)

fldNameLast = CStr(fld_col(fld_col.Count))

'   count each day without communication
Do While Not rs.EOF
    updFld = "commcount"
    rs.Edit
    rs(updFld).Value = 1
    rs.Update
    For i = 2 To fld_col.Count
        fldName = CStr(fld_col(i))
        fldnamePrev = CStr(fld_col(i - 1))
        If IsNull(rs(fldNameLast).Value) Then
            rs.Edit
            rs(updFld).Value = -1
            rs.Update
            Exit For
        ElseIf rs(fldName).Value = rs(fldnamePrev).Value Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value + 1
            rs.Update
        ElseIf rs(fldName).Value < rs(fldnamePrev).Value Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value + 1
            rs.Update
        ElseIf rs(fldName).Value < CDate(fld_col(1)) Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value + 1
            rs.Update
        ElseIf rs(fldName).Value > rs(fldnamePrev).Value Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value
            rs.Update
        End If
    Next
    rs.MoveNext
Loop

' only count until first new communication instance
rs.MoveFirst
Do While Not rs.EOF
    updFld = "commcountSeq"
    rs.Edit
    rs(updFld).Value = 1
    rs.Update
    For i = 2 To fld_col.Count
        fldName = CStr(fld_col(i))
        fldnamePrev = CStr(fld_col(i - 1))
        If IsNull(rs(fldNameLast).Value) Then
            rs.Edit
            rs(updFld).Value = -1
            rs.Update
            Exit For
        ElseIf rs(fldName).Value = rs(fldnamePrev).Value Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value + 1
            rs.Update
        ElseIf rs(fldName).Value < rs(fldnamePrev).Value Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value + 1
            rs.Update
        ElseIf rs(fldName).Value < CDate(fld_col(1)) Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value + 1
            rs.Update
        ElseIf rs(fldName).Value > rs(fldnamePrev).Value Then
            Exit For
        End If
    Next
    rs.MoveNext
Loop

'   Count of instances of no communication regardless of number of sequential days
rs.MoveFirst
Do While Not rs.EOF
    updFld = "countInstance"
    rs.Edit
    rs(updFld).Value = 1
    rs.Update
    For i = 2 To fld_col.Count
        fldName = CStr(fld_col(i))
        fldnamePrev = CStr(fld_col(i - 1))
        If IsNull(rs(fldNameLast).Value) Then
            rs.Edit
            rs(updFld).Value = -1
            rs.Update
            Exit For
        ElseIf rs(fldName).Value = rs(fldnamePrev).Value Then
            If i - 2 < 1 Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value
                rs.Update
            ElseIf rs(fldName).Value > rs(CStr(fld_col(i - 2))).Value And rs(CStr(fld_col(i - 2))).Value >= CDate(fld_col(1)) Then 'checks increase in comm date isn't in the past (2011 issue
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            End If
        ' exceptions for 2011 issue meters
        ElseIf rs(fldName).Value < rs(fldnamePrev).Value Then
            If rs(fldnamePrev).Value < CDate(fld_col(1)) Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value
                rs.Update
            ElseIf i - 2 < 1 Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value
                rs.Update
            ElseIf rs(fldnamePrev).Value > rs(CStr(fld_col(i - 2))).Value Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            End If
        End If
    Next
    rs.MoveNext
Loop
        
rs.Close
Set td = Nothing
Set rs = Nothing

DoCmd.SetWarnings True
DoCmd.Hourglass False

MsgBox "Done"

End Sub

Sub CommunicatiePerDag()

Dim fld_col As New Collection
Dim fld As Field
Dim fldName As String
Dim fldnamePrev As String
Dim fldNameLast As String
Dim updFld As String
Dim rsSource As DAO.Recordset
Dim fldDate As Variant

Set db = CurrentDb()
Call closeAllObjects(False)
If TableExists("storingByDays") Then DeleteTable ("storingByDays")
Set td = db.CreateTableDef("storingByDays")

With td
    .Fields.Append .CreateField("daysOfMonth", dbDate)
    .Fields.Append .CreateField("metersNoCommTot", dbLong)
    .Fields.Append .CreateField("metersNoCommE", dbLong)
    .Fields.Append .CreateField("metersNoCommG", dbLong)
End With

db.TableDefs.Append td

Set rs = db.OpenRecordset(td.Name)

Set td = db.TableDefs("StoringenCount")
DoCmd.SetWarnings False
DoCmd.Hourglass True
Call closeAllObjects(True)
For Each fld In td.Fields
    If IsNumeric(Left(CStr(fld.Name), 1)) Then
        fld_col.Add CStr(fld.Name)
    End If
Next

For i = 1 To fld_col.Count
    rs.AddNew
    rs("daysofMonth").Value = CDate(fld_col(i))
    rs.Update
Next

rs.MoveFirst
    For i = 1 To fld_col.Count
        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
            fldDate = CDate(fld_col(i))
            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] WHERE [" & fld_col(i) & "] < #" & Format(CDate(fldDate) - 1, "yyyy-mm-dd") & "# AND [" & fld_col(i) & "] is not null ")
            rs.Edit
          rs("metersNoCommTot").Value = rsSource("countset").Value
            rs.Update
            rsSource.Close
        End If
        rs.MoveNext
    Next

rs.MoveFirst
    For i = 1 To fld_col.Count
        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
            fldDate = CDate(fld_col(i))
            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] WHERE [" & fld_col(i) & "] < #" & Format(CDate(fldDate) - 1, "yyyy-mm-dd") & "# AND [" & fld_col(i) & "] is not null AND type = ""E""")
            rs.Edit
            rs("metersNoCommE").Value = rsSource("countset").Value
            rs.Update
            rsSource.Close
        End If
        rs.MoveNext
    Next
    
rs.MoveFirst
    For i = 1 To fld_col.Count
        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
            fldDate = CDate(fld_col(i))
            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] WHERE [" & fld_col(i) & "] < #" & Format(CDate(fldDate) - 1, "yyyy-mm-dd") & "# AND [" & fld_col(i) & "] is not null AND type = ""G""")
            rs.Edit
            rs("metersNoCommG").Value = rsSource("countset").Value
            rs.Update
            rsSource.Close
        End If
        rs.MoveNext
    Next
    
Set td = Nothing
Set rs = Nothing

DoCmd.SetWarnings True
DoCmd.Hourglass False
Application.RefreshDatabaseWindow

MsgBox "Done"

End Sub

Sub countLastCommDay()

Dim fld_col As New Collection
Dim fld As Field
Dim fldName As String
Dim fldnamePrev As String
Dim fldNameLast As String
Dim updFld As String
Dim rsSource As DAO.Recordset

Set db = CurrentDb()
Set rs = db.OpenRecordset("doorlooptijden")

Set td = db.TableDefs("StoringenCount")
DoCmd.SetWarnings False
DoCmd.Hourglass True
Call closeAllObjects(True)

For Each fld In td.Fields
    If IsNumeric(Left(CStr(fld.Name), 1)) Then
        fld_col.Add CStr(fld.Name)
    End If
Next

Set rsSource = db.OpenRecordset("SELECT * FROM storingencount where commcount = -1")

rs.MoveFirst
Do While Not rs.EOF
    rs.Edit
    rs("LastCommDate").Value = Null
    rs.Update
    rsSource.MoveFirst
    Do While Not rsSource.EOF
        If rs("meternummer").Value = rsSource("meternummer").Value Then
            For i = 1 To fld_col.Count
                If IsNull(rsSource(fld_col(i)).Value) Then
                    If IsNull(rs("lastcommdate").Value) Then
                        rs.Edit
                        rs("LastCommDate").Value = CDate(fld_col(i))
                        rs.Update
                    End If
                End If
            Next
        End If
        rsSource.MoveNext
    Loop
    rs.MoveNext
Loop
 
rs.Close
Set rs = Nothing
rsSource.Close
Set rsSource = Nothing

Set td = Nothing
Set db = Nothing

DoCmd.SetWarnings True
DoCmd.Hourglass False

MsgBox "done"

End Sub
