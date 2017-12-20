Sub TEST()

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
Call deleteTable("Temp")

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
    Call deleteTable("temp")
    MFEDate = Format(DateSerial(Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 14, 4), Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 10, 2), Mid(strFileList(intFile), InStr(strFileList(intFile), ".csv") - 8, 2)), "yyyy-mm-dd")
    MFEtable = "MFE stats_" & Format(MFEDate, "dd-mm-yyyy")
    If fieldExists(tblName, CStr(Format(MFEDate, "dd-mm-yyyy"))) = False Then
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
        If fieldExists(tblName, dateField) = False Then td.Fields.Append td.CreateField(dateField, dbDate)

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


Sub countDays()

Dim fld_col As New Collection
Dim fld As Field
Dim fldname As String
Dim fldNamePrev As String
Dim fldNamePrevTwo As String
Dim fldNameLast As String
Dim updFld As String
Dim nDays As Integer

Call closeAllObjects(True)
Set db = CurrentDb()

nDays = 0   ' if 0, all date fields are used
Set td = db.TableDefs("storingencount")
    If fieldExists("storingencount", "commCountSeq") = False Then
        td.Fields.Append td.CreateField("commCountSeq", dbInteger)
    End If
    If fieldExists("storingencount", "commCount") = False Then
        td.Fields.Append td.CreateField("commCount", dbInteger)
    End If
    If fieldExists("storingencount", "countInstance") = False Then
        td.Fields.Append td.CreateField("countInstance", dbInteger)
    End If
    If fieldExists("storingencount", "maxDaysAvailable") = False Then
        td.Fields.Append td.CreateField("maxDaysAvailable", dbInteger)
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

If nDays = 0 Then
    nDays = fld_col.Count
    fldNameLast = CStr(fld_col(fld_col.Count))
Else
    nDays = nDays + 1
    fldNameLast = CStr(fld_col(nDays))
End If


playSound "chimes", &H1
'   count each day with communication
rs.MoveLast
N = 0
SysCmd acSysCmdInitMeter, "maxDaysAvailable", rs.RecordCount
rs.MoveFirst

Do While Not rs.EOF
    updFld = "maxDaysAvailable"
    rs.Edit
    rs(updFld).Value = 0
    rs.Update
    For i = 2 To nDays
        fldname = CStr(fld_col(i))
        If Not IsNull(rs(fldname).Value) Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value + 1
            rs.Update
        End If
    Next
    N = N + 1
    SysCmd acSysCmdUpdateMeter, N
    If N Mod 10000 = 0 Then
        SysCmd acSysCmdUpdateMeter, N
        DoEvents
    End If
    rs.MoveNext
Loop
    
playSound "chimes", &H1
'   count each day with communication
rs.MoveLast
N = 0
SysCmd acSysCmdInitMeter, "commcount", rs.RecordCount
rs.MoveFirst

Do While Not rs.EOF
    updFld = "commcount"
    rs.Edit
    rs(updFld).Value = 0
    rs.Update
    For i = 1 To nDays
        fldname = CStr(fld_col(i))
        If i = 1 Then
            If rs(fldname).Value > CDate(fldname) Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            End If
        Else
        fldNamePrev = CStr(fld_col(i - 1))
'        If IsNull(rs(fldNameLast).Value) Then
'               rs.Edit
'               rs(updFld).Value = -1
'               rs.Update
'               Exit For
            If rs(fldname).Value = rs(fldNamePrev).Value Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value
                rs.Update
            ElseIf rs(fldname).Value < rs(fldNamePrev).Value Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value
                rs.Update
            ElseIf rs(fldname).Value < CDate(fld_col(1)) Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value
                rs.Update
            ElseIf IsNull(rs(fldNamePrev).Value) And rs(fldname).Value > CDate(fldname) Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            ElseIf rs(fldname).Value > rs(fldNamePrev).Value Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            End If
        End If
    Next
    N = N + 1
    SysCmd acSysCmdUpdateMeter, N
    If N Mod 100 = 0 Then
        SysCmd acSysCmdUpdateMeter, N
        DoEvents
    End If
    rs.MoveNext
Loop

playSound "chimes", &H1

' only count until first new NO communication instance
rs.MoveLast
N = 0
SysCmd acSysCmdInitMeter, "commcountSeq", rs.RecordCount
rs.MoveFirst

Do While Not rs.EOF
    updFld = "commcountSeq"
    rs.Edit
    rs(updFld).Value = 0
    rs.Update
    For i = 1 To nDays
        fldname = CStr(fld_col(i))
        If i = 1 Then
            If rs(fldname).Value > CDate(fldname) Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            End If
        Else
        fldNamePrev = CStr(fld_col(i - 1))
'           If IsNull(rs(fldNameLast).Value) Then
'               rs.Edit
'               rs(updFld).Value = -1
'               rs.Update
'               Exit For
            If rs(fldname).Value > rs(fldNamePrev).Value Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            ElseIf IsNull(rs(fldNamePrev).Value) And rs(fldname).Value > CDate(fldname) Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            ElseIf rs(fldname).Value = rs(fldNamePrev).Value Then
                Exit For
            ElseIf rs(fldname).Value < CDate(fld_col(1)) Then
                Exit For
            ElseIf rs(fldname).Value < rs(fldNamePrev).Value Then
                Exit For
            End If
        End If
    Next
    N = N + 1
    SysCmd acSysCmdUpdateMeter, N
    If N Mod 100 = 0 Then
        SysCmd acSysCmdUpdateMeter, N
        DoEvents
    End If
    rs.MoveNext
Loop

'   Count of instances of communication regardless of number of sequential days

playSound "chimes", &H1

rs.MoveLast
N = 0
SysCmd acSysCmdInitMeter, "commcountInstance", rs.RecordCount
rs.MoveFirst

Do While Not rs.EOF
    updFld = "countInstance"
    rs.Edit
    rs(updFld).Value = 0
    rs.Update
    For i = 1 To nDays
        fldname = CStr(fld_col(i))
        If i > 2 Then
            fldNamePrevTwo = CStr(fld_col(i - 2))
        ElseIf i > 1 Then
            fldNamePrevTwo = CStr(Format(CDate(fld_col(i - 1)) - 1, "dd-mm-yyyy"))
        End If
        If i = 1 Then
            If rs(fldname).Value > CDate(fldname) Then
                rs.Edit
                rs(updFld).Value = rs(updFld).Value + 1
                rs.Update
            End If
        Else
        fldNamePrev = CStr(fld_col(i - 1))
'            If IsNull(rs(fldNameLast).Value) Then
'               rs.Edit
'               rs(updFld).Value = -1
'               rs.Update
'               Exit For
            If rs(fldname).Value = rs(fldNamePrev).Value Then
                If i - 2 < 1 Then
                    rs.Edit
                    rs(updFld).Value = rs(updFld).Value
                    rs.Update
                ElseIf rs(fldname).Value > rs(fldNamePrevTwo).Value And rs(fldNamePrevTwo).Value >= CDate(fld_col(1)) Then 'checks increase in comm date isn't in the past (2011 issue
                    rs.Edit
                    rs(updFld).Value = rs(updFld).Value + 1
                    rs.Update
                End If
            ' exceptions for 2011 issue meters
            ElseIf rs(fldname).Value < rs(fldNamePrev).Value Then
                If rs(fldNamePrev).Value < CDate(fld_col(1)) Then
                    rs.Edit
                    rs(updFld).Value = rs(updFld).Value
                    rs.Update
                ElseIf i - 2 < 1 Then
                    rs.Edit
                    rs(updFld).Value = rs(updFld).Value
                    rs.Update
                ElseIf rs(fldNamePrev).Value > rs(fldNamePrevTwo).Value Then
                    rs.Edit
                    rs(updFld).Value = rs(updFld).Value + 1
                    rs.Update
                End If
            End If
        End If
    Next
    N = N + 1
    If N Mod 100 = 0 Then
        SysCmd acSysCmdUpdateMeter, N
        DoEvents
    End If
    rs.MoveNext
Loop

SysCmd acSysCmdRemoveMeter

rs.Close
Set td = Nothing
Set rs = Nothing

DoCmd.SetWarnings True
DoCmd.Hourglass False

playSound "Alarm05", &H1

Call percentageCalc
Call averageCommCalc
MsgBox "Done"

End Sub

Sub CommunicatiePerDag()

Dim fld_col As New Collection
Dim fld As Field
Dim fldname As String
Dim fldNamePrev As String
Dim fldNameLast As String
Dim updFld As String
Dim rsSource As DAO.Recordset
Dim fldDate As Variant

Set db = CurrentDb()
Call closeAllObjects(False)
If TableExists("storingByDays") Then deleteTable ("storingByDays")
Set td = db.CreateTableDef("storingByDays")

With td
    .Fields.Append .CreateField("daysOfMonth", dbDate)
    .Fields.Append .CreateField("metersCommTot", dbLong)
    .Fields.Append .CreateField("metersCommE", dbLong)
    .Fields.Append .CreateField("metersCommG", dbLong)
    .Fields.Append .CreateField("metersPopTot", dbLong)
    .Fields.Append .CreateField("metersPopE", dbLong)
    .Fields.Append .CreateField("metersPopG", dbLong)
    .Fields.Append .CreateField("metersPercCommTot", dbDouble)
    .Fields.Append .CreateField("metersPercCommE", dbDouble)
    .Fields.Append .CreateField("metersPopCommG", dbDouble)
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

'   Aantal meters in totale populatie
rs.MoveFirst
    For i = 1 To fld_col.Count
        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
            fldDate = CDate(fld_col(i))
            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] WHERE [" & fld_col(i) & "] is not null ")
            rs.Edit
          rs("metersPopTot").Value = rsSource("countset").Value
            rs.Update
            rsSource.Close
        End If
        rs.MoveNext
    Next
    
''   Aantal E-meter in totale populatie
'rs.MoveFirst
'    For i = 1 To fld_col.Count
'        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
'            fldDate = CDate(fld_col(i))
'            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] WHERE [" & fld_col(i) & "] is not null AND supplytype like ""E*""")
'            rs.Edit
'            rs("metersPopE").Value = rsSource("countset").Value
'            rs.Update
'            rsSource.Close
'        End If
'        rs.MoveNext
'    Next
'
''   Aantal G-meter in totale populatie
'rs.MoveFirst
'    For i = 1 To fld_col.Count
'        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
'            fldDate = CDate(fld_col(i))
'            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] WHERE [" & fld_col(i) & "] is not null AND supplytype like ""G*""")
'            rs.Edit
'            rs("metersPopG").Value = rsSource("countset").Value
'            rs.Update
'            rsSource.Close
'        End If
'        rs.MoveNext
'    Next
    
'   Aantal dat communiceert van totaal beschikbare populatie
rs.MoveFirst
    For i = 1 To fld_col.Count
        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
            fldDate = CDate(fld_col(i))
            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] " & _
                                            "WHERE [" & fld_col(i) & "] >= #" & Format(CDate(fldDate), "yyyy-mm-dd") & "# " & _
                                            "AND [" & fld_col(i) & "] is not null ")
            rs.Edit
          rs("metersCommTot").Value = rsSource("countset").Value
            rs.Update
            rsSource.Close
        End If
        rs.MoveNext
    Next

''   Aantal dat communiceert van totaal beschikbare E-meter populatie
'rs.MoveFirst
'    For i = 1 To fld_col.Count
'        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
'            fldDate = CDate(fld_col(i))
'            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] WHERE [" & fld_col(i) & "] >= #" & Format(CDate(fldDate) - 1, "yyyy-mm-dd") & "# AND [" & fld_col(i) & "] is not null AND supplytype like ""E*""")
'            rs.Edit
'            rs("metersCommE").Value = rsSource("countset").Value
'            rs.Update
'            rsSource.Close
'        End If
'        rs.MoveNext
'    Next
'
''   Aantal dat communiceert van totaal beschikbare G-meter populatie
'rs.MoveFirst
'    For i = 1 To fld_col.Count
'        If rs("daysofMonth").Value = CDate(fld_col(i)) Then
'            fldDate = CDate(fld_col(i))
'            Set rsSource = db.OpenRecordset("SELECT count([" & fld_col(i) & "]) as CountSet FROM [" & td.Name & "] WHERE [" & fld_col(i) & "] >= #" & Format(CDate(fldDate) - 1, "yyyy-mm-dd") & "# AND [" & fld_col(i) & "] is not null AND supplyType like ""G*""")
'            rs.Edit
'            rs("metersCommG").Value = rsSource("countset").Value
'            rs.Update
'            rsSource.Close
'        End If
'        rs.MoveNext
'    Next
'
'   Percentage Communicatie
Set td = db.TableDefs("storingByDays")

' percentage totaal
SQL = "UPDATE [" & td.Name & "] SET metersPercCommTot=(metersCommTot/metersPopTot)*100"
    db.Execute SQL
 
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
Dim fldname As String
Dim fldNamePrev As String
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


Sub percentageCalc()


Dim tName As String
Dim fName As String
Dim N As Integer
tName = "nTalByPercentage"
Call closeAllObjects(False)
Set db = CurrentDb()


fName = DMax("commcount", "storingencount")

Set td = db.TableDefs("storingencount")
If fieldExists("storingencount", "percComm") = False Then
        td.Fields.Append td.CreateField("percComm", dbDouble)
End If

SQL = "UPDATE storingencount SET PercComm=commCount/maxDaysAvailable"
    db.Execute SQL
If TableExists(tName) = False Then
    Set td = db.CreateTableDef(tName)
Else
    Set td = db.TableDefs(tName)
End If

If fieldExists(tName, "percentageCommunicatie") = False Then
        td.Fields.Append td.CreateField("percentageCommunicatie", dbInteger)
End If

If TableExists(tName) = False Then
    db.TableDefs.Append td
End If

Set rs = db.OpenRecordset("SELECT * FROM [" & td.Name & "] ORDER BY percentageCommunicatie asc")

If Not rs.EOF Then
    rs.MoveLast
    N = rs("percentageCommunicatie").Value
Else
    N = -1
End If

While N < 100
    rs.AddNew
    rs("percentageCommunicatie").Value = N
    N = N + 1
    rs.Update
Wend
        
rs.Close
fName = "nTal" & fName
If fieldExists(tName, fName) = False Then
    td.Fields.Append td.CreateField(fName, dbDouble)
End If

SQL = " SELECT round((t1.[PercComm])*100) AS percentageCommunicatie, count(round(([PercComm])*100)) AS nTal INTO calcTemp "
SQL = SQL & " FROM StoringenCount as t1 GROUP BY Round(([PercComm])*100)"
    db.Execute SQL
    
SQL = " UPDATE [" & tName & "]as t1 INNER JOIN calcTemp as t2 "
SQL = SQL & " ON t1.percentageCommunicatie=t2.percentageCommunicatie SET t1.[" & fName & "]=t2.nTal"
    db.Execute SQL

DoCmd.DeleteObject acTable, "calcTemp"

playSound "chimes", &H0

End Sub

Sub averageCommCalc()

Dim tName As String
Dim fName As String
Dim nDays As Integer
Dim N As Integer
tName = "averageCommDays"
Call closeAllObjects(False)
Set db = CurrentDb()

nDays = DMax("commcount", "storingencount")
If TableExists(tName) = False Then
    Set td = db.CreateTableDef(tName)
Else
    Set td = db.TableDefs(tName)
End If

Dim strArray() As Variant
Dim strItem As Variant

strArray = Array("avgDaysBetweenInterval", "avgCommSeq", "minIntervals", "maxIntervals", "avgIntervals", "stDevIntervals", "avgInteruption")

If fieldExists(tName, "nTal") = False Then
    td.Fields.Append td.CreateField("nTal", dbText)
End If
For Each strItem In strArray
    fName = strItem
    If fieldExists(tName, fName) = False Then
        td.Fields.Append td.CreateField(fName, dbDouble)
    End If
Next

fName = "nTal" & nDays
If TableExists(tName) = False Then
    db.TableDefs.Append td
End If

Set rs = td.OpenRecordset
For i = 1 To rs.RecordCount + 1
    Dim tmp As DAO.Recordset
    If rs.EOF Then
        rs.AddNew
        Exit For
    ElseIf rs("ntal").Value = fName Then
        rs.Edit
        Exit For
    End If
    rs.MoveNext
Next

Set tmp = db.OpenRecordset("SELECT (sum(commCount)/sum(countInstance))/" & nDays & " as tmpFld, (sum(commCount)/sum(countInstance)) as nDaysFld, " & _
" min(countInstance) as minFld, max(countInstance) as maxFld, avg(countInstance) as avgFld, stDev(countInstance) as stDevFld, " & _
" sum((" & nDays & "-commCount))/sum(countInstance) as disFld " & _
" FROM StoringenCount as t1 WHERE countInstance<>-1")
rs("nTal").Value = fName
rs("avgCommSeq").Value = tmp("tmpFld").Value
rs("avgdaysBetweenInterval").Value = tmp("nDaysFld").Value
rs("minIntervals").Value = tmp("minFld").Value
rs("maxIntervals").Value = tmp("maxFld").Value
rs("avgIntervals").Value = tmp("avgFld").Value
rs("stDevIntervals").Value = tmp("stDevFld").Value
rs("avgInteruption").Value = tmp("disFld").Value
tmp.Close
rs.Update
rs.Close

playSound "chimes", &H0
End Sub

Sub countSequences()

Dim fld_col As New Collection
Dim fld As Field
Dim fldname As String
Dim fldNamePrev As String
Dim fldNamePrevTwo As String
Dim fldNameLast As String
Dim updFld As String
Dim nDays As Integer

Call closeAllObjects(True)
Set db = CurrentDb()

' create new table countSequences
If TableExists("storingencount") = False Then
    cancelMsg = "Geen brontabel gevonden"
    Call cancelAction
End If

If TableExists("countSequences") Then
    msgInput = MsgBox("Tabel bestaat al, wissen?", vbYesNo)
    Select Case msgInput
        Case vbYes
            Dim fldArray() As Variant
            Dim fldInt As Integer
            DoCmd.DeleteObject acTable, "countsequences"
            DoCmd.CopyObject , "countSequences", acTable, "storingenCount"
            Set td = db.TableDefs("countsequences")
            For Each fld In td.Fields
                If IsNumeric(Left(fld.Name, 4)) Then
                    fldInt = fldInt + 1
                    ReDim Preserve fldArray(1 To fldInt)
                    fldArray(fldInt) = fld.Name
                End If
            Next
            For i = 1 To UBound(fldArray)
                td.Fields.Delete (fldArray(i))
            Next
            Erase fldArray
        Case vbNo
        Case Else
            Call cancelAction
        End Select
Else
    DoCmd.CopyObject , "countSequences", acTable, "storingenCount"
    Set td = db.TableDefs("countsequences")
    For Each fld In td.Fields
        If IsNumeric(Left(fld.Name, 4)) Then
            td.Fields.Delete (fld.Name)
        End If
    Next
End If

nDays = 0   ' if 0, all date fields are used
Set td = db.TableDefs("storingencount")
Set td2 = db.TableDefs("countSequences")



'   add missing datefields to countSequences
For Each fld In td.Fields
    If IsNumeric(Left(CStr(fld.Name), 1)) Then
        If fieldExists(td2.Name, fld.Name) = False Then
            td2.Fields.Append td2.CreateField(fld.Name, dbInteger)
        End If
    End If
Next

DoCmd.SetWarnings False
DoCmd.Hourglass True
Call closeAllObjects(True)
For Each fld In td.Fields
    If IsNumeric(Left(CStr(fld.Name), 1)) Then
        fld_col.Add CStr(fld.Name)
    End If
Next

'   Get last field in collection (last day)

If nDays = 0 Then
    nDays = fld_col.Count
    fldNameLast = CStr(fld_col(fld_col.Count))
Else
    nDays = nDays + 1
    fldNameLast = CStr(fld_col(nDays))
End If

Set rs = db.OpenRecordset("SELECT storingencount.* FROM storingencount INNER JOIN countSequences as t2 ON storingencount.equipmentId=t2.equipmentId order by t2.equipmentId asc")
Set rs2 = db.OpenRecordset("countsequences")

playSound "chimes", &H1

'   count each day without communication
rs.MoveLast
N = 0
SysCmd acSysCmdInitMeter, "commcountSequencing", rs.RecordCount
rs.MoveFirst

Do While Not rs.EOF
    If rs("equipmentId").Value > rs2("equipmentId").Value Then
        rs.MovePrevious
    ElseIf rs("equipmentId").Value = rs2("equipmentId").Value Then
        For i = 1 To nDays
            fldname = CStr(fld_col(i))
            updFld = fldname
            rs2.Edit
            rs2(updFld).Value = 0
            rs2.Update
            If i >= 2 Then
                fldNamePrev = CStr(fld_col(i - 1))
            Else
                fldNamePrev = fldname
            End If
            updFldprev = fldNamePrev
            If IsNull(rs(fldNameLast).Value) Then
                rs2.Edit
                rs2(updFld).Value = rs2(updFldprev).Value + 1
                rs2.Update
                Exit For
            ElseIf rs(fldname).Value = rs(fldNamePrev).Value Then
                rs2.Edit
                rs2(updFld).Value = rs2(updFldprev).Value + 1
                rs2.Update
            ElseIf rs(fldname).Value < rs(fldNamePrev).Value Then
                rs2.Edit
                rs2(updFld).Value = rs2(updFldprev).Value + 1
                rs2.Update
            ElseIf rs(fldname).Value < CDate(fld_col(1)) Then
                rs2.Edit
                rs2(updFld).Value = rs2(updFldprev).Value + 1
                rs2.Update
            ElseIf rs(fldname).Value > rs(fldNamePrev).Value Then
                rs2.Edit
                rs2(updFld).Value = 0
                rs2.Update
            End If
        Next
        N = N + 1
        SysCmd acSysCmdUpdateMeter, N
        If N Mod 100 = 0 Then
            SysCmd acSysCmdUpdateMeter, N
            DoEvents
        End If
        rs2.MoveNext
        rs.MoveNext
    Else
        rs.MoveNext
    End If
Loop


playSound "chimes", &H1

SysCmd acSysCmdRemoveMeter

rs.Close
rs2.Close
Set td = Nothing
Set rs = Nothing

DoCmd.SetWarnings True
DoCmd.Hourglass False

playSound "Alarm05", &H1

MsgBox "Done"


End Sub

Sub daySplitcount()

Dim fld_col As New Collection
Dim fld As Field
Dim fldname As String
Dim fldNamePrev As String
Dim updFld As String
Dim nDays As Integer


Call closeAllObjects

DoCmd.SetWarnings False
DoCmd.Hourglass True
Set db = CurrentDb

Set td = db.TableDefs("countSequences")

For Each fld In td.Fields
    If IsNumeric(Left(CStr(fld.Name), 1)) Then
        fld_col.Add CStr(fld.Name)
    End If
Next

'    set the days counter to measure number of intervals that fit criteria
nDays = 2
daysMeasure = "daysIs" & nDays
If fieldExists(td.Name, daysMeasure) = False Then
        td.Fields.Append td.CreateField(daysMeasure, dbInteger)
End If
'   Get last field in collection (last day)

Set rs = db.OpenRecordset("countSequences")

playSound "chimes", &H1
'   count each day without communication
rs.MoveLast
N = 0


SysCmd acSysCmdInitMeter, daysMeasure, rs.RecordCount
rs.MoveFirst

Do While Not rs.EOF
    updFld = daysMeasure
    rs.Edit
    rs(updFld).Value = 0
    rs.Update
    For i = 1 To fld_col.Count
        fldname = CStr(fld_col(i))
        If rs(fldname).Value = nDays Then
            rs.Edit
            rs(updFld).Value = rs(updFld).Value + 1
            rs.Update
        End If
    Next
    N = N + 1
    SysCmd acSysCmdUpdateMeter, N
    If N Mod 100 = 0 Then
        SysCmd acSysCmdUpdateMeter, N
        DoEvents
    End If
    rs.MoveNext
Loop

Call endRoutine
End Sub
