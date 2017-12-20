    Option Compare Database

Sub deploymentImport()

If gEnableErrorHandling Then On Error GoTo errorHandler

DoCmd.SetWarnings False
If CurrentProject.AllForms(frmDep).IsLoaded = False Then
    DoCmd.OpenForm frmDep, acNormal, , , , acHidden
Else
    Forms.Item(frmDep).Visible = False
End If

Set db = CurrentDb()
Call closeAllObjects(frmDep)

'   ---------------------------------------
'   Importeren van MFE, controle op reeds aanwezige MFE Dump en Datum van Dump.
Set rs = db.OpenRecordset("MFE Stats")
If rs.EOF Then
    MsgBox "Selecteer de MFE Dump"
    rs.Close
    Call selectFileMFE(frmDep, "Splunk_meterData Import")
Else
    rs.Close
    response = MsgBox("MFE Dump bijwerken? Huidige datum is: " & Application.DMax("[MFE datum]", "[MFE stats]") & ".", vbYesNoCancel + vbQuestion)
    Select Case response
        Case vbYes
            Call selectFileMFE(frmDep, "Splunk_meterData Import")
        Case vbCancel
            cancelMsg = "Handeling afgebroken door gebruiker"
            Call cancelAction(frmDep)
        Case vbIgnore
            cancelMsg = "Handeling afgebroken door gebruiker"
            Call cancelAction(frmDep)
    End Select
End If

Call deleteImportErrorTables

'If CurrentProject.AllForms(frmDep).IsLoaded = True Then
'    Call weekSelect(frmDep, Forms.Item(frmDep).weeknrCap)
'End If

'   ---------------------------------------
'   Importscript voor het importeren van het Topdesk bestand
MsgBox "Selecteer het te importeren bestand Deployment TD incidenten"
Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .AllowMultiSelect = False
    .Title = "Selecteer het te importeren bestand TD incidenten"
    .Filters.Clear
    .Filters.Add "Excel Bestanden", "*.xlsx, *.xls"
    .InitialFileName = "S:\Stedin\Diensten\SMB\Operations KVB\Datacollectie\Uitvalbakken\07.Deployment TD Proces\"
End With

If fd.Show = True Then
    If fd.SelectedItems(1) <> vbNullString Then
        fileName = fd.SelectedItems(1)
    End If
Else
    
    '   ---------------------------------------
    '   Exit code if no file is selected
        cancelMsg = "Er is geen Deployment TD bestand geselecteerd, handeling afgebroken"
    Call cancelAction(frmDep)
    Exit Sub
End If

'   ---------------------------------------
'   Cleanup old data
SQL = "Delete * FROM Dep_TD_All"
    db.Execute SQL
DoCmd.Hourglass False

'   ---------------------------------------
'   Actual import of new TD data
Call deleteTable("temp")
Call deleteTable("eerstelijns incident")
If Right(fileName, 4) = "xlsx" Then
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "eerstelijns incident", fileName, True
ElseIf Right(fileName, 4) <> "xls" Then
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel9, "eerstelijns incident", fileName, True
End If

SQL = " Select * INTO temp FROM [eerstelijns incident]"
    db.Execute SQL
    
'shorten the fieldnames
Set td = db.TableDefs("temp")
For Each fld In td.Fields
    If fld.Name Like "EAN*" Then
        fld.Name = Left(fld.Name, 10)
    ElseIf fld.Name Like "Meternummer*" Then
        fld.Name = Left(fld.Name, 13)
    ElseIf fld.Name Like "plaatsingsDatum*" Then
        fld.Name = Left(fld.Name, 17)
    ElseIf fld.Name Like "plaatsingsDatum*" Then
        fld.Name = Left(fld.Name, 17)
    ElseIf fld.Name Like "Status Storing MFE*" Then
        fld.Name = "Splunk Status (Splunk Informatie PRODUCTIE)"
    End If
Next

SQL = "INSERT INTO DEP_TD_All SELECT * FROM temp"
    db.Execute SQL
    
Call deleteTable("temp")

'   ---------------------------------------
'   Processing the data, checking what currently communicates and what doesn't

        'check to add fields if they do not exist
Set td = db.TableDefs("dep_TD_all")

For Each fld In td.Fields
    If fld.Name Like "EAN*" Then
        fld.Name = Left(fld.Name, 10)
    ElseIf fld.Name Like "Meternummer*" Then
        fld.Name = Left(fld.Name, 13)
    ElseIf fld.Name Like "plaatsingsDatum*" Then
        fld.Name = Left(fld.Name, 17)
    End If
Next

If fieldExists("Dep_TD_all", "opmerking") = False Then
    td.Fields.Append td.CreateField("opmerking", dbText, 255)
End If
If fieldExists("Dep_TD_all", "actie") = False Then
    td.Fields.Append td.CreateField("actie", dbText, 255)
End If

        'Update current status based on MFE stats
SQL = "UPDATE Dep_TD_all as t1 INNER JOIN [MFE STATS] as t2 ON t1.[Meternummer E] = t2.[e-meternummer] "
SQL = SQL & "SET t1.[EAN code E] = t2.[EAN code E] "
    db.Execute SQL
    
SQL = "UPDATE Dep_TD_all as t1 INNER JOIN [MFE STATS] as t2 ON t1.[Meternummer G] = t2.[g-meternummer] "
SQL = SQL & "SET t1.[EAN code G] = t2.[EAN code G], t1.[ean code E] = iif(t1.[ean code e] not like ""87*"", t2.[ean code E], t1.[ean code E])"
    db.Execute SQL
    
SQL = "UPDATE DEP_TD_ALL SET opmerking = iif(subcategorie like ""nacht*"", ""Verkeerdom aangesloten"", "
SQL = SQL & " iif(subcategorie like ""Gas meter sleeping"", ""Gasmeter slaapt"", ""Slimme Data""))"
    db.Execute SQL
    
SQL = "UPDATE Dep_TD_all as t1 INNER JOIN [MFE STATS] as t2 ON t1.[Meternummer G] = t2.[g-meternummer] "
SQL = SQL & " SET t1.[EAN code G] = t2.[EAN code G], t1.[ean code E] = iif(t1.[ean code e] not like ""87*"", t2.[ean code E], t1.[ean code E]), "
SQL = SQL & " t1.[meternummer e] = iif(t1.[meternummer e] like ""Unknown"", t2.[e-meternummer], t1.[meternummer E]) "
    db.Execute SQL
    
SQL = "UPDATE DEP_TD_ALL SET opmerking = iif(subcategorie like ""nacht*"", ""Verkeerdom aangesloten"", "
SQL = SQL & " iif(subcategorie like ""Gas meter sleeping"", ""Gasmeter slaapt"", ""Slimme Data""))"
    db.Execute SQL
    
SQL = "UPDATE DEP_TD_all as t1 LEFT JOIN [MFE STATS] as t2 On t1.[ean code e] = t2.[ean code E] and t1.[meternummer e] = t2.[e-meternummer] "
SQL = SQL & "SET opmerking = ""afgenomen"" WHERE t2.[ean code E] is null"
    db.Execute SQL
  
SQL = "UPDATE DEP_TD_ALL SET Actie = iif(opmerking like ""Afgenomen"",""Incident Sluiten"","
SQL = SQL & " iif(subcategorie like ""DNV bij de intervalstanden."", ""Start Controleroutine"","
SQL = SQL & " iif(subcategorie like ""Start Controleroutine"" or subcategorie like ""Te weinig goede sessies*"", ""Slechte Communicatie"","
SQL = SQL & " iif(subcategorie like ""Gas meter uncoupled"", ""Ontkoppelde Gasmeter"","
SQL = SQL & " iif(opmerking not like ""slimme Data"", ""Uitzetten"", ""Onderzoek oorzaak"")))))"
    db.Execute SQL


'   ---------------------------------------
'   End sequence for code execution to reset environments and variables
Err.Clear
Call endRoutine(frmDep)
Exit Sub

errorHandler:
Call endRoutine(frmDep)
End Sub

Sub deploymentSharepointUpdate()

If gEnableErrorHandling Then On Error GoTo errorHandler

Dim idMax As Long
Dim idNew As Long
Dim newFlowE As Integer
Dim newFlowG As Integer
Dim newFlowT As Integer
Dim gcom As Integer

Set db = CurrentDb

'   ---------------------------------------
'   step 1: Sharepoint vullen met TopDesk Data verrijkt met MFE adres gegevens
SysCmd acSysCmdSetStatus, "Sharepoint vullen met TopDesk Data verrijkt met MFE adres gegevens"
MsgBox "bijwerken sharepoint met nieuwe incidenten", vbInformation + vbOKOnly
DoEvents

Set rs = db.OpenRecordset("work-around deployment")
rs.MoveLast
idMax = rs!Id

'   Slapende Gasmeters
SQL = "INSERT INTO [Work-around deployment] ("
SQL = SQL & " Incidentnummer, Status, [G-meter EAN], InstallatieDatum, Postcode, Plaatsnaam, straat_Naam, Huisnummer, Huisnr_Toevoeging,[G-meter nummer], meterType) "
SQL = SQL & " SELECT t1.Incidentnummer, t1.opmerking, t1.[EAN code G], t1.[Plaatsingsdatum G],"
SQL = SQL & " t2.[MFE postcode], t2.[MFE Plaatsnaam], [straat naam], getHuisnummer(t2.[MFE huisnummer]), "
SQL = SQL & " getToevoeging(t2.[MFE huisnummer]), [meternummer G], t2.[G-Meter type] "
SQL = SQL & " FROM (DEP_TD_all as t1 INNER JOIN [MFE STATS] as t2 ON t1.[Meternummer G] = t2.[G-meternummer]) "
SQL = SQL & " LEFT JOIN [Work-around deployment] ON t1.Incidentnummer = [Work-around deployment].Incidentnummer "
SQL = SQL & " WHERE [Work-around deployment].Incidentnummer Is Null AND t1.opmerking Like ""Gasmeter slaapt"""
    db.Execute SQL

rs.Requery
rs.MoveLast
idNew = rs!Id
newFlowG = idNew - idMax

'   Verkeerkom aangesloten
SQL = "INSERT INTO [Work-around deployment] ("
SQL = SQL & " Incidentnummer, Status, [E-meter EAN], InstallatieDatum, Postcode, Plaatsnaam, straat_Naam, Huisnummer, Huisnr_Toevoeging,[E-meter nummer], meterType) "
SQL = SQL & " SELECT t1.Incidentnummer, t1.opmerking, t1.[EAN code E], t1.[Plaatsingsdatum E],"
SQL = SQL & " t2.[MFE postcode], t2.[MFE Plaatsnaam], [straat naam], getHuisnummer(t2.[MFE huisnummer]), "
SQL = SQL & " getToevoeging(t2.[MFE huisnummer]), [meternummer E], t2.[E-Meter type]"
SQL = SQL & " FROM (DEP_TD_all as t1 INNER JOIN [MFE STATS] as t2 ON t1.[Meternummer E] = t2.[E-meternummer]) "
SQL = SQL & " LEFT JOIN [Work-around deployment] ON t1.Incidentnummer = [Work-around deployment].Incidentnummer "
SQL = SQL & " WHERE [Work-around deployment].Incidentnummer Is Null AND t1.opmerking Like ""Verkeerdom*"""
    db.Execute SQL
    
rs.Requery
rs.MoveLast
idMax = rs!Id
newFlowE = idMax - idNew
newFlowT = newFlowG + newFlowE

'   ---------------------------------------
'   step 2: Sharepoint vullen met locatienummers
SysCmd acSysCmdSetStatus, "Sharepoint vullen met locatienummers"
DoEvents

'   Slapende gasmeters
SQL = "UPDATE ([work-around deployment] as t1 INNER JOIN [MFE STATS] as t2 ON t1.[G-meter EAN] = t2.[EAN CODE G]) "
SQL = SQL & " INNER JOIN refLocatienummers as t3 ON t2.[EAN CODE E] = t3.EANcode "
SQL = SQL & " SET t1.locatienummer = t3.locatienummer "
SQL = SQL & " WHERE t1.locatienummer is null AND t1.[Terugkoppeling Slimme Data] is null "
    db.Execute SQL

'   Verkeerdom aangesloten
SQL = "UPDATE [work-around deployment] as t1 INNER JOIN refLocatienummers as t2 ON t1.[E-meter EAN] = t2.[EANCode] "
SQL = SQL & " SET t1.locatienummer = t2.locatienummer "
SQL = SQL & " WHERE t1.locatienummer is null AND t1.[Terugkoppeling Slimme Data] is null"
    db.Execute SQL
    
'   ontbrekende locatienummers toevoegen to sharepoint
'   ---------------------------------------

SQL = "SELECT locatienummer, iif([E-meter EAN] is null, [G-meter EAN], [E-meter EaN]) as [Ean Code] FROM [Work-around deployment] WHERE locatieNummer is null "
Qstr = "temp_loc"
Call createQuery(Qstr)
Set qt = db.QueryDefs(Qstr)
qt.SQL = SQL

Set rs = db.OpenRecordset(Qstr)
If Not rs.EOF Then
    MsgBox "Niet alle locatienummers gevonden, handmatig toevoegen", vbInformation
    Set stoCollection = New Collection
    
    rs.MoveFirst
    
    Do While Not rs.EOF
        stoCollection.Add rs("EAN code").Value
        rs.MoveNext
    Loop
    rs.Close
    
    If CurrentProject.AllForms.Item("Sto_Loc_frm").IsLoaded Then DoCmd.Close acForm, "Sto_Loc_frm", acSaveNo
    DoCmd.Hourglass False
    DoCmd.OpenForm "Sto_Loc_frm", acNormal, , , acFormEdit, acDialog, Qstr
    DoCmd.Hourglass True
    
'   Locatienummer bijwerken indien nieuw toegevoegd
'   ---------------------------------------
    Dim rsSource As DAO.Recordset
    Dim LOCNR As String
    Dim EANE As String
    
    Set rsSource = db.OpenRecordset("SELECT DISTINCT [locatieNummer], iif([E-meter EAN] is null,[G-meter EAN],[E-meter EAN]) as [EAN Code] From [Work-around deployment] WHERE locatieNummer is not null")
    Set rs = db.OpenRecordset("refLocatienummers")
    
    On Error Resume Next
    For i = 1 To stoCollection.Count
        rsSource.MoveFirst
        Do While Not rsSource.EOF
            If rsSource("EAN code").Value Like "87*" Then
                If rsSource("EAN code").Value = stoCollection(i) Then
                    If Not IsNull(rsSource("LocatieNummer").Value) Then
                        LOCNR = rsSource("locatieNummer").Value
                        EANE = rsSource("EAN code").Value
                        rs.AddNew
                        rs("EANCode").Value = EANE
                        rs("Locatienummer").Value = LOCNR
                        rs.Update
                    End If
                End If
            End If
            rsSource.MoveNext
        Loop
    Next
    rs.Close
    rsSource.Close
    
    Set rs = Nothing
    Set rsSource = Nothing
    
End If
Call deleteQuery(Qstr)
    
'   ---------------------------------------
'   step 3: Sharepoint vullen met regioindeling
SysCmd acSysCmdSetStatus, "Sharepoint vullen met regioindeling"
SQL = "UPDATE [work-around deployment] as t1 INNER JOIN refRegioindeling as t2 ON left(t1.[postcode],4) = cstr(nz(t2.postcode4)) "
SQL = SQL & " SET t1.regio=t2.[regio indeling] WHERE t1.regio is null "
    db.Execute SQL
         
'   ---------------------------------------
'   step 5: bestaande gegevens bijwerken (controle op gas-meters)
SysCmd acSysCmdSetStatus, "bestaande sharepoint gegevens bijwerken"
DoEvents
SQL = "UPDATE [work-around deployment] as t1 INNER JOIN [MFE STATS] as t2 On t1.[G-meter EAN]=t2.[ean code G] "
SQL = SQL & " SET t1.[Terugkoppeling Slimme Data]=""Communiceert " & Format(Date, "dd-mm-yyyy") & """ "
SQL = SQL & " WHERE t2.[laatste comm G]>t2.[mfe datum]-2 AND t1.[terugkoppeling Slimme Data] is null"
    db.Execute SQL
gcom = db.RecordsAffected

SQL = "UPDATE [work-around deployment] as t1 INNER JOIN [MFE STATS] as t2 On t1.[E-meter EAN]=t2.[ean code E] "
SQL = SQL & " SET t1.[Terugkoppeling Slimme Data]=""Communiceert " & Format(Date, "dd-mm-yyyy") & """ "
SQL = SQL & " WHERE t2.[deploymentStateE]<>""Deployment"" AND t1.[terugkoppeling Slimme Data] is null AND t1.status like ""verkeerd*"" "
    db.Execute SQL
SysCmd acSysCmdClearStatus
ecom = db.RecordsAffected

strMsg = "Nieuw toegevoegd op " & Format(Date, "dd-mm-yyyy") & vbCrLf _
        & "G-meters: " & newFlowG & vbCrLf _
        & "E-meters:  " & newFlowE & vbCrLf _
        & "Totaal:    " & newFlowT & vbCrLf _
        & "Aantal bestaande g-meters status bijgewerkt: " & gcom & vbCrLf _
        & "Aantal bestaande e-meters status bijgewerkt: " & ecom & vbCrLf _

MsgBox strMsg, vbOKOnly + vbInformation, "Overzicht"

strMsg = Replace(strMsg, "Nieuw toegevoegd", "----------------------------------------" & vbCrLf & "Nieuw toegevoegd")
fileName = Environ("temp") & "\Deployment_Cijfers.txt"

txtFileNum = FreeFile
Open fileName For Append As txtFileNum
Print #txtFileNum, strMsg
Close #txtFileNum

Shell "notepad " & fileName, 3

'   Add report data to stats table

SQL = " INSERT INTO STATS(datum,ntalE,nTalG,nTalTot,nTalCommuniceertE,ntalcommuniceertG) "
SQL = SQL & " VALUES(#" & Format(Date, "yyyy-mm-dd") & "#," & newFlowE & "," & newFlowG & "," & newFlowT & "," & ecom & "," & gcom & ")"
    db.Execute SQL

'   ---------------------------------------
'   automatic email generation to be sent (added in v1.2 [new 22-09-2016])
Dim Outlook As Object
Dim Outmail As Object
Dim nTalE As Long
Dim nTalG As Long
Dim nTalCommE As Long
Dim nTalCommG As Long

Set Outlook = CreateObject("outlook.application")
Set Outmail = Outlook.createitem(0)
Set rs = db.OpenRecordset("SELECT * FROM Stats order by datum desc")

nTalE = 0
nTalG = 0
nTalCommE = 0
nTalCommG = 0
rs.MoveFirst
Do While Not rs.EOF
    If rs("datum").Value = CDate(Date) Then
        nTalE = nTalE + rs("nTalE").Value
        nTalG = nTalG + rs("nTalG").Value
        nTalCommE = nTalCommE + rs("nTalCommuniceertE").Value
        nTalCommG = nTalCommG + rs("nTalCommuniceertG").Value
    ElseIf rs("datum").Value < CDate(Date) Then
        Exit Do
    End If
    rs.MoveNext
Loop

rs.Close
strBody = "<HTML><body><font size=3 face=calibri> " _
          & "<p>All,</p>" _
          & "<p>De GSA SharePoint Deployment is bijgewerkt.</p>" _
          & "<p><b>Toegevoegd:</b><br>" _
          & nTalE & " verkeerdom aangesloten(E - meters)<br>" _
          & nTalG & " slapende Gasmeters(G - meters)</p>" _
          & "<p><b>Verkeerdom aangesloten:</b><br>" _
          & nTalCommE & " aansluitingen waar de meter verkeerdom was aangesloten hebben succesvol de deploymentcontrole doorlopen</p>" _
          & "<p><b>Slapende GasMeter</b><br>" _
          & nTalCommG & " aansluitingen met slapende gasmeters zijn gaan communiceren.</p>" _
          & "<p>Met vriendelijke groet,<br>" _
          & "Team Slimme Data</p>" _
          & "</font></body></HTML>"
        
With Outmail
    .to = "ronald.krutzen@stedin.net; raymond.vandenberg@stedin.net; Ramesh.binda@stedin.net; Irving.Aarnoutse@stedin.net"
    .CC = "lidewij.nicolai@stedin.net"
    .Subject = "Sharepoint bijgewerkt"
    .BodyFormat = olFormatHTML
    .HTMLBody = strBody
    .Display
End With
    
'   ---------------------------------------
'   End sequence for code execution to reset environments and variables
Err.Clear
Call endRoutine(frmDep)
Exit Sub

errorHandler:
Call endRoutine(frmDep)

End Sub

Sub depTdImport()

'   ---------------------------------------
If gEnableErrorHandling Then On Error GoTo errorHandler

DoCmd.SetWarnings False
DoCmd.Hourglass True
If CurrentProject.AllForms(frmDep).IsLoaded = False Then
    DoCmd.OpenForm frmDep, acNormal, , , , acHidden
Else
    Forms.Item(frmDep).Visible = False
End If

Set db = CurrentDb()
Call closeAllObjects(frmDep)

'   ---------------------------------------
'   Importeren van MFE, controle op reeds aanwezige MFE Dump en Datum van Dump.
Set rs = db.OpenRecordset("MFE Stats")
If rs.EOF Then
    MsgBox "Selecteer de MFE Dump"
    rs.Close
    Call selectFileMFE(frmDep, "Splunk_meterData Import")
Else
    rs.Close
    response = MsgBox("MFE Dump bijwerken? Huidige datum is: " & Application.DMax("[MFE datum]", "[MFE stats]") & ".", vbYesNoCancel + vbQuestion)
    Select Case response
        Case vbYes
            Call selectFileMFE(frmDep, "Splunk_meterData Import")
        Case vbCancel
            cancelMsg = "Handeling afgebroken door gebruiker"
            Call cancelAction(frmDep)
    End Select
End If

Call deleteImportErrorTables

'   import actual meters in deployment
'   ---------------------------------------

Set rs = db.OpenRecordset("SELECT max(lastCommunicationTs) as fDate FROM metersInDeployment WHERE lastCommunicationTs is not null and lastCommunicationTs < date()")
If Not IsNull(rs!fdate) Then
    tDate = rs!fdate
End If
rs.Close
Set td = db.TableDefs("metersinDeployment")

i = 0
If td.RecordCount = 0 Then
    MsgBox "Selecteer bestand met meters in Deployment"
    Call selectFile(frmDep, "S:\Stedin\Diensten\SMB\Operations KVB\Datacollectie\Uitvalbakken\07.Deployment TD Proces\02. SplunkInDeployment\")
    i = 1
Else
    response = MsgBox("Bestand met meters in Deployment bijwerken? Huidige datum is: " & tDate & ".", vbYesNoCancel + vbQuestion)
    Select Case response
        Case vbYes
            Call selectFile(frmDep, "S:\Stedin\Diensten\SMB\Operations KVB\Datacollectie\Uitvalbakken\07.Deployment TD Proces\02. SplunkInDeployment\")
            i = 1
        Case vbCancel
            cancelMsg = "Handeling afgebroken door gebruiker"
            Call cancelAction(frmDep)
    End Select
End If

If i = 1 Then
    Call deleteTable("metersInDeployment")
    DoCmd.TransferText acImportDelim, "SplunkInDeployment", "metersInDeployment", fileName, True
End If

'   Generate export
'   ---------------------------------------
depStartFile = "S:\Stedin\Diensten\SMB\Operations KVB\Datacollectie\Uitvalbakken\02_Terugkoppelingen\2016\MFE\" & Format(Date, "yyyymmdd") & "_StartControle_MFE.CSV"
If Dir(depStartFile) <> "" Then Kill depStartFile
Qstr = "export_ControleStart"
Call createQuery(Qstr)
Set qt = db.QueryDefs(Qstr)

'   Ook verkeerdom meenemen
If MsgBox("Ook verkeerdom Aangesloten meters meenemen?", vbYesNo, "Uitzetten?") = vbYes Then
    n = 2
    SQL = "SELECT DISTINCT locationEAN as [ean code], equipmentId FROM ("
    SQL = SQL & " SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, t1.startDeploymentTs"
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[E-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer E (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""EL"" AND t2.[Laatste comm E]>=[t2].[MFE datum]-2 AND t2.[Laatste comm E]<=[t2].[MFE datum]+2 "
    SQL = SQL & " AND t1.startDeploymentTs<=[t2].[MFE datum]-10 AND t1.lastCommunicationTs > date()-10 "
    SQL = SQL & " UNION ALL SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, "
    SQL = SQL & " t1.startDeploymentTs"
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[G-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer G (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""GAS""  AND t2.[Laatste comm G]>=[t2].[MFE datum]-2 AND t2.[Laatste comm G]<=[t2].[MFE datum]+2 AND t1.startDeploymentTs<=[t2].[MFE datum]-10 "
    SQL = SQL & " AND t1.lastCommunicationTs > date()-10) "
'   Alleen Nieuw
ElseIf MsgBox("Ook eerder uitgezette meters opnieuw proberen?", vbYesNo, "Uitzetten?") = vbYes Then
    n = 0
    SQL = "SELECT DISTINCT locationEAN as [ean code], equipmentId FROM ("
    SQL = SQL & " SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[E-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer E (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""EL"" AND t2.[Laatste comm E]>=[t2].[MFE datum]-2 AND t2.[Laatste comm E]<=[t2].[MFE datum]+2 AND (t4s.opmerking Not Like ""verkeerdom*"" OR t4s.opmerking is null) "
    SQL = SQL & " AND t1.startDeploymentTs<=[t2].[MFE datum]-10 AND t1.lastCommunicationTs > date()-10 "
    SQL = SQL & " UNION ALL SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, "
    SQL = SQL & " t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[G-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer G (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""GAS""  AND t2.[Laatste comm G]>=[t2].[MFE datum]-2 AND t2.[Laatste comm G]<=[t2].[MFE datum]+2 AND t1.startDeploymentTs<=[t2].[MFE datum]-10 "
    SQL = SQL & " AND t1.lastCommunicationTs > date()-10) "
Else
    n = 1
    SQL = "SELECT DISTINCT locationEAN as [ean code], equipmentId FROM ("
    SQL = SQL & " SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[E-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer E (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""EL"" AND t2.[Laatste comm E]>=[t2].[MFE datum]-2 AND t2.[Laatste comm E]<=[t2].[MFE datum]+2 AND (t4s.opmerking Not Like ""verkeerdom*"" OR t4s.opmerking is null) "
    SQL = SQL & " AND t1.startDeploymentTs<=[t2].[MFE datum]-10 AND t1.lastCommunicationTs > date()-10 AND (t3.datum is null OR t3.datum=#" & Format(Date, "yyyy-mm-dd") & "#) "
    SQL = SQL & " UNION ALL SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, "
    SQL = SQL & " t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[G-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer G (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""GAS"" AND t2.[Laatste comm G]>=[t2].[MFE datum]-2 AND t2.[Laatste comm G]<=[t2].[MFE datum]+2 AND t1.startDeploymentTs<=[t2].[MFE datum]-10 "
    SQL = SQL & " AND t1.lastCommunicationTs > date()-10 AND (t3.datum is null OR t3.datum=#" & Format(Date, "yyyy-mm-dd") & "#)) "
End If

qt.SQL = SQL
Set rs = db.OpenRecordset(Qstr)
If rs.EOF Then
    i = 0
Else
    rs.MoveLast
    i = rs.RecordCount
End If
rs.Close
Set rs = Nothing

DoCmd.TransferText acExportDelim, "Export_ControleStart Spec", Qstr, depStartFile, True

SQL = "DELETE * FROM controleGestartnTal WHERE datum=#" & Format(Date, "yyyy-mm-dd") & "#"
    db.Execute SQL
SQL = " INSERT INTO controleGestartnTal(datum,nTalUitgezet) values(#" & Format(Date, "yyyy-mm-dd") & "#," & i & " )"
    db.Execute SQL
    
Call deleteQuery(Qstr)

SQL = "DELETE * FROM controlegestart WHERE datum=#" & Format(Date, "yyyy-mm-dd") & "#"
    db.Execute SQL
    
'   Insert new records into total list of restarted
'   ---------------------------------------
If n = 0 Then
    SQL = "INSERT INTO ControleGestart([ean code], equipmentId, datum) "
    SQL = SQL & "SELECT DISTINCT locationEAN, equipmentId, #" & Format(Date, "yyyy-mm-dd") & "# FROM ("
    SQL = SQL & " SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[E-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer E (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""EL"" AND t2.[Laatste comm E]>=[t2].[MFE datum]-2 AND t2.[Laatste comm E]<=[t2].[MFE datum]+2 AND (t4s.opmerking Not Like ""verkeerdom*"" OR t4s.opmerking is null) "
    SQL = SQL & " AND t1.startDeploymentTs<=[t2].[MFE datum]-10 AND t1.lastCommunicationTs > date()-10 "
    SQL = SQL & " UNION ALL SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, "
    SQL = SQL & " t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[G-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer G (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""GAS""  AND t2.[Laatste comm G]>=[t2].[MFE datum]-2 AND t2.[Laatste comm G]<=[t2].[MFE datum]+2 AND t1.startDeploymentTs<=[t2].[MFE datum]-10 "
    SQL = SQL & " AND t1.lastCommunicationTs > date()-10) "
        db.Execute SQL
ElseIf n = 1 Then
    SQL = "INSERT INTO ControleGestart([ean code], equipmentId, datum) "
    SQL = SQL & "SELECT DISTINCT locationEAN, equipmentId, #" & Format(Date, "yyyy-mm-dd") & "# FROM ("
    SQL = SQL & " SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[E-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer E (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""EL"" AND t2.[Laatste comm E]>=[t2].[MFE datum]-2 AND t2.[Laatste comm E]<=[t2].[MFE datum]+2 AND (t4s.opmerking Not Like ""verkeerdom*"" OR t4s.opmerking is null) "
    SQL = SQL & " AND t1.startDeploymentTs<=[t2].[MFE datum]-10 AND t1.lastCommunicationTs > date()-10 AND (t3.datum is null OR t3.datum=#" & Format(Date, "yyyy-mm-dd") & "#) "
    SQL = SQL & " UNION ALL SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, "
    SQL = SQL & " t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[G-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer G (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""GAS"" AND t2.[Laatste comm G]>=[t2].[MFE datum]-2 AND t2.[Laatste comm G]<=[t2].[MFE datum]+2 AND t1.startDeploymentTs<=[t2].[MFE datum]-10 "
    SQL = SQL & " AND t1.lastCommunicationTs > date()-10 AND (t3.datum is null OR t3.datum=#" & Format(Date, "yyyy-mm-dd") & "#)) "
        db.Execute SQL
ElseIf n = 2 Then
    SQL = "INSERT INTO ControleGestart([ean code], equipmentId, datum) "
    SQL = SQL & "SELECT DISTINCT locationEAN, equipmentId, #" & Format(Date, "yyyy-mm-dd") & "# FROM ("
    SQL = SQL & " SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[E-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer E (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""EL"" AND t2.[Laatste comm E]>=[t2].[MFE datum]-2 AND t2.[Laatste comm E]<=[t2].[MFE datum]+2  "
    SQL = SQL & " AND t1.startDeploymentTs<=[t2].[MFE datum]-10 AND t1.lastCommunicationTs > date()-10 "
    SQL = SQL & " UNION ALL SELECT t1.equipmentId, t1.locationEan, t4s.opmerking, t4.subcategorie, "
    SQL = SQL & " t1.startDeploymentTs, t3.datum as [Eerder uitgezet], "
    SQL = SQL & " (SELECT count(equipmentId) FROM controlegestart WHERE equipmentId=t1.equipmentId) AS [nTal Keer uitgezet] "
    SQL = SQL & " FROM ((ControleGestart AS t3 RIGHT JOIN metersInDeployment AS t1 ON t3.equipmentId = t1.equipmentId) "
    SQL = SQL & " INNER JOIN [MFE STATS] AS t2 ON t1.equipmentId = t2.[G-Meternummer]) LEFT JOIN ([eerstelijns incident] AS t4 "
    SQL = SQL & " LEFT JOIN DEP_TD_all AS t4s ON t4.Incidentnummer = t4s.Incidentnummer) ON t1.equipmentId = t4.[Meternummer G (Splunk Informatie PRODUCTIE)] "
    SQL = SQL & " WHERE 1 = 1 AND t1.supplytype=""GAS""  AND t2.[Laatste comm G]>=[t2].[MFE datum]-2 AND t2.[Laatste comm G]<=[t2].[MFE datum]+2 AND t1.startDeploymentTs<=[t2].[MFE datum]-10 "
    SQL = SQL & " AND t1.lastCommunicationTs > date()-10) "
        db.Execute SQL
End If

MsgBox "Bestand geëxporteerd: " & vbNewLine & depStartFile & vbNewLine & " Aantal meters opnieuw door controle routine: " & i & "."

'   ---------------------------------------
'   End sequence for code execution to reset environments and variables

Err.Clear
Call endRoutine(frmDep)
Exit Sub

errorHandler:
Call endRoutine(frmDep)

End Sub

Sub depComStatus()
'   ---------------------------------------
If gEnableErrorHandling Then On Error GoTo errorHandler

DoCmd.SetWarnings False
DoCmd.Hourglass True
If CurrentProject.AllForms(frmDep).IsLoaded = False Then
    DoCmd.OpenForm frmDep, acNormal, , , , acHidden
Else
    Forms.Item(frmDep).Visible = False
End If

Set db = CurrentDb()
Call closeAllObjects(frmDep)

'   ---------------------------------------
'   Importeren van MFE, controle op reeds aanwezige MFE Dump en Datum van Dump.
Set rs = db.OpenRecordset("MFE Stats")
If rs.EOF Then
    MsgBox "Selecteer de MFE Dump"
    Call selectFileMFE(frmDep, "Splunk_meterData Import")
Else
    response = MsgBox("MFE Dump bijwerken? Huidige datum is: " & Application.DMax("[MFE datum]", "[MFE stats]") & ".", vbYesNoCancel + vbQuestion)
    Select Case response
        Case vbYes
            Call selectFileMFE(frmDep, "Splunk_meterData Import")
        Case vbCancel
            cancelMsg = "Handeling afgebroken door gebruiker"
            Call cancelAction(frmDep)
    End Select
End If

Call deleteImportErrorTables
rs.Close
'   ---------------------------------------
'   bestaande gegevens bijwerken (controle op gas-meters)

SysCmd acSysCmdSetStatus, "bestaande sharepoint gegevens bijwerken"
DoEvents
SQL = "UPDATE [work-around deployment] as t1 INNER JOIN [MFE STATS] as t2 On t1.[G-meter EAN]=t2.[ean code G] "
SQL = SQL & " SET t1.[Terugkoppeling Slimme Data]=""Communiceert " & Format(Date, "dd-mm-yyyy") & """ "
SQL = SQL & " WHERE t2.[laatste comm G]>t2.[mfe datum]-2 AND t1.[terugkoppeling Slimme Data] is null"
    db.Execute SQL
gcom = db.RecordsAffected

SysCmd acSysCmdClearStatus

MsgBox "Aantal bestaande meters status bijgewerkt:(communiceren) " & gcom _
        , vbOKOnly + vbInformation, "Overzicht"

'   ---------------------------------------

'   ---------------------------------------
'   End sequence for code execution to reset environments and variables
Err.Clear
Call endRoutine(frmDep)
Exit Sub

errorHandler:
Call endRoutine(frmDep)

End Sub
Sub depUitvoerder()

'   ---------------------------------------
If gEnableErrorHandling Then On Error GoTo errorHandler

DoCmd.SetWarnings False
DoCmd.Hourglass True
If CurrentProject.AllForms(frmDep).IsLoaded = False Then
    DoCmd.OpenForm frmDep, acNormal, , , , acHidden
Else
    Forms.Item(frmDep).Visible = False
End If

Set db = CurrentDb()
Call closeAllObjects(frmDep)

 If CurrentProject.AllForms.Item("Loc_BW").IsLoaded Then DoCmd.Close acForm, "Loc_BW", acSaveNo
    DoCmd.Hourglass False
    DoCmd.OpenForm "Loc_BW", acNormal, , , acFormEdit, acDialog
    DoCmd.Hourglass True

SQL = " UPDATE [Work-around deployment] as t1 INNER JOIN results_BW  as t2 ON t1.Locatienummer=clng(t2.Locatienummer) "
SQL = SQL & " SET t1.uitvoerder=t2.uitvoerdernaam, t1.monteur=t2.monteurnaam WHERE t1.uitvoerder Is Null AND (t2.[Nieuw meternummer Elektra] & t2.[Nieuw meternummer GAS]<> ""##"")  "
    db.Execute SQL
    
SQL = "UPDATE [work-around deployment] as t1 INNER JOIN refUitvoerders as t2 ON t1.uitvoerder=t2.uitvoerder "
SQL = SQL & " SET Kavel_Niet_Kavel = afdeling"
SQL = SQL & " WHERE kavel_Niet_Kavel is null"
    db.Execute SQL

'   ---------------------------------------
'   End sequence for code execution to reset environments and variables
Err.Clear
Call endRoutine(frmDep)
Exit Sub

errorHandler:
Call endRoutine(frmDep)
End Sub
