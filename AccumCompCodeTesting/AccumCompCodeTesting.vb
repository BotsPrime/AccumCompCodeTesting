Option Explicit On

Imports System.Deployment.Application
Imports System.IO
Imports System.Diagnostics
Imports System.Net.Mail
Imports bgw = System.ComponentModel
Imports pcom = AutPSTypeLibrary             'Need to add this reference for this (PCOMM autECLPS Automation Object 1.0 Library)

Imports System.Runtime.InteropServices      'for Marshal.ReleaseComObject


Public Class AccumCompCodeTesting

    Dim objRx As pcom.AutPS
    Dim objWait As Object
    Dim objMgr, objMgr2 As Object
    Dim ObjSessionHandle As Integer
    Dim intSessions As Integer, x As Integer
    Dim autECLConnList As Object

    'Excel Object Variables
    Dim objExcel
    Dim objWorkbook1
    Dim objWorksheet1
    Dim objWorksheet2

    'Dim objExcelAppObject
    Dim objExcelfolder
    Dim objExcelDirectory
    Dim objExcelFilePath
    Dim msoFileDialogFolderPicker
    Dim var_SplitString
    Dim int_UBoundIndex
    Dim var_FileName As String

    Dim StartTime_Overall As DateTime
    Dim EndTime_Overall As DateTime

    Dim StartTime_ForRow As DateTime
    Dim EndTime_ForRow As DateTime

    Dim iRxNumberCounter As Long
    Dim sFillDate As String
    Dim iMemID As String

    Dim Dev_Prod       'Database Source Name  ... this will help control Prod vs Dev

    Dim rowNum As Integer       'This will walk us thru each Row on the spreadsheet

    'Global Variables that we will now display on the form
    Dim iQty As Integer = 30
    Dim iDaySupply As Integer = 30
    Dim iCost As Integer = 100


    Private Sub AccumCompCodeTesting_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.cmbEnv.SelectedIndex = 0

        lblQty.Text = iQty.ToString
        lblDaySupply.Text = iDaySupply.ToString
        lblCost.Text = iCost.ToString
    End Sub

    Public Sub getMemberID(RxClaimID As String)
        GoHome()

        IsRightScreenName("CCT600", 1, 2, 60000)
        TypeMe("1")         'Eligibility/Claim Transaction
        MoveMe("enter", 1)

        waitForMe()

        IsRightScreenName("CCT610", 1, 2, 60000)
        TypeMe("6")         'Claim Transaction
        MoveMe("enter", 1)

        waitForMe()

        IsRightScreenName("RCRXS", 1, 2, 60000)
        MoveMe("pf8", 1)    'Search by RxClaim#

        waitForMe()

        SettingText(RxClaimID, 4, 4)    'Text, row, col
        MoveMe("enter", 1)

        waitForMe()

        If RxClaimID = Trim(objRx.GetText(10, 4, 15)) Then
            SettingText("5", 10, 2)    'Text, row, col
            MoveMe("enter", 1)

            waitForMe()

            iMemID = Trim(objRx.GetText(8, 12, 22))
        Else
            MsgBox("We could NOT find the MemberId of - " & iMemID)
            iMemID = "0"
        End If

        waitForMe()

        GoHome()
    End Sub

    Private Sub btnRun_Click(sender As System.Object, e As System.EventArgs) Handles btnRun.Click
        Try
            Me.btnRun.Enabled = False

            If GetSpreadsheet() = False Then
                MsgBox("Sorry...we are experiencing difficulties with opening up the spreadsheet.")
                Exit Sub
            End If

            OpenRxClaim_Session()

            'make a for loop here to go thru each member
            Dim int_RecordCount = objWorkbook1.Worksheets(1).Range("A1").CurrentRegion.Rows.Count

            If int_RecordCount > 1 Then
                Dim int_Counter_Main As Integer

                'Set the RxNum Counter...this will now be unique for EVERY claim
                iRxNumberCounter = txtRxNum.Text

                'Start overall timer
                StartTime_Overall = DateTime.Now

                objWorkbook1.Worksheets(1).Cells(2, 51).Value = GetUsername()   'Username
                objWorkbook1.Worksheets(1).Cells(2, 52).Value = cmbEnv.SelectedItem   'Environment

                For int_Counter_Main = 2 To int_RecordCount '+ 1		'Need the plus 1 if we only do one record
                    'Start timer for this row
                    StartTime_ForRow = DateTime.Now

                    rowNum = int_Counter_Main       'rowNum will be used throughout the program

                    iMemID = objWorkbook1.Worksheets(1).Cells(rowNum, 1).Value

                    'getMemberID(objWorkbook1.Worksheets(1).Cells(rowNum, 1).Value)

                    'Preset the value 
                    'iRxNumberCounter = 50151101080
                    'sFillDate = "01-01-2015"

                    'iRxNumberCounter = txtRxNum.Text
                    sFillDate = FormatDate(Me.DTP_FillDate.Value)

                    'ChangeMemberPlan(1)     '1 = 1st time thru             'Leave Alone…we are not changing Member Plan 1st anymore (as of 12/30/15)

                    SubmitClaim(3)          '3 = 3 Times this will be run

                    '1st Scrape
                    ScrapeAccumulators(1)   '1 = 1st time thru

                    ChangeMemberPlan(2)     '2 = 2nd time thru

                    '2nd scrape
                    ScrapeAccumulators(2)   '2 = 2nd time thru

                    SubmitClaim(1)          '1 = 1 Time this will be run

                    '3rd scrape
                    ScrapeAccumulators(3)   '3 = 3rd time thru

                    'Compare 1st, 2nd, & 3rd Deduction scrapes
                    CompareDeductions()

                    'End the timer for this row
                    EndTime_ForRow = DateTime.Now

                    objWorkbook1.Worksheets(1).Cells(rowNum, 44).Value = (EndTime_ForRow - StartTime_ForRow).ToString("hh':'mm':'ss")
                Next

                'End Overall timer
                EndTime_Overall = DateTime.Now

                objWorkbook1.Worksheets(1).Cells(2, 48).Value = StartTime_Overall
                objWorkbook1.Worksheets(1).Cells(2, 49).Value = EndTime_Overall

                objWorkbook1.Worksheets(1).Cells(2, 46).Value = (EndTime_Overall - StartTime_Overall).ToString("hh':'mm':'ss")

                'Close down shop
                CloseApp()

                Me.btnRun.Enabled = True
            End If

        Catch ex As Exception
            MsgBox("Error in:  btnRun_Click()  ...  " & ex.ToString)
        End Try
    End Sub

    Public Sub CompareDeductions()

        Dim i1, i2 As Integer

        'compare if numeric
        If IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 20).Value)) And IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 25).Value)) Then
            i1 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 20).Value)
            i2 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 25).Value)

            objWorkbook1.Worksheets(1).Cells(rowNum, 35).Value = i1 - i2        'Diff for 1st and 2nd Ded
        End If

        If IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 21).Value)) And IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 26).Value)) Then
            i1 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 21).Value)
            i2 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 26).Value)

            objWorkbook1.Worksheets(1).Cells(rowNum, 36).Value = i1 - i2        'Diff for 1st and 2nd OOP
        End If

        '****************************************

        If IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 20).Value)) And IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 32).Value)) Then
            i1 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 20).Value)
            i2 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 32).Value)

            objWorkbook1.Worksheets(1).Cells(rowNum, 38).Value = i1 - i2        'Diff for 1st and 3rd Ded
        End If

        If IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 21).Value)) And IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 33).Value)) Then
            i1 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 21).Value)
            i2 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 33).Value)

            objWorkbook1.Worksheets(1).Cells(rowNum, 39).Value = i1 - i2        'Diff for 1st and 3rd OOP
        End If

        '****************************************

        If IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 25).Value)) And IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 32).Value)) Then
            i1 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 25).Value)
            i2 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 32).Value)

            objWorkbook1.Worksheets(1).Cells(rowNum, 41).Value = i1 - i2        'Diff for 2nd and 3rd Ded
        End If

        If IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 26).Value)) And IsNumeric(Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 33).Value)) Then
            i1 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 26).Value)
            i2 = Trim(objWorkbook1.Worksheets(1).Cells(rowNum, 33).Value)

            objWorkbook1.Worksheets(1).Cells(rowNum, 42).Value = i1 - i2        'Diff for 2nd and 3rd OOP
        End If

    End Sub

    Public Sub CloseApp()
        ''Close Results Spreadsheet
        objExcel.DisplayAlerts = False

        'If Len(Trim(var_TimeStamp)) > 0 Then
        'objExcel.ActiveWorkbook.SaveAs("C:\Users\Public\Reverse_Resubmit Reports\R_and_R_ " & var_TimeStamp & ".xlsx")
        objExcel.ActiveWorkbook.Save()

        'End If

        objExcelFilePath = Nothing
        objExcelfolder = Nothing

        '****   Close Excel *********************************************************
        objExcel.ActiveWorkbook.Close()
        objExcel.Quit()
        '****************************************************************************

        Marshal.ReleaseComObject(objExcel)

        'Clean up
        objExcel = Nothing
        objWorkbook1 = Nothing
        objWorksheet1 = Nothing
        objWorksheet2 = Nothing

        GC.Collect()

        '*****  Close RxClaim session  **********************************************
        objMgr2.StopConnection(ObjSessionHandle)
        ''***************************************************************************

        MsgBox("All Done")
    End Sub

    Public Function GetSpreadsheet() As Boolean
        GetSpreadsheet = False

        Try
            Dim fileDialog As New OpenFileDialog

            fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            fileDialog.Filter = "Excel Worksheets|*.xls;*.xlsx;*.xlsm"
            fileDialog.FilterIndex = 2
            fileDialog.RestoreDirectory = True

            If fileDialog.ShowDialog() <> DialogResult.OK Then
                Exit Function
            Else
                GetSpreadsheet = True
            End If

            objExcelFilePath = fileDialog.FileName

            lblSpreadsheetLoc.Text = objExcelFilePath
            lblSpreadsheetName.Text = System.IO.Path.GetFileName(objExcelFilePath)

            objExcel = CreateObject("Excel.Application")
            objWorkbook1 = objExcel.Workbooks.Open(objExcelFilePath)

            objExcelFilePath = Nothing      'now that we are done with it...clear it.

            objWorksheet1 = objWorkbook1.Worksheets(1)
            'objWorksheet2 = objWorkbook1.Worksheets(2)

            'objExcel.Visible = False        '--only do this if you want to see the progress
            objExcel.Visible = True        '--only do this if you want to see the progress

        Catch ex As Exception
            MsgBox("Error in:  btnRun_Click()  ...  " & ex.ToString)
        End Try
    End Function

    Public Sub OpenRxClaim_Session()
        Try
            'Dim var_ScreenID
            Dim intSessions, x, y As Integer
            Dim Envir As String

            objRx = CreateObject("PCOMM.autECLPS")
            objWait = CreateObject("PCOMM.autECLOIA")
            objMgr = CreateObject("PCOMM.autECLConnMgr")
            autECLConnList = CreateObject("PCOMM.autECLConnList")

            intSessions = objMgr.autECLConnList.Count

            '** So...if we have session 1 (a), 2, (b) open and I close session 1 (a)...and now I want to open a new session...
            'the new session will be A...which is 1...that is the one I want to use.
            If intSessions > 0 Then

                For x = 1 To intSessions
                    y = 0

                    If x = 1 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "a" Then
                        y = 1
                    ElseIf x = 2 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "b" Then
                        y = 2
                    ElseIf x = 3 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "c" Then
                        y = 3
                    ElseIf x = 4 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "d" Then
                        y = 4
                    ElseIf x > 4 Then
                        'shouldn't have more than 5 sessions open...right?!
                        MsgBox("Sorry you have too many RxClaim sessions open." & Chr(13) & "Please close 1 or more and try again.")
                        Exit Sub
                    End If

                    If y > 0 Then
                        Exit For
                    End If
                Next

                If y = 0 Then y = intSessions + 1

            ElseIf intSessions = 0 Then
                y = 1
            Else
                MsgBox("SOMETHING IS WRONG...")
                Exit Sub
            End If

            'Now find the "File name" to open up based on their selection
            If LCase(cmbEnv.SelectedItem) = "dev01" Then
                Envir = "Dev01.AS4"
            ElseIf LCase(cmbEnv.SelectedItem) = "dev02" Then
                Envir = "Dev02.AS4"
            ElseIf LCase(cmbEnv.SelectedItem) = "prod03" Then
                Envir = "PROD03.AS4"
            ElseIf LCase(cmbEnv.SelectedItem) = "prod01" Then
                Envir = "PROD01.AS4"
                'MsgBox("Sorry...not available to use at this time")
                'Exit Sub
            Else
                MsgBox("Environment was not found ... exiting.")
                Exit Sub
            End If

            Dim sDir As String = getMyDocs()

            'MsgBox("getMyDocs is:  " & sDir)

            If sDir.Length > 1 Then
                'Now we are trying to open up a session
                Try
                    Process.Start(sDir & "RxClaims Sessions\" & Envir)
                Catch
                    Try
                        Process.Start("C:\Users\Public\Desktop\RxClaims Sessions\" & Envir)
                    Catch ex As Exception
                        MsgBox("Please open up an RxClaim session and then press 'OK' to this message")
                    End Try
                End Try
            Else
                MsgBox("couldn't find Desktop")
            End If

            'Now try connecting to that session ... we will wait 7 seconds
            'This is a hard wait to ensure that the RxClaim session has started
            waitOnMe(7000)

            objMgr2 = CreateObject("PCOMM.autECLConnMgr")

            ObjSessionHandle = objMgr2.autECLConnList(y).Handle

            objRx.SetConnectionByHandle(ObjSessionHandle)
            objWait.SetConnectionByHandle(ObjSessionHandle)


            '*****  Initialize  *******************************************************************

            'IF 19,2 for 11 = "Press Enter"  ...  This is usually the 1st screen that shows if you already have another session open
            If Trim(objRx.GetText(19, 2, 11)) = "Press Enter" Then
                objRx.SendKeys("[Enter]")
                waitOnMe(1000)
            End If

            'If Days until password expires message appears...
            If Trim(objRx.GetText(7, 2, 28)) = "Days until password expires" Then
                objRx.SendKeys("[Enter]")
                waitOnMe(1000)
            End If

            'Need to account for a "Display Messages" screen
            If Trim(objRx.GetText(1, 2, 60)) = "Display Messages" Then
                objRx.SendKeys("[Enter]")
                waitOnMe(1000)
            End If

            waitForMe()

            IsRightScreenName("Prime", 1, 33, 5000)

            waitForMe()

            If LCase(cmbEnv.SelectedItem) = "prod03" Then
                objRx.SetText("PPF", 21, 7)
            Else
                objRx.SetText("RX6", 21, 7)
            End If

            waitForMe()
            MoveMe("enter", 1)

            intSessions = objMgr.autECLConnList.Count

            waitForMe()

            Dim var_ScreenID
            var_ScreenID = Trim(objRx.GetText(1, 2, 10))

            waitForMe()

            Dev_Prod = cmbEnv.SelectedItem.ToString

            waitForMe()

            IsRightScreenName("CCT6", 1, 2, 60000)
            '**************************************************************************************

            waitForMe()
        Catch ex As Exception
            MsgBox("Error in:  OpenRxClaim_Session()  ...  " & ex.ToString)
        End Try
    End Sub

    Public Sub SubmitClaim(iTimes As Integer)
        Try
            GoHome()

            IsRightScreenName("CCT600", 1, 2, 60000)
            TypeMe("3")         'Manual Claim
            MoveMe("enter", 1)

            waitForMe()

            IsRightScreenName("CCT630", 1, 2, 60000)
            TypeMe("2")         'DO Manual Claim
            MoveMe("enter", 1)

            waitForMe()

            IsRightScreenName("CCT632", 1, 2, 60000)
            TypeMe("1")         'Transaction
            MoveMe("enter", 1)

            waitForMe()

            IsRightScreenName("RCNCP050", 1, 2, 60000)
            MoveMe("pf6", 1)    'Add

            'Now start adding values to the "Add Transaction" page *****************************************
            '**This is a screen on top of a screen...there are 2 fields that need to be tabbed out of (BIN and Fill Dt)

            waitForMe()

            'Bin	(*this has a problem if we do NOT tab out*)
            SettingText("610455", 11, 14)  'Text, row, col
            MoveMe("tab", 1) 'Do NOT remove this...This tab needs to be here!  ... will get this error:  "The key used to exit field not valid"

            'Proc Ctrl
            SettingText("PGIGN", 11, 41)  'Text, row, col

            'Grp
            SettingText("0", 11, 59)    'Text, row, col

            'Pharmacy
            SettingText("2408474", 12, 14)    'Text, row, col

            'Rx Nbr
            'distinguish between do and 51...lengths vary!!!!
            SettingText(iRxNumberCounter, 12, 41) 'Text, row, col

            'adjust counter for next time thru
            iRxNumberCounter = iRxNumberCounter + 1

            'Rf	
            SettingText("00", 12, 59)    'Text, row, col

            'Fill Dt
            MoveMe2("eraseeof", 14, 14)
            SettingText(sFillDate, 14, 14)    'Text, row, col
            MoveMe("tab", 1) 'DoNOT remove this...This tab needs to be here!  ... will get this error:  "The key used to exit field not valid"

            'Member Id
            'SettingText(iMemID, 14, 41)    'Text, row, col
            SettingText("?", 14, 41)    'Text, row, col

            MoveMe("enter", 1)

            'Enter MemberId
            MoveMe2("eraseeof", 3, 4)
            SettingText(iMemID, 3, 4)    'Text, row, col

            MoveMe("enter", 1)

            If iMemID = Trim(objRx.GetText(8, 4, 15)) Then
                SettingText("1", 8, 2)    'Text, row, col
                MoveMe("enter", 1)
            Else
                MsgBox("We could NOT find the MemberId of - " & iMemID)
            End If

            'LN 2119

            waitForMe()

            '**********************************************************************************************

            'ProdId
            MoveMe2("eraseeof", 11, 20)
            SettingText(txtProdID.Text, 11, 20)    'Text, row, col

            'Disp Qty
            MoveMe2("eraseeof", 12, 11)
            'SettingText("30", 12, 11)    'Text, row, col
            SettingText(iQty, 12, 11)    'Text, row, col

            'DS
            MoveMe2("eraseeof", 12, 26)
            'SettingText("30", 12, 26)    'Text, row, col
            SettingText(iDaySupply, 12, 26)    'Text, row, col

            'Wrtn Dt
            MoveMe2("eraseeof", 13, 10)
            SettingText("01-01-2015", 13, 10)    'Text, row, col

            'PSC
            MoveMe2("eraseeof", 14, 6)
            SettingText("0", 14, 6)    'Text, row, col

            'Compound Code
            MoveMe2("eraseeof", 14, 14)
            SettingText("1", 14, 14)    'Text, row, col

            'Prescriber Qual
            MoveMe2("eraseeof", 18, 19)
            SettingText("01", 18, 19)    'Text, row, col

            'Prescriber ID
            MoveMe2("eraseeof", 18, 26)
            SettingText("1457467904", 18, 26)    'Text, row, col

            'Cost
            MoveMe2("eraseeof", 10, 47)
            'SettingText("100", 10, 47)    'Text, row, col
            SettingText(iCost, 10, 47)    'Text, row, col

            'DUE
            MoveMe2("eraseeof", 18, 47)
            SettingText("100", 18, 47)    'Text, row, col

            'UC/W
            MoveMe2("eraseeof", 19, 47)
            SettingText("100", 19, 47)    'Text, row, col

            'page down to the 2nd page and enter "1" in for the Rx Orign field  *********
            MoveMe("roll up", 1)

            waitForMe()

            MoveMe2("eraseeof", 4, 55)
            SettingText("1", 4, 55)    'Text, row, col

            MoveMe("enter", 1)

            waitForMe()

            'Page back up to see the status of the change
            MoveMe("roll down", 1)

            waitForMe()

            'If iTimes > 1 then do it again
            For i = 1 To iTimes

                'Submit to Router
                MoveMe("enter", 1)
                MoveMe("pf18", 1)

                'Scrape the RxClaim #, Status, & Rej
                If i = 1 And iTimes = 3 Then        'This meaning that it is the 1st claim
                    objWorksheet1.Cells(rowNum, 8).Value = Trim(objRx.GetText(20, 12, 15))
                    objWorksheet1.Cells(rowNum, 9).Value = Trim(objRx.GetText(21, 6, 1))
                    objWorksheet1.Cells(rowNum, 10).Value = Trim(objRx.GetText(21, 12, 20))
                ElseIf i = 2 Then
                    objWorksheet1.Cells(rowNum, 12).Value = Trim(objRx.GetText(20, 12, 15))
                    objWorksheet1.Cells(rowNum, 13).Value = Trim(objRx.GetText(21, 6, 1))
                    objWorksheet1.Cells(rowNum, 14).Value = Trim(objRx.GetText(21, 12, 20))
                ElseIf i = 3 Then
                    objWorksheet1.Cells(rowNum, 16).Value = Trim(objRx.GetText(20, 12, 15))
                    objWorksheet1.Cells(rowNum, 17).Value = Trim(objRx.GetText(21, 6, 1))
                    objWorksheet1.Cells(rowNum, 18).Value = Trim(objRx.GetText(21, 12, 20))
                ElseIf i = 1 And iTimes = 1 Then    'This would fall into the 4th claim
                    objWorksheet1.Cells(rowNum, 28).Value = Trim(objRx.GetText(20, 12, 15))
                    objWorksheet1.Cells(rowNum, 29).Value = Trim(objRx.GetText(21, 6, 1))
                    objWorksheet1.Cells(rowNum, 30).Value = Trim(objRx.GetText(21, 12, 20))
                End If

                AdjustFillDate()

                'Fill Dt
                MoveMe2("eraseeof", 4, 65)
                SettingText(sFillDate, 4, 65)    'Text, row, col
                MoveMe("tab", 1) 'DoNOT remove this...This tab needs to be here!  ... will get this error:  "The key used to exit field not valid"

                'RxNumber
                MoveMe2("eraseeof", 5, 29)
                SettingText(iRxNumberCounter, 5, 29)    'Text, row, col

                'adjust counter for next time thru
                iRxNumberCounter = iRxNumberCounter + 1
            Next

        Catch ex As Exception
            MsgBox("Error in:  GetSpreadsheet()  ...  " & ex.ToString)
        End Try
    End Sub

    Public Sub AdjustFillDate()
        'If sFillDate = "01-01-2015" Then
        '    sFillDate = "02-01-2015"
        'ElseIf sFillDate = "02-01-2015" Then
        '    sFillDate = "03-01-2015"
        'ElseIf sFillDate = "03-01-2015" Then
        '    sFillDate = "04-01-2015"
        'Else
        '    sFillDate = "01-01-2015"
        'End If

        'cool

        ' Calculate what day of the week is 31 days from this instant. 
        Dim today As System.DateTime
        Dim duration As System.TimeSpan
        Dim answer As System.DateTime

        Dim ry As System.DateTime = sFillDate

        today = ry
        duration = New System.TimeSpan(31, 0, 0, 0)
        answer = today.Add(duration)


        'now split out the date the way we want.  EX: 03-01-2015
        Dim m As String = Microsoft.VisualBasic.Right("0" & answer.Month.ToString, 2)
        Dim d As String = Microsoft.VisualBasic.Right("0" & answer.Day, 2)

        sFillDate = m & "-" & d & "-" & answer.Year

    End Sub

    Public Sub ChangeMemberPlan(iTimeThru As Integer)
        Try
            Dim iPlanColumn As Integer
            Dim iChangeMemberPlanStatusColumn As Integer

            If iTimeThru = 1 Then
                iPlanColumn = 3
                iChangeMemberPlanStatusColumn = 6
            ElseIf iTimeThru = 2 Then
                iPlanColumn = 4
                iChangeMemberPlanStatusColumn = 23
            End If

            GoHome()    'Start at the beginning

            IsRightScreenName("CCT600", 1, 2, 60000)
            TypeMe("1")         'Eligibility/Claim Transaction
            MoveMe("enter", 1)

            waitForMe()

            IsRightScreenName("CCT610", 1, 2, 60000)
            TypeMe("1")         'Carrier/Account/Group
            MoveMe("enter", 1)

            waitForMe()

            IsRightScreenName("CCT610A", 1, 2, 60000)
            TypeMe("3")         'Group/Member
            MoveMe("enter", 1)

            waitForMe()

            IsRightScreenName("RCGRP027", 1, 2, 60000)

            'Enter GroupId
            SettingText(Trim(objWorksheet1.Cells(rowNum, 2).Value), 4, 4)  'Text, row, col
            MoveMe("enter", 1)

            waitForMe()

            If Trim(objRx.GetText(9, 4, 15)) = Trim(objWorksheet1.Cells(rowNum, 2).Value) Then
                SettingText("2", 9, 2)  'Text, row, col
                MoveMe("enter", 1)

                waitForMe()

                IsRightScreenName("RCGRP004", 1, 2, 60000)
                MoveMe("pf7", 1)        'Elig

                waitForMe()

                IsRightScreenName("RCGEL002", 1, 2, 60000)

                waitForMe()

                'Currently adding "2" to the 1st one
                SettingText("2", 13, 2)  'Text, row, col
                MoveMe("enter", 1)

                IsRightScreenName("RCGEL005", 1, 2, 60000)

                waitForMe()

                MoveMe2("eraseeof", 10, 16)

                'grab the member plan
                'SettingText(Trim(objWorksheet1.Cells(rowNum, iPlanColumn).Value), 10, 16)    'Text, row, col


                'If adding for the 1st time... OPT #1 *****
                'If iTimeThru = 1 Then
                '    SettingText(Trim(objWorksheet1.Cells(rowNum, iPlanColumn).Value), 10, 16)    'Text, row, col
                '    SettingText("01-01-15", 10, 37)  'Text, row, col
                'End If
                '******************************************

                'If adding for the 1st time... OPT #2 *****
                If iTimeThru = 1 Then

                    ' did not work -->  SettingText("", 10, 16)  'Text, row, col

                    'objRx.SetCursorPos(10, 16)

                    'maybe then try to tab into this field and then try f4


                    objRx.SetCursorPos(8, 37)  '(row, column)	'set the focus to Thru Dt field
                    MoveMe("tab", 1)            'tab to the Plan
                    waitForMe()

                    MoveMe("pf4", 1)

                    waitForMe()

                    IsRightScreenName("RCPLN018", 1, 2, 60000)      'Select Active Plan by Plan Code

                    waitForMe()

                    SettingText(Trim(objWorksheet1.Cells(rowNum, iPlanColumn).Value), 4, 5)    'Text, row, col
                    MoveMe("enter", 1)

                    If Trim(objWorksheet1.Cells(rowNum, iPlanColumn).Value) = Trim(objRx.GetText(10, 5, 12)) Then
                        'Currently adding "2" to the 1st one
                        SettingText("1", 10, 2)  'Text, row, col
                        MoveMe("enter", 1)
                    End If
                Else
                    SettingText(Trim(objWorksheet1.Cells(rowNum, iPlanColumn).Value), 10, 16)
                End If
                '******************************************

                MoveMe("enter", 1)

                '*******   Confirm Prompt  ************************************

                Dim sCurText As String
                Dim iCurRow, iCurCol As Integer

                iCurRow = objRx.CursorPosRow
                iCurCol = objRx.CursorPosCol

                sCurText = objRx.GetText(iCurRow, iCurCol, 1)

                If sCurText = "N" Then
                    TypeMe("Y")     'Confirm Prompt
                    objWorksheet1.Cells(rowNum, iChangeMemberPlanStatusColumn).Value = "Successful"
                Else
                    objWorksheet1.Cells(rowNum, iChangeMemberPlanStatusColumn).Value = "UN-Successful"
                End If

                '**************************************************************

            End If

        Catch ex As Exception
            MsgBox("Error in:  ChangeMemberPlan()  ...  " & ex.ToString)
        End Try
    End Sub

    Public Sub ScrapeAccumulators(iTimeThru As Integer)
        Try
            Dim iDedCol As Integer
            Dim iOOPCol As Integer

            If iTimeThru = 1 Then
                iDedCol = 20
                iOOPCol = 21
            ElseIf iTimeThru = 2 Then
                iDedCol = 25
                iOOPCol = 26
            ElseIf iTimeThru = 3 Then
                iDedCol = 32
                iOOPCol = 33
            End If

            GoHome()    'Start at the beginning

            IsRightScreenName("CCT600", 1, 2, 60000)
            TypeMe("1")         'Eligibility/Claim Transaction
            MoveMe("enter", 1)

            waitForMe()

            IsRightScreenName("CCT610", 1, 2, 60000)
            TypeMe("2")         'Member
            MoveMe("enter", 1)

            waitForMe()

            IsRightScreenName("RCMBR004", 1, 2, 60000)

            'Enter MemberId
            TypeMe(iMemID)
            MoveMe("enter", 1)

            waitForMe()

            'Select the right member and type "5" and [enter]
            If Trim(objRx.GetText(10, 4, 20)) = iMemID Then
                SettingText("5", 10, 2)    'Text, row, col
                MoveMe("enter", 1)

                waitForMe()

                IsRightScreenName("RCMBR010B", 1, 2, 60000)

                'F8 (Detail)
                MoveMe("pf8", 1)

                waitForMe()

                IsRightScreenName("RCMBR069", 1, 2, 60000)

                'Type "19" and [enter]  (Accumulator)
                SettingText("19", 4, 20)    'Text, row, col
                MoveMe("enter", 1)

                waitForMe()

                IsRightScreenName("RCMA1001", 1, 2, 60000)

                '1-Deductible  **********************************************************************
                SettingText("1", 4, 20)    'Text, row, col
                MoveMe("enter", 1)

                waitForMe()

                IsRightScreenName("RCMA1010", 1, 2, 60000)

                If Len(Trim(objRx.GetText(10, 6, 20))) < 1 Then
                    'Means nothing was found on the "Deductible Period" screen
                    objWorksheet1.Cells(rowNum, iDedCol).Value = "Nothing Found (Deductible Period screen)"

                    'Back out to the Accumulator screen
                    MoveMe("pf12", 1)
                Else
                    'Figure out which one to select...we will select the 1st one for now
                    SettingText("7", 10, 2)    'Text, row, col
                    MoveMe("enter", 1)

                    waitForMe()

                    IsRightScreenName("RCMA1011", 1, 2, 60000)

                    objWorksheet1.Cells(rowNum, iDedCol).Value = Trim(objRx.GetText(12, 69, 10))

                    'Back out to the Accumulator screen
                    MoveMe("pf12", 2)
                End If

                waitForMe()

                IsRightScreenName("RCMA1001", 1, 2, 60000)

                '2-Out of Pocket  *******************************************************************
                SettingText("2", 4, 20)    'Text, row, col
                MoveMe("enter", 1)

                waitForMe()

                IsRightScreenName("RCMA1020", 1, 2, 60000)

                If Len(Trim(objRx.GetText(10, 6, 20))) < 1 Then
                    'Means nothing was found on the "Deductible Period" screen
                    objWorksheet1.Cells(rowNum, iOOPCol).Value = "Nothing Found (OOP Maximum Period screen)"
                Else
                    'Figure out which one to select...we will select the 1st one for now
                    SettingText("7", 10, 2)    'Text, row, col
                    MoveMe("enter", 1)

                    waitForMe()

                    IsRightScreenName("RCMA1021", 1, 2, 60000)

                    objWorksheet1.Cells(rowNum, iOOPCol).Value = Trim(objRx.GetText(12, 69, 10))
                End If

            End If

        Catch ex As Exception
            MsgBox("Error in:  ScrapeDeductible()  ...  " & ex.ToString)
        End Try
    End Sub

    Function getMyDocs() As String
        Try
            Dim WshShell As Object
            WshShell = CreateObject("WScript.Shell")
            getMyDocs = WshShell.SpecialFolders("Desktop") & "\"        'This will get the folder path for my desktop
        Catch ex As Exception
            MsgBox("Error in:  getMyDocs()  ...  " & ex.ToString)
            getMyDocs = ""
        End Try
    End Function

    Sub waitForMe()
        Try
            objWait.WaitForAppAvailable()
            System.Threading.Thread.Sleep(10)
            objWait.WaitForInputReady()
        Catch ex As Exception
            MsgBox("Error in:  waitForMe()  ...  " & ex.ToString)
        End Try
    End Sub

    Public Sub waitOnMe(intHowLong)
        Try
            objRx.Wait(intHowLong)
        Catch ex As Exception
            MsgBox("Error in:  waitOnMe()  ...  " & ex.ToString)
        End Try
    End Sub

    Public Sub IsRightScreenName(scrName, row, col, mil)
        If (objRx.WaitForString(scrName, row, col, mil, True)) Then    'This will wait up to the Milliseconds provided
            'Do Nothing...because we are on the desired screen
        Else
            MsgBox("stop...we have detected that you are not on the expected screen.  Please look into.  scrName is:  " & scrName & " row is:  " & row & " col is:  " & col & " mil is:  " & mil)
        End If
    End Sub

    Sub TypeMe(value)
        Err.Clear()       'This will clear any pre-existing errors

        waitForMe()
        'Enter in the value provided
        objRx.SetText(value)

        'Check here if we have a RED X

        waitForMe()

        'This will display only if we encountered any issues...
        If Err.Number <> 0 Then
            'An Exception occurred
            MsgBox("Exception in TypeMe(): " & vbCrLf & "    Error number: " & Err.Number & vbCrLf & "    Error description: " & Err.Description & vbCrLf)
        End If
    End Sub

    Sub MoveMe(command, amount)
        Err.Clear()       'This will clear any pre-existing errors

        'Do what the command says and do it as many times as the amount says
        'Most common commands will be "tab" and "pf12"

        Dim i As Integer

        For i = 1 To amount
            waitForMe()
            objRx.SendKeys("[" & command & "]")

            'MsgBox("Check here if we have a RED X")

            waitForMe()
        Next

        'This will display only if we encountered any issues...
        If Err.Number <> 0 Then
            'An Exception occurred
            MsgBox("Exception in MoveMe(): " & vbCrLf & "    Error number: " & Err.Number & vbCrLf & "    Error description: " & Err.Description & vbCrLf)
        End If
    End Sub

    Sub MoveMe2(command, r, c)
        waitForMe()
        objRx.SendKeys("[" & command & "]", r, c)
        waitForMe()
    End Sub

    Sub GoHome()
        Try
            'This Subroutine will continue to check to see what screen we are on and
            'get us back to the "Home" screen (CCT600)

            Dim iCounter
            iCounter = 0

            waitForMe()

            'If Trim(objRx.GetText(19, 2, 11)) = "Press Enter" Then

            Do While Trim(objRx.GetText(1, 2, 6)) <> "CCT600"
                waitForMe()
                MoveMe("pf3", 1)
                waitForMe()

                iCounter = iCounter + 1

                If iCounter > 20 Then   'Just in-case it would get stuck in the loop...I wanted a semi-clean way to get out
                    MsgBox("we are exiting GoHome() Subroutine...We probably encountered an error.")
                    Exit Sub
                End If
            Loop
        Catch ex As Exception
            MsgBox("Error in:  GoHome()  ...  " & ex.ToString)
        End Try
    End Sub

    Sub SettingText(text, row, col)
        Err.Clear()       'This will clear any pre-existing errors

        waitForMe()
        objRx.SetText(text, row, col)
        waitForMe()

        'This will display only if we encountered any issues...
        If Err.Number <> 0 Then
            'An Exception occurred
            MsgBox("Exception in SettingText(): " & vbCrLf & "    Error number: " & Err.Number & vbCrLf & "    Error description: " & Err.Description & vbCrLf)
        End If
    End Sub

    Function GetUsername() As String
        Dim objNet      'This will get the username of the person logged into the PC running this Macro
        objNet = CreateObject("WScript.NetWork")
        GetUsername = objNet.UserName
    End Function

    Function FormatDate(myDate)
        'Lets zero fill the day and month

        If IsDate(myDate) Then
            Dim m, d, y

            'm = Right("0" & DatePart("m", myDate), 2)
            m = ("0" & DatePart("m", myDate)).ToString.Substring(("0" & DatePart("m", myDate)).Length - 2, 2)
            'd = Right("0" & DatePart("d", myDate), 2)
            d = ("0" & DatePart("d", myDate)).ToString.Substring(("0" & DatePart("d", myDate)).Length - 2, 2)
            y = DatePart("yyyy", myDate)

            FormatDate = m & "-" & d & "-" & y
        Else
            FormatDate = "          "
        End If
    End Function
End Class
