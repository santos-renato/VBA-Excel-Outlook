Attribute VB_Name = "Email_Creator"
Sub Email_creator()

    ' Macro to send each Sales Order template in worksheets to countries in scope

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .Application.StatusBar = "Macro is running, please wait..."
    End With

    Dim EmailList() As String
    Dim ArrayPaths() As String
    Dim Country_Current As String, Country_Next As String, PO_Number As String, SavePath As String, SavePathGroup As String, Sign_def As String
    Dim EmailLR, EmailLC As Byte
    Dim StartSheet As Byte, TotalSheets As Byte
    Dim MacroBook As Workbook, NewBook As Workbook
    Dim Attachment As Variant
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    Dim objEmail As Object
    
    Set MacroBook = ActiveWorkbook
    
    EmailLR = ShMailList.Range("A" & Rows.Count).End(xlUp).row
    EmailLC = ShMailList.Cells(1, Columns.Count).End(xlToLeft).Column
    StartSheet = 6
    TotalSheets = MacroBook.Worksheets.Count
    
    ' loop to order sheets alphabetically -> will be useful to later on send templates per country
    For i = StartSheet To TotalSheets - 1
        For j = i + 1 To TotalSheets
            If UCase(MacroBook.Sheets(j).Name) < UCase(MacroBook.Sheets(i).Name) Then
                MacroBook.Sheets(j).Move before:=MacroBook.Sheets(i)
            End If
        Next j
    Next i
    
    ' fill array for email distribution list
    ReDim EmailList(2 To EmailLR, 2 To EmailLC)
    For i = 2 To EmailLR
        For j = 2 To EmailLC
            EmailList(i, j) = ShMailList.Cells(i, j).Value
        Next j
    Next i
    
    ' check if directory to save temporary files exist. if not create it
    If Dir("C:\IMAC_Templates_Email_Temp", vbDirectory) = "" Then
        MkDir ("C:\IMAC_Templates_Email_Temp")
    Else
    End If

    ' loop to create emails per country
    For i = StartSheet To TotalSheets
        Country_Current = Left(MacroBook.Sheets(i).Name, 2)
        If i = TotalSheets Then GoTo CountryLimit:
        Country_Next = Left(MacroBook.Sheets(i + 1).Name, 2)
CountryLimit:
        PO_Number = MacroBook.Sheets(i).Cells(5, 4)
        SavePath = "C:\IMAC_Templates_Email_Temp\IMAC_Pricing_" & Country_Current & " " & "PO_" & PO_Number & ".xlsx"
        MacroBook.Sheets(i).Copy
        Set NewBook = ActiveWorkbook
        NewBook.SaveAs SavePath
        NewBook.Close False
        If i = TotalSheets Then GoTo Here:
        If Country_Current = Country_Next Then
            ' all paths in one string separated by comma for later split it into array
            SavePathGroup = SavePathGroup & SavePath & ","
        Else
Here:
            SavePathGroup = SavePathGroup & SavePath & ","
            ' clean array for next different country
            Erase ArrayPaths
            ' fill array with split folder paths
            ArrayPaths = Split(Left(SavePathGroup, Len(SavePathGroup) - 1), ",")
            ' email creation starts here
            Set objEmail = objOutlook.CreateItem(olMailItem)
            objEmail.display
            ' get default user signature
            Sign_def = objEmail.HTMLbody
            ' configuration of email
            With objEmail
                .SentOnBehalfOfName = ""
                .To = Application.WorksheetFunction.VLookup(Country_Current, EmailList, 2, False)
                .CC = Application.WorksheetFunction.VLookup(Country_Current, EmailList, 3, False)
                .Subject = "IMAC/HW Rfc_Pricing_" & Country_Current
                .HTMLbody = "Hello,<br><br>" & _
                        "Please invoice according to the attached file.<br><br>" & _
                        "Thanks.<br><br>" _
                        & Sign_def
                For Each Attachment In ArrayPaths
                .Attachments.Add Attachment
                Next
                .display
            End With
            ' Clear paths for next countries
            SavePath = ""
            SavePathGroup = ""
        End If
    Next i
    
    ' delete all email attachments from temporary file
    Kill "C:\IMAC_Templates_Email_Temp\*.xlsx"
    
    MacroBook.Activate
    Sheet3.Activate
    MsgBox "Emails created with success!", vbInformation, "Task Completed"

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .Application.StatusBar = ""
    End With
    
    End Sub
