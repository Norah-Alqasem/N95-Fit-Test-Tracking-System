Attribute VB_Name = "N95Tracker"
Option Explicit

' ============================================================
' N95 Fit Test Expiration Alert System (Excel VBA + Outlook)
' --------------------------------------------------------------
' Scans every department sheet and emails an alert for any employee
' whose Validation Status is "Expiring Soon".
'
' Setup:
'   1. Update CONFIG constants below.
'   2. Make sure Outlook is installed and signed in.
'   3. Run CheckN95FitTests manually to test, then schedule via
'      Windows Task Scheduler or the Workbook_Open event.
' ============================================================

' ---- CONFIG (edit these values) ----
Private Const INFECTION_CONTROL_EMAIL As String = "infectioncontrol@hospital.com"
Private Const HOSPITAL_NAME As String = "Hospital Name"
Private Const COOLDOWN_DAYS As Integer = 7
Private Const CACHE_SHEET As String = "_AlertCache"
' ------------------------------------

Public Sub CheckN95FitTests()
    Dim ws As Worksheet
    Dim cache As Worksheet
    Dim sentCount As Long, skippedCooldown As Long, errorCount As Long
    Dim lastRow As Long, i As Long
    Dim status As String, empName As String, empId As String
    Dim deptEmail As String, maskType As String, dept As String
    Dim expDate As Date
    Dim cacheKey As String, lastSent As Date
    Dim excludeList As String

    excludeList = "|Dashboard|README|" & CACHE_SHEET & "|Settings|Logs|"

    ' Get or create the hidden cache sheet
    Set cache = GetOrCreateCacheSheet()

    sentCount = 0
    skippedCooldown = 0
    errorCount = 0

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, excludeList, "|" & ws.Name & "|", vbTextCompare) = 0 Then
            On Error Resume Next
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            On Error GoTo 0
            If lastRow < 3 Then GoTo NextSheet

            For i = 3 To lastRow
                empName = SafeStr(ws.Cells(i, 1).Value)
                If Len(empName) > 0 Then
                    status = SafeStr(ws.Cells(i, 7).Value)
                    If status = "Expiring Soon" Then
                        empId = SafeStr(ws.Cells(i, 2).Value)
                        deptEmail = SafeStr(ws.Cells(i, 3).Value)
                        maskType = SafeStr(ws.Cells(i, 4).Value)
                        dept = ws.Name

                        If IsDate(ws.Cells(i, 6).Value) Then
                            expDate = ws.Cells(i, 6).Value
                        Else
                            expDate = 0
                        End If

                        cacheKey = dept & "|" & empId & "|" & Format(expDate, "yyyy-mm-dd")
                        lastSent = GetCacheEntry(cache, cacheKey)

                        If COOLDOWN_DAYS > 0 And lastSent > 0 _
                           And (Now - lastSent) < COOLDOWN_DAYS Then
                            skippedCooldown = skippedCooldown + 1
                        Else
                            If SendAlertEmail(empName, empId, deptEmail, _
                                              maskType, expDate, dept) Then
                                SetCacheEntry cache, cacheKey, Now
                                sentCount = sentCount + 1
                            Else
                                errorCount = errorCount + 1
                            End If
                        End If
                    End If
                End If
            Next i
        End If
NextSheet:
    Next ws

    Application.ScreenUpdating = True

    LogToCache cache, "Run @ " & Format(Now, "yyyy-mm-dd hh:nn") & _
                       " | sent: " & sentCount & _
                       " | cooldown: " & skippedCooldown & _
                       " | errors: " & errorCount

    MsgBox "N95 daily check complete." & vbCrLf & vbCrLf & _
           "Sent: " & sentCount & vbCrLf & _
           "Cooldown skipped: " & skippedCooldown & vbCrLf & _
           "Errors: " & errorCount, _
           vbInformation, "N95 Tracker"
End Sub

Private Function SendAlertEmail(empName As String, empId As String, _
                                deptEmail As String, maskType As String, _
                                expDate As Date, dept As String) As Boolean
    Dim olApp As Object, olMail As Object
    Dim recipients As String, subject As String, body As String
    Dim daysRemaining As Long

    On Error GoTo Fail

    If Len(deptEmail) = 0 Or InStr(deptEmail, "@") = 0 Then
        Debug.Print "Skipped " & empName & " (" & dept & "): missing/invalid email"
        SendAlertEmail = False
        Exit Function
    End If

    daysRemaining = CLng(expDate - Date)
    If daysRemaining < 0 Then daysRemaining = 0

    recipients = deptEmail
    If Len(INFECTION_CONTROL_EMAIL) > 0 Then
        recipients = recipients & ";" & INFECTION_CONTROL_EMAIL
    End If

    subject = "N95 Fit Test Expiring Soon - " & empName & " (" & dept & ")"

    body = BuildEmailBody(empName, empId, deptEmail, maskType, _
                          expDate, dept, daysRemaining)

    Set olApp = GetOutlookApp()
    Set olMail = olApp.CreateItem(0) ' olMailItem = 0
    With olMail
        .To = recipients
        .Subject = subject
        .HTMLBody = body
        .Send
    End With

    Debug.Print "Sent: " & empName & " (" & dept & ")"
    SendAlertEmail = True
    Exit Function
Fail:
    Debug.Print "ERROR for " & empName & ": " & Err.Description
    SendAlertEmail = False
End Function

Private Function BuildEmailBody(empName As String, empId As String, _
                                deptEmail As String, maskType As String, _
                                expDate As Date, dept As String, _
                                daysRemaining As Long) As String
    Dim s As String
    s = "<div style='font-family:Arial,Helvetica,sans-serif;max-width:620px;" & _
        "margin:0 auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;'>"
    s = s & "<div style='background-color:#f39c12;color:#ffffff;" & _
            "padding:20px;text-align:center;'>"
    s = s & "<h2 style='margin:0;font-size:20px;'>" & _
            "N95 Fit Test Expiration Alert</h2>"
    s = s & "<p style='margin:6px 0 0;font-size:13px;'>" & _
            EscapeHtml(HOSPITAL_NAME) & " &mdash; Infection Control</p></div>"
    s = s & "<div style='padding:24px;background-color:#ffffff;color:#333;'>"
    s = s & "<p style='font-size:15px;margin:0 0 12px;'>" & _
            "Dear Department Lead and Infection Control Team,</p>"
    s = s & "<p style='font-size:15px;margin:0 0 16px;'>" & _
            "This is an automated reminder that the N95 mask fit test for " & _
            "the following employee is " & _
            "<strong style='color:#e67e22;'>expiring within " & daysRemaining & _
            " day(s)</strong>. Please schedule a re-test to ensure continued " & _
            "compliance and staff safety.</p>"
    s = s & "<table style='width:100%;border-collapse:collapse;" & _
            "margin:18px 0;font-size:14px;'>"
    s = s & RowHtml("Employee Name", empName, True)
    s = s & RowHtml("Employee ID", empId, False)
    s = s & RowHtml("Department", dept, True)
    s = s & RowHtml("N95 Mask Type", maskType, False)
    s = s & "<tr style='background-color:#fff3cd;'>" & _
            "<td style='padding:12px;border:1px solid #dee2e6;font-weight:bold;width:40%;'>" & _
            "Expiration Date</td>" & _
            "<td style='padding:12px;border:1px solid #dee2e6;color:#856404;font-weight:bold;'>" & _
            EscapeHtml(Format(expDate, "dd mmm yyyy")) & "</td></tr>"
    s = s & "</table>"
    s = s & "<p style='font-size:14px;background-color:#e7f3fe;padding:12px;" & _
            "border-left:4px solid #2196F3;border-radius:4px;margin:0 0 12px;'>" & _
            "<strong>Action Required:</strong> Coordinate with Infection Control to " & _
            "re-schedule the fit test before the expiration date.</p>"
    s = s & "<p style='font-size:12px;color:#999;margin-top:24px;" & _
            "border-top:1px solid #eee;padding-top:12px;'>" & _
            "Automated notification from the N95 Fit Test Tracking System. " & _
            "Sent on " & Format(Now, "dd mmm yyyy hh:nn") & ".</p>"
    s = s & "</div></div>"
    BuildEmailBody = s
End Function

Private Function RowHtml(label As String, value As String, alt As Boolean) As String
    Dim bg As String
    If alt Then bg = "#f8f9fa" Else bg = "#ffffff"
    RowHtml = "<tr style='background-color:" & bg & ";'>" & _
              "<td style='padding:12px;border:1px solid #dee2e6;font-weight:bold;width:40%;'>" & _
              EscapeHtml(label) & "</td>" & _
              "<td style='padding:12px;border:1px solid #dee2e6;'>" & _
              EscapeHtml(IIf(Len(value) = 0, "-", value)) & "</td></tr>"
End Function

Private Function GetOutlookApp() As Object
    Dim ol As Object
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then
        Set ol = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    Set GetOutlookApp = ol
End Function

Private Function GetOrCreateCacheSheet() As Worksheet
    Dim cache As Worksheet
    On Error Resume Next
    Set cache = ThisWorkbook.Sheets(CACHE_SHEET)
    On Error GoTo 0
    If cache Is Nothing Then
        Set cache = ThisWorkbook.Sheets.Add(After:= _
            ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        cache.Name = CACHE_SHEET
        cache.Cells(1, 1).Value = "Key"
        cache.Cells(1, 2).Value = "LastSent"
        cache.Cells(1, 3).Value = "RunLog"
        cache.Visible = xlSheetVeryHidden
    End If
    Set GetOrCreateCacheSheet = cache
End Function

Private Function GetCacheEntry(cache As Worksheet, key As String) As Date
    Dim lastRow As Long, i As Long
    GetCacheEntry = 0
    lastRow = cache.Cells(cache.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If cache.Cells(i, 1).Value = key Then
            If IsDate(cache.Cells(i, 2).Value) Then
                GetCacheEntry = cache.Cells(i, 2).Value
            End If
            Exit Function
        End If
    Next i
End Function

Private Sub SetCacheEntry(cache As Worksheet, key As String, value As Date)
    Dim lastRow As Long, i As Long
    lastRow = cache.Cells(cache.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If cache.Cells(i, 1).Value = key Then
            cache.Cells(i, 2).Value = value
            Exit Sub
        End If
    Next i
    cache.Cells(lastRow + 1, 1).Value = key
    cache.Cells(lastRow + 1, 2).Value = value
End Sub

Private Sub LogToCache(cache As Worksheet, msg As String)
    Dim lastRow As Long
    lastRow = cache.Cells(cache.Rows.Count, 3).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    cache.Cells(lastRow + 1, 3).Value = msg
End Sub

Private Function SafeStr(v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim(CStr(v))
    End If
End Function

Private Function EscapeHtml(s As String) As String
    Dim r As String
    r = Replace(s, "&", "&amp;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, """", "&quot;")
    r = Replace(r, "'", "&#39;")
    EscapeHtml = r
End Function

' --- Optional helpers ---

Public Sub SendSelfTest()
    ' Sends a simple test email to your own Outlook address to verify permissions.
    Dim olApp As Object, olMail As Object
    On Error GoTo Fail
    Set olApp = GetOutlookApp()
    Set olMail = olApp.CreateItem(0)
    With olMail
        .To = olApp.Session.CurrentUser.Address
        .Subject = "N95 Tracker - self-test OK"
        .HTMLBody = "<p>Outlook + VBA email permissions are working correctly. " & _
                    "You can now run CheckN95FitTests().</p>"
        .Send
    End With
    MsgBox "Self-test email sent. Check your Outlook inbox.", vbInformation
    Exit Sub
Fail:
    MsgBox "Self-test failed: " & Err.Description, vbCritical
End Sub

Public Sub ResetAlertCache()
    Dim cache As Worksheet
    Set cache = GetOrCreateCacheSheet()
    cache.Cells.Clear
    cache.Cells(1, 1).Value = "Key"
    cache.Cells(1, 2).Value = "LastSent"
    cache.Cells(1, 3).Value = "RunLog"
    MsgBox "Alert cache cleared.", vbInformation
End Sub

Public Sub TestRun()
    ' Manual run that pops a result message box.
    CheckN95FitTests
End Sub
