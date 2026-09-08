import os
import sys
import glob
import datetime
import win32com.client
import win32com.client.dynamic
import pandas as pd
import json

# Expected standard time slots for renewals reporting
TIME_SLOTS = ["06:12", "08:12", "10:12", "12:12", "14:12", "16:12", "18:12", "20:12", "22:12"]

def parse_subject_datetime(subject):
    """
    Parses the subject to extract datetime if it matches the hourly status pattern.
    Pattern: 'Hourly Status Update - YYYY-MM-DD HH:MM:SS'
    """
    prefix = "Hourly Status Update - "
    if not subject.startswith(prefix):
        return None
    try:
        dt_str = subject[len(prefix):].strip()
        dt = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
        return dt
    except Exception:
        return None

def get_attachment(filtered_messages, target_date, target_time_slot, temp_dir, keyword, dest_prefix):
    """
    Finds the email corresponding to the target date and time slot, extracts the attachment
    matching `keyword`, and saves it to a unique file in `temp_dir`.
    """
    subject_pattern = f"Hourly Status Update - {target_date} {target_time_slot}:"
    for msg in filtered_messages:
        try:
            if subject_pattern in msg.Subject:
                attachments = msg.Attachments
                for i in range(1, attachments.Count + 1):
                    att = attachments.Item(i)
                    fname = att.FileName.lower()
                    matched = False
                    if keyword == "partner":
                        if "partnerwise_test" in fname or "recent_renewal_partner" in fname or "recent_renew_partner" in fname:
                            matched = True
                    elif keyword == "outlet":
                        if "outletwise" in fname or "recent_renewal_outlet" in fname:
                            matched = True
                    elif keyword in fname:
                        matched = True
                        
                    if matched:
                        safe_time = target_time_slot.replace(":", "-")
                        dest_file = f"{dest_prefix}_{target_date}_{safe_time}.csv"
                        dest_path = os.path.join(temp_dir, dest_file)
                        att.SaveAsFile(dest_path)
                        return dest_path
        except Exception:
            pass
    return None

def create_macro_dashboard(excel_path, macro_path):
    """
    Automates Excel via COM to build a two-sheet interactive macro workbook:
    1. 'Partnerwise': 12-column table with Region & Partner dropdowns, powered by recent_renewal_partnerwise_test.csv.
    2. 'Outletwise': 13-column table with Region & Partner dropdowns, powered by recent_renewal_outletwise.csv.
    3. Hidden 'Data_Partner' and 'Data_Outlet' sheets.
    """
    print("Connecting to Excel COM to build two-sheet macro dashboard...")
    excel = win32com.client.dynamic.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        print(f"Loading raw workbook: {excel_path}")
        wb = excel.Workbooks.Open(excel_path)
        
        # Reference raw data sheets
        raw_partner = wb.Sheets("Data_Partner")
        raw_outlet = wb.Sheets("Data_Outlet")
        
        # 1. Create Partnerwise sheet as the first sheet
        print("Creating 'Partnerwise' UI sheet...")
        dash_partner = wb.Sheets.Add(Before=raw_partner)
        dash_partner.Name = "Partnerwise"
        
        # Title block
        dash_partner.Range("A1:L2").Merge()
        dash_partner.Range("A1").Value = "Partner Renewals Performance Dashboard"
        title_p = dash_partner.Range("A1")
        title_p.Font.Bold = True
        title_p.Font.Name = "Calibri"
        title_p.Font.Size = 16
        title_p.Font.Color = 0xFFFFFF
        title_p.Interior.Color = 0x784E1F  # Navy Blue
        title_p.HorizontalAlignment = -4108
        title_p.VerticalAlignment = -4108
            
        dash_partner.Rows(1).RowHeight = 25
        dash_partner.Rows(2).RowHeight = 25
        
        # Dropdown labels
        dash_partner.Range("B4").Value = "Region:"
        dash_partner.Range("B4").Font.Bold = True
        dash_partner.Range("B4").HorizontalAlignment = -4152
        
        dash_partner.Range("B5").Value = "Partner:"
        dash_partner.Range("B5").Font.Bold = True
        dash_partner.Range("B5").HorizontalAlignment = -4152
        
        for cell_name in ["C4", "C5"]:
            c = dash_partner.Range(cell_name)
            c.Interior.Color = 0xF2F2F2
            c.Borders.LineStyle = 1
            c.Borders.Weight = 2
            
        dash_partner.Columns("A").ColumnWidth = 5
        dash_partner.Columns("B").ColumnWidth = 16
        dash_partner.Columns("C").ColumnWidth = 38
        
        btn_p_reset = dash_partner.Buttons().Add(330, 48, 100, 24)
        btn_p_reset.OnAction = "ResetPartnerFilters"
        btn_p_reset.Caption = "Reset Filters"
        
        btn_p_refresh = dash_partner.Buttons().Add(440, 48, 130, 24)
        btn_p_refresh.OnAction = "DisplayPartnerwise"
        btn_p_refresh.Caption = "Refresh Partnerwise"

        # 2. Create Outletwise sheet
        print("Creating 'Outletwise' UI sheet...")
        dash_outlet = wb.Sheets.Add(After=dash_partner)
        dash_outlet.Name = "Outletwise"
        
        # Title block
        dash_outlet.Range("A1:M2").Merge()
        dash_outlet.Range("A1").Value = "Outlet Renewals Performance Dashboard"
        title_o = dash_outlet.Range("A1")
        title_o.Font.Bold = True
        title_o.Font.Name = "Calibri"
        title_o.Font.Size = 16
        title_o.Font.Color = 0xFFFFFF
        title_o.Interior.Color = 0x784E1F  # Navy Blue
        title_o.HorizontalAlignment = -4108
        title_o.VerticalAlignment = -4108
            
        dash_outlet.Rows(1).RowHeight = 25
        dash_outlet.Rows(2).RowHeight = 25
        
        # Dropdown labels
        dash_outlet.Range("B4").Value = "Region:"
        dash_outlet.Range("B4").Font.Bold = True
        dash_outlet.Range("B4").HorizontalAlignment = -4152
        
        dash_outlet.Range("B5").Value = "Partner:"
        dash_outlet.Range("B5").Font.Bold = True
        dash_outlet.Range("B5").HorizontalAlignment = -4152
        
        for cell_name in ["C4", "C5"]:
            c = dash_outlet.Range(cell_name)
            c.Interior.Color = 0xF2F2F2
            c.Borders.LineStyle = 1
            c.Borders.Weight = 2
            
        dash_outlet.Columns("A").ColumnWidth = 5
        dash_outlet.Columns("B").ColumnWidth = 16
        dash_outlet.Columns("C").ColumnWidth = 38
        
        btn_o_reset = dash_outlet.Buttons().Add(330, 48, 100, 24)
        btn_o_reset.OnAction = "ResetOutletFilters"
        btn_o_reset.Caption = "Reset Filters"
        
        btn_o_refresh = dash_outlet.Buttons().Add(440, 48, 130, 24)
        btn_o_refresh.OnAction = "DisplayOutletwise"
        btn_o_refresh.Caption = "Refresh Outletwise"
        
        # 3. Inject VBA standard module
        print("Injecting VBA macro subroutines...")
        module_code = """
' -------------------------------------------------------------
' PARTNERWISE MACROS
' -------------------------------------------------------------
Sub UpdatePartnerDropdownPartnerwise(ByVal selectedRegion As String)
    Dim dataSheet As Worksheet
    Dim dashSheet As Worksheet
    Dim partnerCell As Range
    Dim lastRow As Long
    Dim partnerDict As Object
    Dim i As Long
    Dim rVal As String
    Dim pVal As String
    Dim key As Variant
    Dim writeIdx As Long
    Dim lastPartnerRow As Long
    
    Set dataSheet = ThisWorkbook.Sheets("Data_Partner")
    Set dashSheet = ThisWorkbook.Sheets("Partnerwise")
    Set partnerCell = dashSheet.Range("C5")
    
    partnerCell.Validation.Delete
    dataSheet.Range("U:U").Clear
    
    lastRow = dataSheet.UsedRange.Rows.Count
    If lastRow <= 1 Then
        lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    End If
    
    Set partnerDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        rVal = Trim(dataSheet.Cells(i, 1).Value)
        pVal = Trim(dataSheet.Cells(i, 2).Value)
        If selectedRegion = "" Or LCase(rVal) = LCase(selectedRegion) Then
            If pVal <> "" Then partnerDict(pVal) = True
        End If
    Next i
    
    If partnerDict.Count = 0 Then Exit Sub
    
    writeIdx = 2
    For Each key In partnerDict.Keys
        dataSheet.Cells(writeIdx, 21).Value = key
        writeIdx = writeIdx + 1
    Next key
    
    lastPartnerRow = writeIdx - 1
    
    With partnerCell.Validation
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=Data_Partner!$U$2:$U$" & lastPartnerRow
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Sub InitializePartnerValidation()
    Dim dataSheet As Worksheet
    Dim dashSheet As Worksheet
    Dim regionCell As Range
    Dim lastRow As Long
    Dim regionDict As Object
    Dim i As Long
    Dim rVal As String
    Dim key As Variant
    Dim writeIdx As Long
    Dim lastRegionRow As Long
    
    Set dataSheet = ThisWorkbook.Sheets("Data_Partner")
    Set dashSheet = ThisWorkbook.Sheets("Partnerwise")
    Set regionCell = dashSheet.Range("C4")
    
    regionCell.Validation.Delete
    dataSheet.Range("T:T").Clear
    
    lastRow = dataSheet.UsedRange.Rows.Count
    If lastRow <= 1 Then
        lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    End If
    
    Set regionDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        rVal = Trim(dataSheet.Cells(i, 1).Value)
        If rVal <> "" And LCase(rVal) <> "total" And LCase(rVal) <> "all" And LCase(rVal) <> "region" Then
            regionDict(rVal) = True
        End If
    Next i
    
    If regionDict.Count = 0 Then Exit Sub
    
    writeIdx = 2
    For Each key In regionDict.Keys
        dataSheet.Cells(writeIdx, 20).Value = key
        writeIdx = writeIdx + 1
    Next key
    
    lastRegionRow = writeIdx - 1
    
    With regionCell.Validation
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=Data_Partner!$T$2:$T$" & lastRegionRow
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Call UpdatePartnerDropdownPartnerwise("")
End Sub

Sub DisplayPartnerwise()
    Dim dataSheet As Worksheet
    Dim dashSheet As Worksheet
    Dim regionFilter As String
    Dim partnerFilter As String
    Dim lastRowData As Long
    Dim colIdx As Long
    Dim headers(1 To 12) As String
    Dim writeRow As Long
    Dim i As Long
    Dim rVal As String
    Dim pVal As String
    Dim matchRow As Boolean
    Dim sumToday As Long
    Dim sumPrevToday As Long
    Dim sumYestSame As Long
    Dim sumYest2hr As Long
    Dim sumYest10pm As Long
    Dim sumTarget As Long
    Dim has2hr As Boolean
    Dim dataRange As Range
    Dim rowIdx As Long
    Dim valGap As Variant
    Dim valNeeded As Variant
    Dim valBtd As Variant
    Dim valLast2h As Variant
    
    Set dataSheet = ThisWorkbook.Sheets("Data_Partner")
    Set dashSheet = ThisWorkbook.Sheets("Partnerwise")
    
    dashSheet.Range("A8:L10000").Clear
    
    regionFilter = Trim(dashSheet.Range("C4").Value)
    partnerFilter = Trim(dashSheet.Range("C5").Value)
    
    lastRowData = dataSheet.UsedRange.Rows.Count
    If lastRowData <= 1 Then
        lastRowData = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    headers(1) = "Region"
    headers(2) = "Partner"
    headers(3) = "Renewals today"
    headers(4) = "Renewals in last 2 hours"
    headers(5) = "Renewals Yesterday - Same time"
    headers(6) = "Renewal Gap - Same time"
    headers(7) = "Renewals yesterday - 2 hrs later"
    headers(8) = "Renewals needed in 2 hours"
    headers(9) = "Renewals Yesterday - 10:12PM"
    headers(10) = "Target"
    headers(11) = "BTD"
    headers(12) = "Achieved %"
    
    For colIdx = 1 To 12
        dashSheet.Cells(8, colIdx).Value = headers(colIdx)
    Next colIdx
    
    With dashSheet.Range("A8:L8")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = True
    End With
    dashSheet.Rows(8).RowHeight = 48
    
    writeRow = 9
    sumToday = 0
    sumPrevToday = 0
    sumYestSame = 0
    sumYest2hr = 0
    sumYest10pm = 0
    sumTarget = 0
    has2hr = False
    
    For i = 2 To lastRowData
        rVal = Trim(dataSheet.Cells(i, 1).Value)
        pVal = Trim(dataSheet.Cells(i, 2).Value)
        matchRow = True
        If regionFilter <> "" And LCase(rVal) <> LCase(regionFilter) Then matchRow = False
        If partnerFilter <> "" And LCase(pVal) <> LCase(partnerFilter) Then matchRow = False
        
        If matchRow Then
            dashSheet.Cells(writeRow, 1).Value = dataSheet.Cells(i, 1).Value
            dashSheet.Cells(writeRow, 2).Value = dataSheet.Cells(i, 2).Value
            dashSheet.Cells(writeRow, 3).Value = dataSheet.Cells(i, 3).Value
            dashSheet.Cells(writeRow, 4).Value = dataSheet.Cells(i, 4).Value
            dashSheet.Cells(writeRow, 5).Value = dataSheet.Cells(i, 5).Value
            dashSheet.Cells(writeRow, 6).Value = dataSheet.Cells(i, 6).Value
            dashSheet.Cells(writeRow, 7).Value = dataSheet.Cells(i, 7).Value
            dashSheet.Cells(writeRow, 8).Value = dataSheet.Cells(i, 8).Value
            dashSheet.Cells(writeRow, 9).Value = dataSheet.Cells(i, 9).Value
            dashSheet.Cells(writeRow, 10).Value = dataSheet.Cells(i, 10).Value
            dashSheet.Cells(writeRow, 11).Value = dataSheet.Cells(i, 11).Value
            
            ' Achieved %
            If Val(dataSheet.Cells(i, 10).Value) > 0 Then
                dashSheet.Cells(writeRow, 12).Value = Val(dataSheet.Cells(i, 3).Value) / Val(dataSheet.Cells(i, 10).Value)
            Else
                dashSheet.Cells(writeRow, 12).Value = 0
            End If
            dashSheet.Cells(writeRow, 12).NumberFormat = "0%"
            
            sumToday = sumToday + Val(dataSheet.Cells(i, 3).Value)
            sumPrevToday = sumPrevToday + Val(dataSheet.Cells(i, 12).Value)
            sumYestSame = sumYestSame + Val(dataSheet.Cells(i, 5).Value)
            
            If Not IsEmpty(dataSheet.Cells(i, 7).Value) And dataSheet.Cells(i, 7).Value <> "" Then
                sumYest2hr = sumYest2hr + Val(dataSheet.Cells(i, 7).Value)
                has2hr = True
            End If
            
            sumYest10pm = sumYest10pm + Val(dataSheet.Cells(i, 9).Value)
            sumTarget = sumTarget + Val(dataSheet.Cells(i, 10).Value)
            writeRow = writeRow + 1
        End If
    Next i
    
    If writeRow = 9 Then
        dashSheet.Cells(9, 1).Value = "No data matched the selected filters."
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If
    
    dashSheet.Cells(writeRow, 1).Value = "Total"
    dashSheet.Cells(writeRow, 2).Value = ""
    dashSheet.Cells(writeRow, 3).Value = sumToday
    dashSheet.Cells(writeRow, 4).Value = sumToday - sumPrevToday
    dashSheet.Cells(writeRow, 5).Value = sumYestSame
    dashSheet.Cells(writeRow, 6).Value = sumToday - sumYestSame
    
    If has2hr Then
        dashSheet.Cells(writeRow, 7).Value = sumYest2hr
        dashSheet.Cells(writeRow, 8).Value = sumYest2hr - sumToday
    Else
        dashSheet.Cells(writeRow, 7).Value = ""
        dashSheet.Cells(writeRow, 8).Value = ""
    End If
    
    dashSheet.Cells(writeRow, 9).Value = sumYest10pm
    dashSheet.Cells(writeRow, 10).Value = sumTarget
    dashSheet.Cells(writeRow, 11).Value = sumTarget - sumToday
    
    If sumTarget > 0 Then
        dashSheet.Cells(writeRow, 12).Value = sumToday / sumTarget
    Else
        dashSheet.Cells(writeRow, 12).Value = 0
    End If
    dashSheet.Cells(writeRow, 12).NumberFormat = "0%"
    
    Set dataRange = dashSheet.Range(dashSheet.Cells(9, 1), dashSheet.Cells(writeRow, 12))
    With dataRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Borders.LineStyle = 1
        .Borders.Weight = 2
        .Borders.Color = RGB(217, 217, 217)
    End With
    
    For rowIdx = 9 To writeRow
        dashSheet.Rows(rowIdx).RowHeight = 20
        For colIdx = 1 To 12
            If colIdx >= 3 Then
                dashSheet.Cells(rowIdx, colIdx).HorizontalAlignment = -4108
            Else
                dashSheet.Cells(rowIdx, colIdx).HorizontalAlignment = -4131
            End If
            dashSheet.Cells(rowIdx, colIdx).VerticalAlignment = -4108
        Next colIdx
        
        valLast2h = dashSheet.Cells(rowIdx, 4).Value
        valGap = dashSheet.Cells(rowIdx, 6).Value
        valNeeded = dashSheet.Cells(rowIdx, 8).Value
        valBtd = dashSheet.Cells(rowIdx, 11).Value
        
        If IsNumeric(valLast2h) And valLast2h = 0 Then
            dashSheet.Cells(rowIdx, 4).Interior.Color = RGB(255, 120, 120)
            dashSheet.Cells(rowIdx, 4).Font.Bold = True
        End If
        
        If IsNumeric(valGap) And valGap < 0 Then
            dashSheet.Cells(rowIdx, 6).Interior.Color = RGB(255, 120, 120)
            dashSheet.Cells(rowIdx, 6).Font.Bold = True
        End If
        
        If IsNumeric(valNeeded) And valNeeded > 0 Then
            dashSheet.Cells(rowIdx, 8).Interior.Color = RGB(255, 120, 120)
            dashSheet.Cells(rowIdx, 8).Font.Bold = True
        End If
        
        If IsNumeric(valBtd) Then
            If valBtd > 2 Then
                dashSheet.Cells(rowIdx, 11).Interior.Color = RGB(255, 120, 120)
                dashSheet.Cells(rowIdx, 11).Font.Bold = True
            ElseIf valBtd > 0 And valBtd <= 2 Then
                dashSheet.Cells(rowIdx, 11).Interior.Color = RGB(217, 225, 242)
                dashSheet.Cells(rowIdx, 11).Font.Bold = True
            Else
                dashSheet.Cells(rowIdx, 11).Interior.Color = RGB(226, 239, 218)
                dashSheet.Cells(rowIdx, 11).Font.Bold = True
            End If
        End If
    Next rowIdx
    
    With dashSheet.Range(dashSheet.Cells(writeRow, 1), dashSheet.Cells(writeRow, 12))
        .Font.Bold = True
        .Borders(8).LineStyle = -4119
        .Borders(9).LineStyle = -4119
        .Borders(8).Color = RGB(71, 85, 105)
        .Borders(9).Color = RGB(71, 85, 105)
        .Interior.Color = RGB(248, 250, 252)
    End With
    
    dashSheet.Columns("A").AutoFit
    dashSheet.Columns("B").AutoFit
    If dashSheet.Columns("A").ColumnWidth < 12 Then dashSheet.Columns("A").ColumnWidth = 12
    If dashSheet.Columns("B").ColumnWidth < 12 Then dashSheet.Columns("B").ColumnWidth = 12
    
    For colIdx = 3 To 12
        dashSheet.Columns(Chr(64 + colIdx)).ColumnWidth = 11
    Next colIdx
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub ResetPartnerFilters()
    Dim dashSheet As Worksheet
    Set dashSheet = ThisWorkbook.Sheets("Partnerwise")
    Application.EnableEvents = False
    dashSheet.Range("C4").Value = ""
    dashSheet.Range("C5").Value = ""
    Application.EnableEvents = True
    Call UpdatePartnerDropdownPartnerwise("")
    Call DisplayPartnerwise
End Sub

' -------------------------------------------------------------
' OUTLETWISE MACROS
' -------------------------------------------------------------
Sub UpdatePartnerDropdownOutletwise(ByVal selectedRegion As String)
    Dim dataSheet As Worksheet
    Dim dashSheet As Worksheet
    Dim partnerCell As Range
    Dim lastRow As Long
    Dim partnerDict As Object
    Dim i As Long
    Dim rVal As String
    Dim pVal As String
    Dim oVal As String
    Dim key As Variant
    Dim writeIdx As Long
    Dim lastPartnerRow As Long
    
    Set dataSheet = ThisWorkbook.Sheets("Data_Outlet")
    Set dashSheet = ThisWorkbook.Sheets("Outletwise")
    Set partnerCell = dashSheet.Range("C5")
    
    partnerCell.Validation.Delete
    dataSheet.Range("U:U").Clear
    
    lastRow = dataSheet.UsedRange.Rows.Count
    If lastRow <= 1 Then
        lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    End If
    
    Set partnerDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        rVal = Trim(dataSheet.Cells(i, 1).Value)
        pVal = Trim(dataSheet.Cells(i, 2).Value)
        oVal = Trim(dataSheet.Cells(i, 3).Value)
        If oVal <> "" And LCase(oVal) <> "none" And LCase(oVal) <> "nan" Then
            If selectedRegion = "" Or LCase(rVal) = LCase(selectedRegion) Then
                If pVal <> "" Then partnerDict(pVal) = True
            End If
        End If
    Next i
    
    If partnerDict.Count = 0 Then Exit Sub
    
    writeIdx = 2
    For Each key In partnerDict.Keys
        dataSheet.Cells(writeIdx, 21).Value = key
        writeIdx = writeIdx + 1
    Next key
    
    lastPartnerRow = writeIdx - 1
    
    With partnerCell.Validation
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=Data_Outlet!$U$2:$U$" & lastPartnerRow
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Sub InitializeOutletValidation()
    Dim dataSheet As Worksheet
    Dim dashSheet As Worksheet
    Dim regionCell As Range
    Dim lastRow As Long
    Dim regionDict As Object
    Dim i As Long
    Dim rVal As String
    Dim oVal As String
    Dim key As Variant
    Dim writeIdx As Long
    Dim lastRegionRow As Long
    
    Set dataSheet = ThisWorkbook.Sheets("Data_Outlet")
    Set dashSheet = ThisWorkbook.Sheets("Outletwise")
    Set regionCell = dashSheet.Range("C4")
    
    regionCell.Validation.Delete
    dataSheet.Range("T:T").Clear
    
    lastRow = dataSheet.UsedRange.Rows.Count
    If lastRow <= 1 Then
        lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    End If
    
    Set regionDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        rVal = Trim(dataSheet.Cells(i, 1).Value)
        oVal = Trim(dataSheet.Cells(i, 3).Value)
        If oVal <> "" And LCase(oVal) <> "none" And LCase(oVal) <> "nan" Then
            If rVal <> "" And LCase(rVal) <> "total" And LCase(rVal) <> "region" Then
                regionDict(rVal) = True
            End If
        End If
    Next i
    
    If regionDict.Count = 0 Then Exit Sub
    
    writeIdx = 2
    For Each key In regionDict.Keys
        dataSheet.Cells(writeIdx, 20).Value = key
        writeIdx = writeIdx + 1
    Next key
    
    lastRegionRow = writeIdx - 1
    
    With regionCell.Validation
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=Data_Outlet!$T$2:$T$" & lastRegionRow
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Call UpdatePartnerDropdownOutletwise("")
End Sub

Sub DisplayOutletwise()
    Dim dataSheet As Worksheet
    Dim dashSheet As Worksheet
    Dim regionFilter As String
    Dim partnerFilter As String
    Dim lastRowData As Long
    Dim colIdx As Long
    Dim headers(1 To 13) As String
    Dim writeRow As Long
    Dim i As Long
    Dim rVal As String
    Dim pVal As String
    Dim oVal As String
    Dim matchRow As Boolean
    Dim sumToday As Long
    Dim sumPrevToday As Long
    Dim sumYestSame As Long
    Dim sumYest2hr As Long
    Dim sumYest10pm As Long
    Dim sumTarget As Long
    Dim has2hr As Boolean
    Dim dataRange As Range
    Dim rowIdx As Long
    Dim valGap As Variant
    Dim valNeeded As Variant
    Dim valBtd As Variant
    Dim valLast2h As Variant
    
    Set dataSheet = ThisWorkbook.Sheets("Data_Outlet")
    Set dashSheet = ThisWorkbook.Sheets("Outletwise")
    
    dashSheet.Range("A8:M10000").Clear
    
    regionFilter = Trim(dashSheet.Range("C4").Value)
    partnerFilter = Trim(dashSheet.Range("C5").Value)
    
    lastRowData = dataSheet.UsedRange.Rows.Count
    If lastRowData <= 1 Then
        lastRowData = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    headers(1) = "region"
    headers(2) = "Partner"
    headers(3) = "Outlet"
    headers(4) = "Renewals today"
    headers(5) = "Renewals in last 2 hours"
    headers(6) = "Renewals Yesterday - Same time"
    headers(7) = "Renewal Gap - Same time"
    headers(8) = "Renewals yesterday - 2 hrs later"
    headers(9) = "Renewals needed in 2 hours"
    headers(10) = "Renewals Yesterday - 10:12PM"
    headers(11) = "Target"
    headers(12) = "BTD"
    headers(13) = "Achieved %"
    
    For colIdx = 1 To 13
        dashSheet.Cells(8, colIdx).Value = headers(colIdx)
    Next colIdx
    
    With dashSheet.Range("A8:M8")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = True
    End With
    dashSheet.Rows(8).RowHeight = 48
    
    writeRow = 9
    sumToday = 0
    sumPrevToday = 0
    sumYestSame = 0
    sumYest2hr = 0
    sumYest10pm = 0
    sumTarget = 0
    has2hr = False
    
    For i = 2 To lastRowData
        rVal = Trim(dataSheet.Cells(i, 1).Value)
        pVal = Trim(dataSheet.Cells(i, 2).Value)
        oVal = Trim(dataSheet.Cells(i, 3).Value)
        matchRow = True
        If oVal = "" Or LCase(oVal) = "none" Or LCase(oVal) = "nan" Then matchRow = False
        If regionFilter <> "" And LCase(rVal) <> LCase(regionFilter) Then matchRow = False
        If partnerFilter <> "" And LCase(pVal) <> LCase(partnerFilter) Then matchRow = False
        
        If matchRow Then
            dashSheet.Cells(writeRow, 1).Value = dataSheet.Cells(i, 1).Value
            dashSheet.Cells(writeRow, 2).Value = dataSheet.Cells(i, 2).Value
            dashSheet.Cells(writeRow, 3).Value = dataSheet.Cells(i, 3).Value
            dashSheet.Cells(writeRow, 4).Value = dataSheet.Cells(i, 4).Value
            dashSheet.Cells(writeRow, 5).Value = dataSheet.Cells(i, 5).Value
            dashSheet.Cells(writeRow, 6).Value = dataSheet.Cells(i, 6).Value
            dashSheet.Cells(writeRow, 7).Value = dataSheet.Cells(i, 7).Value
            dashSheet.Cells(writeRow, 8).Value = dataSheet.Cells(i, 8).Value
            dashSheet.Cells(writeRow, 9).Value = dataSheet.Cells(i, 9).Value
            dashSheet.Cells(writeRow, 10).Value = dataSheet.Cells(i, 10).Value
            dashSheet.Cells(writeRow, 11).Value = dataSheet.Cells(i, 11).Value
            dashSheet.Cells(writeRow, 12).Value = dataSheet.Cells(i, 12).Value
            
            If Val(dataSheet.Cells(i, 11).Value) > 0 Then
                dashSheet.Cells(writeRow, 13).Value = Val(dataSheet.Cells(i, 4).Value) / Val(dataSheet.Cells(i, 11).Value)
            Else
                dashSheet.Cells(writeRow, 13).Value = 0
            End If
            dashSheet.Cells(writeRow, 13).NumberFormat = "0%"
            
            sumToday = sumToday + Val(dataSheet.Cells(i, 4).Value)
            sumPrevToday = sumPrevToday + Val(dataSheet.Cells(i, 13).Value)
            sumYestSame = sumYestSame + Val(dataSheet.Cells(i, 6).Value)
            
            If Not IsEmpty(dataSheet.Cells(i, 8).Value) And dataSheet.Cells(i, 8).Value <> "" Then
                sumYest2hr = sumYest2hr + Val(dataSheet.Cells(i, 8).Value)
                has2hr = True
            End If
            
            sumYest10pm = sumYest10pm + Val(dataSheet.Cells(i, 10).Value)
            sumTarget = sumTarget + Val(dataSheet.Cells(i, 11).Value)
            writeRow = writeRow + 1
        End If
    Next i
    
    If writeRow = 9 Then
        dashSheet.Cells(9, 1).Value = "No data matched the selected filters."
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If
    
    dashSheet.Cells(writeRow, 1).Value = "Total"
    dashSheet.Cells(writeRow, 2).Value = ""
    dashSheet.Cells(writeRow, 3).Value = ""
    dashSheet.Cells(writeRow, 4).Value = sumToday
    dashSheet.Cells(writeRow, 5).Value = sumToday - sumPrevToday
    dashSheet.Cells(writeRow, 6).Value = sumYestSame
    dashSheet.Cells(writeRow, 7).Value = sumToday - sumYestSame
    
    If has2hr Then
        dashSheet.Cells(writeRow, 8).Value = sumYest2hr
        dashSheet.Cells(writeRow, 9).Value = sumYest2hr - sumToday
    Else
        dashSheet.Cells(writeRow, 8).Value = ""
        dashSheet.Cells(writeRow, 9).Value = ""
    End If
    
    dashSheet.Cells(writeRow, 10).Value = sumYest10pm
    dashSheet.Cells(writeRow, 11).Value = sumTarget
    dashSheet.Cells(writeRow, 12).Value = sumTarget - sumToday
    
    If sumTarget > 0 Then
        dashSheet.Cells(writeRow, 13).Value = sumToday / sumTarget
    Else
        dashSheet.Cells(writeRow, 13).Value = 0
    End If
    dashSheet.Cells(writeRow, 13).NumberFormat = "0%"
    
    Set dataRange = dashSheet.Range(dashSheet.Cells(9, 1), dashSheet.Cells(writeRow, 13))
    With dataRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Borders.LineStyle = 1
        .Borders.Weight = 2
        .Borders.Color = RGB(217, 217, 217)
    End With
    
    For rowIdx = 9 To writeRow
        dashSheet.Rows(rowIdx).RowHeight = 20
        For colIdx = 1 To 13
            If colIdx >= 4 Then
                dashSheet.Cells(rowIdx, colIdx).HorizontalAlignment = -4108
            Else
                dashSheet.Cells(rowIdx, colIdx).HorizontalAlignment = -4131
            End If
            dashSheet.Cells(rowIdx, colIdx).VerticalAlignment = -4108
        Next colIdx
        
        valLast2h = dashSheet.Cells(rowIdx, 5).Value
        valGap = dashSheet.Cells(rowIdx, 7).Value
        valNeeded = dashSheet.Cells(rowIdx, 9).Value
        valBtd = dashSheet.Cells(rowIdx, 12).Value
        
        If IsNumeric(valLast2h) And valLast2h = 0 Then
            dashSheet.Cells(rowIdx, 5).Interior.Color = RGB(255, 120, 120)
            dashSheet.Cells(rowIdx, 5).Font.Bold = True
        End If
        
        If IsNumeric(valGap) And valGap < 0 Then
            dashSheet.Cells(rowIdx, 7).Interior.Color = RGB(255, 120, 120)
            dashSheet.Cells(rowIdx, 7).Font.Bold = True
        End If
        
        If IsNumeric(valNeeded) And valNeeded > 0 Then
            dashSheet.Cells(rowIdx, 9).Interior.Color = RGB(255, 120, 120)
            dashSheet.Cells(rowIdx, 9).Font.Bold = True
        End If
        
        If IsNumeric(valBtd) Then
            If valBtd > 2 Then
                dashSheet.Cells(rowIdx, 12).Interior.Color = RGB(255, 120, 120)
                dashSheet.Cells(rowIdx, 12).Font.Bold = True
            ElseIf valBtd > 0 And valBtd <= 2 Then
                dashSheet.Cells(rowIdx, 12).Interior.Color = RGB(217, 225, 242)
                dashSheet.Cells(rowIdx, 12).Font.Bold = True
            Else
                dashSheet.Cells(rowIdx, 12).Interior.Color = RGB(226, 239, 218)
                dashSheet.Cells(rowIdx, 12).Font.Bold = True
            End If
        End If
    Next rowIdx
    
    With dashSheet.Range(dashSheet.Cells(writeRow, 1), dashSheet.Cells(writeRow, 13))
        .Font.Bold = True
        .Borders(8).LineStyle = -4119
        .Borders(9).LineStyle = -4119
        .Borders(8).Color = RGB(71, 85, 105)
        .Borders(9).Color = RGB(71, 85, 105)
        .Interior.Color = RGB(248, 250, 252)
    End With
    
    dashSheet.Columns("A").AutoFit
    dashSheet.Columns("B").AutoFit
    dashSheet.Columns("C").AutoFit
    If dashSheet.Columns("A").ColumnWidth < 12 Then dashSheet.Columns("A").ColumnWidth = 12
    If dashSheet.Columns("B").ColumnWidth < 12 Then dashSheet.Columns("B").ColumnWidth = 12
    If dashSheet.Columns("C").ColumnWidth < 12 Then dashSheet.Columns("C").ColumnWidth = 12
    
    For colIdx = 4 To 13
        dashSheet.Columns(Chr(64 + colIdx)).ColumnWidth = 11
    Next colIdx
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub ResetOutletFilters()
    Dim dashSheet As Worksheet
    Set dashSheet = ThisWorkbook.Sheets("Outletwise")
    Application.EnableEvents = False
    dashSheet.Range("C4").Value = ""
    dashSheet.Range("C5").Value = ""
    Application.EnableEvents = True
    Call UpdatePartnerDropdownOutletwise("")
    Call DisplayOutletwise
End Sub
"""

        vba_project = wb.VBProject
        vba_module = vba_project.VBComponents.Add(1)
        vba_module.CodeModule.AddFromString(module_code)
        
        # 4. Inject worksheet change event code into Partnerwise and Outletwise sheets
        sheet_code_partner = """
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim regionCell As Range
    Dim partnerCell As Range
    Set regionCell = Me.Range("C4")
    Set partnerCell = Me.Range("C5")
    
    If Not Intersect(Target, regionCell) Is Nothing Then
        Application.EnableEvents = False
        partnerCell.Value = ""
        Call UpdatePartnerDropdownPartnerwise(regionCell.Value)
        Application.EnableEvents = True
        Call DisplayPartnerwise
    ElseIf Not Intersect(Target, partnerCell) Is Nothing Then
        Call DisplayPartnerwise
    End If
End Sub
"""

        sheet_code_outlet = """
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim regionCell As Range
    Dim partnerCell As Range
    Set regionCell = Me.Range("C4")
    Set partnerCell = Me.Range("C5")
    
    If Not Intersect(Target, regionCell) Is Nothing Then
        Application.EnableEvents = False
        partnerCell.Value = ""
        Call UpdatePartnerDropdownOutletwise(regionCell.Value)
        Application.EnableEvents = True
        Call DisplayOutletwise
    ElseIf Not Intersect(Target, partnerCell) Is Nothing Then
        Call DisplayOutletwise
    End If
End Sub
"""

        for comp in vba_project.VBComponents:
            try:
                sheet_tab_name = comp.Properties("Name").Value
                if sheet_tab_name == "Partnerwise":
                    comp.CodeModule.AddFromString(sheet_code_partner)
                    print("Successfully injected event code into Partnerwise sheet.")
                elif sheet_tab_name == "Outletwise":
                    comp.CodeModule.AddFromString(sheet_code_outlet)
                    print("Successfully injected event code into Outletwise sheet.")
            except Exception:
                pass

        # 5. Inject Workbook_Open trigger
        workbook_code = """
Private Sub Workbook_Open()
    Call InitializePartnerValidation
    Call InitializeOutletValidation
End Sub
"""
        for comp in vba_project.VBComponents:
            if comp.Name == "ThisWorkbook":
                comp.CodeModule.AddFromString(workbook_code)
                break

        # 6. Initialize dropdowns and pre-render initial tables on both sheets
        print("Initializing Partnerwise dropdowns and rendering table...")
        excel.Run("InitializePartnerValidation")
        excel.Run("DisplayPartnerwise")
        
        print("Initializing Outletwise dropdowns and rendering table...")
        excel.Run("InitializeOutletValidation")
        excel.Run("DisplayOutletwise")
        
        # Hide raw data sheets
        raw_partner.Visible = 0  # xlSheetHidden
        raw_outlet.Visible = 0   # xlSheetHidden
        
        # Activate Partnerwise as the primary starting view
        dash_partner.Activate()
        
        # Save as Macro-Enabled Workbook (.xlsm)
        print(f"Saving macro enabled workbook: {macro_path}")
        wb.SaveAs(macro_path, FileFormat=52)
        wb.Close()
        print("Macro dashboard successfully created and closed.")
        
    except Exception as e:
        print(f"Error during COM automation: {e}")
        if "SaveAs" in str(e) or "-2146827284" in str(e) or "2147352567" in str(e):
            print("\n[IMPORTANT] The output file 'renewals_comparison.xlsm' might be open in Excel.")
            print("Please close renewals_comparison.xlsm and rerun the script to update the report.\n")
        raise e
    finally:
        excel.Quit()
        del excel



def build_web_dashboard(partner_rows, outlet_rows, metadata):
    html_template = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Subisu Renewals Performance Tracker</title>
  <style>
    :root {
      --primary: #1f4e78;
      --primary-dark: #143554;
      --primary-light: #2d6ca5;
      --accent: #0284c7;
      --bg: #f8fafc;
      --surface: #ffffff;
      --text-main: #0f172a;
      --text-muted: #64748b;
      --border: #e2e8f0;
      --border-dark: #cbd5e1;
      --badge-red-bg: #fee2e2;
      --badge-red-text: #991b1b;
      --cell-red-bg: #ff7878;
      --cell-red-text: #5a0000;
      --cell-blue-bg: #d9e1f2;
      --cell-blue-text: #1e3a8a;
      --cell-green-bg: #e2efda;
      --cell-green-text: #14532d;
      --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
      --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -2px rgba(0, 0, 0, 0.1);
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      background-color: var(--bg);
      color: var(--text-main);
      line-height: 1.5;
      padding-bottom: 60px;
    }

    /* Header & Navigation */
    .header {
      background: linear-gradient(135deg, var(--primary), var(--primary-dark));
      color: #ffffff;
      padding: 24px 32px;
      box-shadow: var(--shadow-md);
      position: sticky;
      top: 0;
      z-index: 50;
    }

    .header-content {
      max-width: 1600px;
      margin: 0 auto;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
    }

    .header-title h1 {
      font-size: 24px;
      font-weight: 700;
      letter-spacing: -0.025em;
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .header-title p {
      font-size: 13px;
      color: #cbd5e1;
      margin-top: 2px;
    }

    .header-badges {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 12px;
    }

    .badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: rgba(255, 255, 255, 0.12);
      border: 1px solid rgba(255, 255, 255, 0.25);
      padding: 6px 14px;
      border-radius: 9999px;
      font-size: 13px;
      font-weight: 500;
      backdrop-filter: blur(4px);
    }

    .badge-pulse {
      width: 8px;
      height: 8px;
      background-color: #22c55e;
      border-radius: 50%;
      box-shadow: 0 0 0 2px rgba(34, 197, 94, 0.4);
      animation: pulse 2s infinite;
    }

    @keyframes pulse {
      0%, 100% { transform: scale(1); opacity: 1; }
      50% { transform: scale(1.3); opacity: 0.6; }
    }

    /* Container */
    .container {
      max-width: 1600px;
      margin: 24px auto;
      padding: 0 24px;
    }

    /* KPI Summary Cards */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }

    .kpi-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 16px 20px;
      box-shadow: var(--shadow-sm);
      transition: transform 0.15s ease, box-shadow 0.15s ease;
    }

    .kpi-card:hover {
      transform: translateY(-2px);
      box-shadow: var(--shadow-md);
    }

    .kpi-label {
      font-size: 12px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      color: var(--text-muted);
      margin-bottom: 6px;
    }

    .kpi-value {
      font-size: 26px;
      font-weight: 800;
      color: var(--primary-dark);
      display: flex;
      align-items: baseline;
      gap: 6px;
    }

    .kpi-sub {
      font-size: 11px;
      color: var(--text-muted);
      margin-top: 4px;
    }

    .kpi-value.positive { color: #16a34a; }
    .kpi-value.negative { color: #dc2626; }

    /* Control Panel */
    .control-panel {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 16px 20px;
      margin-bottom: 20px;
      box-shadow: var(--shadow-sm);
      display: flex;
      flex-direction: column;
      gap: 16px;
    }

    .tabs-row {
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      border-bottom: 1px solid var(--border);
      padding-bottom: 16px;
    }

    .tab-group {
      display: inline-flex;
      background: #f1f5f9;
      padding: 4px;
      border-radius: 10px;
      gap: 4px;
    }

    .tab-btn {
      padding: 8px 20px;
      border: none;
      background: transparent;
      font-size: 14px;
      font-weight: 600;
      color: var(--text-muted);
      border-radius: 8px;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .tab-btn.active {
      background: #ffffff;
      color: var(--primary);
      box-shadow: var(--shadow-sm);
    }

    .tab-badge {
      background: #e2e8f0;
      color: #334155;
      font-size: 11px;
      padding: 2px 8px;
      border-radius: 9999px;
      font-weight: 700;
    }

    .tab-btn.active .tab-badge {
      background: var(--primary);
      color: #ffffff;
    }

    .filters-row {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 14px;
    }

    .filter-item {
      display: flex;
      align-items: center;
      gap: 8px;
      flex: 1;
      min-width: 200px;
    }

    .filter-label {
      font-size: 13px;
      font-weight: 600;
      color: var(--text-main);
      white-space: nowrap;
    }

    .filter-select, .search-input {
      width: 100%;
      padding: 8px 12px;
      border: 1px solid var(--border-dark);
      border-radius: 8px;
      font-size: 14px;
      background-color: #ffffff;
      color: var(--text-main);
      outline: none;
      transition: border-color 0.15s ease, box-shadow 0.15s ease;
    }

    .filter-select:focus, .search-input:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(2, 132, 199, 0.15);
    }

    .btn {
      padding: 8px 18px;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.15s ease;
      white-space: nowrap;
      border: none;
    }

    .btn-secondary {
      background: #f1f5f9;
      color: #334155;
      border: 1px solid var(--border-dark);
    }

    .btn-secondary:hover {
      background: #e2e8f0;
    }

    .btn-primary {
      background: var(--primary);
      color: #ffffff;
    }

    .btn-primary:hover {
      background: var(--primary-light);
    }

    /* Table Container */
    .table-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      box-shadow: var(--shadow-sm);
      overflow: hidden;
    }

    .table-wrapper {
      overflow-x: auto;
      max-height: 720px;
      position: relative;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
      text-align: left;
    }

    thead th {
      background-color: var(--primary);
      color: #ffffff;
      font-weight: 600;
      padding: 12px 10px;
      text-align: center;
      white-space: normal;
      vertical-align: middle;
      position: sticky;
      top: 0;
      z-index: 20;
      border-right: 1px solid rgba(255, 255, 255, 0.15);
      border-bottom: 2px solid var(--primary-dark);
      font-size: 12px;
      line-height: 1.3;
    }

    thead th:first-child, thead th:nth-child(2), thead th:nth-child(3) {
      text-align: left;
      padding-left: 14px;
    }

    tbody td {
      padding: 8px 10px;
      border-bottom: 1px solid var(--border);
      border-right: 1px solid #f1f5f9;
      vertical-align: middle;
      font-variant-numeric: tabular-nums;
    }

    tbody tr:hover {
      background-color: #f8fafc;
    }

    .text-left { text-align: left !important; }
    .text-center { text-align: center !important; }
    .text-right { text-align: right !important; }

    /* Conditional Formatting Styles matching Excel */
    .cell-red {
      background-color: var(--cell-red-bg) !important;
      color: var(--cell-red-text) !important;
      font-weight: 700;
    }

    .cell-blue {
      background-color: var(--cell-blue-bg) !important;
      color: var(--cell-blue-text) !important;
      font-weight: 700;
    }

    .cell-green {
      background-color: var(--cell-green-bg) !important;
      color: var(--cell-green-text) !important;
      font-weight: 700;
    }

    /* Footer / Totals Row */
    tfoot tr {
      background-color: #f8fafc;
      font-weight: 700;
      position: sticky;
      bottom: 0;
      z-index: 10;
      box-shadow: 0 -2px 4px rgba(0,0,0,0.05);
    }

    tfoot td {
      padding: 12px 10px;
      border-top: 2px solid var(--border-dark);
      border-bottom: 3px double var(--primary);
      border-right: 1px solid var(--border);
    }

    .empty-state {
      text-align: center;
      padding: 48px 24px;
      color: var(--text-muted);
      font-size: 15px;
    }

    /* Footer info */
    .dashboard-footer {
      text-align: center;
      margin-top: 28px;
      font-size: 12px;
      color: var(--text-muted);
    }

    @media (max-width: 768px) {
      .header-content { flex-direction: column; align-items: flex-start; }
      .filters-row { flex-direction: column; align-items: stretch; }
      .filter-item { min-width: 100%; }
      .kpi-grid { grid-template-columns: repeat(2, 1fr); }
    }
  </style>
</head>
<body>

  <!-- Header -->
  <header class="header">
    <div class="header-content">
      <div class="header-title">
        <h1>
          <span>⚡</span> Subisu Renewals Performance Tracker
        </h1>
        <p>Hourly Comparative Renewals Intelligence & Performance Monitoring</p>
      </div>
      <div class="header-badges">
        <div class="badge">
          <span class="badge-pulse"></span>
          <span>Target Slot: <strong id="slot-badge">""" + metadata["target_date"] + " " + metadata["target_time_slot"] + """</strong></span>
        </div>
        <div class="badge">
          <span>🕒 Last Updated: <strong id="updated-badge">""" + metadata["generated_at"] + """</strong></span>
        </div>
      </div>
    </div>
  </header>

  <div class="container">

    <!-- KPI Summary Grid -->
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">Renewals Today</div>
        <div class="kpi-value" id="kpi-today">0</div>
        <div class="kpi-sub">Total renewals achieved today</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Last 2 Hours</div>
        <div class="kpi-value" id="kpi-last2h">0</div>
        <div class="kpi-sub">Velocity in recent 2hr window</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Yesterday Same Time</div>
        <div class="kpi-value" id="kpi-yest-same">0</div>
        <div class="kpi-sub">Benchmark at same hour yesterday</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Renewal Gap</div>
        <div class="kpi-value" id="kpi-gap">0</div>
        <div class="kpi-sub">Today vs Yesterday Same Time</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Target</div>
        <div class="kpi-value" id="kpi-target">0</div>
        <div class="kpi-sub">Calculated base target</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Behind To Date (BTD)</div>
        <div class="kpi-value" id="kpi-btd">0</div>
        <div class="kpi-sub">Remaining to hit target</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Achieved %</div>
        <div class="kpi-value" id="kpi-achieved">0%</div>
        <div class="kpi-sub">Overall progress to target</div>
      </div>
    </div>

    <!-- Control Panel -->
    <div class="control-panel">
      <div class="tabs-row">
        <div class="tab-group">
          <button class="tab-btn active" id="tab-partner" onclick="switchTab('partner')">
            <span>🏢 Partnerwise</span>
            <span class="tab-badge" id="badge-count-partner">""" + str(len(partner_rows)) + """</span>
          </button>
          <button class="tab-btn" id="tab-outlet" onclick="switchTab('outlet')">
            <span>📍 Outletwise</span>
            <span class="tab-badge" id="badge-count-outlet">""" + str(len(outlet_rows)) + """</span>
          </button>
        </div>

        <div style="display: flex; gap: 10px; align-items: center;">
          <button class="btn btn-secondary" onclick="resetFilters()">
            <span>↺</span> Reset Filters
          </button>
          <button class="btn btn-primary" onclick="exportToCSV()">
            <span>📥</span> Export to CSV
          </button>
        </div>
      </div>

      <div class="filters-row">
        <div class="filter-item">
          <label class="filter-label" for="filter-region">Region:</label>
          <select id="filter-region" class="filter-select" onchange="onRegionChange()">
            <option value="">All Regions</option>
          </select>
        </div>

        <div class="filter-item">
          <label class="filter-label" for="filter-partner">Partner:</label>
          <select id="filter-partner" class="filter-select" onchange="onPartnerChange()">
            <option value="">All Partners</option>
          </select>
        </div>

        <div class="filter-item" style="flex: 1.5;">
          <label class="filter-label" for="filter-search">Search:</label>
          <input type="text" id="filter-search" class="search-input" placeholder="Quick search by name..." oninput="onSearchChange()">
        </div>
      </div>
    </div>

    <!-- Table Card -->
    <div class="table-card">
      <div class="table-wrapper">
        <table id="dashboard-table">
          <thead id="table-head">
            <!-- Rendered by JS -->
          </thead>
          <tbody id="table-body">
            <!-- Rendered by JS -->
          </tbody>
          <tfoot id="table-foot">
            <!-- Rendered by JS -->
          </tfoot>
        </table>
      </div>
    </div>

    <div class="dashboard-footer">
      Subisu Cablenet Ltd. • Automated Operations Intelligence Dashboard • Hosted on GitHub Pages
    </div>

  </div>

  <!-- Embedded Dashboard Data -->
  <script id="dashboard-data" type="application/json">
""" + json.dumps({
        "metadata": metadata,
        "partner_data": partner_rows,
        "outlet_data": outlet_rows
    }, indent=2) + """
  </script>

  <!-- Dashboard Controller Script -->
  <script>
    let currentTab = 'partner';
    const rawData = JSON.parse(document.getElementById('dashboard-data').textContent);
    const partnerData = rawData.partner_data || [];
    const outletData = rawData.outlet_data || [];

    function init() {
      populateRegionDropdown();
      populatePartnerDropdown();
      renderCurrentView();
    }

    function switchTab(tab) {
      if (currentTab === tab) return;
      currentTab = tab;
      
      document.getElementById('tab-partner').classList.toggle('active', tab === 'partner');
      document.getElementById('tab-outlet').classList.toggle('active', tab === 'outlet');
      
      // Reset partner dropdown options for active tab
      populateRegionDropdown();
      populatePartnerDropdown();
      renderCurrentView();
    }

    function getActiveDataset() {
      return currentTab === 'partner' ? partnerData : outletData;
    }

    function populateRegionDropdown() {
      const data = getActiveDataset();
      const regionSelect = document.getElementById('filter-region');
      const currentSelected = regionSelect.value;
      
      const regions = new Set();
      data.forEach(row => {
        const r = (row.Region || row.region || '').trim();
        if (r && r.toLowerCase() !== 'total' && r.toLowerCase() !== 'all') {
          regions.add(r);
        }
      });
      
      const sorted = Array.from(regions).sort();
      regionSelect.innerHTML = '<option value="">All Regions</option>';
      sorted.forEach(r => {
        const opt = document.createElement('option');
        opt.value = r;
        opt.textContent = r;
        if (r === currentSelected) opt.selected = true;
        regionSelect.appendChild(opt);
      });
    }

    function populatePartnerDropdown() {
      const data = getActiveDataset();
      const selectedRegion = document.getElementById('filter-region').value.trim().toLowerCase();
      const partnerSelect = document.getElementById('filter-partner');
      const currentSelected = partnerSelect.value;
      
      const partners = new Set();
      data.forEach(row => {
        const r = (row.Region || row.region || '').trim().toLowerCase();
        const p = (row.Partner || '').trim();
        if (!selectedRegion || r === selectedRegion) {
          if (p) partners.add(p);
        }
      });
      
      const sorted = Array.from(partners).sort();
      partnerSelect.innerHTML = '<option value="">All Partners</option>';
      sorted.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p;
        opt.textContent = p;
        if (p === currentSelected) opt.selected = true;
        partnerSelect.appendChild(opt);
      });
    }

    function onRegionChange() {
      populatePartnerDropdown();
      renderCurrentView();
    }

    function onPartnerChange() {
      renderCurrentView();
    }

    function onSearchChange() {
      renderCurrentView();
    }

    function resetFilters() {
      document.getElementById('filter-region').value = '';
      document.getElementById('filter-partner').value = '';
      document.getElementById('filter-search').value = '';
      populatePartnerDropdown();
      renderCurrentView();
    }

    function getFilteredData() {
      const data = getActiveDataset();
      const selRegion = document.getElementById('filter-region').value.trim().toLowerCase();
      const selPartner = document.getElementById('filter-partner').value.trim().toLowerCase();
      const searchQuery = document.getElementById('filter-search').value.trim().toLowerCase();
      
      return data.filter(row => {
        const r = (row.Region || row.region || '').trim().toLowerCase();
        const p = (row.Partner || '').trim().toLowerCase();
        const o = (row.Outlet || '').trim().toLowerCase();
        
        if (selRegion && r !== selRegion) return false;
        if (selPartner && p !== selPartner) return false;
        if (searchQuery) {
          const matched = r.includes(searchQuery) || p.includes(searchQuery) || o.includes(searchQuery);
          if (!matched) return false;
        }
        return true;
      });
    }

    function renderCurrentView() {
      const filtered = getFilteredData();
      renderKPIs(filtered);
      renderTableHeaders();
      renderTableBody(filtered);
      renderTableFooter(filtered);
    }

    function renderKPIs(rows) {
      let sumToday = 0;
      let sumLast2h = 0;
      let sumYestSame = 0;
      let sumTarget = 0;

      rows.forEach(r => {
        sumToday += Number(r['Renewals today'] || 0);
        sumLast2h += Number(r['Renewals in last 2 hours'] || 0);
        sumYestSame += Number(r['Renewals Yesterday - Same time'] || 0);
        sumTarget += Number(r['Target'] || 0);
      });

      const gap = sumToday - sumYestSame;
      const btd = sumTarget - sumToday;
      const pct = sumTarget > 0 ? Math.round((sumToday / sumTarget) * 100) : 0;

      document.getElementById('kpi-today').textContent = sumToday.toLocaleString();
      document.getElementById('kpi-last2h').textContent = sumLast2h.toLocaleString();
      document.getElementById('kpi-yest-same').textContent = sumYestSame.toLocaleString();
      
      const gapEl = document.getElementById('kpi-gap');
      gapEl.textContent = (gap >= 0 ? '+' : '') + gap.toLocaleString();
      gapEl.className = 'kpi-value ' + (gap < 0 ? 'negative' : 'positive');

      document.getElementById('kpi-target').textContent = sumTarget.toLocaleString();
      
      const btdEl = document.getElementById('kpi-btd');
      btdEl.textContent = btd.toLocaleString();
      btdEl.className = 'kpi-value ' + (btd > 2 ? 'negative' : (btd <= 0 ? 'positive' : ''));

      document.getElementById('kpi-achieved').textContent = pct + '%';
    }

    function renderTableHeaders() {
      const thead = document.getElementById('table-head');
      if (currentTab === 'partner') {
        thead.innerHTML = `
          <tr>
            <th style="min-width: 140px;">Region</th>
            <th style="min-width: 220px;">Partner</th>
            <th style="min-width: 100px;">Renewals<br>today</th>
            <th style="min-width: 110px;">Renewals in<br>last 2 hours</th>
            <th style="min-width: 120px;">Renewals Yesterday<br>- Same time</th>
            <th style="min-width: 110px;">Renewal Gap<br>- Same time</th>
            <th style="min-width: 120px;">Renewals yesterday<br>- 2 hrs later</th>
            <th style="min-width: 120px;">Renewals needed<br>in 2 hours</th>
            <th style="min-width: 120px;">Renewals Yesterday<br>- 10:12PM</th>
            <th style="min-width: 90px;">Target</th>
            <th style="min-width: 90px;">BTD</th>
            <th style="min-width: 100px;">Achieved %</th>
          </tr>
        `;
      } else {
        thead.innerHTML = `
          <tr>
            <th style="min-width: 130px;">Region</th>
            <th style="min-width: 200px;">Partner</th>
            <th style="min-width: 160px;">Outlet</th>
            <th style="min-width: 100px;">Renewals<br>today</th>
            <th style="min-width: 110px;">Renewals in<br>last 2 hours</th>
            <th style="min-width: 120px;">Renewals Yesterday<br>- Same time</th>
            <th style="min-width: 110px;">Renewal Gap<br>- Same time</th>
            <th style="min-width: 120px;">Renewals yesterday<br>- 2 hrs later</th>
            <th style="min-width: 120px;">Renewals needed<br>in 2 hours</th>
            <th style="min-width: 120px;">Renewals Yesterday<br>- 10:12PM</th>
            <th style="min-width: 90px;">Target</th>
            <th style="min-width: 90px;">BTD</th>
            <th style="min-width: 100px;">Achieved %</th>
          </tr>
        `;
      }
    }

    function renderTableBody(rows) {
      const tbody = document.getElementById('table-body');
      if (rows.length === 0) {
        const colSpan = currentTab === 'partner' ? 12 : 13;
        tbody.innerHTML = `<tr><td colspan="${colSpan}" class="empty-state">No records found matching the selected filters.</td></tr>`;
        return;
      }

      let html = '';
      rows.forEach(r => {
        const today = Number(r['Renewals today'] || 0);
        const last2h = Number(r['Renewals in last 2 hours'] || 0);
        const yestSame = Number(r['Renewals Yesterday - Same time'] || 0);
        const gap = Number(r['Renewal Gap - Same time'] || 0);
        const yest2hr = r['Renewals yesterday - 2 hrs later'];
        const needed = r['Renewals needed in 2 hours'];
        const yest10pm = Number(r['Renewals Yesterday - 10:12PM'] || 0);
        const target = Number(r['Target'] || 0);
        const btd = Number(r['BTD'] || 0);
        const achPct = target > 0 ? Math.round((today / target) * 100) : 0;

        // Conditional classes matching Excel
        const last2hClass = last2h === 0 ? 'cell-red' : '';
        const gapClass = gap < 0 ? 'cell-red' : '';
        const neededClass = (needed !== null && needed !== undefined && Number(needed) > 0) ? 'cell-red' : '';
        
        let btdClass = '';
        if (btd > 2) btdClass = 'cell-red';
        else if (btd > 0 && btd <= 2) btdClass = 'cell-blue';
        else btdClass = 'cell-green';

        if (currentTab === 'partner') {
          html += `
            <tr>
              <td class="text-left font-medium">${r.Region || ''}</td>
              <td class="text-left">${r.Partner || ''}</td>
              <td class="text-center">${today}</td>
              <td class="text-center ${last2hClass}">${last2h}</td>
              <td class="text-center">${yestSame}</td>
              <td class="text-center ${gapClass}">${gap}</td>
              <td class="text-center">${yest2hr !== null && yest2hr !== undefined ? yest2hr : ''}</td>
              <td class="text-center ${neededClass}">${needed !== null && needed !== undefined ? needed : ''}</td>
              <td class="text-center">${yest10pm}</td>
              <td class="text-center">${target}</td>
              <td class="text-center ${btdClass}">${btd}</td>
              <td class="text-center font-bold">${achPct}%</td>
            </tr>
          `;
        } else {
          html += `
            <tr>
              <td class="text-left font-medium">${r.region || r.Region || ''}</td>
              <td class="text-left">${r.Partner || ''}</td>
              <td class="text-left">${r.Outlet || ''}</td>
              <td class="text-center">${today}</td>
              <td class="text-center ${last2hClass}">${last2h}</td>
              <td class="text-center">${yestSame}</td>
              <td class="text-center ${gapClass}">${gap}</td>
              <td class="text-center">${yest2hr !== null && yest2hr !== undefined ? yest2hr : ''}</td>
              <td class="text-center ${neededClass}">${needed !== null && needed !== undefined ? needed : ''}</td>
              <td class="text-center">${yest10pm}</td>
              <td class="text-center">${target}</td>
              <td class="text-center ${btdClass}">${btd}</td>
              <td class="text-center font-bold">${achPct}%</td>
            </tr>
          `;
        }
      });

      tbody.innerHTML = html;
    }

    function renderTableFooter(rows) {
      const tfoot = document.getElementById('table-foot');
      if (rows.length === 0) {
        tfoot.innerHTML = '';
        return;
      }

      let sumToday = 0;
      let sumPrevToday = 0;
      let sumYestSame = 0;
      let sumYest2hr = 0;
      let sumYest10pm = 0;
      let sumTarget = 0;
      let has2hr = false;

      rows.forEach(r => {
        sumToday += Number(r['Renewals today'] || 0);
        sumPrevToday += Number(r['today_prev_renewal'] || 0);
        sumYestSame += Number(r['Renewals Yesterday - Same time'] || 0);
        if (r['Renewals yesterday - 2 hrs later'] !== null && r['Renewals yesterday - 2 hrs later'] !== undefined && r['Renewals yesterday - 2 hrs later'] !== '') {
          sumYest2hr += Number(r['Renewals yesterday - 2 hrs later']);
          has2hr = true;
        }
        sumYest10pm += Number(r['Renewals Yesterday - 10:12PM'] || 0);
        sumTarget += Number(r['Target'] || 0);
      });

      const totLast2h = sumToday - sumPrevToday;
      const totGap = sumToday - sumYestSame;
      const totNeeded = has2hr ? (sumYest2hr - sumToday) : '';
      const totBtd = sumTarget - sumToday;
      const totAchPct = sumTarget > 0 ? Math.round((sumToday / sumTarget) * 100) : 0;

      const last2hClass = totLast2h === 0 ? 'cell-red' : '';
      const gapClass = totGap < 0 ? 'cell-red' : '';
      const neededClass = (has2hr && totNeeded > 0) ? 'cell-red' : '';
      
      let btdClass = '';
      if (totBtd > 2) btdClass = 'cell-red';
      else if (totBtd > 0 && totBtd <= 2) btdClass = 'cell-blue';
      else btdClass = 'cell-green';

      if (currentTab === 'partner') {
        tfoot.innerHTML = `
          <tr>
            <td class="text-left">Total</td>
            <td class="text-left"></td>
            <td class="text-center">${sumToday}</td>
            <td class="text-center ${last2hClass}">${totLast2h}</td>
            <td class="text-center">${sumYestSame}</td>
            <td class="text-center ${gapClass}">${totGap}</td>
            <td class="text-center">${has2hr ? sumYest2hr : ''}</td>
            <td class="text-center ${neededClass}">${totNeeded}</td>
            <td class="text-center">${sumYest10pm}</td>
            <td class="text-center">${sumTarget}</td>
            <td class="text-center ${btdClass}">${totBtd}</td>
            <td class="text-center font-bold">${totAchPct}%</td>
          </tr>
        `;
      } else {
        tfoot.innerHTML = `
          <tr>
            <td class="text-left">Total</td>
            <td class="text-left"></td>
            <td class="text-left"></td>
            <td class="text-center">${sumToday}</td>
            <td class="text-center ${last2hClass}">${totLast2h}</td>
            <td class="text-center">${sumYestSame}</td>
            <td class="text-center ${gapClass}">${totGap}</td>
            <td class="text-center">${has2hr ? sumYest2hr : ''}</td>
            <td class="text-center ${neededClass}">${totNeeded}</td>
            <td class="text-center">${sumYest10pm}</td>
            <td class="text-center">${sumTarget}</td>
            <td class="text-center ${btdClass}">${totBtd}</td>
            <td class="text-center font-bold">${totAchPct}%</td>
          </tr>
        `;
      }
    }

    function exportToCSV() {
      const rows = getFilteredData();
      if (rows.length === 0) {
        alert('No data available to export.');
        return;
      }

      let headers = [];
      if (currentTab === 'partner') {
        headers = [
          'Region', 'Partner', 'Renewals today', 'Renewals in last 2 hours',
          'Renewals Yesterday - Same time', 'Renewal Gap - Same time',
          'Renewals yesterday - 2 hrs later', 'Renewals needed in 2 hours',
          'Renewals Yesterday - 10:12PM', 'Target', 'BTD', 'Achieved %'
        ];
      } else {
        headers = [
          'Region', 'Partner', 'Outlet', 'Renewals today', 'Renewals in last 2 hours',
          'Renewals Yesterday - Same time', 'Renewal Gap - Same time',
          'Renewals yesterday - 2 hrs later', 'Renewals needed in 2 hours',
          'Renewals Yesterday - 10:12PM', 'Target', 'BTD', 'Achieved %'
        ];
      }

      const csvRows = [headers.join(',')];
      rows.forEach(r => {
        const today = Number(r['Renewals today'] || 0);
        const target = Number(r['Target'] || 0);
        const pct = target > 0 ? Math.round((today / target) * 100) + '%' : '0%';
        
        let rowVals = [];
        if (currentTab === 'partner') {
          rowVals = [
            `"${r.Region || ''}"`,
            `"${r.Partner || ''}"`,
            today,
            Number(r['Renewals in last 2 hours'] || 0),
            Number(r['Renewals Yesterday - Same time'] || 0),
            Number(r['Renewal Gap - Same time'] || 0),
            r['Renewals yesterday - 2 hrs later'] ?? '',
            r['Renewals needed in 2 hours'] ?? '',
            Number(r['Renewals Yesterday - 10:12PM'] || 0),
            target,
            Number(r['BTD'] || 0),
            `"${pct}"`
          ];
        } else {
          rowVals = [
            `"${r.region || r.Region || ''}"`,
            `"${r.Partner || ''}"`,
            `"${r.Outlet || ''}"`,
            today,
            Number(r['Renewals in last 2 hours'] || 0),
            Number(r['Renewals Yesterday - Same time'] || 0),
            Number(r['Renewal Gap - Same time'] || 0),
            r['Renewals yesterday - 2 hrs later'] ?? '',
            r['Renewals needed in 2 hours'] ?? '',
            Number(r['Renewals Yesterday - 10:12PM'] || 0),
            target,
            Number(r['BTD'] || 0),
            `"${pct}"`
          ];
        }
        csvRows.push(rowVals.join(','));
      });

      const blob = new Blob([csvRows.join('\\n')], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.setAttribute('href', url);
      link.setAttribute('download', `renewals_${currentTab}_export.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }

    // Initialize on load
    window.addEventListener('DOMContentLoaded', init);
  </script>
</body>
</html>
"""
    return html_template

def generate_web_dashboard(output_partner_df, output_outlet_df, target_date, target_time_slot, next_time_slot, output_dir):
    """
    Generates a responsive, modern standalone HTML dashboard (index.html)
    with embedded JSON data, ready for immediate hosting on GitHub Pages
    or offline browser viewing without CORS issues.
    """
    p_records = output_partner_df.fillna("").to_dict(orient="records")
    o_records = output_outlet_df.fillna("").to_dict(orient="records")
    
    metadata = {
        "target_date": target_date,
        "target_time_slot": target_time_slot,
        "next_time_slot": next_time_slot or "",
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    html_content = build_web_dashboard(p_records, o_records, metadata)
    project_root = os.path.dirname(output_dir)
    root_html = os.path.join(project_root, "index.html")
    docs_dir = os.path.join(project_root, "docs")
    os.makedirs(docs_dir, exist_ok=True)
    docs_html = os.path.join(docs_dir, "index.html")
    
    with open(root_html, "w", encoding="utf-8") as f:
        f.write(html_content)
    with open(docs_html, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    print(f"Successfully generated Web Dashboard: {root_html} and {docs_html}")

def sync_to_github():
    """
    Automatically commits and pushes the updated web dashboard to GitHub Pages
    if a git remote 'origin' is configured.
    """
    import subprocess
    git_exe = r"C:\Users\hp\AppData\Local\Programs\Git\cmd\git.exe"
    if not os.path.exists(git_exe):
        import shutil
        git_exe = shutil.which("git")
    if not git_exe:
        print("[INFO] Git executable not found. Skipping GitHub sync.")
        return
        
    project_root = r"F:\targetpend"
    print("\n----------------------------------------------------")
    print("3. Syncing Web Dashboard to GitHub Pages...")
    print("----------------------------------------------------")
    try:
        chk = subprocess.run([git_exe, "rev-parse", "--is-inside-work-tree"], cwd=project_root, capture_output=True, text=True)
        if chk.returncode != 0:
            print("[INFO] Not a git repository. Skipping GitHub sync.")
            return

        subprocess.run([git_exe, "add", "index.html", "docs/", "README.md", ".gitignore", "generate_renewals_report.py", "generate_report.py", "run_report.bat"], cwd=project_root, check=False)
        
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        subprocess.run([git_exe, "commit", "-m", f"Auto-update renewals tracker web dashboard: {now_str}"], cwd=project_root, capture_output=True)
        
        res = subprocess.run([git_exe, "remote", "get-url", "origin"], cwd=project_root, capture_output=True, text=True)
        if res.returncode == 0:
            print("Pushing latest web dashboard to GitHub Pages...")
            push_res = subprocess.run([git_exe, "push", "origin", "main"], cwd=project_root, capture_output=True, text=True)
            if push_res.returncode == 0:
                print("[SUCCESS] Web dashboard successfully synced to GitHub Pages!")
            else:
                print(f"[WARNING] Git push failed: {push_res.stderr.strip()}")
        else:
            print("[INFO] Local Git commit created. Web dashboard (index.html) is ready.")
            print("       To enable automatic GitHub Pages push, configure your remote:")
            print("       git remote add origin https://github.com/<username>/<repo>.git")
    except Exception as e:
        print(f"[WARNING] Git sync encountered an issue: {e}")

def main():
    source_dir = r"F:\targetpend\source"
    output_dir = r"F:\targetpend\output"

    # Find the daily source CSV file (excluding temp directory)
    csv_files = [f for f in glob.glob(os.path.join(source_dir, "*.csv")) if "temp" not in f]
    if not csv_files:
        print("Error: No daily CSV source files found in source/ directory.")
        return
    daily_source = max(csv_files, key=os.path.getmtime)
    print(f"Loading daily user status source: {daily_source}")
    
    # Load daily user status data
    df_daily = pd.read_csv(daily_source)
    df_daily.columns = df_daily.columns.str.strip()
    
    cols_map = {c.lower(): c for c in df_daily.columns}
    region_col = cols_map.get('region', 'Region')
    partner_col = cols_map.get('partner', 'Partner')
    outlet_col = cols_map.get('outlet', 'Outlet')
    
    for col in [region_col, partner_col, outlet_col]:
        df_daily[col] = df_daily[col].astype(str).str.strip().replace('nan', '')

    print("Connecting to Outlook...")
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        outlook = win32com.client.Dispatch("Outlook.Application")
        
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6) # olFolderInbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    
    sender_email = "noreply.s3@subisu.net.np"
    filtered_messages = messages.Restrict(f"[SenderEmailAddress] = '{sender_email}'")
    
    print(f"Total emails from {sender_email}: {len(filtered_messages)}")
    if len(filtered_messages) == 0:
        print("Error: No emails found from the specified sender.")
        return

    # Find the latest email received today to determine the target date and time slot
    target_date = None
    target_time_slot = None
    
    for msg in filtered_messages:
        try:
            dt = parse_subject_datetime(msg.Subject)
            if dt:
                target_date = dt.strftime("%Y-%m-%d")
                time_str = dt.strftime("%H:%M")
                for slot in TIME_SLOTS:
                    if slot == time_str:
                        target_time_slot = slot
                        break
                if target_time_slot:
                    break
        except Exception:
            pass
            
    if not target_date or not target_time_slot:
        print("Error: Could not identify a valid target date and time slot from recent emails.")
        return
        
    print(f"Targeting Date: {target_date} | Time Slot: {target_time_slot}")
    
    target_dt = datetime.datetime.strptime(target_date, "%Y-%m-%d")
    yesterday_dt = target_dt - datetime.timedelta(days=1)
    yesterday_date = yesterday_dt.strftime("%Y-%m-%d")
    
    try:
        slot_idx = TIME_SLOTS.index(target_time_slot)
        next_time_slot = TIME_SLOTS[slot_idx + 1] if slot_idx < len(TIME_SLOTS) - 1 else None
    except ValueError:
        next_time_slot = None

    try:
        slot_idx = TIME_SLOTS.index(target_time_slot)
        prev_time_slot = TIME_SLOTS[slot_idx - 1] if slot_idx > 0 else None
    except ValueError:
        prev_time_slot = None

    end_time_slot = "22:12"
    
    temp_dir = os.path.join(source_dir, "temp")
    os.makedirs(temp_dir, exist_ok=True)
    
    print("Downloading attachments from Outlook...")
    # 1. Partner attachments (recent_renewal_partnerwise_test.csv)
    p_path_today = get_attachment(filtered_messages, target_date, target_time_slot, temp_dir, "partner", "partner_today")
    p_path_prev = get_attachment(filtered_messages, target_date, prev_time_slot, temp_dir, "partner", "partner_prev") if prev_time_slot else None
    p_path_ys = get_attachment(filtered_messages, yesterday_date, target_time_slot, temp_dir, "partner", "partner_ys")
    p_path_y2 = get_attachment(filtered_messages, yesterday_date, next_time_slot, temp_dir, "partner", "partner_y2") if next_time_slot else None
    p_path_y10 = get_attachment(filtered_messages, yesterday_date, end_time_slot, temp_dir, "partner", "partner_y10")

    # 2. Outlet attachments (recent_renewal_outletwise.csv)
    o_path_today = get_attachment(filtered_messages, target_date, target_time_slot, temp_dir, "outletwise", "outlet_today")
    o_path_prev = get_attachment(filtered_messages, target_date, prev_time_slot, temp_dir, "outletwise", "outlet_prev") if prev_time_slot else None
    o_path_ys = get_attachment(filtered_messages, yesterday_date, target_time_slot, temp_dir, "outletwise", "outlet_ys")
    o_path_y2 = get_attachment(filtered_messages, yesterday_date, next_time_slot, temp_dir, "outletwise", "outlet_y2") if next_time_slot else None
    o_path_y10 = get_attachment(filtered_messages, yesterday_date, end_time_slot, temp_dir, "outletwise", "outlet_y10")

    # Helper: Preprocess Partner data
    def load_partner_df(filepath, prefix):
        if not filepath or not os.path.exists(filepath):
            return pd.DataFrame(columns=['Region', 'Partner', f'{prefix}_renewal'])
        df = pd.read_csv(filepath)
        df.columns = df.columns.str.strip()
        if 'Region' in df.columns:
            df = df[~df['Region'].astype(str).str.strip().str.lower().isin(['total', 'all'])].copy()
        
        renewal_col = None
        if 'Renewal' in df.columns:
            renewal_col = 'Renewal'
        elif len(df.columns) >= 13:
            renewal_col = df.columns[12]
            
        clean_df = pd.DataFrame()
        clean_df['Region'] = df['Region'].astype(str).str.strip().replace('nan', '')
        clean_df['Partner'] = df['Partner'].astype(str).str.strip().replace('nan', '')
        if renewal_col and renewal_col in df.columns:
            clean_df[f'{prefix}_renewal'] = pd.to_numeric(df[renewal_col], errors='coerce').fillna(0).astype(int)
        else:
            clean_df[f'{prefix}_renewal'] = 0
            
        return clean_df.groupby(['Region', 'Partner'], dropna=False)[f'{prefix}_renewal'].sum().reset_index()

    # Helper: Preprocess Outlet data
    def load_outlet_df(filepath, prefix):
        if not filepath or not os.path.exists(filepath):
            return pd.DataFrame(columns=['Partner', 'Outlet', f'{prefix}_renewal'])
        df = pd.read_csv(filepath)
        df.columns = df.columns.str.strip()
        clean_df = df[['Partner', 'Outlet', 'Renewal']].copy()
        for col in ['Partner', 'Outlet']:
            clean_df[col] = clean_df[col].astype(str).str.strip().replace('nan', '')
        clean_df['Renewal'] = pd.to_numeric(clean_df['Renewal'], errors='coerce').fillna(0).astype(int)
        clean_df = clean_df.rename(columns={'Renewal': f'{prefix}_renewal'})
        return clean_df.groupby(['Partner', 'Outlet'], dropna=False)[f'{prefix}_renewal'].sum().reset_index()

    # -------------------------------------------------------------
    # BUILD PARTNERWISE DATASET
    # -------------------------------------------------------------
    print("Processing Partnerwise dataset...")
    pt = load_partner_df(p_path_today, "today")
    ptp = load_partner_df(p_path_prev, "today_prev")
    pys = load_partner_df(p_path_ys, "yest_same")
    py2 = load_partner_df(p_path_y2, "yest_2hr")
    py10 = load_partner_df(p_path_y10, "yest_10pm")

    mp1 = pd.merge(pt, pys, on=['Region', 'Partner'], how='outer')
    mp2 = pd.merge(mp1, py2, on=['Region', 'Partner'], how='outer')
    mp3 = pd.merge(mp2, py10, on=['Region', 'Partner'], how='outer')
    partner_merged = pd.merge(mp3, ptp, on=['Region', 'Partner'], how='outer')

    renewal_cols = ['today_renewal', 'yest_same_renewal', 'yest_2hr_renewal', 'yest_10pm_renewal', 'today_prev_renewal']
    partner_merged[renewal_cols] = partner_merged[renewal_cols].fillna(0).astype(int)

    # Join daily user count per Partner for Target calculation
    daily_p_counts = df_daily.groupby('Partner', dropna=False).size().reset_index(name='Total_users')
    daily_p_counts['Partner'] = daily_p_counts['Partner'].astype(str).str.strip()
    
    partner_merged['Partner'] = partner_merged['Partner'].astype(str).str.strip()
    partner_merged = pd.merge(partner_merged, daily_p_counts, on='Partner', how='left')
    partner_merged['Total_users'] = partner_merged['Total_users'].fillna(0).astype(int)
    
    partner_merged['Target'] = (partner_merged['Total_users'] / 3.0).round().astype(int)
    partner_merged['BTD'] = partner_merged['Target'] - partner_merged['today_renewal']
    
    partner_merged['Renewals today'] = partner_merged['today_renewal']
    partner_merged['Renewals in last 2 hours'] = partner_merged['today_renewal'] - partner_merged['today_prev_renewal']
    partner_merged['Renewals Yesterday - Same time'] = partner_merged['yest_same_renewal']
    partner_merged['Renewal Gap - Same time'] = partner_merged['today_renewal'] - partner_merged['yest_same_renewal']
    
    if next_time_slot:
        partner_merged['Renewals yesterday - 2 hrs later'] = partner_merged['yest_2hr_renewal']
        partner_merged['Renewals needed in 2 hours'] = partner_merged['yest_2hr_renewal'] - partner_merged['today_renewal']
    else:
        partner_merged['Renewals yesterday - 2 hrs later'] = None
        partner_merged['Renewals needed in 2 hours'] = None
        
    partner_merged['Renewals Yesterday - 10:12PM'] = partner_merged['yest_10pm_renewal']
    
    output_partner_df = partner_merged[[
        'Region', 'Partner',
        'Renewals today',
        'Renewals in last 2 hours',
        'Renewals Yesterday - Same time',
        'Renewal Gap - Same time',
        'Renewals yesterday - 2 hrs later',
        'Renewals needed in 2 hours',
        'Renewals Yesterday - 10:12PM',
        'Target',
        'BTD',
        'today_prev_renewal'
    ]].copy()

    # -------------------------------------------------------------
    # BUILD OUTLETWISE DATASET
    # -------------------------------------------------------------
    print("Processing Outletwise dataset...")
    ot = load_outlet_df(o_path_today, "today")
    otp = load_outlet_df(o_path_prev, "today_prev")
    oys = load_outlet_df(o_path_ys, "yest_same")
    oy2 = load_outlet_df(o_path_y2, "yest_2hr")
    oy10 = load_outlet_df(o_path_y10, "yest_10pm")

    mo1 = pd.merge(ot, oys, on=['Partner', 'Outlet'], how='outer')
    mo2 = pd.merge(mo1, oy2, on=['Partner', 'Outlet'], how='outer')
    mo3 = pd.merge(mo2, oy10, on=['Partner', 'Outlet'], how='outer')
    outlet_renewals = pd.merge(mo3, otp, on=['Partner', 'Outlet'], how='outer')

    # Filter daily data to rows where outlet is present and not blank
    df_daily_outlets = df_daily[
        df_daily[outlet_col].notna() &
        (df_daily[outlet_col].astype(str).str.strip() != '') &
        (~df_daily[outlet_col].astype(str).str.strip().str.lower().isin(['none', 'nan']))
    ].copy()

    # Master base from daily CSV
    df_base = df_daily_outlets.groupby([region_col, partner_col, outlet_col], dropna=False).size().reset_index(name='Total_users')
    df_base = df_base.rename(columns={region_col: 'Region', partner_col: 'Partner', outlet_col: 'Outlet'})
    
    outlet_merged = pd.merge(df_base, outlet_renewals, on=['Partner', 'Outlet'], how='left')
    outlet_merged[renewal_cols] = outlet_merged[renewal_cols].fillna(0).astype(int)

    for col in ['Region', 'Partner', 'Outlet']:
        outlet_merged[col] = outlet_merged[col].astype(str).str.strip().replace('nan', '')

    outlet_merged['Total_users'] = outlet_merged['Total_users'].fillna(0).astype(int)
    outlet_merged['Target'] = (outlet_merged['Total_users'] / 3.0).round().astype(int)
    outlet_merged['BTD'] = outlet_merged['Target'] - outlet_merged['today_renewal']

    outlet_merged['Renewals today'] = outlet_merged['today_renewal']
    outlet_merged['Renewals in last 2 hours'] = outlet_merged['today_renewal'] - outlet_merged['today_prev_renewal']
    outlet_merged['Renewals Yesterday - Same time'] = outlet_merged['yest_same_renewal']
    outlet_merged['Renewal Gap - Same time'] = outlet_merged['today_renewal'] - outlet_merged['yest_same_renewal']

    if next_time_slot:
        outlet_merged['Renewals yesterday - 2 hrs later'] = outlet_merged['yest_2hr_renewal']
        outlet_merged['Renewals needed in 2 hours'] = outlet_merged['yest_2hr_renewal'] - outlet_merged['today_renewal']
    else:
        outlet_merged['Renewals yesterday - 2 hrs later'] = None
        outlet_merged['Renewals needed in 2 hours'] = None

    outlet_merged['Renewals Yesterday - 10:12PM'] = outlet_merged['yest_10pm_renewal']

    output_outlet_df = outlet_merged[[
        'Region', 'Partner', 'Outlet',
        'Renewals today',
        'Renewals in last 2 hours',
        'Renewals Yesterday - Same time',
        'Renewal Gap - Same time',
        'Renewals yesterday - 2 hrs later',
        'Renewals needed in 2 hours',
        'Renewals Yesterday - 10:12PM',
        'Target',
        'BTD',
        'today_prev_renewal'
    ]].copy().rename(columns={'Region': 'region'})

    # Keep only rows where Outlet is available (not empty, not None, not nan)
    output_outlet_df = output_outlet_df[
        output_outlet_df['Outlet'].notna() &
        (output_outlet_df['Outlet'].astype(str).str.strip() != '') &
        (~output_outlet_df['Outlet'].astype(str).str.strip().str.lower().isin(['none', 'nan']))
    ].copy()

    # -------------------------------------------------------------
    # EXPORT RAW TO MULTI-SHEET EXCEL
    # -------------------------------------------------------------
    temp_excel = os.path.join(output_dir, "renewals_comparison_temp.xlsx")
    with pd.ExcelWriter(temp_excel, engine='openpyxl') as writer:
        output_partner_df.to_excel(writer, sheet_name="Data_Partner", index=False)
        output_outlet_df.to_excel(writer, sheet_name="Data_Outlet", index=False)

    final_xlsm = os.path.join(output_dir, "renewals_comparison.xlsm")
    create_macro_dashboard(temp_excel, final_xlsm)

    # Generate standalone web dashboard for GitHub Pages
    generate_web_dashboard(output_partner_df, output_outlet_df, target_date, target_time_slot, next_time_slot, output_dir)

    # Cleanup temp files
    if os.path.exists(temp_excel):
        os.remove(temp_excel)
        
    for p in [p_path_today, p_path_prev, p_path_ys, p_path_y2, p_path_y10,
              o_path_today, o_path_prev, o_path_ys, o_path_y2, o_path_y10]:
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass
                
    try:
        os.rmdir(temp_dir)
    except Exception:
        pass

    print(f"\nSuccessfully generated interactive Excel macro dashboard: {final_xlsm}\n")
    
    # Automatically sync updated web dashboard to GitHub Pages
    sync_to_github()

if __name__ == "__main__":
    main()
