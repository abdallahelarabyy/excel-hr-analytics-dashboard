' Builds "Employees HR Dashboard.xlsx" from the raw Employees.xlsx using Excel automation.
' Usage: cscript //nologo build_hr.vbs <pdf preview path>
Option Explicit
Const SRC = "F:\Data Analysis\Digital Hub\Excel task\Raw Data\Employees.xlsx"
Const DST = "F:\Data Analysis\Digital Hub\Excel task\Output\Employees HR Dashboard.xlsx"
Dim PDF: PDF = WScript.Arguments(0)

Dim xl, wb, wsA, wsD, pc, stage
Dim ptKPI, ptDept, ptGender, ptCountry, ptHire, ptCenter, ptSal, ptRate, ptSum
Dim NAVY, BLUE, ORANGE, TEAL, GRAYTXT, BORDER, BG, WHITE
NAVY = RGB(31, 42, 68): BLUE = RGB(47, 110, 186): ORANGE = RGB(232, 131, 58)
TEAL = RGB(59, 162, 122): GRAYTXT = RGB(96, 106, 122): BORDER = RGB(218, 223, 230)
BG = RGB(242, 244, 247): WHITE = RGB(255, 255, 255)

Set xl = CreateObject("Excel.Application")
xl.Visible = False: xl.DisplayAlerts = False: xl.ScreenUpdating = False
On Error Resume Next
Main
If Err.Number <> 0 Then
    WScript.Echo "ERROR at stage [" & stage & "]: " & Err.Number & " " & Err.Description
    Err.Clear
    If IsObject(wb) Then wb.Close False
    xl.Quit
    WScript.Quit 1
End If
On Error GoTo 0
xl.Quit
WScript.Echo "DONE"

' ---------------------------------------------------------------- helpers
Sub AddCol(lo, nm, f)
    Dim c: Set c = lo.ListColumns.Add
    c.Name = nm
    c.DataBodyRange.Formula = f
End Sub

Function NewPT(anchor, nm, title)
    wsA.Range(anchor).Offset(-1, 0).Value = title
    wsA.Range(anchor).Offset(-1, 0).Font.Bold = True
    wsA.Range(anchor).Offset(-1, 0).Font.Color = NAVY
    Dim pt: Set pt = pc.CreatePivotTable(wsA.Range(anchor), nm, , 6)
    pt.RowAxisLayout 1
    pt.TableStyle2 = "PivotStyleLight16"
    Set NewPT = pt
End Function

Sub AddDF(pt, fld, cap, fn, fmt)
    Dim d: Set d = pt.AddDataField(pt.PivotFields(fld), cap, fn)
    d.NumberFormat = fmt
End Sub

Function CountPT(anchor, nm, title, rowField, sortDesc)
    Dim pt: Set pt = NewPT(anchor, nm, title)
    pt.PivotFields(rowField).Orientation = 1
    AddDF pt, "No", "Headcount", -4112, "#,##0"
    If sortDesc Then pt.PivotFields(rowField).AutoSort 2, "Headcount"
    Set CountPT = pt
End Function

Sub StyleChart(ch, title)
    ch.ShowAllFieldButtons = False
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    With ch.ChartTitle.Format.TextFrame2.TextRange.Font
        .Size = 12: .Bold = True: .Fill.ForeColor.RGB = NAVY
    End With
    ch.ChartArea.Format.Fill.ForeColor.RGB = WHITE
    ch.ChartArea.Format.Line.Visible = True
    ch.ChartArea.Format.Line.ForeColor.RGB = BORDER
    ch.ChartArea.Format.Line.Weight = 0.75
    ch.ChartArea.RoundedCorners = True
End Sub

Function NewChart(pt, ctype, l, t, w, h, title)
    Dim s: Set s = wsD.Shapes.AddChart2(-1, ctype, l, t, w, h)
    s.Chart.SetSourceData pt.TableRange1
    s.Placement = 3
    StyleChart s.Chart, title
    Set NewChart = s.Chart
End Function

Sub CleanAxes(ch)
    ch.HasLegend = False
    ch.Axes(2).HasMajorGridlines = False
    ch.Axes(2).Delete
    ch.Axes(1).TickLabels.Font.Size = 9
    ch.Axes(1).TickLabels.Font.Color = GRAYTXT
    ch.Axes(1).Format.Line.ForeColor.RGB = BORDER
End Sub

Sub SetLabels(ser, fmt)
    ser.HasDataLabels = True
    ser.DataLabels.NumberFormat = fmt
    ser.DataLabels.Font.Size = 9
    ser.DataLabels.Font.Color = NAVY
End Sub

Sub FmtText(tf, sz, bold, clr)
    With tf.TextRange.Font
        .Size = sz: .Bold = bold: .Fill.ForeColor.RGB = clr
    End With
End Sub

Function NewBox(l, t, w, h)
    Dim s: Set s = wsD.Shapes.AddTextbox(1, l, t, w, h)
    s.Fill.Visible = False: s.Line.Visible = False
    s.TextFrame2.MarginLeft = 0: s.TextFrame2.MarginRight = 0
    s.TextFrame2.MarginTop = 0: s.TextFrame2.MarginBottom = 0
    s.TextFrame2.WordWrap = True
    s.Placement = 3
    Set NewBox = s
End Function

Sub NewSlicer(field, cap, t, h, cols)
    Dim sc, sl, p
    Set sc = wb.SlicerCaches.Add2(ptKPI, field)
    For Each p In Array(ptDept, ptGender, ptCountry, ptHire, ptCenter, ptSal, ptRate, ptSum)
        sc.PivotTables.AddPivotTable p
    Next
    Set sl = sc.Slicers.Add(wsD, , "sl" & Replace(field, " ", ""), cap, t, 10, 170, h)
    sl.Style = "SlicerStyleLight1"
    sl.NumberOfColumns = cols
    sl.RowHeight = 18
    sl.Shape.Placement = 3
End Sub

' ---------------------------------------------------------------- main
Sub Main()
    Dim ws, lo, i, c, r, s, ch, area, x, cw, labels, refs, accents, a, g, tf, wsN, notes, cardW

    stage = "open"
    Set wb = xl.Workbooks.Open(SRC, 0, True)

    stage = "theme colours"
    With wb.Theme.ThemeColorScheme
        .Colors(3).RGB = NAVY
        .Colors(5).RGB = BLUE
        .Colors(6).RGB = ORANGE
        .Colors(7).RGB = TEAL
        .Colors(8).RGB = RGB(142, 108, 196)
        .Colors(9).RGB = RGB(217, 83, 79)
        .Colors(10).RGB = RGB(127, 140, 154)
    End With

    stage = "remove empty map sheet"
    wb.Worksheets("Sheet1").Delete
    For i = wb.Names.Count To 1 Step -1
        If Left(wb.Names(i).Name, 9) = "_xlchart." Then wb.Names(i).Delete
    Next

    ' ---------------- Data sheet
    stage = "data sheet"
    Set ws = wb.Worksheets("Employees")
    ws.Name = "Data"
    Set lo = ws.ListObjects("EmpTable")
    AddCol lo, "Full Name", "=[@[First Name]]&"" ""&[@[Last Name]]"
    AddCol lo, "Hire Year", "=YEAR([@[Start Date]])"
    AddCol lo, "Salary Band", "=LOOKUP([@[Monthly Salary]],{0,1000,1500,2000,2500,3000},{""0-999"",""1000-1499"",""1500-1999"",""2000-2499"",""2500-2999"",""3000+""})"
    AddCol lo, "Performance", "=IF([@[Job Rate]]>=4.5,""High (4.5-5)"",IF([@[Job Rate]]>=3,""Medium (3)"",""Low (1-2)""))"
    AddCol lo, "Total Leaves", "=[@[Sick Leaves]]+[@[Unpaid Leaves]]"
    AddCol lo, "Overtime Level", "=IF([@[Overtime Hours]]>50,""High (>50 h)"",""Normal (0-50 h)"")"
    lo.ListColumns("Start Date").DataBodyRange.NumberFormat = "dd-mmm-yyyy"
    lo.ListColumns("Monthly Salary").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Annual Salary").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Job Rate").DataBodyRange.NumberFormat = "0.0"
    lo.TableStyle = "TableStyleMedium2"
    lo.Range.Columns.AutoFit
    lo.HeaderRowRange.Font.Bold = True

    ' ---------------- Analysis sheet (pivots, one shared cache)
    stage = "pivot cache"
    Set wsA = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
    wsA.Name = "Analysis"
    wsA.Range("B1").Value = "Analysis - pivot tables that feed the Dashboard (all share one pivot cache and are filtered by the dashboard slicers)"
    wsA.Range("B1").Font.Bold = True: wsA.Range("B1").Font.Size = 13: wsA.Range("B1").Font.Color = NAVY
    Set pc = wb.PivotCaches.Create(1, "EmpTable", 6)

    stage = "pivot KPI"
    Set ptKPI = NewPT("B4", "ptKPI", "KPI pivot")
    AddDF ptKPI, "No", "Employees", -4112, "#,##0"
    AddDF ptKPI, "Annual Salary", "Total Payroll", -4157, "#,##0"
    AddDF ptKPI, "Monthly Salary", "Avg Monthly Salary", -4106, "#,##0"
    AddDF ptKPI, "Job Rate", "Avg Job Rate", -4106, "0.00"
    AddDF ptKPI, "Years", "Avg Tenure", -4106, "0.0"
    AddDF ptKPI, "Overtime Hours", "Total Overtime", -4157, "#,##0"
    ptKPI.PivotCache.RefreshOnFileOpen = True

    stage = "pivots"
    Set ptDept = CountPT("B19", "ptDept", "Headcount by Department", "Department", True)
    Set ptGender = CountPT("E19", "ptGender", "Gender split", "Gender", False)
    Set ptCountry = CountPT("H19", "ptCountry", "Headcount by Country", "Country", True)
    Set ptHire = CountPT("K19", "ptHire", "Hires by Start Year", "Hire Year", False)
    Set ptCenter = CountPT("N19", "ptCenter", "Headcount by Center", "Center", True)
    Set ptSal = CountPT("Q19", "ptSal", "Monthly salary bands", "Salary Band", False)
    Set ptRate = CountPT("T19", "ptRate", "Employees by Job Rate", "Job Rate", False)

    stage = "pivot department summary"
    Set ptSum = NewPT("W19", "ptDeptSummary", "Department summary")
    ptSum.PivotFields("Department").Orientation = 1
    AddDF ptSum, "No", "Headcount", -4112, "#,##0"
    AddDF ptSum, "Annual Salary", "Total Payroll", -4157, "#,##0"
    AddDF ptSum, "Monthly Salary", "Avg Monthly Salary", -4106, "#,##0"
    AddDF ptSum, "Job Rate", "Avg Job Rate", -4106, "0.00"
    AddDF ptSum, "Sick Leaves", "Avg Sick Leaves", -4106, "0.0"
    AddDF ptSum, "Unpaid Leaves", "Avg Unpaid Leaves", -4106, "0.0"
    AddDF ptSum, "Overtime Hours", "Avg Overtime Hours", -4106, "0.0"
    ptSum.PivotFields("Department").AutoSort 2, "Headcount"

    stage = "KPI formulas"
    a = ptKPI.TableRange1.Cells(1, 1).Address
    g = ptGender.TableRange1.Cells(1, 1).Address
    wsA.Range("B9").Value = "Dashboard KPIs (the KPI cards are linked to these cells)"
    wsA.Range("B9").Font.Bold = True: wsA.Range("B9").Font.Color = NAVY
    labels = Array("Employees", "Annual Payroll", "Avg Monthly Salary", "Avg Job Rate", "Avg Tenure (years)", "Female Share")
    wsA.Range("C10").Formula = "=IFERROR(GETPIVOTDATA(""Employees""," & a & "),0)"
    wsA.Range("C11").Formula = "=IFERROR(GETPIVOTDATA(""Total Payroll""," & a & "),0)"
    wsA.Range("C12").Formula = "=IFERROR(GETPIVOTDATA(""Avg Monthly Salary""," & a & "),0)"
    wsA.Range("C13").Formula = "=IFERROR(GETPIVOTDATA(""Avg Job Rate""," & a & "),0)"
    wsA.Range("C14").Formula = "=IFERROR(GETPIVOTDATA(""Avg Tenure""," & a & "),0)"
    wsA.Range("C15").Formula = "=IFERROR(IFERROR(GETPIVOTDATA(""Headcount""," & g & ",""Gender"",""Female""),0)/C10,0)"
    wsA.Range("C10").NumberFormat = "#,##0"
    wsA.Range("C11").NumberFormat = "#,##0.00,,""M"""
    wsA.Range("C12").NumberFormat = "#,##0"
    wsA.Range("C13").NumberFormat = "0.00"" / 5"""
    wsA.Range("C14").NumberFormat = "0.0"" yrs"""
    wsA.Range("C15").NumberFormat = "0.0%"
    For i = 0 To 5
        wsA.Range("B" & (10 + i)).Value = labels(i)
    Next
    wsA.Range("B10:C15").Borders.Color = BORDER
    wsA.Columns("A").ColumnWidth = 2
    wsA.Range("B3:AE60").Columns.AutoFit
    wsA.Columns("B").ColumnWidth = 24

    ' ---------------- Dashboard
    stage = "dashboard canvas"
    Set wsD = wb.Worksheets.Add(wb.Worksheets(1))
    wsD.Name = "Dashboard"
    c = 1: Do While wsD.Cells(1, c).Left < 1112: c = c + 1: Loop
    r = 1: Do While wsD.Cells(r, 1).Top < 922: r = r + 1: Loop
    Set area = wsD.Range(wsD.Cells(1, 1), wsD.Cells(r, c))
    area.Interior.Color = BG

    stage = "banner"
    Set s = wsD.Shapes.AddShape(1, 10, 10, 1090, 52)
    s.Fill.ForeColor.RGB = NAVY: s.Line.Visible = False: s.Placement = 3
    Set tf = s.TextFrame2
    tf.MarginLeft = 16: tf.VerticalAnchor = 3
    tf.TextRange.Text = "Employee HR Dashboard" & vbCr & "Workforce overview  |  689 employees in 5 countries and 20 departments  |  Use the slicers on the left to filter every KPI and chart"
    tf.TextRange.ParagraphFormat.Alignment = 1
    With tf.TextRange.Paragraphs(1).Font
        .Size = 20: .Bold = True: .Fill.ForeColor.RGB = WHITE
    End With
    With tf.TextRange.Paragraphs(2).Font
        .Size = 10: .Bold = False: .Fill.ForeColor.RGB = RGB(200, 210, 228)
    End With

    stage = "KPI cards"
    refs = Array("$C$10", "$C$11", "$C$12", "$C$13", "$C$14", "$C$15")
    labels = Array("EMPLOYEES", "ANNUAL PAYROLL", "AVG MONTHLY SALARY", "AVG JOB RATE", "AVG TENURE", "FEMALE SHARE")
    accents = Array(BLUE, TEAL, BLUE, ORANGE, TEAL, ORANGE)
    cardW = (910 - 50) / 6
    For i = 0 To 5
        x = 190 + i * (cardW + 10)
        Set s = wsD.Shapes.AddShape(5, x, 70, cardW, 68)
        s.Adjustments.Item(1) = 0.08
        s.Fill.ForeColor.RGB = WHITE: s.Line.ForeColor.RGB = BORDER: s.Line.Weight = 0.75
        s.Placement = 3
        Set s = wsD.Shapes.AddShape(1, x + 8, 80, 4, 48)
        s.Fill.ForeColor.RGB = accents(i): s.Line.Visible = False: s.Placement = 3
        Set s = NewBox(x + 20, 81, cardW - 26, 14)
        s.TextFrame2.TextRange.Text = labels(i)
        FmtText s.TextFrame2, 8.5, True, GRAYTXT
        Set s = NewBox(x + 20, 97, cardW - 26, 30)
        s.DrawingObject.Formula = "=Analysis!" & refs(i)
        FmtText s.TextFrame2, 20, True, NAVY
    Next

    stage = "charts"
    cw = (910 - 20) / 3
    Set ch = NewChart(ptDept, 57, 190, 148, cw, 510, "Headcount by Department")
    CleanAxes ch
    ch.Axes(1).ReversePlotOrder = True
    ch.ChartGroups(1).GapWidth = 40
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = BLUE
    SetLabels ch.SeriesCollection(1), "#,##0"

    stage = "chart gender"
    Set ch = NewChart(ptGender, -4120, 190 + cw + 10, 148, cw, 250, "Gender Split")
    ch.ChartGroups(1).DoughnutHoleSize = 58
    ch.HasLegend = True: ch.Legend.Position = -4107
    ch.Legend.Font.Size = 9
    With ch.SeriesCollection(1)
        .HasDataLabels = True
        .DataLabels.ShowValue = False
        .DataLabels.ShowPercentage = True
        .DataLabels.NumberFormat = "0.0%"
        .DataLabels.Font.Size = 10: .DataLabels.Font.Bold = True
        .DataLabels.Font.Color = WHITE
        .Format.Line.ForeColor.RGB = WHITE
    End With

    stage = "chart country"
    Set ch = NewChart(ptCountry, 57, 190 + 2 * (cw + 10), 148, cw, 250, "Headcount by Country")
    CleanAxes ch
    ch.Axes(1).ReversePlotOrder = True
    ch.ChartGroups(1).GapWidth = 50
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = BLUE
    SetLabels ch.SeriesCollection(1), "#,##0"

    stage = "chart hires"
    Set ch = NewChart(ptHire, 65, 190 + cw + 10, 408, cw, 250, "Hires by Start Year")
    CleanAxes ch
    With ch.SeriesCollection(1)
        .Format.Line.ForeColor.RGB = BLUE
        .Format.Line.Weight = 2.5
        .MarkerStyle = 8: .MarkerSize = 7
        .MarkerBackgroundColor = BLUE: .MarkerForegroundColor = BLUE
    End With
    SetLabels ch.SeriesCollection(1), "#,##0"
    ch.SeriesCollection(1).DataLabels.Position = 0

    stage = "chart center"
    Set ch = NewChart(ptCenter, 51, 190 + 2 * (cw + 10), 408, cw, 250, "Headcount by Center")
    CleanAxes ch
    ch.ChartGroups(1).GapWidth = 60
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = BLUE
    SetLabels ch.SeriesCollection(1), "#,##0"

    stage = "chart salary"
    Set ch = NewChart(ptSal, 51, 190, 668, cw, 242, "Monthly Salary Distribution")
    CleanAxes ch
    ch.ChartGroups(1).GapWidth = 25
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = TEAL
    SetLabels ch.SeriesCollection(1), "#,##0"

    stage = "chart rate"
    Set ch = NewChart(ptRate, 51, 190 + cw + 10, 668, cw, 242, "Employees by Job Rate (1-5)")
    CleanAxes ch
    ch.ChartGroups(1).GapWidth = 50
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = ORANGE
    SetLabels ch.SeriesCollection(1), "#,##0"

    stage = "insights box"
    x = 190 + 2 * (cw + 10)
    Set s = wsD.Shapes.AddShape(5, x, 668, cw, 242)
    s.Adjustments.Item(1) = 0.04
    s.Fill.ForeColor.RGB = WHITE: s.Line.ForeColor.RGB = BORDER: s.Line.Weight = 0.75
    s.Placement = 3
    Set tf = s.TextFrame2
    tf.MarginLeft = 12: tf.MarginRight = 10: tf.MarginTop = 10: tf.VerticalAnchor = 1
    tf.TextRange.Text = "Key Insights (all 689 employees)" & vbCr & _
        "Egypt holds 55% of staff; Manufacturing is the largest department (140, 20%)." & vbCr & _
        "4 departments (Manufacturing, Quality Control, Account Mgmt, Quality Assurance) hold 55% of headcount." & vbCr & _
        "65% of staff are male, but average pay is almost equal (2,073 vs 2,059, a 0.7% gap)." & vbCr & _
        "Hiring peaked in 2019 with 248 hires (36%), then fell to 162 in 2020." & vbCr & _
        "Pay is not linked to job rate or tenure (correlation ~0): rate-5 staff average 2,009 vs 2,076 for rate 1." & vbCr & _
        "52 employees (7.5%) work 50+ overtime hours and generate 52% of all 9,441 overtime hours." & vbCr & _
        "49% of staff are rated 4.5-5, while 21% (142) are rated 1-2 and may need performance plans."
    tf.TextRange.ParagraphFormat.Alignment = 1
    With tf.TextRange.Font
        .Size = 9: .Bold = False: .Fill.ForeColor.RGB = NAVY
    End With
    tf.TextRange.ParagraphFormat.SpaceAfter = 3
    For i = 2 To 8
        tf.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Visible = True
        tf.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Character = 8226
        tf.TextRange.Paragraphs(i).ParagraphFormat.LeftIndent = 10
        tf.TextRange.Paragraphs(i).ParagraphFormat.FirstLineIndent = -10
    Next
    With tf.TextRange.Paragraphs(1).Font
        .Size = 12: .Bold = True
    End With

    stage = "slicers"
    NewSlicer "Gender", "Gender", 70, 56, 2
    NewSlicer "Center", "Center", 134, 100, 2
    NewSlicer "Country", "Country", 242, 142, 1
    NewSlicer "Hire Year", "Hire Year", 392, 80, 3
    NewSlicer "Department", "Department", 480, 430, 1

    ' ---------------- Notes sheet
    stage = "notes sheet"
    Set wsN = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
    wsN.Name = "Notes"
    notes = Array( _
        "T|Workbook notes", _
        "H|Structure", _
        "|Dashboard - KPI cards, 7 pivot charts, key insights and 5 slicers (Gender, Center, Country, Hire Year, Department) connected to every pivot.", _
        "|Data - the original EmpTable (689 employees) with 6 helper columns added at the right.", _
        "|Analysis - 9 pivot tables built on one shared pivot cache, plus the KPI cells the dashboard cards link to.", _
        "H|Data quality checks (raw file)", _
        "|689 rows x 15 columns; no blank cells, no leading/trailing or double spaces, consistent category spelling.", _
        "|Annual Salary = 12 x Monthly Salary on every row. Employee No is unique (1-689).", _
        "|14 full names appear more than once, but each has a different start date, department and salary - different people, so nothing was removed.", _
        "|The raw file's Sheet1 held an empty map chart pointing at a blank cell; it was removed.", _
        "|Years = DATEDIF(Start Date, TODAY(), ""Y"") is kept as in the source, so tenure updates each day; the pivot cache refreshes when the file opens.", _
        "|Overtime Hours is heavily skewed: median 7 h, 52 employees above 50 h, maximum 198 h (flagged in the Overtime Level column).", _
        "H|Helper columns added", _
        "|Full Name, Hire Year, Salary Band (monthly), Performance (Low 1-2 / Medium 3 / High 4.5-5), Total Leaves (sick + unpaid), Overtime Level (>50 h).", _
        "H|Key insights (all employees)", _
        "|Headcount 689; annual payroll 17,099,892; average monthly salary 2,068; average job rate 3.59; 35% female.", _
        "|Egypt 379 (55%), UAE 156, Saudi Arabia 90, Syria 53, Lebanon 11. Main center is the largest (251).", _
        "|Manufacturing (140), Quality Control (89), Account Management (84) and Quality Assurance (67) = 55% of staff.", _
        "|Male vs female average monthly salary: 2,073 vs 2,059 (0.7% gap) - pay is effectively equal by gender.", _
        "|Hires by year: 2016 51, 2017 77, 2018 151, 2019 248 (peak), 2020 162.", _
        "|Salary shows no relationship with job rate, tenure or overtime (all correlations between -0.08 and 0.04).", _
        "|49% of staff are rated 4.5-5, while 21% (142) are rated 1-2 - a group to review for performance plans.", _
        "|52 employees (7.5%) generate 52% of all overtime hours; Quality Control and Account Management have the most of them (9 each).", _
        "|Highest average pay: Human Resources (2,556) and Environmental Compliance (2,508); lowest: Research Center (1,887).", _
        "H|How to use", _
        "|Click slicer buttons on the Dashboard (Ctrl+click for several); use the funnel-x icon on a slicer to clear it.", _
        "|If the data changes, use Data > Refresh All to update every pivot, chart and KPI.")
    r = 2
    For i = 0 To UBound(notes)
        a = Split(notes(i), "|", 2)
        wsN.Cells(r, 2).Value = a(1)
        If a(0) = "T" Then
            wsN.Cells(r, 2).Font.Size = 16: wsN.Cells(r, 2).Font.Bold = True: wsN.Cells(r, 2).Font.Color = NAVY
            r = r + 1
        ElseIf a(0) = "H" Then
            r = r + 1
            wsN.Cells(r, 2).Value = a(1): wsN.Cells(r - 1, 2).ClearContents
            wsN.Cells(r, 2).Font.Bold = True: wsN.Cells(r, 2).Font.Size = 12: wsN.Cells(r, 2).Font.Color = BLUE
        Else
            wsN.Cells(r, 2).Value = "- " & a(1)
        End If
        r = r + 1
    Next
    wsN.Columns("A").ColumnWidth = 2
    wsN.Columns("B").ColumnWidth = 130

    ' ---------------- windows, print setup, preview, save
    stage = "window settings"
    ws.Activate
    xl.ActiveWindow.SplitRow = 1: xl.ActiveWindow.SplitColumn = 0
    xl.ActiveWindow.FreezePanes = True
    wsA.Activate: xl.ActiveWindow.DisplayGridlines = False
    wsN.Activate: xl.ActiveWindow.DisplayGridlines = False
    wsD.Activate
    xl.ActiveWindow.DisplayGridlines = False
    xl.ActiveWindow.DisplayHeadings = False
    xl.ActiveWindow.Zoom = 85
    wsD.Range("A1").Select
    wsD.Tab.Color = NAVY

    stage = "page setup"
    With wsD.PageSetup
        .PrintArea = area.Address
        .Orientation = 2
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = 18: .RightMargin = 18: .TopMargin = 18: .BottomMargin = 18
        .CenterHorizontally = True
    End With

    stage = "save"
    xl.ScreenUpdating = True
    wb.SaveAs DST, 51
    stage = "pdf preview"
    wsD.ExportAsFixedFormat 0, PDF
    wb.Close False
End Sub
