# ==============================================================================
# modules/Export-FinalResultSheet.ps1
# Institutional Multi-Sheet Result Generator (CW, End Sem, Nptel Sheet)
# Generates SGSITS Indore 70 / 30 Scheme Workbooks for College & University Grade Submission
# ==============================================================================

if (-not (Get-Command 'Get-CourseExamResultsStore' -ErrorAction SilentlyContinue)) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Join-Path (Get-Location) "modules" }
    $storeScript = Join-Path $scriptDir "Import-StudentSheet.ps1"
    if (Test-Path -LiteralPath $storeScript) {
        . $storeScript
    }
}

function Get-CourseCollegeResultSheetsFilePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CourseId,
        [string]$CourseName = $null,
        [string]$DataDir = $null
    )

    if (-not $DataDir) {
        $appRoot = Split-Path -Parent $PSScriptRoot
        $DataDir = Join-Path $appRoot "data"
    }

    $coursesParent = Join-Path $DataDir "Courses"
    $cleanCourseName = if ($CourseName) { ($CourseName -replace '[\\/:*?"<>|]', '_').Trim() } else { $CourseId }
    $courseDir = Join-Path $coursesParent $cleanCourseName

    # Institutional sheets stored in dedicated 'Result_Sheets' folder under the course
    $resultSheetsDir = Join-Path $courseDir "Result_Sheets"
    if (-not (Test-Path -LiteralPath $resultSheetsDir)) {
        $null = New-Item -ItemType Directory -Path $resultSheetsDir -Force
    }

    return (Join-Path $resultSheetsDir "${cleanCourseName}_College_Result_Sheets.xlsx")
}

function Get-CourseCollegeSeparateSheetFilePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CourseId,
        [string]$CourseName = $null,
        [Parameter(Mandatory = $true)]
        [ValidateSet('CW', 'EndSem', 'End sem', 'End Sem', 'Cumulative', 'Cumulative Sheet', 'Main Result', 'Nptel', 'NPTEL', 'Nptel Sheet')]
        [string]$SheetType,
        [string]$DataDir = $null
    )

    if (-not $DataDir) {
        $appRoot = Split-Path -Parent $PSScriptRoot
        $DataDir = Join-Path $appRoot "data"
    }

    $coursesParent = Join-Path $DataDir "Courses"
    $cleanCourseName = if ($CourseName) { ($CourseName -replace '[\\/:*?"<>|]', '_').Trim() } else { $CourseId }
    $courseDir = Join-Path $coursesParent $cleanCourseName

    # Institutional sheets stored in dedicated 'Result_Sheets' folder under the course
    $resultSheetsDir = Join-Path $courseDir "Result_Sheets"
    if (-not (Test-Path -LiteralPath $resultSheetsDir)) {
        $null = New-Item -ItemType Directory -Path $resultSheetsDir -Force
    }

    $suffix = switch ($SheetType) {
        'CW'                                                      { "CW" }
        { $_ -in 'EndSem', 'End sem', 'End Sem' }                 { "End_Sem" }
        { $_ -in 'Cumulative', 'Cumulative Sheet', 'Main Result' }{ "Cumulative" }
        { $_ -in 'Nptel', 'Nptel Sheet', 'NPTEL' }                { "NPTEL" }
        default                                                   { $SheetType }
    }

    return (Join-Path $resultSheetsDir "${cleanCourseName}_${suffix}.xlsx")
}

function Test-CollegeFileLocked {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($stream) {
            $stream.Close()
            $stream.Dispose()
        }
        return $false
    } catch {
        return $true
    }
}

function Get-CourseCollegeSheetData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Course,
        [Parameter(Mandatory = $true)]
        [ValidateSet('CW', 'EndSem', 'End sem', 'End Sem', 'Cumulative', 'Cumulative Sheet', 'Main Result', 'Nptel', 'NPTEL', 'Nptel Sheet')]
        [string]$SheetType,
        [string]$SessionName = "Jan-June 2026",
        [string]$DataDir = $null
    )

    $cId   = [string]$Course.Id
    $cName = if ($Course.Name) { [string]$Course.Name } else { "Course" }
    $cCode = if ($Course.Code) { [string]$Course.Code } else { "" }
    $sem   = if ($Course.Semester) { [string]$Course.Semester } else { "8" }

    # Verify Result Verification Sheet exists for course
    $vSheetPath = if ($Course.ExamVerificationSheet) { [string]$Course.ExamVerificationSheet } else { "" }
    if (-not $vSheetPath -or -not (Test-Path -LiteralPath $vSheetPath -ErrorAction SilentlyContinue)) {
        return [PSCustomObject]@{
            Header1  = "SHRI G. S. INSTITUTE OF TECHNOLOGY & SCIENCE, INDORE"
            Header2  = "SESSION : $SessionName"
            Header3  = ""
            Title    = ""
            Subtitle = ""
            Rows     = @()
        }
    }

    $store = Get-CourseExamResultsStore -CourseId $cId -CourseName $cName -DataDir $DataDir
    $allStudents = if ($store -and $store.Students) { @($store.Students) } else { @() }
    $verifiedStudents = @($allStudents | Where-Object { $_.ExamVerificationStatus -eq 'Verified' })

    # Sort alphabetically by RollNo
    $sortedStudents = @($verifiedStudents | Sort-Object { [string]$_.RollNo })

    # Resolve dynamic course metadata with safe fallbacks
    $collegeName = if ($Course.CollegeName) { [string]$Course.CollegeName } else { "SHRI G. S. INSTITUTE OF TECHNOLOGY & SCIENCE, INDORE" }
    $sessionStr  = if ($Course.Session) { [string]$Course.Session } elseif ($SessionName) { [string]$SessionName } else { "Jan-June 2026" }
    $yearStr     = if ($Course.AcademicYear) { [string]$Course.AcademicYear } else {
        $semInt = 8
        [int]::TryParse($sem, [ref]$semInt) | Out-Null
        switch ($semInt) {
            { $_ -in 7, 8 } { "IV YEAR" }
            { $_ -in 5, 6 } { "III YEAR" }
            { $_ -in 3, 4 } { "II YEAR" }
            { $_ -in 1, 2 } { "I YEAR" }
            default         { "IV YEAR" }
        }
    }
    $electiveStr = if ($Course.ElectiveType) { [string]$Course.ElectiveType } else { "Elective VI" }
    $tagSuffix   = if ($Course.HeaderTag) { [string]$Course.HeaderTag } else { "(NPTEL)" }
    $tagTitle    = if ($tagSuffix) { " $tagSuffix" } else { "" }

    $header1 = $collegeName
    $header2 = "SESSION : $sessionStr"
    $codeTitle = if ($cCode) { "$cCode " } else { "" }

    $title = ""
    $subtitle = ""
    $header3 = ""
    $rows = [System.Collections.ArrayList]@()

    $srNo = 1
    foreach ($st in $sortedStudents) {
        # Assignment marks: prefer verified, fallback to declared
        $assignRaw = if ($st.VerifiedAssignmentMarks) { [string]$st.VerifiedAssignmentMarks } elseif ($st.DeclaredAssignmentMarks) { [string]$st.DeclaredAssignmentMarks } else { "0" }
        $assignVal = 0.0
        [double]::TryParse($assignRaw, [ref]$assignVal) | Out-Null

        # Exam marks: prefer verified, fallback to declared
        $examRaw = if ($st.VerifiedExamMarks) { [string]$st.VerifiedExamMarks } elseif ($st.DeclaredExamMarks) { [string]$st.DeclaredExamMarks } else { "0" }
        $examVal = 0.0
        [double]::TryParse($examRaw, [ref]$examVal) | Out-Null

        if ($SheetType -eq 'CW') {
            # Subsheet 2: CW - Sr. No. | Enrollment | Name | CW(30)
            $cw30 = [Math]::Round(($assignVal / 25.0 * 30.0), 0, [System.MidpointRounding]::AwayFromZero)

            $null = $rows.Add([PSCustomObject]@{
                "Sr. No."    = $srNo
                "Enrollment" = [string]$st.RollNo
                "Name"       = [string]$st.Name
                "CW(30)"     = [int]$cw30
            })
        }
        elseif ($SheetType -in 'EndSem', 'End sem', 'End Sem') {
            # Subsheet 3: End Sem - Sr. No. | Enrollment | Name | End Sem(70)
            $marks70 = [Math]::Round(($examVal / 75.0 * 70.0), 0, [System.MidpointRounding]::AwayFromZero)

            $null = $rows.Add([PSCustomObject]@{
                "Sr. No."      = $srNo
                "Enrollment"   = [string]$st.RollNo
                "Name"         = [string]$st.Name
                "End Sem(70)"  = [int]$marks70
            })
        }
        elseif ($SheetType -in 'Cumulative', 'Cumulative Sheet', 'Main Result') {
            # Subsheet 1: Main Result - Sr. No. | Enrollment | Name | Assignment(25) | CW(30) | Final Exam(75) | End Sem (70) | Result | Total(30+70)
            $cw30 = [Math]::Round(($assignVal / 25.0 * 30.0), 0, [System.MidpointRounding]::AwayFromZero)
            $marks70 = [Math]::Round(($examVal / 75.0 * 70.0), 0, [System.MidpointRounding]::AwayFromZero)
            $tot = [int]($cw30 + $marks70)
            $passFail = if ($examVal -ge 30 -and $tot -ge 40) { "PASS" } else { "FAIL" }

            $null = $rows.Add([PSCustomObject]@{
                "Sr. No."          = $srNo
                "Enrollment"       = [string]$st.RollNo
                "Name"             = [string]$st.Name
                "Assignment(25)"   = $assignVal
                "CW(30)"           = [int]$cw30
                "Final Exam(75)"   = $examVal
                "End Sem (70)"     = [int]$marks70
                "Result"           = $passFail
                "Total(30+70)"     = $tot
            })
        }
        elseif ($SheetType -in 'Nptel', 'NPTEL', 'Nptel Sheet') {
            # Subsheet 4: NPTEL - Sr. No. | Enrollment | Name | Assignment | Total
            $totalRaw = if ($st.VerifiedTotalMarks) { [string]$st.VerifiedTotalMarks } elseif ($st.DeclaredTotalMarks) { [string]$st.DeclaredTotalMarks } else { "" }
            $totalVal = 0.0
            if (-not [double]::TryParse($totalRaw, [ref]$totalVal) -or $totalVal -le 0) {
                $totalVal = $assignVal + $examVal
            }

            $null = $rows.Add([PSCustomObject]@{
                "Sr. No."     = $srNo
                "Enrollment"  = [string]$st.RollNo
                "Name"        = [string]$st.Name
                "Assignment"  = $assignVal
                "Total"       = $totalVal
            })
        }
        $srNo++
    }

    if ($SheetType -eq 'CW') {
        $title    = "Course Work (CW) Sheet Preview"
        $subtitle = "Showing $($sortedStudents.Count) Verified student records $([string][char]0x2022) SGSITS 70/30 Scheme $([string][char]0x2022) 30 Marks"
        $header3  = "$yearStr CW  $electiveStr  ${codeTitle}${cName}${tagTitle}"
    }
    elseif ($SheetType -in 'EndSem', 'End sem', 'End Sem') {
        $title    = "End Sem Theory Sheet Preview"
        $subtitle = "Showing $($sortedStudents.Count) Verified student records $([string][char]0x2022) SGSITS 70/30 Scheme $([string][char]0x2022) 70 Marks"
        $header3  = "$yearStr End Sem  $electiveStr  ${codeTitle}${cName}${tagTitle}"
    }
    elseif ($SheetType -in 'Cumulative', 'Cumulative Sheet', 'Main Result') {
        $title    = "Cumulative Result Sheet Preview"
        $subtitle = "Showing $($sortedStudents.Count) Verified student records $([string][char]0x2022) Master Evaluation Ledger $([string][char]0x2022) 100 Marks"
        $header3  = "$yearStr Cumulative  $electiveStr  ${codeTitle}${cName}${tagTitle}"
    }
    elseif ($SheetType -in 'Nptel', 'NPTEL', 'Nptel Sheet') {
        $title    = "NPTEL Raw Certification Sheet Preview"
        $subtitle = "Showing $($sortedStudents.Count) Verified student records $([string][char]0x2022) NPTEL Official Raw Scores"
        $header3  = "$yearStr NPTEL  $electiveStr  ${codeTitle}${cName}${tagTitle}"
    }

    return [PSCustomObject]@{
        SheetType      = $SheetType
        Title          = $title
        Subtitle       = $subtitle
        Header1        = $header1
        Header2        = $header2
        Header3        = $header3
        Rows           = @($rows)
        TotalVerified  = $sortedStudents.Count
        TotalStudents  = $allStudents.Count
        UnderReview    = ($allStudents.Count - $sortedStudents.Count)
    }
}

function Format-AndPopulateWorksheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,
        [Parameter(Mandatory = $true)]
        [string]$SheetName,
        [Parameter(Mandatory = $true)]
        $SortedStudents,
        [string]$CollegeName = "SHRI G. S. INSTITUTE OF TECHNOLOGY & SCIENCE, INDORE",
        [string]$SessionStr = "Jan-June 2026",
        [string]$YearStr = "IV YEAR",
        [string]$ElectiveStr = "Elective VI",
        [string]$CodeTitle = "",
        [string]$CName = "Course",
        [string]$TagTitle = " (NPTEL)"
    )

    $ws = $Worksheet

    # Determine column headers and count
    $headers = switch ($SheetName) {
        'CW'                                                      { @("Sr. No.", "Enrollment", "Name", "CW(30)") }
        { $_ -in 'End Sem', 'End sem', 'EndSem' }                 { @("Sr. No.", "Enrollment", "Name", "End Sem(70)") }
        { $_ -in 'Main Result', 'Cumulative', 'Cumulative Sheet' }{ @("Sr. No.", "Enrollment", "Name", "Assignment(25)", "CW(30)", "Final Exam(75)", "End Sem (70)", "Result", "Total(30+70)") }
        { $_ -in 'NPTEL', 'Nptel', 'Nptel Sheet' }                { @("Sr. No.", "Enrollment", "Name", "Assignment", "Total") }
    }
    $colCount = [int]$headers.Count

    # Institutional Header Rows 1 to 3
    $ws.SetValue(1, 1, $CollegeName)
    $ws.SetValue(2, 1, "SESSION : $SessionStr")

    $h3 = switch -Wildcard ($SheetName) {
        'CW*'                                     { "$YearStr CW  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
        'End*sem*'                                { "$YearStr End Sem  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
        'EndSem*'                                 { "$YearStr End Sem  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
        'End Sem*'                                { "$YearStr End Sem  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
        'Main Result*'                            { "$YearStr Cumulative  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
        'Cumulative*'                             { "$YearStr Cumulative  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
        'NPTEL*'                                  { "$YearStr NPTEL  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
        'Nptel*'                                  { "$YearStr NPTEL  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
        default                                   { "$YearStr Result  $ElectiveStr  ${CodeTitle}${CName}${TagTitle}" }
    }
    $ws.SetValue(3, 1, $h3)

    # Merge Rows 1 to 3 across columns 1..colCount and apply centered styling
    # Compute adaptive font size based on column count and text length to guarantee zero truncation
    $h1Size = 12
    $h2Size = 11
    $h3Size = 11

    if ($colCount -le 4) {
        $h1Size = if ($CollegeName.Length -gt 55) { 10 } elseif ($CollegeName.Length -gt 45) { 10.5 } else { 11 }
        $h2Size = 10
        $h3Size = if ($h3.Length -gt 55) { 9.5 } elseif ($h3.Length -gt 45) { 10 } else { 10.5 }
    }
    elseif ($colCount -eq 5) {
        $h1Size = if ($CollegeName.Length -gt 55) { 10.5 } else { 11.5 }
        $h2Size = 10.5
        $h3Size = if ($h3.Length -gt 55) { 10 } else { 10.5 }
    }

    for ($hr = 1; $hr -le 3; $hr++) {
        $hRange = $ws.Cells[$hr, 1, $hr, $colCount]
        $hRange.Merge = $true
        $hRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $hRange.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $hRange.Style.Font.Name           = "Times New Roman"
        $hRange.Style.Font.Bold           = $true
        $hRange.Style.ShrinkToFit         = $true
        $hRange.Style.Font.Size           = switch ($hr) {
            1 { $h1Size }
            2 { $h2Size }
            3 { $h3Size }
        }
    }

    # Set executive row heights
    $ws.Row(1).Height = 24
    $ws.Row(2).Height = 18
    $ws.Row(3).Height = 18
    $ws.Row(4).Height = 10
    $ws.Row(5).Height = 26

    # Column Headers on Row 5
    for ($colIdx = 0; $colIdx -lt $headers.Count; $colIdx++) {
        $cInt = [int]($colIdx + 1)
        $ws.SetValue(5, $cInt, $headers[$colIdx])
        $cell = $ws.Cells[5, $cInt]
        $cell.Style.Font.Name = "Times New Roman"
        $cell.Style.Font.Size = 11
        $cell.Style.Font.Bold = $true
        $cell.Style.WrapText  = $true
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    }

    # Populate student rows starting at Row 6
    $currRow = 6
    $srNum = 1
    foreach ($st in $SortedStudents) {
        $rInt = [int]$currRow
        $ws.Row($rInt).Height = 20

        $assignRaw = if ($st.VerifiedAssignmentMarks) { [string]$st.VerifiedAssignmentMarks } elseif ($st.DeclaredAssignmentMarks) { [string]$st.DeclaredAssignmentMarks } else { "0" }
        $assignVal = 0.0
        [double]::TryParse($assignRaw, [ref]$assignVal) | Out-Null

        $examRaw = if ($st.VerifiedExamMarks) { [string]$st.VerifiedExamMarks } elseif ($st.DeclaredExamMarks) { [string]$st.DeclaredExamMarks } else { "0" }
        $examVal = 0.0
        [double]::TryParse($examRaw, [ref]$examVal) | Out-Null

        $ws.SetValue($rInt, 1, $srNum)
        $ws.SetValue($rInt, 2, [string]$st.RollNo)
        $ws.SetValue($rInt, 3, [string]$st.Name)

        if ($SheetName -in 'Main Result', 'Cumulative', 'Cumulative Sheet') {
            # Subsheet 1: Main Result (9 Columns)
            # Col 4: Assignment(25)
            $ws.SetValue($rInt, 4, $assignVal)
            $ws.Cells[$rInt, 4].Style.Numberformat.Format = "0.##"

            # Col 5: CW(30) - Formula: =ROUND(D{r}/25*30, 0)
            $ws.Cells[$rInt, 5].Formula = "ROUND(D${rInt}/25*30, 0)"
            $ws.Cells[$rInt, 5].Style.Numberformat.Format = "0"

            # Col 6: Final Exam(75)
            $ws.SetValue($rInt, 6, $examVal)
            $ws.Cells[$rInt, 6].Style.Numberformat.Format = "0.##"

            # Col 7: End Sem (70) - Formula: =ROUND(F{r}/75*70, 0)
            $ws.Cells[$rInt, 7].Formula = "ROUND(F${rInt}/75*70, 0)"
            $ws.Cells[$rInt, 7].Style.Numberformat.Format = "0"

            # Col 8: Result - Formula: =IF(AND(F{r}>=30, I{r}>=40), "PASS", "FAIL")
            $ws.Cells[$rInt, 8].Formula = "IF(AND(F${rInt}>=30, I${rInt}>=40), ""PASS"", ""FAIL"")"

            # Col 9: Total(30+70) - Formula: =E{r}+G{r}
            $ws.Cells[$rInt, 9].Formula = "E${rInt}+G${rInt}"
            $ws.Cells[$rInt, 9].Style.Numberformat.Format = "0"
        }
        elseif ($SheetName -eq 'CW') {
            # Subsheet 2: CW (4 Columns: Sr. No., Enrollment, Name, CW(30))
            $cw30 = [Math]::Round(($assignVal / 25.0 * 30.0), 0, [System.MidpointRounding]::AwayFromZero)
            $ws.SetValue($rInt, 4, [int]$cw30)
            $ws.Cells[$rInt, 4].Style.Numberformat.Format = "0"
        }
        elseif ($SheetName -in 'End Sem', 'End sem', 'EndSem') {
            # Subsheet 3: End Sem (4 Columns: Sr. No., Enrollment, Name, End Sem(70))
            $marks70 = [Math]::Round(($examVal / 75.0 * 70.0), 0, [System.MidpointRounding]::AwayFromZero)
            $ws.SetValue($rInt, 4, [int]$marks70)
            $ws.Cells[$rInt, 4].Style.Numberformat.Format = "0"
        }
        elseif ($SheetName -in 'NPTEL', 'Nptel', 'Nptel Sheet') {
            # Subsheet 4: NPTEL (5 Columns: Sr. No., Enrollment, Name, Assignment, Total)
            $ws.SetValue($rInt, 4, $assignVal)
            $ws.Cells[$rInt, 4].Style.Numberformat.Format = "0.##"

            $totalRaw = if ($st.VerifiedTotalMarks) { [string]$st.VerifiedTotalMarks } elseif ($st.DeclaredTotalMarks) { [string]$st.DeclaredTotalMarks } else { "" }
            $totalVal = 0.0
            if (-not [double]::TryParse($totalRaw, [ref]$totalVal) -or $totalVal -le 0) {
                $totalVal = $assignVal + $examVal
            }
            $ws.SetValue($rInt, 5, $totalVal)
            $ws.Cells[$rInt, 5].Style.Numberformat.Format = "0.##"
        }

        $currRow++
        $srNum++
    }

    $lastDataRow = [int][Math]::Max(5, $currRow - 1)

    # Apply cell borders and alignment
    for ($r = 5; $r -le $lastDataRow; $r++) {
        $rIdx = [int]$r
        for ($c = 1; $c -le $colCount; $c++) {
            $cIdx = [int]$c
            $cell = $ws.Cells[$rIdx, $cIdx]
            $cell.Style.Font.Name = "Times New Roman"
            $cell.Style.Font.Size = 11

            $cell.Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $cell.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $cell.Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $cell.Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

            # Alignments
            if ($rIdx -gt 5) {
                if ($cIdx -in 1, 2) {
                    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                } elseif ($cIdx -eq 3) {
                    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
                } else {
                    if ($SheetName -in 'Main Result', 'Cumulative', 'Cumulative Sheet' -and $cIdx -eq 8) {
                        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                    } else {
                        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right
                    }
                }
            }
        }
    }

    # Set exact institutional column widths tailored to sheet column count
    if ($colCount -le 4) {
        # 4-column sheets (CW, End Sem): widened so headers and student names breathe luxuriously with zero clipping
        $ws.Column(1).Width = 7.0   # Sr. No.
        $ws.Column(2).Width = 22.0  # Enrollment
        $ws.Column(3).Width = 34.0  # Student Name
        $ws.Column(4).Width = 20.0  # CW(30) or End Sem(70)
        # Total combined width = 83.0 units (~626 px), guaranteeing 170+ px margins for 52-char headers!
    }
    elseif ($colCount -eq 5) {
        # 5-column sheets (NPTEL):
        $ws.Column(1).Width = 7.0   # Sr. No.
        $ws.Column(2).Width = 20.0  # Enrollment
        $ws.Column(3).Width = 32.0  # Student Name
        $ws.Column(4).Width = 16.0  # Assignment
        $ws.Column(5).Width = 16.0  # Total
        # Total combined width = 91.0 units (~685 px)
    }
    else {
        # Master evaluation ledger (Main Result - 9 columns, total width ~136.0 units):
        $ws.Column(1).Width = 6.0   # Sr. No.
        $ws.Column(2).Width = 18.0  # Enrollment
        $ws.Column(3).Width = 26.0  # Student Name
        for ($c = 4; $c -le $colCount; $c++) {
            $headerText = [string]$headers[$c - 1]
            if ($headerText -match 'Assignment' -or $headerText -match 'Final Exam') {
                $ws.Column($c).Width = 16.0  # Room for 'Assignment(25)' and 'Final Exam(75)'
            } elseif ($headerText -match 'Result') {
                $ws.Column($c).Width = 12.0  # Room for 'PASS' / 'FAIL'
            } elseif ($headerText -match 'Total') {
                $ws.Column($c).Width = 15.0  # Room for 'Total(30+70)'
            } else {
                $ws.Column($c).Width = 13.5  # Marks and formula columns
            }
        }
    }
}

function Export-CourseCollegeResultSheets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Course,
        [string]$SessionName = "Jan-June 2026",
        [string]$TargetDirectory = $null,
        [string]$TargetSheet = 'All',
        [string]$DataDir = $null
    )

    $result = [PSCustomObject]@{
        Success          = $false
        FilePath         = $null
        VerifiedCount    = 0
        TotalStudents    = 0
        UnderReviewCount = 0
        GeneratedSheets  = @()
        Error            = $null
    }

    $cId   = [string]$Course.Id
    $cName = if ($Course.Name) { [string]$Course.Name } else { "Course" }
    $cCode = if ($Course.Code) { [string]$Course.Code } else { "" }
    $sem   = if ($Course.Semester) { [string]$Course.Semester } else { "8" }

    # Resolve target file path
    $targetPath = if ($TargetDirectory) {
        Join-Path $TargetDirectory ("$($cName -replace '[\\/:*?"<>|]', '_')_College_Result_Sheets.xlsx")
    } else {
        Get-CourseCollegeResultSheetsFilePath -CourseId $cId -CourseName $cName -DataDir $DataDir
    }

    # Proactively check if Result Verification Sheet exists for course
    $vSheetPath = if ($Course.ExamVerificationSheet) { [string]$Course.ExamVerificationSheet } else { "" }
    if (-not $vSheetPath -or -not (Test-Path -LiteralPath $vSheetPath -ErrorAction SilentlyContinue)) {
        $result.Error = "Result Verification Sheet has not been generated for this course yet."
        return $result
    }

    # Verify ImportExcel module availability
    if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
        $result.Error = "The 'ImportExcel' module is required. Please install it with: Install-Module ImportExcel -Scope CurrentUser"
        return $result
    }
    Import-Module ImportExcel -ErrorAction SilentlyContinue

    # Load verified students from store
    $store = Get-CourseExamResultsStore -CourseId $cId -CourseName $cName -DataDir $DataDir
    $allStudents = if ($store -and $store.Students) { @($store.Students) } else { @() }
    $verifiedStudents = @($allStudents | Where-Object { $_.ExamVerificationStatus -eq 'Verified' })

    # Sort alphabetically by RollNo
    $sortedStudents = @($verifiedStudents | Sort-Object { [string]$_.RollNo })

    $result.TotalStudents    = $allStudents.Count
    $result.VerifiedCount    = $sortedStudents.Count
    $result.UnderReviewCount = ($allStudents.Count - $sortedStudents.Count)

    # Determine sheets to generate
    $sheetsToGen = switch ($TargetSheet) {
        'CW'                                                      { @('CW') }
        { $_ -in 'EndSem', 'End sem', 'End Sem' }                 { @('End Sem') }
        { $_ -in 'Cumulative', 'Cumulative Sheet', 'Main Result' }{ @('Main Result') }
        { $_ -in 'Nptel', 'NPTEL', 'Nptel Sheet' }                { @('NPTEL') }
        default                                                   { @('Main Result', 'CW', 'End Sem', 'NPTEL') }
    }

    # Proactively check if consolidated file is currently open in Excel
    if (Test-CollegeFileLocked -Path $targetPath) {
        $fName = [System.IO.Path]::GetFileName($targetPath)
        $result.Error = "FILE_LOCKED:$fName"
        return $result
    }

    # Proactively check if any separate sheet files being generated are currently open in Excel
    foreach ($sName in $sheetsToGen) {
        $sepCheckPath = if ($TargetDirectory) {
            $suffix = switch ($sName) {
                'CW'                                                      { "CW" }
                { $_ -in 'End Sem', 'End sem', 'EndSem' }                 { "End_Sem" }
                { $_ -in 'Main Result', 'Cumulative', 'Cumulative Sheet' }{ "Cumulative" }
                { $_ -in 'NPTEL', 'Nptel', 'Nptel Sheet' }                { "NPTEL" }
                default                                                   { $sName }
            }
            Join-Path $TargetDirectory ("$($cName -replace '[\\/:*?""<>|]', '_')_${suffix}.xlsx")
        } else {
            Get-CourseCollegeSeparateSheetFilePath -CourseId $cId -CourseName $cName -SheetType $sName -DataDir $DataDir
        }

        if (Test-CollegeFileLocked -Path $sepCheckPath) {
            $fName = [System.IO.Path]::GetFileName($sepCheckPath)
            $result.Error = "FILE_LOCKED:$fName"
            return $result
        }
    }

    # Resolve dynamic course metadata with safe fallbacks
    $collegeName = if ($Course.CollegeName) { [string]$Course.CollegeName } else { "SHRI G. S. INSTITUTE OF TECHNOLOGY & SCIENCE, INDORE" }
    $sessionStr  = if ($Course.Session) { [string]$Course.Session } elseif ($SessionName) { [string]$SessionName } else { "Jan-June 2026" }
    $yearStr     = if ($Course.AcademicYear) { [string]$Course.AcademicYear } else {
        $semInt = 8
        [int]::TryParse($sem, [ref]$semInt) | Out-Null
        switch ($semInt) {
            { $_ -in 7, 8 } { "IV YEAR" }
            { $_ -in 5, 6 } { "III YEAR" }
            { $_ -in 3, 4 } { "II YEAR" }
            { $_ -in 1, 2 } { "I YEAR" }
            default         { "IV YEAR" }
        }
    }
    $electiveStr = if ($Course.ElectiveType) { [string]$Course.ElectiveType } else { "Elective VI" }
    $tagSuffix   = if ($Course.HeaderTag) { [string]$Course.HeaderTag } else { "(NPTEL)" }
    $codeTitle   = if ($cCode) { "$cCode " } else { "" }
    $tagTitle    = if ($tagSuffix) { " $tagSuffix" } else { "" }

    try {
        # 1. Open or create consolidated ExcelPackage (contains all 4 subsheets)
        $pkg = Open-ExcelPackage -Path $targetPath -Create
        $pkg.Workbook.View.ShowSheetTabs = $true

        if ($TargetSheet -in 'All', $null, '') {
            # Clean up any legacy or duplicate worksheets so master workbook has exactly the 4 fresh subsheets
            foreach ($legacyName in @('Main Result', 'CW', 'End Sem', 'End sem', 'Cumulative', 'Cumulative Sheet', 'Nptel Sheet', 'NPTEL', 'Nptel')) {
                $oldWs = $pkg.Workbook.Worksheets[$legacyName]
                if ($oldWs) {
                    $pkg.Workbook.Worksheets.Delete($legacyName)
                }
            }
        }

        foreach ($sheetName in $sheetsToGen) {
            # If sheet already exists, delete it so fresh formatted sheet is created
            $existingWs = $pkg.Workbook.Worksheets[$sheetName]
            if ($existingWs) {
                $pkg.Workbook.Worksheets.Delete($sheetName)
            }

            $ws = $pkg.Workbook.Worksheets.Add($sheetName)

            Format-AndPopulateWorksheet -Worksheet $ws -SheetName $sheetName `
                -SortedStudents $sortedStudents -CollegeName $collegeName `
                -SessionStr $sessionStr -YearStr $yearStr -ElectiveStr $electiveStr `
                -CodeTitle $codeTitle -CName $cName -TagTitle $tagTitle

            $null = $result.GeneratedSheets += $sheetName
        }

        # Ensure first sheet (Main Result) is active and selected
        if ($pkg.Workbook.Worksheets.Count -gt 0) {
            $pkg.Workbook.Worksheets[1].View.TabSelected = $true
        }

        # Save and close consolidated package
        Close-ExcelPackage $pkg

        # 2. Also generate separate standalone Excel files for each sheet
        foreach ($sheetName in $sheetsToGen) {
            $sepPath = if ($TargetDirectory) {
                $suffix = switch ($sheetName) {
                    'CW'                                                      { "CW" }
                    { $_ -in 'End Sem', 'End sem', 'EndSem' }                 { "End_Sem" }
                    { $_ -in 'Main Result', 'Cumulative', 'Cumulative Sheet' }{ "Cumulative" }
                    { $_ -in 'NPTEL', 'Nptel', 'Nptel Sheet' }                { "NPTEL" }
                    default                                                   { $sheetName }
                }
                Join-Path $TargetDirectory ("$($cName -replace '[\\/:*?""<>|]', '_')_${suffix}.xlsx")
            } else {
                Get-CourseCollegeSeparateSheetFilePath -CourseId $cId -CourseName $cName -SheetType $sheetName -DataDir $DataDir
            }

            if (Test-Path -LiteralPath $sepPath) {
                Remove-Item -LiteralPath $sepPath -Force -ErrorAction SilentlyContinue
            }

            $sepPkg = Open-ExcelPackage -Path $sepPath -Create
            $ws = $sepPkg.Workbook.Worksheets.Add($sheetName)

            Format-AndPopulateWorksheet -Worksheet $ws -SheetName $sheetName `
                -SortedStudents $sortedStudents -CollegeName $collegeName `
                -SessionStr $sessionStr -YearStr $yearStr -ElectiveStr $electiveStr `
                -CodeTitle $codeTitle -CName $cName -TagTitle $tagTitle

            Close-ExcelPackage $sepPkg
        }

        $result.Success  = $true
        $result.FilePath = $targetPath

        # Persist generated path in Course object
        if ($Course.PSObject.Properties['CollegeResultSheets']) {
            $Course.CollegeResultSheets = $targetPath
        } else {
            $Course | Add-Member -NotePropertyName 'CollegeResultSheets' -NotePropertyValue $targetPath -Force
        }
        if (Get-Command -Name 'Save-Courses' -ErrorAction SilentlyContinue) {
            Save-Courses
        }

        return $result
    }
    catch {
        $st = $_.ScriptStackTrace
        $lockedTarget = $targetPath
        if (Test-CollegeFileLocked -Path $targetPath) {
            $lockedTarget = $targetPath
        } else {
            foreach ($sName in $sheetsToGen) {
                $sepCheck = if ($TargetDirectory) {
                    $suf = switch ($sName) {
                        'CW'                                                      { 'CW' }
                        { $_ -in 'End Sem', 'End sem', 'EndSem' }                 { 'End_Sem' }
                        { $_ -in 'Main Result', 'Cumulative', 'Cumulative Sheet' }{ 'Cumulative' }
                        { $_ -in 'NPTEL', 'Nptel', 'Nptel Sheet' }                { 'NPTEL' }
                        default                                                   { $sName }
                    }
                    Join-Path $TargetDirectory ("$($cName -replace '[\\/:*?""<>|]', '_')_${suf}.xlsx")
                } else {
                    Get-CourseCollegeSeparateSheetFilePath -CourseId $cId -CourseName $cName -SheetType $sName -DataDir $DataDir
                }
                if (Test-CollegeFileLocked -Path $sepCheck) {
                    $lockedTarget = $sepCheck
                    break
                }
            }
        }
        if ((Test-CollegeFileLocked -Path $lockedTarget) -or $_.Exception.Message -match "being used by another process|Error saving file|Save") {
            $fName = [System.IO.Path]::GetFileName($lockedTarget)
            $result.Error = "FILE_LOCKED:$fName"
        } else {
            $result.Error = "$($_.Exception.Message) AT $st"
        }
        return $result
    }
}
