param(
    [string]$TemplatePath,
    [string]$DataFilePath
)

try {
    # Read Data from File
    if (-not (Test-Path $DataFilePath)) {
        Write-Error "Data file not found: $DataFilePath"
        exit 1
    }
    $JsonContent = Get-Content -Path $DataFilePath -Raw
    $Data = $JsonContent | ConvertFrom-Json
    $Info = $Data.info

    # Create Excel COM Object
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    # Open Template (Read-Only)
    $Workbook = $Excel.Workbooks.Open($TemplatePath, 0, $true)
    $Sheet = $Workbook.Sheets.Item(1) # FRONT sheet

    # Helper to clean text
    function Clean-Text([string]$str) {
        if ([string]::IsNullOrWhiteSpace($str)) { return "" }
        return $str
    }

    # --- Smart Fill: Find label, skip label's merge area, write to next cell ---
    function Fill-Cell([string]$label, [string]$value) {
        if (-not $value) { return }
        try {
            $Found = $Sheet.UsedRange.Find($label)
            if ($Found) {
                # Determine where the label ends (handling merged cells)
                $MergeArea = $Found.MergeArea
                $LabelEndCol = $MergeArea.Column + $MergeArea.Columns.Count - 1
                
                # Target the cell immediately to the right of the label
                $TargetCol = $LabelEndCol + 1
                $Target = $Sheet.Cells.Item($Found.Row, $TargetCol)
                
                $Target.Value2 = $value
            }
        }
        catch {
            Write-Host "Error filling $label"
        }
    }

    # Fill Data
    Fill-Cell "LAST NAME" (Clean-Text $Info.lname)
    Fill-Cell "FIRST NAME" (Clean-Text $Info.fname)
    Fill-Cell "MIDDLE NAME" (Clean-Text $Info.mname)
    Fill-Cell "LRN" (Clean-Text $Info.lrn)
    
    # Partial Matches for long labels
    Fill-Cell "Date of Birth" (Clean-Text $Info.birthdate)
    Fill-Cell "Sex" (Clean-Text $Info.sex)
    # Fill-Cell "Date of SHS Admission" (Clean-Text $Info.admissionDate)

    # Print
    $Workbook.PrintOut()
    
    # Cleanup
    $Workbook.Close($false)
    $Excel.Quit()
    
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    Remove-Variable Excel
    
    Write-Output "SUCCESS"

}
catch {
    Write-Error $_.Exception.Message
    if ($Excel) {
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    }
    exit 1
}
