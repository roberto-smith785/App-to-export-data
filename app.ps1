# Requires modules: SqlServer, ImportExcel (EPPlus)
Add-Type -AssemblyName System.Windows.Forms

# Create the form with improved size
$form = New-Object System.Windows.Forms.Form
$form.Text = "DataSmart"
$form.Size = New-Object System.Drawing.Size(600, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray
$form.Padding = New-Object System.Windows.Forms.Padding(20)

# Function to create labels
function Create-Label {
    param (
        [string]$text,
        [System.Drawing.Point]$location
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Location = $location
    $label.Size = New-Object System.Drawing.Size(100, 20)
    $label.AutoSize = $true
    return $label
}

# Function to create textboxes
function Create-TextBox {
    param (
        [System.Drawing.Point]$location,
        [int]$width = 300
    )
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = $location
    $textBox.Size = New-Object System.Drawing.Size($width, 20)
    return $textBox
}

# Create GroupBox for Database Connection
$groupBoxConnection = New-Object System.Windows.Forms.GroupBox
$groupBoxConnection.Text = "Database Connection"
$groupBoxConnection.Size = New-Object System.Drawing.Size(540, 140)
$groupBoxConnection.Location = New-Object System.Drawing.Point(20, 20)

# Add controls to Connection GroupBox
$labelServer = Create-Label "SQL Server:" (New-Object System.Drawing.Point(20, 30))
$textServer = Create-TextBox (New-Object System.Drawing.Point(120, 27))
$labelDatabase = Create-Label "Database:" (New-Object System.Drawing.Point(20, 60))
$textDatabase = Create-TextBox (New-Object System.Drawing.Point(120, 57))
$labelTableName = Create-Label "Table Name:" (New-Object System.Drawing.Point(20, 90))
$textTableName = Create-TextBox (New-Object System.Drawing.Point(120, 87))

$groupBoxConnection.Controls.AddRange(@($labelServer, $textServer, $labelDatabase, $textDatabase, $labelTableName, $textTableName))

# Create GroupBox for Authentication
$groupBoxAuth = New-Object System.Windows.Forms.GroupBox
$groupBoxAuth.Text = "Authentication"
$groupBoxAuth.Size = New-Object System.Drawing.Size(540, 120)
$groupBoxAuth.Location = New-Object System.Drawing.Point(20, 170)

# Authentication radio buttons
$radioWindowsAuth = New-Object System.Windows.Forms.RadioButton
$radioWindowsAuth.Text = "Windows Authentication"
$radioWindowsAuth.Location = New-Object System.Drawing.Point(20, 30)
$radioWindowsAuth.Size = New-Object System.Drawing.Size(200, 20)
$radioWindowsAuth.Checked = $true

$radioSQLAuth = New-Object System.Windows.Forms.RadioButton
$radioSQLAuth.Text = "SQL Server Authentication"
$radioSQLAuth.Location = New-Object System.Drawing.Point(220, 30)
$radioSQLAuth.Size = New-Object System.Drawing.Size(200, 20)

# SQL Authentication controls
$labelUsername = Create-Label "Username:" (New-Object System.Drawing.Point(20, 60))
$textUsername = Create-TextBox (New-Object System.Drawing.Point(120, 57))
$labelPassword = Create-Label "Password:" (New-Object System.Drawing.Point(20, 90))
$textPassword = Create-TextBox (New-Object System.Drawing.Point(120, 87))
$textPassword.PasswordChar = "*"

$labelUsername.Visible = $false
$textUsername.Visible = $false
$labelPassword.Visible = $false
$textPassword.Visible = $false

# Authentication radio button event handlers
$radioWindowsAuth.Add_CheckedChanged({
    if ($radioWindowsAuth.Checked) {
        $radioWindowsAuth.Checked = $true
        $radioSQLAuth.Checked = $false
        $labelUsername.Visible = $false
        $textUsername.Visible = $false
        $labelPassword.Visible = $false
        $textPassword.Visible = $false
    }
})

$radioSQLAuth.Add_CheckedChanged({
     if ($radioSQLAuth.Checked) {
        $radioSQLAuth.Checked = $true
        $radioWindowsAuth.Checked = $false
        $labelUsername.Visible = $true
        $textUsername.Visible = $true
        $labelPassword.Visible = $true
        $textPassword.Visible = $true
    }
})

$groupBoxAuth.Controls.AddRange(@($radioWindowsAuth, $radioSQLAuth, $labelUsername, $textUsername, $labelPassword, $textPassword))

# Create GroupBox for Execution Type
$groupBoxExecution = New-Object System.Windows.Forms.GroupBox
$groupBoxExecution.Text = "Execution Type"
$groupBoxExecution.Size = New-Object System.Drawing.Size(540, 120)
$groupBoxExecution.Location = New-Object System.Drawing.Point(20, 300)

# Execution type radio buttons
$radioStoredProcedure = New-Object System.Windows.Forms.RadioButton
$radioStoredProcedure.Text = "Stored Procedure"
$radioStoredProcedure.Location = New-Object System.Drawing.Point(20, 30)
$radioStoredProcedure.Size = New-Object System.Drawing.Size(200, 20)
$radioStoredProcedure.Checked = $true

$radioSqlScript = New-Object System.Windows.Forms.RadioButton
$radioSqlScript.Text = "SQL Script"
$radioSqlScript.Location = New-Object System.Drawing.Point(220, 30)
$radioSqlScript.Size = New-Object System.Drawing.Size(200, 20)

# Stored procedure/SQL Script controls
$labelStoredProcedure = Create-Label "Stored Procedure:" (New-Object System.Drawing.Point(20, 70))
$textStoredProcedure = Create-TextBox (New-Object System.Drawing.Point(120, 67))
$labelSQLScript = Create-Label "SQL Script:" (New-Object System.Drawing.Point(20, 70))
$textSQLScriptPath = Create-TextBox (New-Object System.Drawing.Point(120, 67))
$buttonSelectSQLScript = New-Object System.Windows.Forms.Button
$buttonSelectSQLScript.Text = "Browse..."
$buttonSelectSQLScript.Location = New-Object System.Drawing.Point(430, 66)
$buttonSelectSQLScript.Size = New-Object System.Drawing.Size(80, 23)

$labelSQLScript.Visible = $false
$textSQLScriptPath.Visible = $false
$buttonSelectSQLScript.Visible = $false

# Script type radio button event handlers
$radioStoredProcedure.Add_CheckedChanged({
    if ($radioStoredProcedure.Checked) {
        $radioStoredProcedure.Checked = $true
        $radioSqlScript.Checked = $false
        $labelStoredProcedure.Visible = $true
        $textStoredProcedure.Visible = $true
        $labelSQLScript.Visible = $false
        $textSQLScriptPath.Visible = $false
        $buttonSelectSQLScript.Visible = $false
    }
})

$radioSqlScript.Add_CheckedChanged({
    if ($radioSqlScript.Checked) {
        $radioSqlScript.Checked = $true
        $radioStoredProcedure.Checked = $false
        $labelStoredProcedure.Visible = $false
        $textStoredProcedure.Visible = $false
        $labelSQLScript.Visible = $true
        $textSQLScriptPath.Visible = $true
        $buttonSelectSQLScript.Visible = $true
    }
})

$groupBoxExecution.Controls.AddRange(@($radioStoredProcedure, $radioSqlScript, $labelStoredProcedure, $textStoredProcedure, 
                                     $labelSQLScript, $textSQLScriptPath, $buttonSelectSQLScript))

# Create GroupBox for File Selection
$groupBoxFiles = New-Object System.Windows.Forms.GroupBox
$groupBoxFiles.Text = "File Selection"
$groupBoxFiles.Size = New-Object System.Drawing.Size(540, 140)
$groupBoxFiles.Location = New-Object System.Drawing.Point(20, 430)

# CSV and Output file selection
$labelCSVFile = Create-Label "Data File(CSV):" (New-Object System.Drawing.Point(20, 30))
$textCSVPath = Create-TextBox (New-Object System.Drawing.Point(120, 27))
$buttonSelectCSV = New-Object System.Windows.Forms.Button
$buttonSelectCSV.Text = "Browse..."
$buttonSelectCSV.Location = New-Object System.Drawing.Point(430, 26)
$buttonSelectCSV.Size = New-Object System.Drawing.Size(80, 23)

$buttonSelectCSV.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textCSVPath.Text = $openFileDialog.FileName
    }
})

# CSV and Output file selection
$labelOutputFile = Create-Label "Output File(Excel):" (New-Object System.Drawing.Point(20, 70))
$textOutputFilePath = Create-TextBox (New-Object System.Drawing.Point(120, 67))
$buttonSelectOutputFile = New-Object System.Windows.Forms.Button
$buttonSelectOutputFile.Text = "Browse..."
$buttonSelectOutputFile.Location = New-Object System.Drawing.Point(430, 66)
$buttonSelectOutputFile.Size = New-Object System.Drawing.Size(80, 23)

$buttonSelectOutputFile.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "XLSX files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textOutputFilePath.Text = $openFileDialog.FileName
    }
})

$labelLogFile = Create-Label "Log File(txt):" (New-Object System.Drawing.Point(20, 110))
$textLogFilePath = Create-TextBox (New-Object System.Drawing.Point(120, 107))
$buttonSelectLogFile = New-Object System.Windows.Forms.Button
$buttonSelectLogFile.Text = "Browse..."
$buttonSelectLogFile.Location = New-Object System.Drawing.Point(430, 106)
$buttonSelectLogFile.Size = New-Object System.Drawing.Size(80, 23)

$buttonSelectLogFile.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textLogFilePath.Text = $saveFileDialog.FileName
    }
})

$groupBoxFiles.Controls.AddRange(@($labelCSVFile, $textCSVPath, $buttonSelectCSV, 
                                 $labelOutputFile, $textOutputFilePath, $buttonSelectOutputFile,
                                 $labelLogFile, $textLogFilePath, $buttonSelectLogFile))

# Progress bar and Execute button
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 590)
$progressBar.Size = New-Object System.Drawing.Size(540, 20)

$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = "Execute"
$buttonExecute.Location = New-Object System.Drawing.Point(20, 610)
$buttonExecute.Size = New-Object System.Drawing.Size(540, 30)

$buttonExecute.Add_Click({
    try {
        # Validate required fields
        $requiredFields = @(
            @{ Field = $textServer.Text; Name = "SQL Server" }
            @{ Field = $textDatabase.Text; Name = "Database" }
            @{ Field = $textTableName.Text; Name = "Table Name" }
            @{ Field = $textCSVPath.Text; Name = "CSV File" }
            @{ Field = $textOutputFilePath.Text; Name = "Output File" }
            @{ Field = $textLogFilePath.Text; Name = "Log File" }
        )

        if ($radioSQLAuth.Checked) {
            $requiredFields += @(
                @{ Field = $textUsername.Text; Name = "Username" }
                @{ Field = $textPassword.Text; Name = "Password" }
            )
        }

        if ($radioStoredProcedure.Checked) {
            $requiredFields += @{ Field = $textStoredProcedure.Text; Name = "Stored Procedure" }
        } else {
            $requiredFields += @{ Field = $textSQLScriptPath.Text; Name = "SQL Script File" }
        }

        foreach ($field in $requiredFields) {
            if ([string]::IsNullOrWhiteSpace($field.Field)) {
                $form = New-Object System.Windows.Forms.Form
                $form.Text = "Validation Error"
                $form.Size = New-Object System.Drawing.Size(500, 500)
    
                $panel = New-Object System.Windows.Forms.Panel
                $panel.Dock = 'Fill'
                $panel.AutoScroll = $true

                $textBox = New-Object System.Windows.Forms.TextBox
                $textBox.Multiline = $true
                $textBox.ReadOnly = $true
                $textBox.ScrollBars = 'Vertical'
                $textBox.Dock = 'Fill'
                $textBox.Text = "$($field.Name) is required."

                $panel.Controls.Add($textBox)
                $form.Controls.Add($panel)
                $form.ShowDialog()
                return
            }
        }

        # Build connection string
        $connectionString = if ($radioWindowsAuth.Checked) {
            "Server=$($textServer.Text);Database=$($textDatabase.Text);Integrated Security=True;"
        } else {
            "Server=$($textServer.Text);Database=$($textDatabase.Text);User Id=$($textUsername.Text);Password=$($textPassword.Text);"
        }

        # Import CSV data
        $csvData = Import-Csv -Path $textCSVPath.Text
        $totalRows = $csvData.Count
        $progressBar.Maximum = $totalRows
        $progressBar.Value = 0

        # Start transcript for logging
        Start-Transcript -Path $textOutputFilePath.Text -Force

        # Truncate existing table
        $truncateQuery = "TRUNCATE TABLE $($textTableName.Text)"
        Write-Host "Clearing existing data from table..."
        Invoke-Sqlcmd -ConnectionString $connectionString -Query $truncateQuery

        # Insert CSV data
        Write-Host "Inserting CSV data..."
        $processedRows = 0
        $batchSize = 1000
        $currentBatch = @()

          foreach ($row in $csvData) {
                # Properly format column names with brackets
                $columns = ($row.PSObject.Properties.Name | ForEach-Object { "[$_]" }) -join ", "
    
                # Handle values with proper formatting
                $values = ($row.PSObject.Properties.Value | ForEach-Object {
                    if ($null -eq $_) {
                            "NULL"
                        }else{
                            "'[$_]'"
                        }
                        <#elseif ($_ -is [int] -or $_ -is [decimal] -or $_ -is [double]) {
                            "$_" # Convert numeric value to string without quotes
                        }
                        elseif ($_ -is [datetime]) {
                            # Format datetime in SQL Server compatible format (yyyy-MM-dd HH:mm:ss)
                            "'$($_.ToString("yyyy-MM-dd HH:mm:ss"))'"
                        }
                        else {
                            # Escape single quotes and wrap in quotes
                            "'$($_ -replace "'", "''")'"
                        }#>
                }) -join ", "
    
                # Build the INSERT statement
                $currentBatch += "INSERT INTO [$($textTableName.Text)] ($columns) VALUES ($values);"
    
                $processedRows++
                $progressBar.Value = $processedRows
    
                # Handle batch processing
                if ($currentBatch.Count -eq $batchSize -or $processedRows -eq $totalRows) {
                    $batchQuery = $currentBatch -join "`n"
                    try {
                        Invoke-Sqlcmd -ConnectionString $connectionString -Query $batchQuery
                    }
                    catch {
                        Write-Host "Error inserting batch: $_"
                    }
                    $currentBatch = @()
                }
            }

# Execute stored procedure or SQL script
        if ($radioStoredProcedure.Checked) {
            Write-Host "Executing stored procedure $($textStoredProcedure.Text)..."
            try {
                $query = "EXEC $($textStoredProcedure.Text)"
                $results = Invoke-Sqlcmd -ConnectionString $connectionString -Query $query -As DataTables

                if ($results.Count -eq 0) {
                    Write-Host "Stored procedure executed successfully. No results returned."
                } else {
                    Write-Host "Stored procedure executed successfully."
                    # Create a new Excel application
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $workbook = $excel.Workbooks.Add()
                $saveLocation = $textOutputFilePath.Text

                # Counter for worksheet naming
                $sheetCounter = 1

                foreach ($table in $results) {
                $form = New-Object System.Windows.Forms.Form
                $form.Text = "Data Info"
                $form.Size = New-Object System.Drawing.Size(500, 500)
    
                $panel = New-Object System.Windows.Forms.Panel
                $panel.Dock = 'Fill'
                $panel.AutoScroll = $true

                $textBox = New-Object System.Windows.Forms.TextBox
                $textBox.Multiline = $true
                $textBox.ReadOnly = $true
                $textBox.ScrollBars = 'Vertical'
                $textBox.Dock = 'Fill'
                $textBox.Text = "Result set contains $($table.Rows.Count) rows: $(($table | Format-Table -AutoSize -Wrap | Out-String))"

                $panel.Controls.Add($textBox)
                $form.Controls.Add($panel)
                $form.ShowDialog()
    
                    # Create a new worksheet for each result set
                    if ($sheetCounter -gt 1) {
                        $worksheet = $workbook.Worksheets.Add()
                    } else {
                        $worksheet = $workbook.Worksheets.Item(1)
                    }
    
                    # Name the worksheet
                    $worksheet.Name = "Sheet$($sheetCounter)"
    
                    # Get column headers
                    $headers = $table | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
    
                    # Write headers
                    for ($col = 0; $col -lt $headers.Count; $col++) {
                        $worksheet.Cells(1, $col + 1) = $headers[$col]
                    }
    
                    # Write data
                    for ($row = 0; $row -lt $table.Rows.Count; $row++) {
                        for ($col = 0; $col -lt $headers.Count; $col++) {
                            $worksheet.Cells($row + 2, $col + 1) = $table.Rows[$row].$($headers[$col])
                        }
                    }
    
                    # Auto-fit columns
                    $usedRange = $worksheet.UsedRange
                    $usedRange.EntireColumn.AutoFit() | Out-Null
    
                    $sheetCounter++
                }

                # Save and close
                try {
                    $workbook.SaveAs($saveLocation)
                    Write-Host "Excel file saved successfully to: $saveLocation" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error saving file: $_" -ForegroundColor Red
                }
                finally {
                    $workbook.Close($true)
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    Remove-Variable excel
                }
                }
            } catch {
                Write-Host "Error executing stored procedure: $_"
                throw
            }
        } elseif ($radioSqlScript.Checked) {
            Write-Host "Executing SQL script from $($textSQLScriptPath.Text)..."
            try {
                if (-not (Test-Path $textSQLScriptPath.Text)) {
                    throw "SQL script file not found"
                }
                
                $script = Get-Content $textSQLScriptPath.Text -Raw
                Write-Host "SQL Script contents:`n$script"
                
                $results = Invoke-Sqlcmd -ConnectionString $connectionString -InputFile $textSQLScriptPath.Text -As DataTables
                
                if ($results.Count -eq 0) {
                    Write-Host "SQL script executed successfully. No results returned."
                } else {
                    Write-Host "SQL script executed successfully."
                    foreach ($table in $results) {
                        Write-Host "Result set contains $($table.Rows.Count) rows"
                        $table | Format-Table -AutoSize | Out-String | Write-Host
                    }
                }
            } catch {
                Write-Host "Error executing SQL script: $_"
                throw
            }
        }

        Stop-Transcript
        $progressBar.Value = $progressBar.Maximum
        [System.Windows.Forms.MessageBox]::Show(
            "Operation completed successfully. Check the output file for details.",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        Write-Host "Error: $_"
        Stop-Transcript
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred: $_`nCheck the output file for details.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    } finally {
        if ($progressBar.Value -ne $progressBar.Maximum) {
            $progressBar.Value = 0
        }
    }
})


# Add all main controls to form
$form.Controls.AddRange(@($groupBoxConnection, $groupBoxAuth, $groupBoxExecution, $groupBoxFiles, $progressBar, $buttonExecute))

# [Rest of the event handlers and execution logic remains the same...]

# Show the form
$form.ShowDialog()