
# Requires modules: SqlServer, ImportExcel (EPPlus)
Add-Type -AssemblyName System.Windows.Forms

# Create the form and controls
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV to SQL Execution"
$form.Size = New-Object System.Drawing.Size(500, 650)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

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
    return $label
}

# Function to create textboxes
function Create-TextBox {
    param (
        [System.Drawing.Point]$location
    )
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = $location
    $textBox.Size = New-Object System.Drawing.Size(200, 20)
    return $textBox
}

# Function to create radio buttons
function Create-RadioButton {
    param (
        [string]$text,
        [System.Drawing.Point]$location
    )
    $radioButton = New-Object System.Windows.Forms.RadioButton
    $radioButton.Text = $text
    $radioButton.Location = $location
    $radioButton.Size = New-Object System.Drawing.Size(150, 30)
    return $radioButton
}

# Create labels and textboxes for SQL Server, database, table name, authentication, etc.
$labelServer =  Create-Label "SQL Server:" (New-Object System.Drawing.Point(10, 20))
$form.Controls.Add($labelServer)

$textServer = Create-TextBox (New-Object System.Drawing.Point(120, 20))
$form.Controls.Add($textServer)

$labelDatabase = Create-Label "Database:" (New-Object System.Drawing.Point(10, 60))
$form.Controls.Add($labelDatabase)

$textDatabase = Create-TextBox (New-Object System.Drawing.Point(120, 60))
$form.Controls.Add($textDatabase)

$labelTableName = Create-Label "Table Name:" (New-Object System.Drawing.Point(10, 100))
$form.Controls.Add($labelTableName)

$textTableName = Create-TextBox (New-Object System.Drawing.Point(120, 100))
$form.Controls.Add($textTableName)


# Radio buttons for authentication
$radioWindowsAuth = Create-RadioButton "Windows Authentication" (New-Object System.Drawing.Point(10, 140))
$radioWindowsAuth.Checked = $true
$form.Controls.Add($radioWindowsAuth)

$radioSQLAuth = Create-RadioButton "SQL Server Authentication" (New-Object System.Drawing.Point(180, 140))
$form.Controls.Add($radioSQLAuth)

# SQL Authentication controls
$labelUsername = Create-Label "Username:" (New-Object System.Drawing.Point(10, 180))
$labelUsername.Visible = $false
$form.Controls.Add($labelUsername)

$textUsername = Create-TextBox (New-Object System.Drawing.Point(120, 180))
$textUsername.Visible = $false
$form.Controls.Add($textUsername)

$labelPassword = Create-Label "Password:" (New-Object System.Drawing.Point(10, 220))
$labelPassword.Visible = $false
$form.Controls.Add($labelPassword)

$textPassword = Create-TextBox (New-Object System.Drawing.Point(120, 220))
$textPassword.PasswordChar = "*"
$textPassword.Visible = $false
$form.Controls.Add($textPassword)

# Authentication radio button event handlers
$radioWindowsAuth.Add_CheckedChanged({
    $labelUsername.Visible = !$radioWindowsAuth.Checked
    $textUsername.Visible = !$radioWindowsAuth.Checked
    $labelPassword.Visible = !$radioWindowsAuth.Checked
    $textPassword.Visible = !$radioWindowsAuth.Checked
})

$radioSQLAuth.Add_CheckedChanged({
    $labelUsername.Visible = $radioSQLAuth.Checked
    $textUsername.Visible = $radioSQLAuth.Checked
    $labelPassword.Visible = $radioSQLAuth.Checked
    $textPassword.Visible = $radioSQLAuth.Checked
})

# Execution type radio buttons
$radioStoredProcedure = Create-RadioButton "Stored Procedure" (New-Object System.Drawing.Point(10, 260))
$radioStoredProcedure.Checked = $true
$form.Controls.Add($radioStoredProcedure)

$radioSqlScript = Create-RadioButton "SQL Script" (New-Object System.Drawing.Point(180, 260))
$form.Controls.Add($radioSqlScript)

# Stored procedure controls
$labelStoredProcedure = Create-Label "Stored Procedure:" (New-Object System.Drawing.Point(10, 300))
$form.Controls.Add($labelStoredProcedure)

$textStoredProcedure = Create-TextBox (New-Object System.Drawing.Point(120, 300))
$form.Controls.Add($textStoredProcedure)

# SQL Script controls
$labelSQLScript = Create-Label "SQL Script File:" (New-Object System.Drawing.Point(10, 300))
$labelSQLScript.Visible = $false
$form.Controls.Add($labelSQLScript)

$textSQLScriptPath = Create-TextBox (New-Object System.Drawing.Point(120, 300))
$textSQLScriptPath.Visible = $false
$form.Controls.Add($textSQLScriptPath)

$buttonSelectSQLScript = New-Object System.Windows.Forms.Button
$buttonSelectSQLScript.Text = "Select Script"
$buttonSelectSQLScript.Location = New-Object System.Drawing.Point(330, 300)
$buttonSelectSQLScript.Size = New-Object System.Drawing.Size(100, 20)
$buttonSelectSQLScript.Visible = $false
$form.Controls.Add($buttonSelectSQLScript)

# Script type radio button event handlers
$radioStoredProcedure.Add_CheckedChanged({
    $labelStoredProcedure.Visible = $radioStoredProcedure.Checked
    $textStoredProcedure.Visible = $radioStoredProcedure.Checked
    $labelSQLScript.Visible = !$radioStoredProcedure.Checked
    $textSQLScriptPath.Visible = !$radioStoredProcedure.Checked
    $buttonSelectSQLScript.Visible = !$radioStoredProcedure.Checked
})

$radioSqlScript.Add_CheckedChanged({
    $labelStoredProcedure.Visible = !$radioSqlScript.Checked
    $textStoredProcedure.Visible = !$radioSqlScript.Checked
    $labelSQLScript.Visible = $radioSqlScript.Checked
    $textSQLScriptPath.Visible = $radioSqlScript.Checked
    $buttonSelectSQLScript.Visible = $radioSqlScript.Checked
})

# SQL Script file selection
$buttonSelectSQLScript.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "SQL files (*.sql)|*.sql|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textSQLScriptPath.Text = $openFileDialog.FileName
    }
})

# CSV file selection
$labelCSVFile = Create-Label "CSV File:" (New-Object System.Drawing.Point(10, 350))
$form.Controls.Add($labelCSVFile)

$textCSVPath = New-Object System.Windows.Forms.TextBox
$textCSVPath.Location = New-Object System.Drawing.Point(120, 350)
$textCSVPath.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textCSVPath)

$buttonSelectCSV = New-Object System.Windows.Forms.Button
$buttonSelectCSV.Text = "Select CSV"
$buttonSelectCSV.Location = New-Object System.Drawing.Point(330, 350)
$buttonSelectCSV.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($buttonSelectCSV)

$buttonSelectCSV.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textCSVPath.Text = $openFileDialog.FileName
    }
})

# Output file selection
$labelOutputFile = Create-Label "Output File:" (New-Object System.Drawing.Point(10, 420))
$form.Controls.Add($labelOutputFile)

$textOutputFilePath = New-Object System.Windows.Forms.TextBox
$textOutputFilePath.Location = New-Object System.Drawing.Point(120, 420)
$textOutputFilePath.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textOutputFilePath)

$buttonSelectOutputFile = New-Object System.Windows.Forms.Button
$buttonSelectOutputFile.Text = "Select Output"
$buttonSelectOutputFile.Location = New-Object System.Drawing.Point(330, 420)
$buttonSelectOutputFile.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($buttonSelectOutputFile)

$buttonSelectOutputFile.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textOutputFilePath.Text = $saveFileDialog.FileName
    }
})

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 460)
$progressBar.Size = New-Object System.Drawing.Size(420, 20)
$form.Controls.Add($progressBar)

# Execute button
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = "Execute"
$buttonExecute.Location = New-Object System.Drawing.Point(10, 500)
$buttonExecute.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($buttonExecute)

$buttonExecute.Add_Click({
    try {
        # Validate required fields
        $requiredFields = @(
            @{ Field = $textServer.Text; Name = "SQL Server" }
            @{ Field = $textDatabase.Text; Name = "Database" }
            @{ Field = $textTableName.Text; Name = "Table Name" }
            @{ Field = $textCSVPath.Text; Name = "CSV File" }
            @{ Field = $textOutputFilePath.Text; Name = "Output File" }
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
                [System.Windows.Forms.MessageBox]::Show("$($field.Name) is required.", "Validation Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
            $columns = $row.PSObject.Properties.Name -join ', '
            $values = ($row.PSObject.Properties.Value | ForEach-Object { 
                if ($_ -eq $null) { "NULL" } else { "'$($_ -replace "'", "''")'" }
            }) -join ', '

            $currentBatch += "INSERT INTO $($textTableName.Text) ($columns) VALUES ($values)"
            
            $processedRows++
            $progressBar.Value = $processedRows

            if ($currentBatch.Count -eq $batchSize -or $processedRows -eq $totalRows) {
                $batchQuery = $currentBatch -join "`n"
                try {
                    Invoke-Sqlcmd -ConnectionString $connectionString -Query $batchQuery
                } catch {
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
                    foreach ($table in $results) {
                        Write-Host "Result set contains $($table.Rows.Count) rows"
                        $table | Format-Table -AutoSize | Out-String | Write-Host
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

# Show the form
$form.ShowDialog()