# Requires modules: SqlServer, ImportExcel (EPPlus)
Add-Type -AssemblyName System.Windows.Forms

# Create the form and controls
$form = New-Object System.Windows.Forms.Form
$form.Text = "CSV to SQL Execution"
$form.Size = New-Object System.Drawing.Size(500, 650)
$form.StartPosition = "CenterScreen"

# Create labels and textboxes for SQL Server, database, table name, authentication, etc.
$labelServer = New-Object System.Windows.Forms.Label
$labelServer.Text = "SQL Server:"
$labelServer.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($labelServer)

$textServer = New-Object System.Windows.Forms.TextBox
$textServer.Location = New-Object System.Drawing.Point(120, 20)
$textServer.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textServer)

$labelDatabase = New-Object System.Windows.Forms.Label
$labelDatabase.Text = "Database:"
$labelDatabase.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($labelDatabase)

$textDatabase = New-Object System.Windows.Forms.TextBox
$textDatabase.Location = New-Object System.Drawing.Point(120, 60)
$textDatabase.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textDatabase)

$labelTableName = New-Object System.Windows.Forms.Label
$labelTableName.Text = "Table Name:"
$labelTableName.Location = New-Object System.Drawing.Point(10, 100)
$form.Controls.Add($labelTableName)

$textTableName = New-Object System.Windows.Forms.TextBox
$textTableName.Location = New-Object System.Drawing.Point(120, 100)
$textTableName.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textTableName)

# Radio buttons to select between windows authentication or SQL server authentication
$radioWindowsAuth = New-Object System.Windows.Forms.RadioButton
$radioWindowsAuth.Text = "Windows Authentication"
$radioWindowsAuth.Location = New-Object System.Drawing.Point(10, 140)
$radioWindowsAuth.Size = New-Object System.Drawing.Size(100, 30)
$radioWindowsAuth.Checked = $true
$form.Controls.Add($radioWindowsAuth)

$radioSQLAuth = New-Object System.Windows.Forms.RadioButton
$radioSQLAuth.Text = "SQL Server Authentication"
$radioSQLAuth.Location = New-Object System.Drawing.Point(180, 140)
$radioSQLAuth.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($radioSQLAuth)

# Create labels and textboxes for Username and Password
$labelUsername = New-Object System.Windows.Forms.Label
$labelUsername.Text = "Username:"
$labelUsername.Location = New-Object System.Drawing.Point(10, 180)
$labelUsername.Visible = $false
$form.Controls.Add($labelUsername)

$textUsername = New-Object System.Windows.Forms.TextBox
$textUsername.Location = New-Object System.Drawing.Point(120, 180)
$textUsername.Size = New-Object System.Drawing.Size(200, 20)
$textUsername.Visible = $false
$form.Controls.Add($textUsername)

$labelPassword = New-Object System.Windows.Forms.Label
$labelPassword.Text = "Password:"
$labelPassword.Location = New-Object System.Drawing.Point(10, 220)
$labelPassword.Visible = $false
$form.Controls.Add($labelPassword)

$textPassword = New-Object System.Windows.Forms.TextBox
$textPassword.Location = New-Object System.Drawing.Point(120, 220)
$textPassword.Size = New-Object System.Drawing.Size(200, 20)
$textPassword.PasswordChar = "*"
$textPassword.Visible = $false
$form.Controls.Add($textPassword)

$radioWindowsAuth.Add_CheckedChanged({
    if ($radioWindowsAuth.Checked) {
        $labelUsername.Visible = $false
        $textUsername.Visible = $false
        $labelPassword.Visible = $false
        $textPassword.Visible = $false
    }
})

$radioSQLAuth.Add_CheckedChanged({
    if ($radioSQLAuth.Checked) {
        $labelUsername.Visible = $true
        $textUsername.Visible = $true
        $labelPassword.Visible = $true
        $textPassword.Visible = $true
    }
})

# Radio buttons to select between stored procedure or SQL script
$radioStoredProcedure = New-Object System.Windows.Forms.RadioButton
$radioStoredProcedure.Text = "Stored Procedure"
$radioStoredProcedure.Location = New-Object System.Drawing.Point(10,260)
$radioStoredProcedure.Size = New-Object System.Drawing.Size(100, 30)
$radioStoredProcedure.Checked = $true
$form.Controls.Add($radioStoredProcedure)

$radioSqlScript = New-Object System.Windows.Forms.RadioButton
$radioSqlScript.Text = "SQL Script"
$radioSqlScript.Location = New-Object System.Drawing.Point(180, 260)
$radioSqlScript.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($radioSqlScript)

# Stored procedure input
$labelStoredProcedure = New-Object System.Windows.Forms.Label
$labelStoredProcedure.Text = "Stored Procedure:"
$labelStoredProcedure.Location = New-Object System.Drawing.Point(10, 300)
$form.Controls.Add($labelStoredProcedure)

$textStoredProcedure = New-Object System.Windows.Forms.TextBox
$textStoredProcedure.Location = New-Object System.Drawing.Point(120, 300)
$textStoredProcedure.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textStoredProcedure)

# SQL script file selection
$labelSQLScript = New-Object System.Windows.Forms.Label
$labelSQLScript.Text = "SQL Script File:"
$labelSQLScript.Location = New-Object System.Drawing.Point(10, 300)
$labelSQLScript.Visible = $false
$form.Controls.Add($labelSQLScript)

$textSQLScriptPath = New-Object System.Windows.Forms.TextBox
$textSQLScriptPath.Location = New-Object System.Drawing.Point(120, 300)
$textSQLScriptPath.Size = New-Object System.Drawing.Size(200, 20)
$textSQLScriptPath.Visible = $false
$form.Controls.Add($textSQLScriptPath)

$buttonSelectSQLScript = New-Object System.Windows.Forms.Button
$buttonSelectSQLScript.Text = "Select Script"
$buttonSelectSQLScript.Location = New-Object System.Drawing.Point(330, 300)
$buttonSelectSQLScript.Size = New-Object System.Drawing.Size(100, 20)
$buttonSelectSQLScript.Visible = $false
$form.Controls.Add($buttonSelectSQLScript)

# Toggle between stored procedure and SQL script input
$radioStoredProcedure.Add_CheckedChanged({
    if ($radioStoredProcedure.Checked) {
        $labelStoredProcedure.Visible = $true
        $textStoredProcedure.Visible = $true
        $labelSQLScript.Visible = $false
        $textSQLScriptPath.Visible = $false
        $buttonSelectSQLScript.Visible = $false
    }
})

$radioSqlScript.Add_CheckedChanged({
    if ($radioSqlScript.Checked) {
        $labelStoredProcedure.Visible = $false
        $textStoredProcedure.Visible = $false
        $labelSQLScript.Visible = $true
        $textSQLScriptPath.Visible = $true
        $buttonSelectSQLScript.Visible = $true
    }
})

$buttonSelectSQLScript.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "SQL files (*.sql)|*.sql"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textSQLScriptPath.Text = $openFileDialog.FileName
    }
})

# Button to select the CSV file
$labelCSVFile = New-Object System.Windows.Forms.Label
$labelCSVFile.Text = "CSV File:"
$labelCSVFile.Location = New-Object System.Drawing.Point(10, 350)
$form.Controls.Add($labelCSVFile)

$textCSVPath = New-Object System.Windows.Forms.TextBox
$textCSVPath.Location = New-Object System.Drawing.Point(120, 350)
$textCSVPath.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textCSVPath)

$buttonSelectCSV = New-Object System.Windows.Forms.Button
$buttonSelectCSV.Text = "Select CSV File"
$buttonSelectCSV.Location = New-Object System.Drawing.Point(330, 350)
$buttonSelectCSV.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($buttonSelectCSV)

$buttonSelectCSV.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textCSVPath.Text = $openFileDialog.FileName
    }
})

# Output file selection
$labelOutputFile = New-Object System.Windows.Forms.Label
$labelOutputFile.Text = "Output File:"
$labelOutputFile.Location = New-Object System.Drawing.Point(10, 420)
$form.Controls.Add($labelOutputFile)

$textOutputFilePath = New-Object System.Windows.Forms.TextBox
$textOutputFilePath.Location = New-Object System.Drawing.Point(120, 420)
$textOutputFilePath.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textOutputFilePath)

$buttonSelectOutputFile = New-Object System.Windows.Forms.Button
$buttonSelectOutputFile.Text = "Select Output"
$buttonSelectOutputFile.Location = New-Object System.Drawing.Point(333, 420)
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

# Submit button to execute the stored procedure or script
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = "Execute"
$buttonExecute.Location = New-Object System.Drawing.Point(10, 500)
$form.Controls.Add($buttonExecute)

$buttonExecute.Add_Click({
    try {
        # Check required fields
        $server = $textServer.Text
        $database = $textDatabase.Text
        $tableName = $textTableName.Text
        $username = $textUsername.Text
        $password = $textPassword.Text

        $csvFilePath = $textCSVPath.Text
        $auth = $textAuth.Text
        $storedProcedure = $textStoredProcedure.Text
        $sqlScriptPath = $textSQLScriptPath.Text
        $outputFilePath = $textOutputFilePath.Text

        if ([string]::IsNullOrWhiteSpace($server) -or [string]::IsNullOrWhiteSpace($database) -or 
            [string]::IsNullOrWhiteSpace($tableName) -or [string]::IsNullOrWhiteSpace($csvFilePath) -or 
            ($radioStoredProcedure.Checked -and [string]::IsNullOrWhiteSpace($storedProcedure)) -or 
            ($radioSqlScript.Checked -and [string]::IsNullOrWhiteSpace($sqlScriptPath)) -or
            [string]::IsNullOrWhiteSpace($outputFilePath)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all required fields.")
            return
        }

        # Determine connection string
        if ($radioWindowsAuth.Checked) {
            $connectionString = "Server=$server;Database=$database;Integrated Security=True;"
        } else {
            $username = $textUsername.Text
            $password = $textPassword.Text
            $connectionString = "Server=$server;Database=$database;User Id=$username;Password=$password;"
        }

        # Load CSV data
        $csvData = Import-Csv -Path $csvFilePath

        # Execute stored procedure or SQL script based on user selection
        $results = @()
        if ($radioStoredProcedure.Checked) {
            $query = "EXEC $storedProcedure"
            $results = Invoke-Sqlcmd -ConnectionString $connString -Query $query -As DataTables
        } elseif ($radioSqlScript.Checked -and (Test-Path $sqlScriptPath)) {
            $script = Get-Content $sqlScriptPath -Raw
            $host.UI.RawUI.WindowTitle = "SQL Script Execution"
            $host.UI.WriteLine("Executing the following SQL script :`n$script")
            $results = Invoke-Sqlcmd -ConnectionString $connString -InputFile $script -As DataTables
        }
        [System.Windows.Forms.MessageBox]::Show("Total results : $($results.count)")
        # Export each result set to a separate worksheet
        if ($results.Count -gt 0) {
            $index = 1
            foreach ($resultSet in $results) {
                $worksheetName = "ResultSet$index"
                $resultSet | Export-Csv -Path $outputFilePath -WorksheetName $worksheetName -AutoSize
                $index++
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Execution completed successfully. Output saved to $outputFilePath.")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_")
    }
})

$form.ShowDialog()
