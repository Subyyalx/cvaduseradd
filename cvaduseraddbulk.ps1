
#ADD OU TOO
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Add User to AD"
$form.Size = New-Object System.Drawing.Size(400,300)

# Create labels
$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "CSV File Path:"
$label1.Location = New-Object System.Drawing.Point(10,20)
$label1.Size = New-Object System.Drawing.Size(100,20)

# Create textbox for CSV file path
$textBoxCSV = New-Object System.Windows.Forms.TextBox
$textBoxCSV.Location = New-Object System.Drawing.Point(120,20)
$textBoxCSV.Size = New-Object System.Drawing.Size(200,20)

# Create button to browse for CSV file
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = New-Object System.Drawing.Point(330,20)
$buttonBrowse.Size = New-Object System.Drawing.Size(50,20)
$buttonBrowse.Text = "Browse"
$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Select a CSV File"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $textBoxCSV.Text = $openFileDialog.FileName
    }
})

# Create button to add users from CSV
$buttonAddCSV = New-Object System.Windows.Forms.Button
$buttonAddCSV.Location = New-Object System.Drawing.Point(120,60)
$buttonAddCSV.Size = New-Object System.Drawing.Size(150,40)
$buttonAddCSV.Text = "Add Users from CSV"
$buttonAddCSV.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonAddCSV.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)  # Blue color
$buttonAddCSV.ForeColor = [System.Drawing.Color]::White
$buttonAddCSV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonAddCSV.Add_Click({
    $csvPath = $textBoxCSV.Text
    if ([string]::IsNullOrEmpty($csvPath) -or -not (Test-Path $csvPath)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid CSV file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
        return
    }

    # Read CSV file and add users
    $csvData = Import-Csv $csvPath
    if ($csvData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("CSV file is empty.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
        return
    }

    # Define credentials for an Active Directory admin user
    $adminUsername = "Administrator"
    $adminPassword = ConvertTo-SecureString "Xavor123" -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($adminUsername, $adminPassword)

    foreach ($user in $csvData) {
        $firstName = $user.'First Name'
        $lastName = $user.'Last Name'
        $initials = $user.Initials
        $username = $user.Username
        $displayName = "$firstName $initials $lastName"
        $email = "$username@xavorlms.local"

        try {
            New-ADUser -Name $displayName -GivenName $firstName -Surname $lastName -Initials $initials -SamAccountName $username -UserPrincipalName $email -EmailAddress $email -AccountPassword (ConvertTo-SecureString "XvrLhr123" -AsPlainText -Force) -Enabled $true -DisplayName $displayName -Credential $credential -Server "xavorlms.local"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error adding user '$displayName': $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Users added from CSV file.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK)
})

# Add controls to form
$form.Controls.Add($label1)
$form.Controls.Add($textBoxCSV)
$form.Controls.Add($buttonBrowse)
$form.Controls.Add($buttonAddCSV)

# Show the form
$form.ShowDialog()
