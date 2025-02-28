# Checks if the script is running as an administrator
$isAdmin = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $isAdmin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Requests elevated privileges
    $argList = @("powershell.exe", "-NoProfile", "-ExecutionPolicy", "ByPass", "-Command", "& '" + $myinvocation.mycommand.definition + "'")
    Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Verb RunAs
    Exit
}

# Load the library to create the graphical interface
Add-Type -AssemblyName System.Windows.Forms

# Create the main window
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Macro Settings - Word"
$Form.Size = New-Object System.Drawing.Size(400, 200)
$Form.StartPosition = "CenterScreen"

# Title for the user
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Click to enable macros in Word"
$Label.Size = New-Object System.Drawing.Size(350, 40)
$Label.Location = New-Object System.Drawing.Point(20, 20)
$Form.Controls.Add($Label)

# Button to execute the configuration
$Button = New-Object System.Windows.Forms.Button
$Button.Text = "Enable Macros"
$Button.Size = New-Object System.Drawing.Size(100, 40)
$Button.Location = New-Object System.Drawing.Point(150, 70)
$Form.Controls.Add($Button)

# User feedback box
$FeedbackLabel = New-Object System.Windows.Forms.Label
$FeedbackLabel.Text = ""
$FeedbackLabel.Size = New-Object System.Drawing.Size(350, 40)
$FeedbackLabel.Location = New-Object System.Drawing.Point(20, 120)
$Form.Controls.Add($FeedbackLabel)

# Button function
$Button.Add_Click({
    # Enables all macros in Word
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "VBAWarnings" -Value 1

    # Enables access to the VBA project object model
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AccessVBOM" -Value 1

    # Updates the interface to give feedback to the user
    $FeedbackLabel.Text = "Settings applied successfully! Restart Word to apply."
    $FeedbackLabel.ForeColor = [System.Drawing.Color]::Green
})

# Displays the window
$Form.ShowDialog()
