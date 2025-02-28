# Office Snippets - Word

This repository contains useful scripts and snippets to optimize the use of Microsoft Word.

## 🖥️ Script to Enable Macros in Word

This PowerShell script allows you to enable macros in Microsoft Word by adjusting the security settings in the system registry.

### 📋 Requirements
- PowerShell 5.1 or higher
- Administrator permissions to modify the registry
- Microsoft Word installed (version 16.0 or higher)

### 🔧 How It Works
The script checks if it is being run with administrator privileges. If not, it requests privilege elevation. Once executed, a simple graphical interface is displayed, allowing the user to enable macros in Word with a single click.

### 🔄 Steps Performed by the Script
1. Checks if the script is being run with administrator privileges.
2. If not, it requests privilege elevation.
3. Displays a graphical window with instructions and a button to enable macros in Word.
4. Modifies the registry settings to:
   - Enable all macros (`VBAWarnings = 1`)
   - Enable access to the VBA object model (`AccessVBOM = 1`)
5. Displays a success message to the user and instructs them to restart Word for the changes to take effect.

<br>
<br>

### 📝 How to Use
1. Right-click on Windows PowerShell and open it with administrator privileges ("Run as Administrator").
2. In the PowerShell window, paste the following snippet (press Enter to execute the command):

   ```powershell
   irm "https://raw.githubusercontent.com/masalles/office-snippets-word/refs/heads/main/enable_macros_word.ps1" | iex
   ```
A graphical window will appear with an "Enable Macros" button. Click it to apply the changes.

After execution, a success message will be displayed.

<code style="color: darkorange">Restart Microsoft Word to apply the new settings.</code>
<br>
<br>

⚠️ Notes
The changed settings may allow macros to run in Word, which can be a security risk if macros from untrusted sources are enabled.

The script has been tested on Microsoft Word 16.0 but may work on earlier versions (the registry key may vary).
---

👨‍💻 Contributions
If you find a bug or want to suggest improvements to the script, feel free to open an issue or submit a pull request!

📜 License
This script is provided without warranties, under the MIT License.
