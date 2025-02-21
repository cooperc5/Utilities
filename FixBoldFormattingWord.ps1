Add-Type -TypeDefinition @"
using System;
using System.Windows.Forms;
public class OpenFileDialogHelper {
    public static string SelectFile() {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "Word Documents|*.docx";
        openFileDialog.Title = "Select a Word Document";
        return openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : null;
    }
}
"@ -Language CSharp

# Open File Dialog to Select Word Document
$originalFilePath = [OpenFileDialogHelper]::SelectFile()
if (-not $originalFilePath) {
    Write-Host "No file selected. Exiting..."
    exit
}

# Create a copy of the original file
$directory = [System.IO.Path]::GetDirectoryName($originalFilePath)
$originalFileName = [System.IO.Path]::GetFileNameWithoutExtension($originalFilePath)
$extension = [System.IO.Path]::GetExtension($originalFilePath)
$copyFilePath = [System.IO.Path]::Combine($directory, "${originalFileName}_cleanedUp$extension")

Copy-Item -Path $originalFilePath -Destination $copyFilePath -Force
Write-Host "Backup copy created: $copyFilePath"

# Open Word Application
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($copyFilePath)

# Define the wildcard pattern to match **any text here**
$find = $doc.Content.Find
$find.Text = "\*\*(*)\*\*"  # Word uses ( ) to capture groups in wildcard searches
$find.MatchWildcards = $true

# Loop to replace all matches
do {
    $found = $find.Execute()
    if ($found) {
        $range = $find.Parent
        $textToReplace = $range.Text -replace "^\*\*(.*)\*\*$", '$1'  # Remove ** while keeping the text
        $range.Text = $textToReplace
        $range.Font.Bold = $true  # Apply bold formatting
    }
} while ($found)

# Save and Close
$doc.Save()
$doc.Close()
$word.Quit()

Write-Host "All placeholders replaced successfully in the cleaned-up copy: $copyFilePath"
