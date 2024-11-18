# Registry Paths
$pathExcel = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$pathWord = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
$pathPowerPoint = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Options"

# Registry Type
$type = "DWord"

# Registry Name
$name = "DefaultFormat"

# Main function
function Main {
    # Excel
    Set-RegistryValueIfNotExist($pathExcel, $name, 51, $type)
    # Word
    Set-RegistryValueIfNotExist($pathWord, $name, 51, $type)
    # PowerPoint
    Set-RegistryValueIfNotExist($pathPowerPoint, $name, 27, $type)
}

# Function to set registry value if it doesn't exist
function Set-RegistryValueIfNotExist {
    param (
        [string]$Path,
        [string]$Name,
        [object]$Value,
        [Microsoft.Win32.RegistryValueKind]$Type
    )

    if (Test-Path $Path) {
        $currentValue = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue).$Name
        if ($currentValue -ne $Value) {
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type
            Write-Output "Set $Name to $Value at $Path"
        } else {
            Write-Output "$Name is already set to $Value at $Path"
        }
    } else {
        Write-Output "The path $Path does not exist. Creating path and setting value."
        New-Item -Path $Path -Force | Out-Null
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type
    }
}

# Call the Main function to start the script
Main

# Set the default file format for Excel
#Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -Name "DefaultFormat" -Value 51 -Type DWord
# Set the default file format for PowerPoint
#Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Options" -Name "DefaultFormat" -Value 27 -Type DWord
# Set the default file format for Word
#Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Options" -Name "DefaultFormat" -Value 51 -Type DWord
# Set the default file format for Acess
#Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Access\Settings" -Name "DefaultFileFormat" -Value 12 -Type DWord