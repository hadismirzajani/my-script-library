# Lists all PowerShell modules that are currently installed and available to be used on this computer.
Get-Module -listAvailable 

# Explicitly loads the AppLocker security management module into your current PowerShell session.
Import-Module -name applocker 

# Lists every command, cmdlet, and function available inside the AppLocker module.
Get-Command -Module applocker 
