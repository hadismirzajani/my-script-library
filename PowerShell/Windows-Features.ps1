# Lists all available roles and features (Note: This specific cmdlet works on Windows Server operating systems only).
Get-WindowsFeature 

# Finds the IIS Web Server role and automatically installs it onto the Windows Server.
Get-WindowsFeature -name Web-Server | Install-WindowsFeature 

# Lists advanced operating system packages (like OpenSSH or RSAT tools) available online or locally.
Get-WindowsCapability -online 

# Lists optional operating system features (like Hyper-V or WSL) that can be turned on or off.
Get-WindowsOptionalFeature -online
