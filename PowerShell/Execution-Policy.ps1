# Checks the current system security rule for running PowerShell scripts.
Get-ExecutionPolicy 

# Changes the system security rule to allow or restrict the running of PowerShell scripts.
# 
# Main execution policy types to choose from:
# 1. Restricted   - Default for Windows clients. Blocks all scripts from running; allows individual commands only.
# 2. AllSigned    - Only allows scripts to run if they are signed by a trusted publisher (even scripts you write yourself).
# 3. RemoteSigned - Default for Windows servers. Local scripts run freely; scripts downloaded from the internet must be digitally signed.
# 4. Unrestricted - Runs all scripts. Warns you before running unsigned files downloaded from the internet.
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
