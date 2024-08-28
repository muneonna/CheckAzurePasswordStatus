# Define the necessary modules
$requiredModules = @("Microsoft.Graph.Users")

# Check and install necessary modules
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        try {
            Install-Module -Name $module -Force -Scope CurrentUser
            Write-Host "Installed module: $module"
        } catch {
            Write-Error "Could not install module: $module. Error: $_"
            return
        }
    } else {
        Write-Host "Module already installed: $module"
    }
}

# Execute the import and user retrieval in a new script block
& {
    Import-Module Microsoft.Graph.Users
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
    $properties = @('DisplayName', 'UserPrincipalName', 'AccountEnabled', 'LastPasswordChangeDateTime')
    $users = Get-MgUser -All -Property $properties | 
             Where-Object { $_.UserPrincipalName -notlike '*#*' -and $_.AccountEnabled -eq $true -and (New-TimeSpan -Start $_.LastPasswordChangeDateTime).Days -gt 90 } |
             Select-Object $properties

    # Create a timestamp for the CSV file name
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $csvFileName = "C:\UsersInfo_$timestamp.csv"

    # Export the user information to a CSV file
    $users | Export-Csv -Path $csvFileName -NoTypeInformation
    
    # Write the file location to the console
    Write-Host "CSV file has been created at: $csvFileName"
    
    # Attempt to open the CSV file
    Try {
        Start-Process "excel.exe" $csvFileName -ErrorAction Stop
    } Catch {
        Write-Error "Failed to open CSV file. Error: $_"
    }
}