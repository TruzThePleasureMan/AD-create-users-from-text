# Define variables
$OUPath = "OU=ORG-UNIT,DC=DOMAIN-NAME,DC=DOMAIN-TLD"
$templateUser = "TEMPLATE-USER"
$defaultPassword = "DEFAULTPASSWD"
$homeFolderPath = "\\SERVERNAME\UserHome$"
$debug = $true

# Function to create or update AD user with login name generation, home folder connection, profile path, and copying "MemberOf" from template user
function Create-ADUser {
    param (
        [string]$fname,
        [string]$lname,
        [string]$email,
        [string]$phone,
        [string]$postnummer,
        [switch]$debug
    )

    # Generate login name
    $loginName = ($fname.Substring(0, [Math]::Min(3, $fname.Length)) + $lname.Substring(0, [Math]::Min(3, $lname.Length))).ToLower()

    $user = Get-ADUser -Filter "sAMAccountName -eq '$loginName'" -ErrorAction SilentlyContinue

    if ($user) {
        if ($debug) {
            Write-Host "User '$loginName' already exists. Updating attributes."
        }

        # Update user's attributes if needed
        Set-ADUser -Identity $user -EmailAddress $email -OfficePhone $phone -PostalCode $postnummer -ErrorAction Stop

        # Update user logon name
        Set-ADUser -Identity $user -UserPrincipalName "$loginName@SERVER.local" -ErrorAction Stop

        # Update user profile path
        Set-ADUser -Identity $user -ProfilePath "\\SERVER-NAME\Profile$\$loginName" -ErrorAction Stop

        if ($debug) {
            Write-Host "Updated user: $($user.Name)"
        }

        return $user
    }

    # Get template user
    $templateUserObject = Get-ADUser -Identity $templateUser -Properties MemberOf

    $user = New-ADUser -Name "$fname $lname" -GivenName $fname -Surname $lname -EmailAddress $email -OfficePhone $phone -AccountPassword (ConvertTo-SecureString -AsPlainText $defaultPassword -Force) -ChangePasswordAtLogon $true -Enabled $true -PassThru -SamAccountName $loginName -UserPrincipalName "$loginName@DOMAIN.local" -DisplayName "$fname $lname" -PostalCode $postnummer -ProfilePath "\\SERVER-NAME\Profile$\$loginName"

    if ($debug) {
        Write-Host "Created new user: $($user.Name)"
    }

    # Set MemberOf attribute
    foreach ($group in $templateUserObject.MemberOf) {
        Add-ADGroupMember -Identity $group -Members $user
    }

    return $user
}

# Read user data from a text file
$users = Get-Content -Path "C:\PATHTOTEXTFILE.txt" | ConvertFrom-Csv -Header "fname", "lname", "email", "phone", "postnummer"

# Create or update users in AD
foreach ($user in $users) {
    $fname = $user.fname
    $lname = $user.lname
    $email = $user.email
    $phone = $user.phone
    $postnummer = $user.postnummer  # Ensure postnummer is correctly assigned

    # Generate login name
    $loginName = ($fname.Substring(0, [Math]::Min(3, $fname.Length)) + $lname.Substring(0, [Math]::Min(3, $lname.Length))).ToLower()

    Write-Host "Postnummer for $($fname) $($lname): $($postnummer)"  # Output postnummer value for verification

    # Create or update user
    $userObject = Create-ADUser -fname $fname -lname $lname -email $email -phone $phone -postnummer $postnummer -debug $debug

    if ($userObject) {
        # Move the new user to the specified OU
        try {
            Move-ADObject -Identity $userObject.DistinguishedName -TargetPath $OUPath -ErrorAction Stop
            if ($debug) {
                Write-Host "Moved user to OU: $OUPath"
            }
        } catch {
            Write-Error "Failed to move user to OU: $_"
        }

        # Create home folder and set permissions
        $homeFolder = Join-Path $homeFolderPath $userObject.SamAccountName
        New-Item -Path $homeFolder -ItemType Directory -ErrorAction Stop

        # Grant user full control over their home folder
        $acl = Get-Acl -Path $homeFolder
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("$($userObject.SamAccountName)","FullControl","ContainerInherit,ObjectInherit","None","Allow")
        $acl.AddAccessRule($rule)
        Set-Acl -Path $homeFolder -AclObject $acl

        # Connect home folder
        Set-ADUser -Identity $userObject.SamAccountName -HomeDrive "Z:" -HomeDirectory $homeFolderPath\$($userObject.SamAccountName) -ErrorAction Stop

    }
}
