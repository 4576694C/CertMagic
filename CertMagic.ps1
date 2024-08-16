Function Admin-Check 
{
	<#
	.SYNOPSIS
	Function used to verify if script is run as administrator. Returns True / False.
	.EXAMPLE
		PS> Admin-Check 
	#>
	try {
        $Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $Principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $Identity
        return $Principal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
    } catch {
        Write-Warning "Could not determine wether current context is elevated or not!"
        Write-Host ""
        Write-Host "Error information below:"
        Write-Host ""
    
        $_ | Format-List * -Force | Out-String -Stream | Write-Host -ForegroundColor Yellow
    
        Write-Host ""
    
        throw "Could not determine wether current context is elevated or not!: 0x{0:x} - {1}" -f $_.Exception.HResult, $_.Exception.Message
    }
}


Function Open-CertStore 
{
<#
    .SYNOPSIS
    Open-CertStore - used to open certificate store to view all certificates including archived ones.
           
    .PARAMETER personalStore
    Output of query passed to other functions
    .EXAMPLE
	PS> Open-CertStore
    
#>
    [CmdletBinding()]
    param (
        $personalStore
    )
    $personalStore = Get-Item cert:\LocalMachine\My 
    $personalStore.Open('ReadWrite,IncludeArchived') 
    $personalStore.Certificates | Select Thumbprint, Subject, Archived, NotAfter | sort NotAfter -Descending
}   



function Get-AllCerts 
{
<#
    .SYNOPSIS
    Get-AllCerts - used to list all of the certificates to screen inside the GUI loop.
    
    .EXAMPLE
    PS> Get-AllCerts 
#>
    
    Open-CertStore | ft
    Pause
}

function Get-ExpiredCerts 
{
<#
    .SYNOPSIS
    Get-ExpiredCerts - used to list all expired certificates to screen inside the GUI loop.
    
    .EXAMPLE
    PS> Get-ExpiredCerts 
#>
    $expired = Open-CertStore | where {$_.Archived -eq $False -and $_.notAfter -lt (Get-Date)} | ft
    if ($expired.count -ne 0)
        {
            $expired
        }
    else
        {
            Write-Host "No expired certificates detected!" -ForegroundColor Yellow
        }
    Pause
}
function Set-ArchivedFlagAll 
{   
<#
    .SYNOPSIS
    Set-ArchivedFlagAll  - used to archive all expired certificates.
    
    .EXAMPLE
    PS> Set-ArchivedFlagAll  
#>
    $personalStore = Get-Item cert:\LocalMachine\My 
    $personalStore.Open('ReadWrite,IncludeArchived') 
    $toArchive = $personalStore.certificates | where {$_.Archived -eq $False -AND $_.notAfter -lt (Get-Date)}
    if (($toArchive.count) -ne 0)
        {
            Write-Host "`nThe following certs will be archived:"`n 
            $toArchive | select Subject, notafter,thumbprint | ft -HideTableHeaders
            $proceed = Read-Host "Do you wish to proceed (Y/N) ?"
            if ($proceed -match "[yY]")
                {
                    foreach ($cert in $personalStore.certificates |  where {$_.Archived -eq $False -and $_.notAfter -lt (Get-Date)}){$cert.Archived=$true}
                    Write-host "Certificates Archived." -ForegroundColor Green
                }
        }               
    else
        {
            Write-Host "No expired certificates detected!" -ForegroundColor Yellow
        }
        Pause
}

function Set-ArchivedFlagSpecific 
{
<#
    .SYNOPSIS
    Set-ArchivedFlagSpecific - used to archive a selected expired certificate.
    
    .EXAMPLE
    PS> Set-ArchivedFlagSpecific 
#>
    $array = New-object system.collections.generic.list[system.object]
    $count = 1
    $personalStore = Get-Item cert:\LocalMachine\My 
    $personalStore.Open('ReadWrite,IncludeArchived') 
    $toArchive = $personalStore.certificates | where {$_.Archived -eq $False -AND $_.notAfter -lt (Get-Date)}
    if (($toArchive.count) -ne 0)
        {
            foreach ($cert in $toarchive)
                {
                    $array.add(($cert | Select-Object `
                    @{name = '#'; expression = { $count }},subject,issuer,notafter))
                    $count++ 
                }
            $array | Out-host
            try {
                    [int]$choice = read-host 'Choose cert to renew (1-99): ' -ErrorAction Stop
                }
            catch 
                {
                    write-Host "Invalid entry" -ForegroundColor yellow
                    Pause
                    GUI
                }
            $choice = $choice-1
            $selected = (($toArchive | select subject,issuer,notafter,ThumbPrint))[$choice]
            if (!$selected){write-Host "Invalid entry" -ForegroundColor yellow ;Pause;GUI}
            Write-Host $selected -ForegroundColor Yellow
            $proceed = Read-Host "Do you wish to proceed (Y/N) ?"
            if ($proceed -match "[yY]")
                {
                    foreach ($cert in $personalStore.certificates |  where {$_.ThumbPrint -eq $selected.Thumbprint})  { $cert.Archived=$true }
                    Write-host "Certificates Archived." -ForegroundColor Green
                }
        }               
    else
        {
            Write-Host "No expired certificates detected!" -ForegroundColor Yellow
        }
    Pause
}
    
Function Renew-Certificate 
{
<#
    .SYNOPSIS
    Script used to renew a certificate selected through Out-GridView to renew with the same key.
    Works only on certs that have not yet expired.
    
    .EXAMPLE
    PS> Renew-Certificate
#>
    
    Write-Host "Use this tool to renew expireing certificates with the same key. Works only on not yet expired certs"
    $array = New-object system.collections.generic.list[system.object]
    $certs = Get-childItem cert:\LocalMachine\My | 
            Select-Object * |
            where notafter -gt (get-date).AddMinutes(1) |
            Sort-Object notafter
    $count = 1; 
    foreach ($cert in $certs) 
        { 
            $array.add(($cert | Select-Object `
                @{name = '#'; expression = { $count }}, 
                subject,
                issuer,
                notafter
                ))
            $count++ 
        }
    
    $array | Out-Host
    try 
        {
            [int]$choice = read-host 'Choose cert to renew (1-99): ' -ErrorAction Stop
        }
    catch 
    {
        write-Host "Invalid entry" -ForegroundColor yellow
        Pause
    }
    $choice = $choice-1
    
    $selected = (($certs | select subject,issuer,notafter,SerialNumber))[$choice]
    if (!$selected)
        {
            write-Host "Invalid entry" -ForegroundColor yellow
            Pause
        }
    Write-Host $selected -ForegroundColor Yellow
    $proceed = Read-Host "Has been chosen to be renewed, proceed (Y/N) ?"
    if ($proceed -match "[yY]")
        {
            Write-host "Renewing" $selected.serialnumber
            &certreq @('-Enroll', '-machine', '-q', '-cert', $certToRenew.SerialNumber, 'Renew', 'ReuseKeys')    
        }
    Pause
}

function Get-DummyCert 
{
<#
    .SYNOPSIS
    Get-DummyCert - used to create a quick dummy (by default expired) certificate

    .PARAMETER CertStoreLocation
    Used for selecting cert store location

    .PARAMETER DnsName
    Used to provide DNS name for the certificate

    .PARAMETER FriendlyName
    Used to provide FriendlyName to the certificate

    .PARAMETER NotAfter
    Used to provide expiration date

    .EXAMPLE
    PS> Get-DummyCert -FriendlyName "BestCertEver" -dnsname "dummy.Shodan" -notAfter (get-date).AddDays(1)

#>
    [CmdletBinding()]
    param (
        $CertStoreLocation = "Cert:\LocalMachine\My",
        $DnsName = "dummy.dummylabs",
        $FriendlyName = "dummy",
        $NotAfter = (get-date)
    )
    New-SelfSignedCertificate -CertStoreLocation $CertStoreLocation -DnsName $DnsName -FriendlyName $FriendlyName -NotAfter $NotAfter -Verbose | out-host
    Pause
}

function GUI 
{
<#
.SYNOPSIS
Gui for the CertMagic Tool

.EXAMPLE
GUI
#>    
    while ($menu -ne 0) 
        {
            Write-Host    "###########################################" -ForegroundColor DarkCyan
            Write-Host    "#            CERT MAGIC 101               #" -ForegroundColor DarkCyan
            Write-Host    "###########################################" -ForegroundColor DarkCyan
            Write-Host    "#                                         #" -ForegroundColor DarkCyan
            Write-Host    "# 1 - List All Certificates               #" -ForegroundColor DarkCyan
            Write-Host    "# 2 - List Expired Certificates           #" -ForegroundColor DarkCyan
            Write-Host    "# 3 - Archive All Expired Certificates    #" -ForegroundColor DarkCyan
            Write-Host    "# 4 - Archive Specific Expired Cert       #" -ForegroundColor DarkCyan
            Write-Host    "# 5 - Renew Certificate with the same key #" -ForegroundColor DarkCyan
            Write-Host    "# 6 - Create dummy cert                   #" -ForegroundColor DarkCyan
            Write-Host    "# 0 - Exit                                #" -ForegroundColor DarkCyan
            Write-Host    "#                                         #" -ForegroundColor DarkCyan
            Write-Host    "###########################################" -ForegroundColor DarkCyan
            
            $menu = Read-Host "Select one of the options: "
            switch ( $menu )
                {
                    1 { Get-AllCerts}
                    2 { Get-ExpiredCerts}
                    3 { Set-ArchivedFlagAll}
                    4 { Set-ArchivedFlagSpecific}
                    5 { Renew-Certificate }
                    6 { Get-Dummycert }
                    0 { Break }
                    Default {"Invalid Selection - Exiting";Break}
                }    
        }
}

# Check if script is run as admin
If (!(Admin-Check)) {"Script has to be run as admin!";break}

#Launch Text GUI
GUI