<#
automate cert expiraton
WARNING: Script Must be run as administrator
#>

#show all certs
function Get-AllCerts 
    {
        $personalStore = Get-Item cert:\LocalMachine\My 
        $personalStore.Open('ReadWrite,IncludeArchived') 
        $personalStore.Certificates | Select Thumbprint, Subject, Archived, NotAfter | sort NotAfter
        Pause
        gui
    }

#show not archived
function Get-ExpiredCerts 
    {
        $personalStore = Get-Item cert:\LocalMachine\My 
        $personalStore.Open('ReadWrite,IncludeArchived') 
        $personalStore.Certificates | Select Thumbprint, Subject, Archived, NotAfter | where {$_.Archived -eq $False } | where {$_.notAfter -lt (Get-Date)} | sort NotAfter 
        Pause
        gui
    }

#archive all expired
function Set-ArchivedFlagAll 
    {
        $personalStore = Get-Item cert:\LocalMachine\My 
        $personalStore.Open('ReadWrite,IncludeArchived') 
        foreach ($cert in $personalStore.certificates |  where {$_.notAfter -lt (Get-Date)})  { $cert.Archived=$true }
        Pause
        gui
    }

#archive specific expired
function Set-ArchivedFlagSpecific 
    {
        $personalStore = Get-Item cert:\LocalMachine\My 
        $personalStore.Open('ReadWrite,IncludeArchived') 
        $selected = $personalStore.Certificates `
            | Select Thumbprint, Subject, Archived, NotAfter `
            | where {$_.Archived -eq $False } `
            | where {$_.notAfter -lt (Get-Date)} `
            | sort NotAfter `
            | Out-GridView -Title "Select certificate you wish to archive" –PassThru        
        Write-Host "Processing archival of " $selected.Subject
        foreach ($cert in $personalStore.certificates |  where {$_.ThumbPrint -eq $selected.Thumbprint})  { $cert.Archived=$true }    
        Pause
        gui
    }

Function GUI 
{
Write-Host 
@"

    ########################################
    #            CERT MAGIC 101            #
    ########################################
    #                                      #
    # 1 - List All Certificates            #
    # 2 - List Expired Certificates        #
    # 3 - Archive All Expired Certificates #
    # 4 - Archive Specific Expired Cert    #
    # 0 - Exit                             #
    #                                      #
    ########################################
"@
$choice = Read-Host "    Select one of the options: "
switch ( $choice )
    {
        1 { Get-AllCerts}
        2 { Get-ExpiredCerts}
        3 { Set-ArchivedFlagAll}
        4 { Set-ArchivedFlagSpecific}
        0 { Break }
        Default {"    Invalid Selection - Exiting"; Break}
    }
}

Clear
GUI

<#
<#
	.SYNOPSIS
    Script used to renew a certificate selected through Out-GridView to renew with the same key.
    Works only on certs that have not yet expired.
#>    

$certToRenew = Get-childItem cert:\LocalMachine\My | 
                    Select-Object subject, issuer, notafter, enhancedkeyusagelist, thumbprint, serialnumber |
                    Sort-Object notAfter | 
                    Out-GridView -PassThru -Title "Choose cert to renew"
#renewal part
&certreq @('-Enroll', '-machine', '-q', '-cert', $certToRenew.SerialNumber, 'Renew', 'ReuseKeys')


#text version without ogv
Clear-Host
Write-Host "Use this tool to renew expireing certificates with the same key. Works only on not yet expired certs"
$array = New-object system.collections.generic.list[system.object]
$certs = Get-childItem cert:\LocalMachine\My | 
        Select-Object * |
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
$choice = read-host 'Choose cert to renew'
$choice = $choice-1
$selected = (($certs | select subject,issuer,notafter,SerialNumber))[$choice]
Write-Host $selected -ForegroundColor Yellow -NoNewline
$proceed = Read-Host " has been chosen to renew, proceed (Y/N) ?"
if ($proceed -match "[yY]")
    {
        Write-host "Archiving" $selected.serialnumber
    }
else
    {
        exit
    }
#>
