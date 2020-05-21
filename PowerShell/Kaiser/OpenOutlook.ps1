Add-Type -assembly "Microsoft.Office.Interop.Outlook"
add-type -assembly "System.Runtime.Interopservices"
try
{
$outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
    $outlookWasAlreadyRunning = $true
}
catch
{
    try
    {
        $Outlook = New-Object -comobject Outlook.Application
        $outlookWasAlreadyRunning = $false
    }
    catch
    {
        write-host "You must exit Outlook first."
        exit
    }
}
$namespace = $Outlook.GetNameSpace("MAPI")

$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)