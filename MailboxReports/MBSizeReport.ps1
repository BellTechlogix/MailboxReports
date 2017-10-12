<#
MBSizeReport.ps1
Created By Kristopher Roy
Date Created 02Oct17
Script purpose - Gather details about an exchange mailbox for reporting
#>

#File Select Function - Lets you select your input file
function Get-FileName
{
  param(
      [Parameter(Mandatory=$false)]
      [string] $Filter,
      [Parameter(Mandatory=$false)]
      [switch]$Obj,
      [Parameter(Mandatory=$False)]
      [string]$Title = "Select A File"
    )
   if(!($Title)) { $Title="Select Input File"}
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.FileName = $Title
	#can be set to filter file types
	IF($Filter -ne $null){
		$FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
		$OpenFileDialog.filter = $FilterString}
	if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
		$OpenFileDialog.filter = $Filter}
	$OpenFileDialog.ShowDialog() | Out-Null
	IF($OBJ){
		$fileobject = GI -Path $OpenFileDialog.FileName.tostring()
		Return $fileObject}
	else{Return $OpenFileDialog.FileName}
}

#This Function creates a dialogue to return a Folder Path
function Get-Folder {
    param([string]$Description="Select Folder to place results in",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

#This function allows you to decide on all users or some users
function Select-UserBase
{
	Param(
	[Parameter(Mandatory=$false)]
		[string] $selection
	)
	$title = "Select User Base"
	$message = "Do you wish to poll all Exchange Mailboxes?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
    "Selects All Mailboxes on Exchange."
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
    "Allows selection from import csv."
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
	$result = $host.ui.PromptForChoice($title, $message, $options, 0)
    switch ($result)
    {
        0 {"Yes"}
        1 {"No"}
    }
    Return $selection
}

#This function lets you build an array of specific list items you wish
Function MultipleSelectionBox ($inputarray,$prompt,$listboxtype) {
 
# Taken from Technet - http://technet.microsoft.com/en-us/library/ff730950.aspx
# This version has been updated to work with Powershell v3.0.
# Had to replace $x with $Script:x throughout the function to make it work. 
# This specifies the scope of the X variable.  Not sure why this is needed for v3.
# http://social.technet.microsoft.com/Forums/en-SG/winserverpowershell/thread/bc95fb6c-c583-47c3-94c1-f0d3abe1fafc
#
# Function has 3 inputs:
#     $inputarray = Array of values to be shown in the list box.
#     $prompt = The title of the list box
#     $listboxtype = system.windows.forms.selectionmode (None, One, MutiSimple, or MultiExtended)
 
$Script:x = @()
 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
 
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = $prompt
$objForm.Size = New-Object System.Drawing.Size(300,600) 
$objForm.StartPosition = "CenterScreen"
 
$objForm.KeyPreview = $True
 
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {
        foreach ($objItem in $objListbox.SelectedItems)
            {$Script:x += $objItem}
        $objForm.Close()
    }
    })
 
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})
 
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,520)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
 
$OKButton.Add_Click(
   {
        foreach ($objItem in $objListbox.SelectedItems)
            {$Script:x += $objItem}
        $objForm.Close()
   })
 
$objForm.Controls.Add($OKButton)
 
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,520)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)
 
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please make a selection from the list below:"
$objForm.Controls.Add($objLabel) 
 
$objListbox = New-Object System.Windows.Forms.Listbox 
$objListbox.Location = New-Object System.Drawing.Size(10,40) 
$objListbox.Size = New-Object System.Drawing.Size(260,20) 
 
$objListbox.SelectionMode = $listboxtype
 
$inputarray | ForEach-Object {[void] $objListbox.Items.Add($_)}
 
$objListbox.Height = 470
$objForm.Controls.Add($objListbox) 
$objForm.Topmost = $True
 
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()
 
Return $Script:x
}

#Select the report type
$options01 = "General Report OnPrem","General Report EOL","PreMigration Report OnPrem","PreMigration Report EOL","PostMigration Report OnPrem","PostMigration Report EOL"
$reportselection = MultipleSelectionBox -listboxtype one -inputarray $options01

IF($reportselection -inotlike "*EOL")
{
	#Add the Exchange Module
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

	#For Exchange 2010
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;
}

#Select the deatils you wish in your report
$options02 = "Display Name","Alias","RecipientType","Recipient OU","Primary SMTP address","Email Addresses","Database","ServerName","TotalItemSize","ItemCount","DeletedItemCount","TotalDeletedItemSize","LastLogonTime","TimeStamp"
$selections = MultipleSelectionBox -listboxtype multisimple -inputarray $options02


#MailUser data import
#Check if On-Prem or Exchange Online
if(($reportselection) -notlike "*EOL")
{
	If((Select-UserBase) -eq "Yes")
	{
	write-host "Get All Mailboxes"
	$folder = Get-Folder
	$AllMailbox = Get-mailbox -resultsize unlimited|select *,@{n='SmtpAddress';e={ $_.EmailAddresses.SmtpAddress }}

}
	ELSE{
	write-host "Get Mailboxes From Input File"
	$MailUserFile = Get-FileName -Filter csv -Title "Select MailUser Import File"  -Obj
	$MailUsers = Import-Csv $MailUserFile
	$MailUsers = Import-Csv $MailUserFile
	$mailboxArray = foreach ($mailbox in $mailusers) {
		$curMailbox = Get-Mailbox $mailbox.EmailAddress
		#$stats = $curMailbox | Get-MailboxStatistics
        $curMailbox |
    		Select-Object DisplayName,
            					Alias,
                      DistinguishedName,
                      RecipientType,
                      OrganizationalUnit,
            					@{n='SmtpAddress';e={ $_.EmailAddresses.SmtpAddress }},
            					PrimarySmtpAddress,
                      Database,
                      ServerName,
                      UseDatabaseQuotaDefaults
	}
	#test
		$AllMailbox = $MailboxArray
}

	$i = 0
	$output=@()
	Foreach($Mbx in $AllMailbox)
	{
		$i++
		If($i -ne 0)
		{Write-Progress -Activity ("Scanning Mailboxes . . ."+$Mbx.displayname.tostring()) -Status "Scanned: $i of $($AllMailbox.tostring().Count)" -PercentComplete ($i/$AllMailbox.tostring().Count*100)}
		$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue
		$userObj = New-Object PSObject
		$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
		$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
		$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientType
		$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
		$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
		$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.SmtpAddress -join ";")
		$userObj | Add-Member NoteProperty -Name "Database" -Value $mbx.Database
		$userObj | Add-Member NoteProperty -Name "ServerName" -Value $mbx.ServerName
		if($Stats)
		{
			$totalsizearray = (($Stats.TotalItemSize.Value).tostring()).split("(").split(" ")
			$totalsize = [float]$totalsizearray[0]
			IF($totalsizearray[1] -eq "GB"){$totalsizeMB = $totalsize*1024}
			IF($totalsizearray[1] -eq "MB"){$totalsizeMB = $totalsize}
			IF($totalsizearray[1] -eq "KB"){$totalsizeMB = $totalsize/1024}
			$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $totalsizeMB
			$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
			$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
			$deletedsizearray = (($Stats.TotalDeletedItemSize.Value).tostring()).split("(").split(" ")
			$deletedsize = [float]$deletedsizearray[0]
			IF($deletedsizearray[1] -eq "GB"){$deletedsizeMB = $deletedsize*1024}
			IF($deletedsizearray[1] -eq "MB"){$deletedsizeMB = $deletedsize}
			IF($totalsizearray[1] -eq "KB"){$deletedsizeMB = $deletedsize/1024}
			$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $deletedsizeMB
		}
		$userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $ProhibitSendReceiveQuota$userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
		$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime
		$userObj | Add-Member NoteProperty -Name "TimeStamp" -Value (get-date -Format "yyyy-MMM-dd HH:mm:ss")
		$output += $UserObj  
		# Update Counters and Write Progress
	}
}

#If Exchange Online
if(($reportselection) -like "*EOL")
{
	If((Select-UserBase) -eq "Yes")
	{
	write-host "Get All Mailboxes"
	$folder = Get-Folder
	$AllMailbox = Get-mailbox -resultsize unlimited|select *,@{n='SmtpAddress';e={ $_.EmailAddresses }}

}
	ELSE{
	write-host "Get Mailboxes From Input File"
	$MailUserFile = Get-FileName -Filter csv -Title "Select MailUser Import File"  -Obj
	$MailUsers = Import-Csv $MailUserFile
	$MailUsers = Import-Csv $MailUserFile
	$mailboxArray = foreach ($mailbox in $mailusers) {
		$curMailbox = Get-Mailbox $mailbox.EmailAddress
		if($curMailbox -eq $null -or $curMailbox -eq ""){$curMailbox = Get-Mailbox $mailbox.PrimarySMTPAddress}
		#$stats = $curMailbox | Get-MailboxStatistics
        $curMailbox |
    		Select-Object DisplayName,
            					Alias,
                      DistinguishedName,
                      RecipientType,
                      OrganizationalUnit,
            					@{n='SmtpAddress';e={ $_.EmailAddresses }},
            					PrimarySmtpAddress,
                      Database,
                      ServerName,
                      UseDatabaseQuotaDefaults
	}
	$AllMailbox = $MailboxArray
}

	$i = 0
	$output=@()
	Foreach($Mbx in $AllMailbox)
	{
		$i++
		If($i -ne 0)
		{Write-Progress -Activity ("Scanning Mailboxes . . ."+$Mbx.displayname.tostring()) -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i/$AllMailbox.Count*100)}
		$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue
		$userObj = New-Object PSObject
		$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
		$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
		$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientType
		$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
		$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
		$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.SmtpAddress -join ";")
		$userObj | Add-Member NoteProperty -Name "Database" -Value $mbx.Database
		$userObj | Add-Member NoteProperty -Name "ServerName" -Value $mbx.ServerName
		if($Stats)
		{
			$totalsizearray = (($Stats.TotalItemSize.Value).tostring()).split("(").split(" ")
			$totalsize = [float]$totalsizearray[0]
			IF($totalsizearray[1] -eq "GB"){$totalsizeMB = $totalsize*1024}
			IF($totalsizearray[1] -eq "MB"){$totalsizeMB = $totalsize}
			IF($totalsizearray[1] -eq "KB"){$totalsizeMB = $totalsize/1024}
			$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $totalsizeMB
			$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
			$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
			$deletedsizearray = (($Stats.TotalDeletedItemSize.Value).tostring()).split("(").split(" ")
			$deletedsize = [float]$deletedsizearray[0]
			IF($deletedsizearray[1] -eq "GB"){$deletedsizeMB = $deletedsize*1024}
			IF($deletedsizearray[1] -eq "MB"){$deletedsizeMB = $deletedsize}
			IF($totalsizearray[1] -eq "KB"){$deletedsizeMB = $deletedsize/1024}
			$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $deletedsizeMB
		}
		$userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $ProhibitSendReceiveQuota$userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
		$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime
		$userObj | Add-Member NoteProperty -Name "TimeStamp" -Value (get-date -Format "yyyy-MMM-dd HH:mm:ss")
		$output += $UserObj  
		# Update Counters and Write Progress
	}
}
$rpttype = @{'General Report OnPrem' = "OnPrem_RPT";'General Report EOL' = "EOL_RPT";'PreMigration Report OnPrem' = "OnPrem_PreMig";'PreMigration Report EOL' = "EOL_PreMig";
	'PostMigration Report OnPrem' = "OnPrem_PostMig";'PostMigration Report EOL' = "EOL_PostMig"}


$output = $output|select $selections
$date = get-date -Format "HHmm-yyyy-MMM-dd"
$type = $rpttype[$reportselection]
IF($MailUserFile -ne $null){$output | Export-csv (($MailUserFile.PSParentPath+"\$type")+"_"+("$date.csv")) -NoTypeInformation}
ELSE{$output | Export-csv (($folder+"\$type")+"_"+("$date.csv")) -NoTypeInformation}