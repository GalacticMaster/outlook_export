

Function Get-OutlookInBox 
{ 
    $DateStart = [DateTime]::Now.AddDays(-28)
    $DateEnd = [DateTime]::Now
    Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
    $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type] 
    $outlook = new-object -comobject outlook.application
    $namespace = $outlook.GetNameSpace("MAPI")
    $sFilter="[ReceivedTime] > '{0:dd/MM/yyyy}' AND [ReceivedTime] < '{1:dd/MM/yyyy}'" -f 
    $DateStart,$DateEnd
    $Inbox = $namespace.pickfolder()

    $Client1 = $Inbox.folders | Where-Object {$_.name -eq "SOC"}
    $Client1Subfolder = $Client1.Folders | Where-Object {$_.name -eq "SEC P1" -or $_.name -eq "SEC P2" -or $_.name -eq "SEC P3" -or $_.name -eq "SEC P4"}
    $Client1Subfolder.items | Select-Object -Property subject,body,ReceivedTime | Sort-Object ReceivedTime -Descending |Where-Object {$_.receivedTime -match "4/6/2018"}
    #$emails = $Client1Subfolder.items | Select-Object -Property subject,body | Sort-Object ReceivedTime -Descending 
    #return $emails
    #$emails = $Client1Subfolder.items | Where-Object {$_.RecivedTime -gt [datetime]$DateStart}
    $Esubject = $email.subject
    return $Client1Subfolder.items | Where-Object {$_.receivedTime -gt $DateStart} 
    #return $emails
    #write-host $email

} #end function Get-OutlookInbox

function find-string-in-email($email, $findString)
{
    #write-host $email
    $found = Select-String -Pattern $findString -InputObject $email
    if ($found)
    {
        return $found.Matches[0].Groups[1].Value
    }
    return "MISSING"
}


function pars-email($email)
{
    $findsecID = "\[SECURITY INCIDENT ID\]: (\w\w\w\d+)"
    $findPLevel = "\[SECURITY PRIORITY LEVEL\]: (\d+)"
    $findCateGory = "\[CATEGORY\]: ([\d\w \.\/]+)"
    $findCompromise = "\[STAGE OF COMPROMISE\]: (\w+\w+)"
    $findActivity = "(?smi)\[DETAILS]:(.*?)(Infected|Affected|Internal|Email|Source)s? (Host|Endpoint|Detail|User)s?:?"
    $finddDateTime = "\[DETECTION DATE & TIME]: (\d+\/\d+\/\d+ \d+\:\d+\:\d+)"
    $findNDateTime = "\[NOTIFICATION DATE & TIME]: (\d+\/\d+\/\d+ \d+\:\d+\:\d+)"
    $findusername = "(?smi)\[DETAILS]:(.*?)(File Details:|Files Details:|Further Details:)"
    $findfilepath = "(?smi)\[DETAILS]:(.*?)Compromise Details:"
    
    
    #$findSummary = "(?smi)\[SUMMARY](.*)\[DETAILS]\:"
    #$findFireeywEmail = "Phish(.*)"
    

    $secID = find-string-in-email $email $findsecID
    $pLevel = find-string-in-email $email $findPLevel
    $CateGory = find-string-in-email $email $findCateGory
    $compromise = find-string-in-email $email $findCompromise
    $DDateTime = find-string-in-email $email $finddDateTime
    $NDateTime = find-string-in-email $email $findNDateTime
    $FireEmail = find-string-in-email $email $findFireeywEmail
    $filepath = find-string-in-email $email $findfilepath
    

    #find user name   
    $usernameblock = find-string-in-email $email $findusername
    $name = [regex]::Matches($usernameblock,'(?<=\,|\().+?(?=\,|\[|\))')

    #find file path
    $file = [regex]::Matches($filepath,'(?smi)(File Path:|Path:).*?(\*)')
     
    #find activity date
    $ActivityBlock = find-string-in-email $email $findActivity
    
    $ActivityDate = ""

    $lines = $ActivityBlock.Split([Environment]::NewLine)
    foreach($line in $lines)
    {
        if ($line.StartsWith("*"))
        {
            $ActivityDate = $line.Trim("* ")
        }
    }

      
    return [SecTicket]::new($secID, $pLevel, $CateGory, $compromise, $ActivityDate, $DDateTime, $NDateTime, $FireEmail, $source, $name, $filepath, $file)


}

# create Excel file

$excel = New-Object -ComObject Excel.Application

$excel.Visible = $true

$workbook = $excel.Workbooks.Add()

$sheet = $workbook.ActiveSheet

$counter = 0


$emails = Get-OutlookInBox
foreach ($email in $emails)
{
    #$secticket = pars-email($email)
    #write-host $secticket.ticketnumber
    #Write-Host $emails
    $secticket = pars-email($email.Body)
    #if($secticket.ticketnumber -eq "MISSING")
    #{
        #write-host "ERROR not title"
       # write-host "email " $email.subject
           
    #}

   

    #Loop through the Array and add data into the excel file created.

      $counter++

      $sheet.cells.Item($counter,1) = $secticket.ticketnumber  

      $sheet.cells.Item($counter,2) = $secticket.pLevel

      $sheet.cells.Item($counter,3) = $secticket.category

      $sheet.cells.Item($counter,4) = $secticket.compromise

      $sheet.cells.Item($counter,5) = $secticket.Activitydate

      $sheet.cells.Item($counter,6) = $secticket.DDateTime

      $sheet.cells.Item($counter,7) = $secticket.NDateTime

      $sheet.cells.Item($counter,8) = $secticket.file

      $sheet.cells.Item($counter,9) = $secticket.name

}

#Write-Host $secticket.ticketnumber.trim(),","$email.subject.trim()","$secticket.pLevel.trim(),","$secticket.category.trim(),","$secticket.compromise.trim(),","$secticket.Activitydate,","$secticket.DDateTime,","$secticket.NDateTime,","$secticket.name.trim() 
