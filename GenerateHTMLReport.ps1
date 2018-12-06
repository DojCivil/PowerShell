#*************************************************************************
# Title: GenerateHTMLReport-OS.ps1
# Author: Department of Justice
# Version: 2.0 
# Date: October 20, 2017
# Modified: December 4, 2018 
#************************************************************************
 
function GenerateHTMLReport {
 <#

  .SYNOPSIS
This is a function that will assist with creating HTML/CSS style reports from powershell. It provides a more appealing report than traditional
text output and will email the report once the generated. This report targets a specific AD group; however, this can be changed to fit your business needs.

.DESCRIPTION
The script will generate an HTML report of customers who belong to a specific AD group. When the script runs the end user is prompted for the name of the AD Group to query.

.EXAMPLE
If you have spaces in your AD Group do not put quotes around it: i.e. GenerateHTMLReport AD Users Group

.NOTES
Be sure to modifiy the variables that will work best with your environment. Modify the email variables such as the sender, recipeint, subject, body, SMTP Server

.LINK
https://www.github.com/DojCivil 


 #>
 
    try{
        $adGroup = Read-Host -Prompt "Please type the AD Group Name"
        $groups = Get-ADGroupMember $adGroup
    }
    Catch{
        Write-Host -ForegroundColor Red "Verify the spelling of your AD group is correct"
        Break
    }
    $users = $groups.samAccountName
    $total = $users.Count  
    $customerArray=@() 
    $date = Get-Date
    $me = $env:USERNAME
    $vReport = "./SampleReport.htm"
    $i = 1
    Write-Host -ForegroundColor Green "Script Processing"
   

    #---------------- Set variables to format automated email to include From, Subject, Body (Message), Title, and Recipient ------
    $vFrom = "Engineering@doj.gov"
    $vTitle = "HTML Report" 
    $vSubject = "Add Subject"
    $toEmail = "Add recipient email address here"
    $smptServer = "Add Server Name"
    $reportPath = "Path to HTML report"
    $vBody = "
        <p> 
            Hello, this is some sample text that can be added to the body of the email. </br></br>
        </p>
        </br>
            Thank you for your attention, </br>

            Engineering Staff 
    "
    #---------------------- Create the CSS that builds the HTML report -----------------------------------

    $vCSS = "
    <style>
        h1, h5, th { text-align: center; } 
        table { margin: auto; font-family: sans-serif; box-shadow: 5px 5px 5px #888; border: thin ridge grey; width: 25% } 
        th { background: #0046c3; color: #fff; max-width: 600px; padding: 5px 10px; } 
        td { font-size: 13px; padding: 5px 10px; color: #000; } 
        tr { background: #b8d1f3; } 
        tr:nth-child(even) { background: #dae5f4; } 
        tr:nth-child(odd) { background: #bff4ff; }
    </style>
    "

    #------------- Build array of users in the AD Group and list their name and department -------------------

    $i = 0
    foreach($line in $users){
       $customer = Get-ADUser $line -Properties * | select name, department
       $customerArray += $customer
       $customer | ForEach-Object{
            Write-Progress -Activity "Please Wait ... Processing" -Status "Customer $i of $($users.count)" -PercentComplete (($i/$total) * 100) 
            $i++ 
       }
    }

    $sorted = $customerArray | sort department
    $htmSorted = $sorted | ConvertTo-Html -Fragment -Property Name, Department 
    
    #------------- Convert the report to HTML File format and email as attachement ------------------------

    ConvertTo-Html -Head $vCSS -Title $vTitle -Body "<h1> $vTitle </h1> `n <h5> This Report Was Generated on $date </h5>  `n $htmSorted " | Out-File $vReport
  
  
   Send-MailMessage -SmtpServer $smtpServer -Subject $vSubject -From $vFrom -To "$toEmail" -Body $vBody -Attachments $reportPath -BodyAsHtml -Priority High
   
    #--- Auto Launch Report ----------
    Invoke-Item $vReport
 
}





   
