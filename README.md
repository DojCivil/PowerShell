# PowerShell
PowerShell Automation


<h1> Description </h1>
<p> This is a function that will assist with creating HTML/CSS style reports from powershell. It provides a more appealing report than traditional text output and will email the report as an attachment once generated. This specific report targets a specific AD group's members; however, this can be changed to fit your business needs. You can extend the functionality to create more complex reports and even change some of the HTML/CSS formatting. 
 </p>

<h2> Examples </h2>
<p> After the function has been loaded into memory </br>
Type: GenerateHTMLReport <AD Group Name> </br> </br>
If you want to run fully automate the report as a scheduled task, you will need to call the fucnction at very bottom of the script by typing GenerateHTMLReport. Example: 

    $sorted = $customerArray | sort department
    $htmSorted = $sorted | ConvertTo-Html -Fragment -Property Name, Department 
    
    #------------- Convert the report to HTML File format and email as attachement ------------------------

    ConvertTo-Html -Head $vCSS -Title $vTitle -Body "<h1> $vTitle </h1> `n <h5> This Report Was Generated on $date </h5>  `n $htmSorted " | Out-File $vReport
   #Send-MailMessage -SmtpServer $smtpServer -Subject $vSubject -From $vFrom -To "$toEmail" -Body $vBody -Attachments $reportPath  -BodyAsHtml -Priority High
   
    #--- Auto Launch Report ----------
    Invoke-Item $vReport
 
}

GenerateHTMLReport

</p> 

<h2> Meta Data </h2>
Author: Department of Justice Civil Division
<a href="https://github.com/DojCivil/PowerShell"> MIT License Information </a>





