$Content = @"
<HTML>
<BODY>
<h1> This is a heading </h1>
<P id='para1'>First Paragraph with some Random text</P>
<P>Second paragraph with more random text</P>
<A href="http://Geekeefy.wordpress.com">Cool Powershell blog</A>
</BODY>
</HTML>
"@

$c = @"
<html><body><div>
<div>					
<div>					
<div>					
<div>					
<div>					
<div>					
<div>					
<div>					
<div>					
<div>					
<div>					
<div>					
<div>RTP ongoing, ops bleeding down to safe operating psi for pump to be restored. open 5/2/18 @1200<br>					
					
<table id="jSanity_hideInPlanner">					
<tr>					
<td style="padding:40px 0 0 0;">					
<table width="224" border="0" cellspacing="0" cellpadding="0" style="width:224px;">					
<tr>					
<td style="background-color:#31752F;padding:10px 30px 12px 30px;"><span style="background-color:#31752F;">					
<div align="center" style="text-align:center;"><a href="https://tasks.office.com/murphyoil.onmicrosoft.com/en-US/Home/Task/HjiXIV_KskOWl5LBfR2Wn2QAK7yR?Type=Comment&amp;Channel=GroupMailbox&amp;CreatedTime=636614888895859546" target="_blank" style="text-decoration:none;"><font face="Segoe UI,sans-serif,serif,EmojiFont" size="2" color="white"><span style="font-size:14px;">Reply					
in Microsoft Planner</span></font></a></div>					
</span></td>					
</tr>					
</table>					
</td>					
</tr>					
<tr>					
<td style="padding:15px 0 0 0;">					
<div><font face="Segoe UI,sans-serif,serif,EmojiFont" size="1" color="#666666"><span style="font-size:10px;font-weight:normal;">You can also reply to this email to add a task comment.</span></font></div>					
</td>					
</tr>					
<tr>					
<td style="padding:10px 0 0 0;">					
<div><font face="Segoe UI,sans-serif,serif,EmojiFont" size="1" color="#666666"><span style="font-size:10px;font-weight:normal;">This task is in the </span></font><a href="https://tasks.office.com/murphyoil.onmicrosoft.com/en-US/Home/PlanViews/2D6a9gJHO0WIhXlSUVvrLmQAAKw9?Type=Comment&amp;Channel=GroupMailbox&amp;CreatedTime=636614629432234653"><font face="Segoe UI,sans-serif,serif,EmojiFont" size="1"><span style="font-size:10px;font-weight:normal;">East					
EFS POT</span></font></a><font face="Segoe UI,sans-serif,serif,EmojiFont" size="1" color="#666666"><span style="font-size:10px;font-weight:normal;"> plan.</span></font></div>					
</td>					
</tr>					
</table>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</div>					
</body></html>					
"@


$c | Select-String -Pattern '<div\w+\s+<table' -AllMatches | % {$_.Matches}

# Create HTML file Object
$htmlParser = New-Object -Com "HTMLFile"

# Write HTML content according to DOM Level2 
$htmlParser.IHTMLDocument2_write($c)
 
$count = $htmlParser.all.tags("div").length

$comment = $htmlParser.getElementById("jSanity_hideInPlanner").parentNode | % innerText | % {$_ -replace ","," "} | % { $_ -replace "`n", " "} |  % { $_ -replace "`r", " "}

$comment

for($i=0; $i -lt $count; $i++) {
  
    #Write-Host $HTML.all.tags("div")[$i]
}
