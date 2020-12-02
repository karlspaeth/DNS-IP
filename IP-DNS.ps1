### This script performs nslookup and ping on all DNS names or IP addresses you list in the text file referenced in $excel.



$InputFile = "F:\Scripts\list.txt"





#write-host "Check 1"


$excel = New-Object -comobject excel.application
$excel.visible = $True


$workbook = $excel.Workbooks.Add()
$workbook.Worksheets.Item(3).Delete()
$workbook.Worksheets.Item(2).Delete()


$wksht1 = $workbook.Worksheets.Item(1)
$wksht1.Name = "IP Results"

$wksht1.Cells.Item(1,1) = "DNS"
$wksht1.Cells.Item(1,2) = "IP"
$wksht1.Cells.Item(1,3) = "Ping Status"


$entry = 2


#write-host "Check 2"



$addresses = get-content $InputFile
$reader = New-Object IO.StreamReader $InputFile

while($reader.ReadLine() -ne $null){ $TotalIPs++ }

write-host    ""    
write-Host "Performing nslookup on each host..."    
foreach($address in $addresses) {
	
	## Progress bar
	$wksht1.Cells.Item($entry,1) = $address
	$i++
	$percentdone = (($i / $TotalIPs) * 100)
	$percentdonerounded = "{0:N0}" -f $percentdone
	Write-Progress -Activity "Performing nslookups" -CurrentOperation "Looking up: $address (IP $i of $TotalIPs)" -Status "$percentdonerounded% complete" -PercentComplete $percentdone
	## End progress bar
	
	try
	{
		$newIP = [System.Net.Dns]::GetHostAddresses($address) | select IPAddressToString
		$IP = $newIP[0].IPAddressToString
		$wksht1.Cells.Item($entry,2) = $IP
	}
	catch
	{
		$IP = "Not Found"
		$wksht1.Cells.Item($entry,2) = $IP
	}
	$entry++

}
	
$entry = 2

#write-host    ""            
#write-Host "Pinging each address..."

foreach($address in $addresses) {
	
	## Progress bar
	$j++
	$percentdone2 = (($j / $TotalIPs) * 100)
	$percentdonerounded2 = "{0:N0}" -f $percentdone2
	Write-Progress -Activity "Performing pings" -CurrentOperation "Pinging: $address (IP $j of $TotalIPs)" -Status "$percentdonerounded2% complete" -PercentComplete $percentdone2
	## End progress bar
	
	if (test-Connection -ComputerName $address -Count 1 -Quiet ) 
	{  
		#write-Host "$address responded" -ForegroundColor Green 
		$wksht1.Cells.Item($entry,3) = "Alive"
	} 
	else 
	{ 
		#Write-Warning "$address does not respond to pings"
		$wksht1.Cells.Item($entry,3) = "Dead"              
	}  
	$entry++

}


write-host    ""        
write-host "Done!"
