
function split($myvar) {
	$myvar = $myvar.split("`n")
	return $myvar
	}

$text = get-content "C:\acl.txt"
$tempcsv = 'C:\acl-rules.csv'
$outfile = 'C:\acl-rules.xlsx'  



$groups = split $text | ? {$_ -match "access-group"}
$array = @()
$currentgroup = 0
$groupslength = $groups.length-1

$groups | % {
	$currentlist = 0
	
	
	$9 = $_ -replace "access-group " -replace " .+$" #  $9 acl-name
	$1 = $_ -replace "^.+?interface " # $1 inerface
	$2 = $_ -replace "^.+?$9 " -replace " .+$" # $2 direction
	$9 = $9 + " "
	$list = split $text | ? {$_ -match $9}
	$list = ($list | ? {($_ -notmatch "remark") -AND ($_ -match "hitcnt") -AND ($_ -notmatch "access-group")}) -replace "^  "
	$listlength = $list.length
	
	$list | % { 
		$table = New-Object -TypeName PSObject
		cls
		write "Processing $currentgroup of $groupslength ACL"
		write "Processing $currentlist of $listlength ACE"
		
		if ($_ -match "object") {
			$_ -match "(line\s\d+)"
			$line = $null
			$line = $matches[0]
			$origin = $_
			return
			}
		else {
			$_ = ($_ | ? {($_ -notmatch "object") -AND ($_ -notmatch "remark") -AND ($_ -match "hitcnt")}) -replace "^  "
		
			$hitcnt = $null
			$hitcnt = $_ -replace "^.+?hitcnt=" -replace "\).+$"
			$_ = $_ -replace "^\s+" -replace " log.+$" -replace "^.+?line ","line " -replace " \(hit.+$"
			$_ = $_ -replace "\s\(.+?\)"
			$_ -match "(line\s\d+)\sextended\s(permit|deny)\s(\w+)\s(any|any4|range\s\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\s\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}|host\s\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}|\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\s\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})?\s(eq\s.+?\s|range\s.+?\s|gt\s.+?\s|lt\s.+?\s|neq\s.+?\s)?(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\s\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}|any|any4|range\s\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\s\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}|host\s\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s?(eq\s.+?$|range\s.+?$|gt\s.+?$|lt\s.+?$|neq\s.+?$)?" > $null
			
			
			
			$3 = $matches[1] # $3 line
			$4 = $matches[2] # $4 action
			$7 = $matches[3] # $7 protocol
			$5 = $matches[4] # $5 source
			$10 = $matches[5] # $5 source port
			$6 = $matches[6] # $6 destination
			$8 = $matches[7] # $8 destination port
			
			
			
			if (($3 -eq $line) -AND ($origin -match $9)) {
			$11 = $origin} else {$11 = $null}
			
			
			
			$table | Add-Member –MemberType NoteProperty –Name Interface –Value $1
			$table | Add-Member –MemberType NoteProperty –Name Direction –Value $2 
			$table | Add-Member –MemberType NoteProperty –Name Line –Value $3
			$table | Add-Member –MemberType NoteProperty –Name Action –Value $4
			$table | Add-Member –MemberType NoteProperty –Name Protocol –Value $7
			$table | Add-Member –MemberType NoteProperty –Name Source –Value $5
			$table | Add-Member –MemberType NoteProperty –Name Src_Port –Value $10
			$table | Add-Member –MemberType NoteProperty –Name Destination –Value $6
			$table | Add-Member –MemberType NoteProperty –Name Port –Value $8
			$table | Add-Member –MemberType NoteProperty –Name ACL-Name –Value $9
			$table | Add-Member –MemberType NoteProperty –Name HitCount –Value $hitcnt
			$table | Add-Member –MemberType NoteProperty –Name Origin –Value $11
			
			$array += $table
			}
		$currentlist = $currentlist + 1
		}
	$currentgroup = $currentgroup + 1
	}
 
 
 
$array | Export-Csv $tempcsv -notype -Delimiter ";"

$xl = new-object -comobject excel.application
$xl.visible = $true
$wb = $xl.workbooks.open($tempcsv)
$table=$wb.ActiveSheet.ListObjects.add( 1,$wb.ActiveSheet.UsedRange,0,1)
$table.TableStyle = "TableStyleLight2"
$table = $wb.ActiveSheet.UsedRange.EntireColumn.AutoFit()
$xl.DisplayAlerts=$False
$wb.SaveAs($outfile,51)
$xl.Quit()
 
 
 
 
