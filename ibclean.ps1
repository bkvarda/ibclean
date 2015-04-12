
#get the address from the web - doesn't work for mil buildings but don't care
function Get-BuildingAddress{
param($building)

$building = $building.toString()

$baseurl = "http://campusbuilding.com/b/microsoft-"

$words = $building -split "\s+"


$url = $baseurl+$words[0]+'-'+$words[1]

$result = Invoke-WebRequest $url



$address =  ($result.links| Where target -eq '_new' | select innerText).innerText

if($address){
  return $address
  }
 else{
 Write-Host -ForegroundColor Red "Building doesn't exist or unexpected format"

 }

}

#read list of labs
function Get-AddressList([string]$file){
  $excelInstance = New-Object -ComObject Excel.Application
  $excelInstance.Visible = $false
  $workbook = $excelInstance.Workbooks.Open($file)
  $worksheet = $workbook.sheets.item("Untitled")

  return $worksheet

  }


}

#read install base

#compare list of labs to potential install base matches

#write list of potential matches to file
