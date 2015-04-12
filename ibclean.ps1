


#get the address from the web - doesn't work for mil or advanta buildings but don't care right now
function Get-BuildingAddress{
param($building)

$building = $building.toString()
$building = $building.toLower()

if($building -eq "studio x")
{
  $building = "building 110"
}

$baseurl = "http://campusbuilding.com/b/microsoft-"

$words = $building -split "\s+"


$url = $baseurl+$words[0]+'-'+$words[1]

$result = Invoke-WebRequest $url



$address =  ($result.links| Where target -eq '_new' | select innerText).innerText

if($address){
  return $address
  }
 else{
  return "error"
 Write-Host -ForegroundColor Red "Building doesn't exist or unexpected format"

 }

}



#take PS Object list of labs, query web for address, append address to the list
function Append-Address([string]$file){
 $list = Import-Csv $file
 $list | ForEach-Object{
 
 $building = $_.Building.toString()
 $_.Address = Get-BuildingAddress $_.Building.toString()
 }

 return $list


}

function Get-CurrentIB([string]$file){

return Import-Csv $file

} 

function Get-AddressCorellations([string]$buildingfile,[string]$ibfile){
  $hash = @()
  $ib = Get-CurrentIB $ibfile
  $buildings = Append-Address $buildingfile
  $lastadded = " "
  $score = 0
   ForEach ($frame in $ib){
          $frameaddress = $frame.Address1.toLower()
     ForEach($building in $buildings){
          $fulladdress = $building.Address
          $splitaddress = $fulladdress -split ","
          $bldgaddress = $splitaddress[0].toLower()
          $roomnum = $building.Room
          $bldgname = $building.Building
          $bldgsplit = $bldgname -split "\s+"
          $num = $bldgsplit[1]
          $bldgshort = "bldg $num"

         
          
        if($frame.ITEM_SERIAL_NUMBER -ne $lastadded){ 
         if($frameaddress -eq $bldgaddress){
          $score = $score + 3

          Write-Host "$frameaddress same as $bldgaddress, score added"
         }

         
         if($frame.CS_CUSTOMER_NAME -like "*$bldgname*" -or $frame.CS_CUSTOMER_NAME -like "*$roomnum*"){
         $score = $score + 2
         $name = $frame.CS_CUSTOMER_NAME

         Write-Host "$name contains building name or room #, score added"
         }

         if($frame.ADDRESS2 -like "*$bldgname*" -or $frame.ADDRESS2 -like "*$roomnum*"){
          $score = $score + 2
          $name = $frame.ADDRESS2
          Write-Host "$name contains building name or room #, score added"
          }

          if($frame.CS_CUSTOMER_NAME -like "*$bldgshort*" -or $frame.ADDRESS2 -like "*bldgshort*" -or $frame.ADDRESS1 -like "*bldgshort*"){
           $score = $score +2
           Write-Host "$bldgshort alias found in an address field or site name, score added"
           }

            if($score -gt 0){

            $obj = New-Object System.Object
            $obj |Add-Member -type NoteProperty -Name 'Site_Name' -Value $frame.CS_CUSTOMER_NAME
            $obj |Add-Member -type NoteProperty -Name  'Site_ID' -value $frame.PARTY_NUMBER
            $obj |Add-Member -type NoteProperty -Name 'Serial_Number' -value $frame.ITEM_SERIAL_NUMBER
            $obj |Add-Member -type NoteProperty -Name 'Family' -value $frame.PRODUCT_FAMILY
            $obj |Add-Member -type NoteProperty -Name 'Model'-value $frame.MODEL
            $obj |Add-Member -type NoteProperty -Name 'IB_Adress' -value $frame.ADDRESS1
            $obj |Add-Member -type NoteProperty -Name 'IB_Address2' -value $frame.ADDRESS2
            $obj |Add-Member -type NoteProperty -Name 'List_Address' -value $building.Address
            $obj |Add-Member -type NoteProperty -Name 'MS_Bldg_Alias' -value $building.Building
            $obj |Add-Member -type NoteProperty -Name 'Score' -value $score
            
            $lastadded = $frame.ITEM_SERIAL_NUMBER
            $score = 0 
            $hash += $obj
         }
        }
       }
    }
    $hash |Export-Csv -Path 'C:\test\output.csv'
 }
