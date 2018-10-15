<#
    NAME: GetIPs.ps1
    DESCRIPTION: Gets Ips for a list of servers in an excel sheet
    AUTHOR: Austin Vargason
    DATE MODIFIED: 10/15/18
#>


$data = Import-Csv -Path .\servers.csv

function Get-IPData () {

    param (
        [Parameter(Mandatory=$true)]
        [PsObject[]]$fileInput
    )

    #create a result array
    $resultArray = @()

    #start a counter
    $i = 0

    #save the total count to help with speed
    $count = $fileInput.Count

    #loop through each server and call the Get-IP function
    foreach($server in $fileInput| Select -ExpandProperty "System Name") {

        #save the output array
        $output = Get-IP -Server $server

        #add to the total array
        $resultArray += $output

        #increase the counter
        $i++

        #update the progress bar
        Write-Progress -Activity "Getting Ips" -Status "Received IP for Server: $server" -PercentComplete (($i / $count) * 100)
    }

    #return the result 
    return $resultArray
}

#function to get ip from host name
function Get-IP () {

    param (
        [Parameter(Mandatory=$true)]
        [String]$Server
    )

    #start a resiult array
    $res = @()

    #get the output from the .net get host addresses
    try {
        
        $output = [System.Net.Dns]::GetHostAddresses($Server)
        
        #create an output format
        foreach ($item in $output) {
            $item | Add-Member -Name ServerName -Value $Server -MemberType NoteProperty
            $res += $item
        }

        #return the result array
        return $res
    }
    catch {
        Write-Host "Could not get address for server: $Server" -ForegroundColor Yellow
    }
}

Get-IPData -fileInput $data | Export-Excel -Path .\IPData.xlsx -WorkSheetname "IP Data" -TableName "IPDataTable" -AutoSize -Show