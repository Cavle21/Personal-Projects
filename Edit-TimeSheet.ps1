using module "Modules\ExcelManagement.psm1"

<#

.PARAMETER $mode
    Get-Hours, Set-Hours

.PARAMETER $contractNumber
    The full number for our contract work Auths

.PARAMETER $hoursToAdd
    Hours to add to your store

.NOTES
HardCodedValues
    Get-Value row 5
        $Column = $Worksheet.Cells.Item(5, $Col).Value().trim()
    Get-Value row 1
        While ($Worksheet.Cells.Item(1,$Col).Value() -ne $null)
#>

<#param (
    [Parameter(Mandatory=$true,ParameterSetName='Mode')]
    [ValidateSet("Get-Hours", "Set-Hours")]
    [ValidateNotNullOrEmpty()]
        [string]$mode = "Get-Hours",
    #[Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern(“\d{5}-\d{3}-\d{3}-\d{3}”)]
        [string]$contractNumber = 11240-001-001-001,
    [Parameter(ParameterSetName='Set-Hours')]
        [uint32]$hoursToAdd,
    [Parameter(ParameterSetName='ValueToChange')]
        [string]$valueToChange,
    [Parameter(ParameterSetName='Value')]
        [string]$value,
    [Parameter()]
    [ValidateScript({
        try {
            Test-Path -Path $_ -ErrorAction stop
        } catch {
            Throw "${_}"
        }
    })]
        [string]$path = "$env:userprofile\documents\personaldocs\Work Authorizations.xlsx"
)#>

function Import-MyExcelFile {

    param(    
        [parameter()]
            [string]$path = "$env:userprofile\documents\personaldocs\Work Authorizations.xlsx",
        [parameter()]
            [string]$page = 'Sheet1',
        [Parameter()]
            [boolean]$visibility = $true,
        [Parameter()]
            [boolean]$displayAlerts = $true
    )

    $excelDocument = [ExcelDocument]::New($path, $visibility, $displayAlerts)

    return $excelDocument
}

function Get-Value {

    param (
        [System.Collections.ArrayList]$contractNumbers,
        [excelDocument]$excelDocument
    )


    $headers = $excelDocument.GetColumnHeaders(1,5)

    $results = $excelDocument.GetValuesInRow(1, 6)

    <# 
    $result = [PSCustomObject]@{}
    $numberOfResults = 1
    $foundArr = new-object System.Collections.ArrayList
    #find all rows that match the first set of numbers
    $found = $worksheet.Cells.Find($contractNumbers[0])
    $firstRow = $found.row
    $loop = $true
    #loop through all results $found and store them in an array
    while($loop -eq $true){
        $foundarr.Add($found)
        $found = $worksheet.Cells.FindNext($found)
        $currentRow = $found.row
        #Write-host "$currentrow and $firstRow"
        if ($currentRow -eq $firstRow){
            $loop = $false
        }else{
            $numberOfResults++
        }
    }

    Write-host "There are $numberOfResults rows with that value."

    #check rows found against second element in array ( hopefully this eliminates all others but one)

    forEach ($item in $foundarr){
        #Write-host $item.row
        if (($workSheet.cells.item($item.row,10).text) -eq $contractNumbers[1]){
            $found = $item
        }
    }

    $col = 1

    $ArrHeaders = new-object System.Collections.ArrayList
    #get all column headers
    Do  { 
        $Column = $Worksheet.Cells.Item(5, $Col).Value() #.trim()  
        #Write-host $Column
        $ArrHeaders += $Column -replace " ", "" 
        $Col++ 
    } While ($Worksheet.Cells.Item(5,$Col).Value() -ne $null)
    $col = 1
    #add data to result object using column headers as variable name
    forEach ($objHeader in $ArrHeaders){  
        If ($Worksheet.Cells.Item($found.Row, $col).Value() -ne $null){  
            $result | Add-Member -type NoteProperty -name $objHeader -value $worksheet.Cells.Item($found.Row, $Col).Value() #.trim()  
        }else{ 
            $result | Add-Member -type NoteProperty -name $objHeader -value $null  
        }

        $Col++ 
    }#>

    $excelDocument.CloseDocument()

    return $results
}

Function Set-Value {
    param(
        [object]$GetValueResultObj,
        [string]$valueToChange = "HoursRemaining",
        [string]$value,
        [System.Collections.ArrayList]$excelObjArr
    )

    if($GetValueResultObj.$valueToChange -eq $null){
        throw "$valueToChange parameter does not exist in the result object"
        exit
    }

    $GetValueResultObj.$valueToChange = $value

    

}

function Main {

    [string]$contractNumber = "11240-001-001-001"

    $contractNumbers = $contractNumber.Split("-")

    
    [string]$mode = "Get-Hours"
    [string]$path = "$env:userprofile\documents\personaldocs\Work Authorizations.xlsx"

    write-host "Initializing..."

    $excelObjArr = Import-MyExcelFile

    switch($mode){
        "Get-Hours"{
            $results = Get-Value -contractNumbers $contractNumbers -excelObjArr $excelObjArr
            $result =  "For $contractNumber you have {0} hours remaining" -f $results.HoursRemaining
            Write-host $result
        }
        "Set-Hours"{
            $results = Get-Value -contractNumbers $contractNumbers
        }
    }
}

Main

