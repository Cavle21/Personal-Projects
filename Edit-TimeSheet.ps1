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

param (
    [Parameter(Mandatory=$true,ParameterSetName='Mode')]
    [ValidateSet("Get-Hours", "Set-Hours")]
    [ValidateNotNullOrEmpty()]
        [string]$mode,
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern(“\d{6}-\d{3}-\d{3}-\d{3}”)]
        [string]$contractNumber,
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
)
function Import-MyExcelFile {

    param(    
        [parameter()]
            $path = "$env:userprofile\documents\personaldocs\Work Authorizations.xlsx",
        [parameter()]
            $page = 'Sheet1'
    )

    $excel = New-Object -ComObject Excel.Application
    try{
        $workBook = $excel.Workbooks.Open($path)
    }catch{
        Write-host "Unable to open excelsheet"
        Write-host $error[0].exception.Message
        exit
    }

    $workSheet = $workBook.worksheets | Where-Object {$_.Name -eq $page}

    $excelObjArr = @($excel, $workbook, $workSheet)

    return $excelObjArr
}

function Get-Value {

    param (
        [System.Collections.ArrayList]$contractNumbers,
        [System.Collections.ArrayList]$excelObjArr
    )

    $result = [PSCustomObject]@{}
    $numberOfResults = 1
    $firstRow = $found.row
    $foundArr = new-object System.Collections.ArrayList
    #find all rows that match the first set of numbers
    $found = $worksheet.Cells.Find($contractNumbers[0])
    $loop = $true
    #loop through all results $found and store them in an array
    while($loop -eq $true){
        $foundarr.Add($found)
        $found = $worksheet.Cells.FindNext($found)
        $currentRow = $found.row
        if ($currentRow -eq $firstRow){
            $loop = $false
        }else{
            $numberOfResults++
        }
    }

    #check rows found against second element in array ( hopefully this eliminates all others but one)

    forEach ($item in $foundarr){
        if ($workSheet.cells.item($item.row,10).text -eq $contractNumbers[1]){
            $found = $item
        }
    }

    $col = 1

    $ArrHeaders = new-object System.Collections.ArrayList
    #get all column headers
    Do  { 
        $Column = $Worksheet.Cells.Item(5, $Col).Value().trim()  
        $ArrHeaders += $Column -replace " ", "" 
        $intCol++ 
    } While ($Worksheet.Cells.Item(1,$Col).Value() -ne $null)
    
    $col = 1
    #add data to result object using column headers as variable name
    forEach ($objHeader in $ArrHeaders){  
        If ($objWorksheet.Cells.Item($found.Row, $col).Value() -ne $null){  
            $result | Add-Member -type NoteProperty -name $objHeader -value $worksheet.Cells.Item($found.Row, $Col).Value().trim()  
        }else{ 
            $result | Add-Member -type NoteProperty -name $objHeader -value $null  
        }

        $Col++ 
    }  
    #pscustomObj
    return $result
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

    $contractNumbers = $contractNumber.Split("-")

    $excelObjArr = Import-MyExcelFile

    switch($mode){
        "Get-Hours"{
            $results = Get-Value -contractNumbers $contractNumbers
            Write-Host "For $contractNumber you have {0} hours remaining" -f $results.HoursRemaining
        }
        "Set-Hours"{
            $results = Get-Value -contractNumbers $contractNumbers
        }
    }

    


}

