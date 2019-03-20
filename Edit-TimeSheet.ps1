
param (
    [Parameter(Mandatory=$true,ParameterSetName='Mode')]
    [ValidateSet("Get-Hours", "Set-Hours")]
    [ValidateNotNullOrEmpty()]
        [string]$mode,
    [Parameter(Mandator=$true)]
    [ValidateNotNullOrEmpty()]
        [string]$contractNumber, ######-###-###-###
    [Parameter(ParameterSetName='Set-Hours')]
        [uint32]$hoursToAdd
)

$path = "$env:userprofile\documents\personaldocs\Work Authorizations.xlsx"

$contractNumbers = $contractNumber.Split("-")

function Import-ExcelFile {

    param(    
        [parameter()]
            $file = "$env:userprofile\documents\personaldocs\Work Authorizations.xlsx",
        [prameter()]
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

    return $worksheet

}

function Get-Value {
    param (
        $contractNumbers
    )

    $result = [PSCustomObject]@{}
    $numberOfResults = 1
    $firstRow = $found.row
    $foundArr = new-object System.Collections.ArrayList

    $found = $worksheet.Cells.Find($contractNumbers[0])

    $loop = $true
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

    forEach ($item in $foundarr){
        if ($workSheet.cells.item($item.row,10).text -eq $contractNumbers[1]){
            $found = $item
        }
    }

    $col = 1

    $ArrHeaders = new-object System.Collections.ArrayList

    Do  { 
        $Column = $Worksheet.Cells.Item(5, $Col).Value().trim()  
        $ArrHeaders += $Column -replace " ", "" 
        $intCol++ 
    }While ($Worksheet.Cells.Item(1,$Col).Value() -ne $null)
    
    $col = 1

    foreach ($objHeader in $ArrHeaders){  
        IF ($objWorksheet.Cells.Item($found.Row, $col).Value() -ne $null)  
        {  
                $result | Add-Member -type NoteProperty -name $objHeader -value $worksheet.Cells.Item($found.Row, $Col).Value().trim()  
        }else{ 
                $result | Add-Member -type NoteProperty -name $objHeader -value $null  
        } 
        $Col++ 
    }  

    return $result
}

