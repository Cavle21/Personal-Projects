#$path = "$env:userprofile\documents\personaldocs\Work Authorizations.xlsx"

<#Do  { 
          $Column = $Worksheet.Cells.Item($Row, $Col).Value().trim()  
          $ArrHeaders += $Column -replace " ", "" 
          $intCol++ 
}While ($Worksheet.Cells.Item(1,$Col).Value() -ne $null) #>


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

Function Get-Value {

    param(
        [parameter()]
            [System.Collections.ArrayList]$contractNumbers

    )

    $contractNums = $contractNum.split("-")

    $loop = $true
    $numberOfResults = 1
    
    $firstRow = $found.row

    $foundArr = new-object System.Collections.ArrayList


    $found = $worksheet.Cells.Find($contractNums[0])


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
        if ($workSheet.cells.item($item.row,10).text -eq "001"){
            $found = $item
        }
    }


}



 