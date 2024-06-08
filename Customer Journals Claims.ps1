<#      Extract current account transactions from spreadsheet to be processed as Pastel payment batch.
#>
$inspreadsheet = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\Claims\Claims.xlsx'
$outfile2 = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\Claims\CHQ_Transactions_1.csv'
$custsheet = 'Claims_2024'                                #Transactions worksheet
#################
$startR = 1                                         #Start row
$endR = 16
$filter = "P"                                       #P=Payments, R=Receipts\deposits
#################
$csvfile = 'SHEET1.csv'
$pathout = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\Claims\'
$startCol = 1                                                                   #Start Col (don't change)
$endCol = 11                                                                     #End Col (don't change)
$outfile1 = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\Claims\CHQTEMP.txt'              #Temp file
$outfileF = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\Claims\Claim_Transactions_pastel_' + $filter + '.txt'  #File to be imported into Pastel             
$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript { $_.P1 -eq $filter -and $_.P11 -eq 'Yes'} | Export-Csv -Path $Outfile -NoTypeInformation

ExcelFormatDate -file $Outfile -sheet 'SHEET1' -column 'E:E'

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile

#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }    

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $outfile2 -header type, GL, Expacc, AccCode, date, ref, description, amt, amt1, genterate, paid     

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods2 -transactiondate $aObj.date
    #$pastelper = '13'

    #Switch ($aObj.AccCode) {
        #NJEN { $aObj.AccCode; $aObj.descr }        #Customer         
        #NJEN { $aObj.AccCode; $aObj.descr }        #Customer         
        
       #Default { $expacc = '9999000'; $aObj.descr }       #Dummy account to generate error on import to Pastel
    #}
    Switch ($aObj.Expacc) {
        CALLOUT { $ExpenseAcc = '3701000'}        #Expense account for claim        
        REFRESHMENTS { $ExpenseAcc = '3700000'}        #Expense account for claim        
        EVENT { $ExpenseAcc = '2100000'}        #Expense account for claim        
        
        Default { $ExpenseAcc = '9999000'; $aObj.descr }       #Dummy account to generate error on import to Pastel
    }

    #Switch ($aObj.vat) {
    #    Y { $VATind = '15' }
   #     N { $VATind = '0' }
    #    Default {$VATind = '15'}
   # }
    #Format Pastel batch   
    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = $aObj.GL                      #GDC - general ledger(G), debtor(D), creditor(C)
        debtor  = $aObj.AccCode                       #debtor account to be credit (CR) - claims
        ref     = $aObj.ref
        comment = $aObj.description
        amount  = $aObj.amt1
        fil1    = ''
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = $ExpenseAcc                     #Expense Account amount is being claimed against
        fil6    = '1'
        fil7    = '1'
        fil8    = '0'
        fil9    = '0'
        fil10   = '0'
        amt2    = $aObj.amt1
    }
      
        $objlist = New-Object -TypeName psobject -Property $props1
        $objlist | Select-Object * | Export-Csv -path $outfile1 -NoTypeInformation -Append
    }  
    #Remove header information so file can be imported into Pastel Accounting.
    Get-Content -Path $outfile1 | Select-Object -skip 1 | Set-Content -path $outfilef
    Remove-Item -Path $outfile1

    $checkfile = Test-Path $outfile2
    if ($checkfile) { Remove-Item $outfile2 }
