<#      Extract current account transactions from spreadsheet to be processed as Pastel payment batch.
#>
$inspreadsheet = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\FNB\Transactions\CHQ Transactions 83103.xlsx'
$outfile2 = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\FNB\Transactions\CHQ_Transactions_1.csv'
$custsheet = 'Transactions'                                #Transactions worksheet
#################
$startR = 1                                         #Start row
$endR = 73
$filter = "P"                                       #P=Payments, R=Receipts\deposits
#################
$csvfile = 'SHEET1.csv'
$pathout = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\FNB\Transactions\'
$startCol = 1                                                                   #Start Col (don't change)
$endCol = 12                                                                     #End Col (don't change)
$outfile1 = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\FNB\Transactions\CHQTEMP.txt'              #Temp file
$outfileF = 'C:\Users\Guy\OneDrive - PRIVATE\SARZA\FNB\Transactions\CHQTransactions_pastel_' + $filter + '.txt'  #File to be imported into Pastel             
$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript { $_.P1 -eq $filter -and $_.P12 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

ExcelFormatDate -file $Outfile -sheet 'SHEET1' -column 'D:D'

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile

#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }    
               
                   

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $outfile2 -header type, GL, Expacc, date, ref, date2, amt, bal, desc, amt1, vat     

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods2 -transactiondate $aObj.date
    #$pastelper = '13'

    Switch ($aObj.Expacc) {
        CTCT { $expacc = $aObj.Expacc; $aObj.descr }        #Customer         
        DBEAU { $expacc = $aObj.Expacc; $aObj.descr }        #Customer         
        BLOUW { $expacc = $aObj.Expacc; $aObj.descr }        #Customer         
        PBURG { $expacc = $aObj.Expacc; $aObj.descr }        #Customer         
        RDELL { $expacc = $aObj.Expacc; $aObj.descr }        #Customer         
        UTCT { $expacc = $aObj.Expacc; $aObj.descr }        #Customer         
        ZCUR { $expacc = $aObj.Expacc; $aObj.descr }        #Customer         
        ZARE { $expacc = $aObj.Expacc; $aObj.descr }        #Customer         
        BANK { $expacc = '3200000'; $aObj.descr }         
        DONATION { $expacc = '1000050'; $aObj.descr }         
        SNAP { $expacc = '1000050'; $aObj.descr }         
        #Advertising { $expacc = '3050000'; $aObj.descr }         
        KIT { $expacc = '2000000'; $aObj.descr }         
        CALLCOST { $expacc = '3701000'; $aObj.descr }          #Callout costs       
        EVENTCOST { $expacc = '2100000'; $aObj.descr }          #Event costs       
        AQSS { $expacc = $aObj.Expacc; $aObj.descr }            #Supplier
        SARZAN { $expacc = $aObj.Expacc; $aObj.descr }         #Supplier
        SUNDRY { $expacc = '2900000'; $aObj.descr }           #Sundry income - COCT\WSAR
        Training { $expacc = '2101000'; $aObj.descr }           #COS - Training
        INTP { $expacc = '3900000'; $aObj.descr }
        INTR { $expacc = '2750000'; $aObj.descr }
        STATIONERY { $expacc = '4200000'; $aObj.descr }
        
        Default { $expacc = $aObj.Expacc; $aObj.descr }       #Dummy account to generate error on import to Pastel
    }

    Switch ($aObj.vat) {
        Y { $VATind = '15' }
        N { $VATind = '0' }
        Default {$VATind = '15'}
    }
    #Format Pastel batch   
    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = $aObj.GL                      #GDC - general ledger(G), debtor(D), creditor(C)
        contra  = $expacc                       #Expense account to be debited (DR)
        ref     = $aObj.ref
        comment = $aObj.desc
        amount  = $aObj.amt1
        fil1    = $VATind
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '8400000'                     #Cheque account contra account number
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
