$year = Get-Date -Format "yyyy"
function backup {
    $compress = @{
        Path = "D:\_Files\BUDGET-BOOK"
        CompressionLevel = "Fastest"
        DestinationPath = "C:\Temp\BudgetBackup.zip"
    }

    Write-Host "Waiting"
    Compress-Archive @compress -Update
    Write-Host "`nBackup Complete" -BackgroundColor DarkCyan
    Write-Host "`n===================="
}

function moveExport {
    if (Test-Path $env:USERPROFILE\Downloads\Export.csv) {
        # Write-Host "FOUND!"
        Copy-Item C:\Users\jgigliotti\Downloads\Export.csv -Destination D:\_Files\BUDGET-BOOK\export.csv -Verbose -Force -Confirm
        Move-Item $env:USERPROFILE\Downloads\Export.csv -Destination D:\_Files\BUDGET-BOOK\.data\Export$year.csv -Verbose -Force -Confirm
        Write-Host "`nExport.csv has been moved" -BackgroundColor DarkCyan
        return $true
    } else {
        Write-Host "`nExport.csv File Missing" -BackgroundColor DarkRed
        return $false
    }
    Pause
}

function editTables {
    # $loadTables = $true
    # $response = $null
    if (!(moveExport)) {
        $response = Read-Host -Prompt 'Load Tables? [y/n]'
        if ($response -eq "y") {
            
            Write-Host "`nLoading Tables..." -BackgroundColor Yellow -ForegroundColor Black
            Write-Host "`n===================="
            Start-Process D:\_Files\BUDGET-BOOK\.tables\Transactions$year.xlsm
            Pause
        } elseif ($response -eq "n") {
            Write-Host "`nSkipping Table Loading..."
            Start-Sleep 3
            <# Action when this condition is true #>
        } else {
            break
        }
    } else {
        Write-Host "`nLoading Tables..." -BackgroundColor Yellow -ForegroundColor Black
        Write-Host "`n===================="
        Start-Process D:\_Files\BUDGET-BOOK\.tables\Transactions$year.xlsm
        Pause
    }

}

function startBudget {
    Write-Host "`nOpening Budget Files..." -BackgroundColor Green -ForegroundColor Black
    Start-Process 'D:\_Files\BUDGET-BOOK\Power Bi\Finance Statistics.pbix'
    Start-Process 'D:\_Files\BUDGET-BOOK\Power Pivot\Finance Power Pivot.xlsm'
}


function removeBackup {
	Remove-Item C:\Temp\BudgetBackup.zip -Confirm -Verbose -Force
}

# Backup Budget Folder
backup


# Begin Adjusting Tables
editTables

# Start Budget Files
startBudget

# Prompt to Remove Backup File
removeBackup

Write-Host "`n===== END ====="
