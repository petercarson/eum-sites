workflow Test-RunbookLock
{
    $AutomationAccountName = Get-AutomationVariable -Name 'AutomationAccountName'
    $ResourceGroupName = Get-AutomationVariable -Name 'ResourceGroupName'

    Write-Verbose "AutomationAccountName: $AutomationAccountName"
    Write-Verbose "ResourceGroupName: $ResourceGroupName"
    
    Set-RunbookLock -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Lock $true

    for ($i=0; $i -lt 10; $i++) {
        Write-Verbose "Sleeping $i"
        Start-Sleep -Seconds 6
        }

    Set-RunbookLock -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Lock $false
}
