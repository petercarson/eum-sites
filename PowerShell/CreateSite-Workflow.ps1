workflow CreateSite-Workflow {
    Param
    (
        [Parameter (Mandatory = $true)][int]$listItemID
    )

    $AutomationAccountName = Get-AutomationVariable -Name 'AutomationAccountName'
    $ResourceGroupName = Get-AutomationVariable -Name 'ResourceGroupName'
    
    Set-RunbookLock -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Lock $true

    if (ProvisionSite -listItemID $listItemID) {
        # Apply and implementation specific customizations
        CreateSite-Customizations -listItemID $listItemID
    }

    Set-RunbookLock -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Lock $false
}