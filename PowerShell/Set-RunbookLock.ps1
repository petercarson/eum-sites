workflow Set-RunbookLock {
    Param ( 
        [Parameter(Mandatory = $true)]
        [String] $AutomationAccountName,
        
        [Parameter(Mandatory = $false)]
        [String] $ServicePrincipalConnectionName = 'AzureRunAsConnection',
        
        [Parameter(Mandatory = $true)]
        [String] $ResourceGroupName,
        
        [Parameter(Mandatory = $true)]
        [Boolean] $Lock
    )

    $AutomationJobID = $PSPrivateMetadata.JobId.Guid
    Write-Verbose "Set-RunbookLock Job ID: $AutomationJobID"

    $ServicePrincipalConnection = Get-AutomationConnection -Name $ServicePrincipalConnectionName   
    if (!$ServicePrincipalConnection) {
        $ErrorString = 
        @"
        Service principal connection $ServicePrincipalConnectionName not found.  Make sure you have created it in Assets. 
        See http://aka.ms/runasaccount to learn more about creating Run As accounts. 
"@
        throw $ErrorString
    }  	
    
    Add-AzureRmAccount `
        -ServicePrincipal `
        -TenantId $ServicePrincipalConnection.TenantId `
        -ApplicationId $ServicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $ServicePrincipalConnection.CertificateThumbprint | Write-Verbose

    # Get the information for this job so we can retrieve the Runbook Id
    $CurrentJob = Get-AzureRmAutomationJob -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Id $AutomationJobID
    Write-Verbose "Set-RunbookLock AutomationAccountName: $($CurrentJob.AutomationAccountName)"
    Write-Verbose "Set-RunbookLock RunbookName: $($CurrentJob.RunbookName)"
    Write-Verbose "Set-RunbookLock ResourceGroupName: $($CurrentJob.ResourceGroupName)"
    
    if ($Lock) {
        $AllJobs = Get-AzureRmAutomationJob -AutomationAccountName $CurrentJob.AutomationAccountName `
            -ResourceGroupName $CurrentJob.ResourceGroupName `
            -RunbookName $CurrentJob.RunbookName | Sort-Object -Property CreationTime, JobId | Select-Object -Last 10

        foreach ($job in $AllJobs) {
            Write-Verbose "JobID: $($job.JobId), CreationTime: $($job.CreationTime), Status: $($job.Status)"
        }

        $AllActiveJobs = Get-AzureRmAutomationJob -AutomationAccountName $CurrentJob.AutomationAccountName `
            -ResourceGroupName $CurrentJob.ResourceGroupName `
            -RunbookName $CurrentJob.RunbookName | Where -FilterScript { ($_.Status -ne "Completed") `
                -and ($_.Status -ne "Failed") `
                -and ($_.Status -ne "Stopped") } 

        Write-Verbose "AllActiveJobs.Count $($AllActiveJobs.Count)"

        # If there are any active jobs for this runbook, suspend this job. If this is the only job
        # running then just continue
        If ($AllActiveJobs.Count -gt 1) {
            # In order to prevent a race condition (although still possible if two jobs were created at the 
            # exact same time), let this job continue if it is the oldest created running job
            $OldestJob = $AllActiveJobs | Sort-Object -Property CreationTime, JobId | Select-Object -First 1
            Write-Verbose "AutomationJobID: $($AutomationJobID), OldestJob.JobId: $($OldestJob.JobId)"

            # If this job is not the oldest created job we will suspend it and let the oldest one go through.
            # When the oldest job completes it will call Set-RunbookLock to make sure the next-oldest job for this runbook is resumed.
            if ($AutomationJobID -ne $OldestJob.JobId) {
                Write-Verbose "Suspending runbook job as there are currently running jobs for this runbook already"
                Suspend-Workflow
                Write-Verbose "Job is resumed"
            }   
        }
        Else {
            Write-Verbose "No other currently running jobs for this runbook"
        }
    }   
    Else {
        # Get the next oldest suspended job if there is one for this Runbook Id
        $OldestSuspendedJob = Get-AzureRmAutomationJob -AutomationAccountName $CurrentJob.AutomationAccountName `
            -ResourceGroupName $CurrentJob.ResourceGroupName `
            -RunbookName $CurrentJob.RunbookName | Where -FilterScript { $_.Status -eq "Suspended" } | Sort-Object -Property CreationTime | Select-Object -First 1   
           
        if ($OldestSuspendedJob) {
            Write-Verbose ("Resuming the next suspended job: " + $OldestSuspendedJob.JobId)
            Resume-AzureRmAutomationJob -ResourceGroupName $CurrentJob.ResourceGroupName -AutomationAccountName $CurrentJob.AutomationAccountName -Id $OldestSuspendedJob.JobId | Write-Verbose 
        }
        Else {
            Write-Verbose "No suspended jobs for this runbook"
        }
    }
}