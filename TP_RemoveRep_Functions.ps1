Function Remove-TPRep {
    <#
    .SYNOPSIS
    Used for removal of Rep from TigerPaw Database
    
    .DESCRIPTION
    Used in conjunction with a search function (Get-TPStaleReps) to multi-select Reps and loop
    Updates Tables where repNumber is necessary
    Removes entries from Tables where possible
    Finally removes Rep from tblReps
    
    .PARAMETER TPSQLserver
    FQDN of TigerPaw SQL Server Instance
    
    .PARAMETER TPSQLdb
    Database Name
    
    .PARAMETER oldrepid
    Existing Rep to Delete
    
    .PARAMETER newrepid
    Replacement Rep to Substitute (where necessary)
    
    .EXAMPLE
    PS> $selectedreps | ForEach-Object {Remove-TPRep -TPSQSserver 'svr-tpsql.domain.local' -TPSQLdb 'ARCHIVETigerPaw' -oldrepid $_.repNumber -newrepid 46}
    
    .NOTES
    Assumes your Windows account has permissions on the Database
    Written on the premise that a 'generic' user exists to catch all replacement references (repNumber 46 in the example)
    Written and tested in Powershell 5.1 Environment Against SQL2012 on Server 2012R2
    Support states that certain tables must have a repNumber present
    Updates and Deletes were translated from an SQL Statement given by Support
    One additional Update was added to theirs (tblViewLayoutRepShare)
    Support states always default to Update instead of Remove

    ***ALWAYS Test in Archive First - I won't be held responsible for anyone's bad outcomes.. use at your own risk

    ***Un-Comment the lines with Invoke-SQLCMD actions from this function if you wish to use it - the script does nothing out-of-the-box

    #>
    
    Param(
        [Parameter(Mandatory=$true)][string]$TPSQLserver,
        [Parameter(Mandatory=$true)][string]$TPSQLdb,
        [Parameter(Mandatory=$true)][int32]$oldrepid,
        [Parameter(Mandatory=$true)][int32]$newrepid
    )
    $Updates = [PSCustomObject]@(
        ([PSCustomObject]@{table='tblAccounts';column='RepNumber'})
        ([PSCustomObject]@{table='tblActivities';column='EmailFromRepNumber'})
        ([PSCustomObject]@{table='tblActivities';column='MergeSignedByRepNumber'})
        ([PSCustomObject]@{table='tblActivities';column='TaskRepNumber'})
        ([PSCustomObject]@{table='tblContracts';column='AssignedTech'})
        ([PSCustomObject]@{table='tblContracts';column='CreatedBy'})
        ([PSCustomObject]@{table='tblContracts';column='RepToCredit'})
        ([PSCustomObject]@{table='tblCreditMemos';column='AccountRepNumber'})
        ([PSCustomObject]@{table='tblCreditMemos';column='RepToDebit'})
        ([PSCustomObject]@{table='tblInvoiceChangeAudit';column='RepNumber'})
        ([PSCustomObject]@{table='tblInvoices';column='OnHoldRep'})
        ([PSCustomObject]@{table='tblInvoices';column='SalesRep'})
        ([PSCustomObject]@{table='tblItemMovement';column='RepNumber'})
        ([PSCustomObject]@{table='tblJournal';column='RepNumber'})
        ([PSCustomObject]@{table='tblOpportunities';column='OpportunityOwner'})
        ([PSCustomObject]@{table='tblOpportunities';column='LastModifiedByRep'})
        ([PSCustomObject]@{table='tblOpportunityStageChangeLog';column='StageChangedByRep'})
        ([PSCustomObject]@{table='tblProjectNotes';column='RepNumber'})
        ([PSCustomObject]@{table='tblProjectPhaseAssignments';column='RepNumber'})
        ([PSCustomObject]@{table='tblPurchaseOrders';column='BuyerNumber'})
        ([PSCustomObject]@{table='tblPurchaseOrders';column='RequestedByRep'})
        ([PSCustomObject]@{table='tblQuickSale';column='SalesRep'})
        ([PSCustomObject]@{table='tblQuoteNotes';column='RepNumber'})
        ([PSCustomObject]@{table='tblQuoteNotes';column='RepNumber'})
        ([PSCustomObject]@{table='tblQuotes';column='RepNumber'})
        ([PSCustomObject]@{table='tblRMAReceipts';column='ReceivedBy'})
        ([PSCustomObject]@{table='tblRMAReceipts';column='ReceivedBy'})
        ([PSCustomObject]@{table='tblScheduledActivities';column='CompletedBy'})
        ([PSCustomObject]@{table='tblScheduledActivities';column='ResponsibleEmployee'})
        ([PSCustomObject]@{table='tblScheduledActivities';column='ScheduledBy'})
        ([PSCustomObject]@{table='tblScheduledActivities';column='EmailOnBehalfOfRep'})
        ([PSCustomObject]@{table='tblScheduledActivities';column='MergeSignedByRep'})
        ([PSCustomObject]@{table='tblScheduledActivities';column='TaskRep'})
        ([PSCustomObject]@{table='tblServiceOrders';column='TakenBy'})
        ([PSCustomObject]@{table='tblServiceOrders';column='TechAssigned'})
        ([PSCustomObject]@{table='tblServiceOrders';column='VoidBy'})
        ([PSCustomObject]@{table='tblServiceOrders';column='RepToCredit'})
        ([PSCustomObject]@{table='tblSOLogsBilled';column='ByRep'})
        ([PSCustomObject]@{table='tblSONotes';column='RepNumber'})
        ([PSCustomObject]@{table='tblSysListViewPrint';column='RepNumber'})
        ([PSCustomObject]@{table='tblSysScreenPrint';column='RepNumber'})
        ([PSCustomObject]@{table='tblTasks';column='ScheduledByRepNumber'})
        ([PSCustomObject]@{table='tblTasks';column='ScheduledForRepNumber'})
        ([PSCustomObject]@{table='tblUserConnections';column='RepNumber'})
        ([PSCustomObject]@{table='tblViewLayoutRepShare';column='FKRepNumber'})
        ([PSCustomObject]@{table='tblWorkFlowEventLogRecipients';column='RepNumber'})
        ([PSCustomObject]@{table='tblWorkFlowEventLogs';column='OldAccountRep'})
        ([PSCustomObject]@{table='tblWorkFlowEventLogs';column='RepCausingEvent'})
        ([PSCustomObject]@{table='tblWorkFlowEventStaticRecipients';column='RepNumber'})
        ([PSCustomObject]@{table='tblWorkFlowQueue';column='RepCausingEvent'})
        ([PSCustomObject]@{table='tblWorkFlowQueue';column='OldAccountRep'})
        ([PSCustomObject]@{table='tblWorkFlowQueue';column='RepProcessingThisEntry'})
    )
    $Deletes = [PSCustomObject]@(
        ([PSCustomObject]@{table='tblAssignedRepGroups';column='RepNumber'})
        ([PSCustomObject]@{table='tblAudit';column='RepNumber'})
        ([PSCustomObject]@{table='tblBackTrack';column='RepNumber'})
        ([PSCustomObject]@{table='tblCalendarPrint';column='RepNumber'})
        ([PSCustomObject]@{table='tblDashBoardPanelLayouts';column='RepNumber'})
        ([PSCustomObject]@{table='tblDealerPortalGroupAssignments';column='RepNumber'})
        ([PSCustomObject]@{table='tblEmailTemplateCategoryFavorites';column='RepNumber'})
        ([PSCustomObject]@{table='tblExplorerRepFavorites';column='RepNumber'})
        ([PSCustomObject]@{table='tblGridSettings';column='RepNumber'})
        ([PSCustomObject]@{table='tblLogActivities';column='RepNumber'})
        ([PSCustomObject]@{table='tblLogEmailGroup';column='RepNumber'})
        ([PSCustomObject]@{table='tblLogExportEmail';column='RepNumber'})
        ([PSCustomObject]@{table='tblLogExportGroup';column='RepNumber'})
        ([PSCustomObject]@{table='tblLogInventoryCounts';column='RepNumber'})
        ([PSCustomObject]@{table='tblLogIClean';column='RepNumber'})
        ([PSCustomObject]@{table='tblMailingLabels';column='RepNumber'})
        ([PSCustomObject]@{table='tblPalmSync';column='RepNumber'})
        ([PSCustomObject]@{table='tblPriceBookLabels';column='RepNumber'})
        ([PSCustomObject]@{table='tblQuoteTemplates';column='CreatedByRep'})
        ([PSCustomObject]@{table='tblRepExpertise';column='RepNumber'})
        ([PSCustomObject]@{table='tblReportsRepFavorites';column='RepNumber'})
        ([PSCustomObject]@{table='tblRepOutlookSyncSettings';column='RepNumber'})
        ([PSCustomObject]@{table='tblRepSOPrintSettings';column='RepNumber'})
        ([PSCustomObject]@{table='tblQuoteLocks';column='RepNumber'})
        ([PSCustomObject]@{table='tblWorkFlowRepNotifications';column='RepNumber'})
        ([PSCustomObject]@{table='tblSecurityRoleRepAssignments';column='RepNumber'})
    )
<#  <----REMOVE THIS AND COMMENT END BELOW TO USE

    $Updates | ForEach-Object {
        $sqlstring= 'UPDATE ' + $_.table + ' SET ' + $_.column + ' = ' + $newrepid + ' WHERE (' + $_.column + ' = ' + $oldrepid + ')'

        Invoke-Sqlcmd `
            -ServerInstance $TPSQLserver `
            -Database $TPSQLdb `
            -Query $sqlstring
    }
        
    $Deletes | ForEach-Object {
        $sqlstring= "DELETE FROM " + $_.table + ' WHERE (' + $_.column + ' = ' + $oldrepid + ')'
               
        Invoke-Sqlcmd `
            -ServerInstance $TPSQLserver `
            -Database $TPSQLdb `
            -Query $sqlstring
    }

    Invoke-Sqlcmd `
    -ServerInstance $TPSQLserver `
    -Database $TPSQLdb `
    -Query "
        DELETE FROM tblReps
        WHERE (RepNumber = $oldrepid)
    "
#>  # <----REMOVE THIS AND COMMENT START ABOVE TO USE
}

Function Get-StaleTPReps {
    <#
    .SYNOPSIS
    Gets Selection list of stale TP Reps
    
    .DESCRIPTION
    Prompts for Cutoff Date - uses LastReadNewsDate (only semi-viable date entry in Reps table)
    Returns list of blank dates and all prior to cutoff in out-gridview
    Multi-Select returns an array of Rep objects
    
    .PARAMETER TPSQLserver
    FQDN of TigerPaw SQL Server Instance
    
    .PARAMETER TPSQLdb
    Database Name
    
    .EXAMPLE
    PS> $selectedreps= @(Get-StaleTPReps -TPSQSserver 'svr-tpsql.domain.local' -TPSQLdb 'ARCHIVETigerPaw')
    
    .NOTES
    Returns array of Rep Objects with 5 prorperties:
        -RepNumber
        -Inactive
        -RepName
        -Status
        -LastReadNewsDate

    ForEach to Remove-TPrep Function to remove Reps using the RepNumber Property (See Get-Help Remove-TPRep -Full)
    #>
    
    Param(
    [Parameter(Mandatory=$true)][string]$TPSQLserver,
    [Parameter(Mandatory=$true)][string]$TPSQLdb
    )

    $CutoffDate = (Read-Host "Cutoff Date?" | Get-Date)

    $reps= Invoke-Sqlcmd `
        -ServerInstance $TPSQLServer `
        -Database $TPSQLdb `
        -Query "`
            SELECT RepNumber,Inactive,RepName,Status,LastReadNewsDate `
            FROM tblReps"     

    $stalereps= @()

    #Status Must Be 'Inactive'
    #Either NewsDate is not blank and less than the cutoff
    #Or NewsDate is blank

    $reps | ForEach-Object {
        if (($_.Status -match 'inactive') -and `
            ((($_.LastReadNewsDate -notlike '') -and ($_.LastReadNewsDate -lt $cutoffdate)) `
            -or ($_.LastReadNewsDate -like ''))) `
            {$stalereps += $_}
    }

    Return $stalereps
}
