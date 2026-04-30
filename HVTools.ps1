$Author = "Hugo Remington"
$AuthorDate = "31/10/2021."
$Version = "1.7.3"
<#Description: This script was crafted with the intents and purposes of performing pre-validaton checks and
migration capacity planning by reporting maximum allocated storage capacity provisioned to all VMs in 
a Hyper-V SCVMM environment.

Changelog
===
Added VMM server name to title bar form of report.
Removed empty var filename.
1.7.3 fixed tab name to Cluster Storage Volumes.
Clipboard copy/paste will include headers.
1.7.2 Enabled datagrid clipboard CTRL + C of selected rows. Achieved this by configuring RunspacePool apartmentState as STA from default.
Removed 2nd form runspace.close calls as they are irrelevant.
1.7.1 Fixed exit bug where runspaces would keep the process open. App now exits cleanly.
Fixed bug where the Login button would remain disabled post inventory completion.
New feature; Zombie VHDs tab is now available, reporting orphaned VHD/VHDX files in SCVMM.
1.7.0 Appended new column in Virtual Machines table called Location. This contains VHD/VHDX path.
Removed automatic export of Virtual Machines table.
1.6.9 ZIP multi-report saving now final. Cleanup feature added.
Rounded off all numbers to single decimal point.
Added experimental save to ZIP feature from line 2509. Exports all tables into a temp path before ZIPing to desired save path via dialog.
1.6.8 New feature: Cluster storage volumes tab between at tab 7.
1.6.7 removed $script:powershell = [powershell]::Create() from line 2322.
1.6.6 cosmetic update, rounding off large capacity numbers to 2 decimal places using the System.Math class [math]::Round($var,2)
1.6.5 cosmetic update, added About/help menu in 2nd results form.
Added reset counter for clusters $c in ClusterNetworks foreach loop.
1.6.4 Added tab filter for all 8 datagrids.
1.6.3 Major update. Packed new features in including Clusters, Hosts, Storage Pools, Storage Arrays, Cluster Disks, Networks and Cluster Networks!
1.6.2 Added hosts feature incliding data table, tab, foreach loop and data grid.
Re-enabled maxthreads to all available processor count on system.
New feature, Cluster info! Added 2nd table for Clusters.
1.6.1 Fixed unprotected memory exceptions by remove Add-OutputBoxLine calls within ForEach loops. This is because there are lots of objects in SCVMM and it overloads the richtexbox output.
1.6.0 Adding more tables and tabs for comprehensive information cluding clusters, hosts, storage and networks.
1.5.9 Attempt to resolve unprotected memory leak by calling $script:powershell.EndInvoke($script:handle) at every exit function.
Changed $maxthreads to 3 in order to resolve unprotected memory exceptions.
Added extra tabs, preparing for future releases.
Re-enabled maxthreads.
Updated colour scheme.
#1.5.6 Attempt at fixing memory leak by converting $VMS to an array by casting as such prior.
Fixed VLAN not displaying by changing column type from int32 to String.
Reduced $maxthreads to static 3, as app crashes on unhandled protected memory exceptions in a multi-user RDP session SCVMM instance.
Further enhancements to GUI, streamlining tab view and changed datagridview colour to moccasin.
Fixed 2nd form colour.
Minor bug fixes including table column types as memory, dynamic memory and vCPU were int and not working.
Major GUI updates.
Converted array to data grid view.
Code-signed application using Sectigo certificate.
#1.5.0 Using multi-threading thanks to runspace pools.
Added Out-Grid GUI view.
Performance improvements.
Using array instead of flat memory.
Added static/dynamic memory optimisations.

#>
#Suppress Errors
$script:ErrorActionPreference = 'SilentlyContinue'
$script:ProgressPreference = 'SilentlyContinue'
#Set-ExecutionPolicy unrestricted
$getPowerShellVersion = $PSVersionTable.PSVersion

#Hash table for runspaces
$hash = [hashtable]::Synchronized(@{})

#Collect meta data for runspace
$hash.Author = $Author
$hash.AuthorDate = $AuthorDate
$hash.Version = $Version

#Desktop path
$DesktopPath = [Environment]::GetFolderPath("Desktop")

#Working Path
$script:workingPath = Get-Location

#Date
$script:datestring = (Get-Date).ToString("s").Replace(":","-")


#FUNCIONS

#Append output to text box display.
<#Function Add-OutputBoxLine 
{
    Param ($Message)
    $hash.outputBox.AppendText("$Message")
    $hash.outputBox.Refresh()
    $script:hash.OutputBox.ScrollToCaret()
    $script:hash.OutputBox.SelectionStart = $hash.outputBox.Text.Length
    $hash.outputBox.Selectioncolor = "WindowText"
    $hash.Form.Refresh()
}#>



#Function for testing runspaces and GUI updates
Function RunspaceTestAppCode{
    #$global:hash.outputBox.AppendText("This is outside of the runspace")

    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = "STA"
    $script:runspace.ThreadOptions = "ReuseThread"
    $script:powershell = [powershell]::Create()
    $script:powershell.Runspace = $script:runspace
    $script:runspace.Open()
    $script:runspace.SessionStateProxy.SetVariable("hash",$hash)
    
    
    $script:powershell.AddScript({

        #$global:hash.runspaceOutputbox = $hash.outputBox.Text
        $hash.outputBox.BeginInvoke([action]{$hash.outputbox.SelectionColor = "Green"})
        $hash.outputBox.BeginInvoke([action]{$hash.outputbox.AppendText("Value from runspace!!! Working son")})
    
    })
    $AsyncObject = $script:powershell.BeginInvoke()
    $script:powershell.EndInvoke($AsyncObject)


}


Function RunAppCode{

    $scriptRun = {
        #Suppress Errors
        $script:ErrorActionPreference = 'SilentlyContinue'
        $script:ProgressPreference = 'SilentlyContinue'

        #Collect metadata
        $Author = $hash.Author
        $AuthorDate = $hash.AuthorDate
        $Version = $hash.Version
        Function Add-OutputBoxLine 
        {
            Param ($Message)
            $hash.outputBox.AppendText("$Message")
            $hash.outputBox.Refresh()
            $hash.OutputBox.ScrollToCaret()
            $hash.OutputBox.SelectionStart = $hash.outputBox.Text.Length
            $hash.outputBox.Selectioncolor = "WindowText"
            $hash.Form.Refresh()
        }
        $hash.outputBox.Selectioncolor = "Green"
        Add-OutputBoxLine -Message "`r`nRunning. Please wait."

        #DECLARE VARIABLES
        #Desktop path
        $DesktopPath = [Environment]::GetFolderPath("Desktop")

        #Working Path
        $script:workingPath = Get-Location

        #Date
        $script:datestring = (Get-Date).ToString("s").Replace(":","-")

        #Get SCVMM Credentials
        $secstr = $hash.passwordBox.Text | ConvertTo-SecureString -AsPlainText -Force
        $psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($hash.usernameBox.Text, $secstr)

        #Grey out buttons during code execution.
    $hash.buttonRun.enabled = $false
    $hash.UsernameBox.ReadOnly = $true
    $hash.PasswordBox.ReadOnly = $true
    $hash.serverBox.ReadOnly = $true

    $hash.WatermarkText = "Enter SCVMM server IP or FQDN."


    #Reset progress bar
    $hash.progressBar1.Value = 0
    

    
    #$VMMServer = Read-Host "Please enter your SCVMM server FQDN or IP: "
    [String]$VMMServer = $hash.serverBox.Text

    #Sanity checks.
    If($VMMServer -eq $hash.WatermarkText)
        {
            #Clear the text
            $VMMServer = ""
            #$hash.serverBox.ForeColor = 'WindowText'
        }

    If(!$VMMServer)
    {
        $hash.outputBox.Selectioncolor = "Red"
        Add-OutputBoxLine -Message "`r`nMust enter SCVMM IP address or FQDN."
        $hash.buttonRun.enabled = $true
        $hash.buttonBrowse.enabled = $true
        $hash.UsernameBox.ReadOnly = $false
        $hash.PasswordBox.ReadOnly = $false
        $hash.serverBox.ReadOnly = $false
        return
    }
    
    $verifyVMMServer = Get-SCVMMServer -ComputerName $VMMServer -Credential $psCred

    If(!$verifyVMMServer)
    {
        $hash.outputBox.Selectioncolor = "Red"
        Add-OutputBoxLine -Message "."
        $hash.outputBox.Selectioncolor = "Red"
        Add-OutputBoxLine -Message "`r`nUnable to login to $VMMServer. Please ensure it is an SCVMM instance, check your credentials and try again.`r`nEnsure you are running HVTools from a server that has VMM PowerShell modules installed."
        $hash.outputBox.Selectioncolor = "Red"
        Add-OutputBoxLine -Message "Fail."

        #Grey out buttons during code execution.
        $hash.buttonRun.enabled = $true
        $hash.UsernameBox.ReadOnly = $false
        $hash.PasswordBox.ReadOnly = $false
        $hash.serverBox.ReadOnly = $false
        return
    }
    #Sanity check end.

    

    ## - Create DataTable for Virtual Machines:
    $table = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1 = "VM Name".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Name';
    $table.Columns.Add($column);

    # - Column2 = "FQDN".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "FQDN";
    $table.Columns.Add($column);

    # - Column3 = "Status".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Status";
    $table.Columns.Add($column);

    # - Column4 = "Operating System".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Operating System";
    $table.Columns.Add($column);

    # - Column5 = "Hyper-V Integration Services".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Hyper-V Integration Services";
    $table.Columns.Add($column);

    # - Column6 = "Snapshots / Checkpoints".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Checkpoints";
    $table.Columns.Add($column);

    # - Column7 = "Total Storage in GB".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Total Storage (GB)";
    $table.Columns.Add($column);

    # - Column8 = "Memory".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Memory";
    $table.Columns.Add($column);

    # - Column9 = "Dynamic Memory".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Dynamic Memory";
    $table.Columns.Add($column);

    # - Column10 = "vCPU".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "vCPU";
    $table.Columns.Add($column);

    # - Column11 = "Hyper-V Host".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Hyper-V Host";
    $table.Columns.Add($column);

    # - Column12 = "VHD Location".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Location";
    $table.Columns.Add($column);

    # - Column13 = "VLAN".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "VLAN";
    $table.Columns.Add($column);

    # - Column14 = "IP Address".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "IP Address";
    $table.Columns.Add($column);

    # - Column15 = "Subnet".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Subnet";
    $table.Columns.Add($column);

    # - Column16 = "Default Gateway".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Default Gateway";
    $table.Columns.Add($column);


    <#Finish virtual machines table#>

    #Create table for Cluster info.
    $tableCluster = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Name';
    $tableCluster.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Group';
    $tableCluster.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Cluster Reserve Details';
    $tableCluster.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is VMware HA Enabled';
    $tableCluster.Columns.Add($column);
    
    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Vmware Drs Enabled';
    $tableCluster.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Available Storage Node';
    $tableCluster.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Nodes';
    $tableCluster.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Available Volumes';
    $tableCluster.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Shared Volumes';
    $tableCluster.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'IP Addresses';
    $tableCluster.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Quorum Disk Resource Name';
    $tableCluster.Columns.Add($column);
    
    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Validation Result';
    $tableCluster.Columns.Add($column);

    # - Column13
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Validation Report Path';
    $tableCluster.Columns.Add($column);

    # - Column14
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Cluster Core Resources';
    $tableCluster.Columns.Add($column);

    # - Column15
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is View Only';
    $tableCluster.Columns.Add($column);

    # - Column16
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Marked For Deletion';
    $tableCluster.Columns.Add($column);

    # - Column17
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Fully Cached';
    $tableCluster.Columns.Add($column);



    #Create table for Host info.
    $tableHosts = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Name';
    $tableHosts.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Paths';
    $tableHosts.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Physical CPU Count';
    $tableHosts.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Logical Processor Count';
    $tableHosts.Columns.Add($column);

    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Processor Manufacturer';
    $tableHosts.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Processor Model';
    $tableHosts.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Processor Speed';
    $tableHosts.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Total Memory';
    $tableHosts.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Available Memory';
    $tableHosts.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Disk Volumes';
    $tableHosts.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Operating System';
    $tableHosts.Columns.Add($column);

    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Host Group';
    $tableHosts.Columns.Add($column);

    # - Column13
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Cluster';
    $tableHosts.Columns.Add($column);

    # - Column14
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Cluster Node Status';
    $tableHosts.Columns.Add($column);

    # - Column15
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'HyperV State';
    $tableHosts.Columns.Add($column);

    # - Column16
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'HyperV Version';
    $tableHosts.Columns.Add($column);

    # - Column17
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'HyperV Version State';
    $tableHosts.Columns.Add($column);

    # - Column18
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Remote Connect Enabled';
    $tableHosts.Columns.Add($column);

    # - Column19
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Remote Connect Port';
    $tableHosts.Columns.Add($column);

    # - Column20
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'SSL Tcp Port';
    $tableHosts.Columns.Add($column);

    # - Column21
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Ssl Certificate Hash';
    $tableHosts.Columns.Add($column);

    # - Column22
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Ssh Tcp Port';
    $tableHosts.Columns.Add($column);

    # - Column23
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is RemoteFX Role Installed';
    $tableHosts.Columns.Add($column);

    # - Column24
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Remote Storage Available Capacity';
    $tableHosts.Columns.Add($column);

    # - Column25
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Remote Storage Total Capacity';
    $tableHosts.Columns.Add($column);

    # - Column26
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Local Storage Available Capacity';
    $tableHosts.Columns.Add($column);

    # - Column27
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Local Storage Total Capacity';
    $tableHosts.Columns.Add($column);


    #Table for Storage Pools
    $tablestoragePools = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'SM Name';
    $tablestoragePools.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'SM Display Name';
    $tablestoragePools.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Description';
    $tablestoragePools.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Pool Id';
    $tablestoragePools.Columns.Add($column);

    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Object Id';
    $tablestoragePools.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Storage Array';
    $tablestoragePools.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Remaining Managed Space';
    $tablestoragePools.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'In Use Capacity';
    $tablestoragePools.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Total Managed Space';
    $tablestoragePools.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Usage';
    $tablestoragePools.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Provisioning Type Default';
    $tablestoragePools.Columns.Add($column);

    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Supported Provisioning Types';
    $tablestoragePools.Columns.Add($column);

    # - Column13
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Health Status';
    $tablestoragePools.Columns.Add($column);

    # - Column14
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Accessibility';
    $tablestoragePools.Columns.Add($column);

    # - Column15
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Enabled';
    $tablestoragePools.Columns.Add($column);

    # - Column16
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Added Time';
    $tablestoragePools.Columns.Add($column);

    # - Column17
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Modified Time';
    $tablestoragePools.Columns.Add($column);

    # - Column18
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Marked For Deletion';
    $tablestoragePools.Columns.Add($column);

    # - Column19
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Fully Cached';
    $tablestoragePools.Columns.Add($column);



    #Table for Storage Arrays
    #Create table for Host info.
    $tableStorageArrays = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Name';
    $tableStorageArrays.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Description';
    $tableStorageArrays.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Management Server';
    $tableStorageArrays.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Manufacturer';
    $tableStorageArrays.Columns.Add($column);

    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Model';
    $tableStorageArrays.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Firmware Version';
    $tableStorageArrays.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Serial Number';
    $tableStorageArrays.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Storage Provider';
    $tableStorageArrays.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Storage Pools';
    $tableStorageArrays.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Remaining Capacity';
    $tableStorageArrays.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'In Use Capacity';
    $tableStorageArrays.Columns.Add($column);

    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Total Capacity';
    $tableStorageArrays.Columns.Add($column);

    # - Column13
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Added Time';
    $tableStorageArrays.Columns.Add($column);

    # - Column14
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Modified Time';
    $tableStorageArrays.Columns.Add($column);

    # - Column15
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Enabled';
    $tableStorageArrays.Columns.Add($column);

    
    #Table for ClusterDisks
    $tableClusterDisks = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Name';
    $tableClusterDisks.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Host Cluster';
    $tableClusterDisks.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Owner Node';
    $tableClusterDisks.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Online';
    $tableClusterDisks.Columns.Add($column);

    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'In Use';
    $tableClusterDisks.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Volume GUIDs';
    $tableClusterDisks.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'ID';
    $tableClusterDisks.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Unique ID';
    $tableClusterDisks.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is View Only';
    $tableClusterDisks.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Object Type';
    $tableClusterDisks.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Marked For Deletion';
    $tableClusterDisks.Columns.Add($column);

    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Fully Cached';
    $tableClusterDisks.Columns.Add($column);


    #Create table for Cluster Volumes.
    $tableClusterVolumes = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Volume Label';
    $tableClusterVolumes.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Host';
    $tableClusterVolumes.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Name';
    $tableClusterVolumes.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Classification';
    $tableClusterVolumes.Columns.Add($column);

    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'ID';
    $tableClusterVolumes.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Storage Volume ID';
    $tableClusterVolumes.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Object ID';
    $tableClusterVolumes.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Volume ID';
    $tableClusterVolumes.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Mount Points';
    $tableClusterVolumes.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Free Space';
    $tableClusterVolumes.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Capacity';
    $tableClusterVolumes.Columns.Add($column);

    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Volume Label File System';
    $tableClusterVolumes.Columns.Add($column);

    # - Column13
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is SAN Migration Possible';
    $tableClusterVolumes.Columns.Add($column);

    # - Column14
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Clustered';
    $tableClusterVolumes.Columns.Add($column);

    # - Column15
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Cluster Shared Volume';
    $tableClusterVolumes.Columns.Add($column);

    # - Column16
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'In Use';
    $tableClusterVolumes.Columns.Add($column);

    # - Column17
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Available For Placement';
    $tableClusterVolumes.Columns.Add($column);

    # - Column18
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Disk';
    $tableClusterVolumes.Columns.Add($column);

    # - Column19
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Disk ID';
    $tableClusterVolumes.Columns.Add($column);

    # - Column20
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Storage Disk';
    $tableClusterVolumes.Columns.Add($column);

    # - Column21
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Storage Disk ID';
    $tableClusterVolumes.Columns.Add($column);

    # - Column22
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Storage Pool';
    $tableClusterVolumes.Columns.Add($column);

    # - Column23
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is View Only';
    $tableClusterVolumes.Columns.Add($column);

    # - Column24
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Object Type';
    $tableClusterVolumes.Columns.Add($column);

    # - Column25
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Marked For Deletion';
    $tableClusterVolumes.Columns.Add($column);

    # - Column26
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Fully Cached';
    $tableClusterVolumes.Columns.Add($column);


    #Create table for Networks info.
    $tableNetworks = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Name';
    $tableNetworks.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Description';
    $tableNetworks.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Logical Network';
    $tableNetworks.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Subnet';
    $tableNetworks.Columns.Add($column);

    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Network Gateways';
    $tableNetworks.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VPN Connections';
    $tableNetworks.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'NAT Connections';
    $tableNetworks.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Routing Domain Id';
    $tableNetworks.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Isolation Type';
    $tableNetworks.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Use GRE';
    $tableNetworks.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'External Name';
    $tableNetworks.Columns.Add($column);

    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Network Entity Access Type';
    $tableNetworks.Columns.Add($column);

    # - Column13
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Assigned';
    $tableNetworks.Columns.Add($column);

    # - Column14
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'IsPrivateVlan';
    $tableNetworks.Columns.Add($column);

    # - Column15
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Has Gateway Connection';
    $tableNetworks.Columns.Add($column);

    # - Column16
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Network Manager';
    $tableNetworks.Columns.Add($column);

    # - Column17
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Port ACL';
    $tableNetworks.Columns.Add($column);

    # - Column18
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Granted To List';
    $tableNetworks.Columns.Add($column);

    # - Column19
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'User Role ID';
    $tableNetworks.Columns.Add($column);

    # - Column20
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'User Role';
    $tableNetworks.Columns.Add($column);

    # - Column21
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Owner';
    $tableNetworks.Columns.Add($column);

    # - Column22
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Object Type';
    $tableNetworks.Columns.Add($column);

    # - Column23
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Accessibility';
    $tableNetworks.Columns.Add($column);

    # - Column24
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is View Only';
    $tableNetworks.Columns.Add($column);

    # - Column25
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Added Time';
    $tableNetworks.Columns.Add($column);

    # - Column26
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Modified Time';
    $tableNetworks.Columns.Add($column);

    # - Column27
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Enabled';
    $tableNetworks.Columns.Add($column);

    # - Column28
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Most Recent Task';
    $tableNetworks.Columns.Add($column);

    #Table for ClusterNetworks
    $tableClusterNetworks = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Name';
    $tableClusterNetworks.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Host Cluster';
    $tableClusterNetworks.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Description';
    $tableClusterNetworks.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Bound To VMHost';
    $tableClusterNetworks.Columns.Add($column);

    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Bound Vlan Id';
    $tableClusterNetworks.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Has Common Logical Networks';
    $tableClusterNetworks.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'LogicalNetworks';
    $tableClusterNetworks.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Virtual Networks';
    $tableClusterNetworks.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'ID';
    $tableClusterNetworks.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is View Only';
    $tableClusterNetworks.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Object Type';
    $tableClusterNetworks.Columns.Add($column);

    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Marked For Deletion';
    $tableClusterNetworks.Columns.Add($column);

     # - Column13
     $column = New-Object System.Data.DataColumn;
     $column.DataType = [System.Type]::GetType("System.String");
     $column.ColumnName = 'Is Fully Cached';
     $tableClusterNetworks.Columns.Add($column);


    #Create table for Zombie VHDs.
    $tableZVHDs = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Name';
    $tableZVHDs.Columns.Add($column);

    # - Column2
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Description';
    $tableZVHDs.Columns.Add($column);

    # - Column3
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'ID';
    $tableZVHDs.Columns.Add($column);

    # - Column4
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Operating System';
    $tableZVHDs.Columns.Add($column);

    # - Column5
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Name';
    $tableZVHDs.Columns.Add($column);

    # - Column6
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VM Host';
    $tableZVHDs.Columns.Add($column);

    # - Column7
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Virtualization Platform';
    $tableZVHDs.Columns.Add($column);

    # - Column8
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Namespace';
    $tableZVHDs.Columns.Add($column);

    # - Column9
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VHD Format Type';
    $tableZVHDs.Columns.Add($column);

    # - Column10
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'VHD Type';
    $tableZVHDs.Columns.Add($column);

    # - Column11
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Object Type';
    $tableZVHDs.Columns.Add($column);

    # - Column12
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'State';
    $tableZVHDs.Columns.Add($column);

    # - Column13
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Size';
    $tableZVHDs.Columns.Add($column);

    # - Column14
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Maximum Size';
    $tableZVHDs.Columns.Add($column);

    # - Column15
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Library Server';
    $tableZVHDs.Columns.Add($column);

    # - Column16
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Parent Disk';
    $tableZVHDs.Columns.Add($column);

    # - Column17
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Volume';
    $tableZVHDs.Columns.Add($column);

    # - Column18
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Host Volume ID';
    $tableZVHDs.Columns.Add($column);

    # - Column19
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Share Path';
    $tableZVHDs.Columns.Add($column);

    # - Column20
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'File Share';
    $tableZVHDs.Columns.Add($column);

    # - Column21
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Directory';
    $tableZVHDs.Columns.Add($column);

    # - Column22
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Family Name';
    $tableZVHDs.Columns.Add($column);

    # - Column23
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Orphaned';
    $tableZVHDs.Columns.Add($column);

    # - Column24
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is Cached Vhd';
    $tableZVHDs.Columns.Add($column);

    # - Column25
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Enabled';
    $tableZVHDs.Columns.Add($column);

    # - Column26
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Accessibility';
    $tableZVHDs.Columns.Add($column);

    # - Column27
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Is View Only';
    $tableZVHDs.Columns.Add($column);

    # - Column28
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Added Time';
    $tableZVHDs.Columns.Add($column);

    # - Column29
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Modified Time';
    $tableZVHDs.Columns.Add($column);

    # - Column30
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Most Recent Task If Local';
    $tableZVHDs.Columns.Add($column);





    #Declare variables
    #$filename = "HVTools - $VMMServer - $datestring.csv"
    #$ProbeVMM = Get-SCVMMServer -ComputerName "$VMMServer"
    $VMS = Get-scvirtualmachine | Sort-Object -Property Name
    #Get clusters
    $Clusters = Get-SCVMHostCluster
    #Get Hosts
    $hvHosts = Get-SCVMHost
    #Get Storage Pools
    $StoragePools = Get-SCStoragePool
    #Get storage arrays
    $StorageArrays = Get-SCStorageArray
    #Get cluster disks
    $ClusterDisks = Get-SCStorageClusterDisk
    #Get cluster volumes
    $ClusterVolumes = Get-SCStorageVolume
    #Get networks
    $Networks = Get-SCVMNetwork
    #Get cluster networks
    #$ClusterNetworks = Get-SCClusterVirtualNetwork
    # Requires foreach loop to travers all clusters. ie foreach($cluster in $Clusters) {$ClusterNetwork = Get-SCClusterVirtualNetwork -VMHostCluster $cluster}

    #Get zombie VHDs.
    $ZVHDS = Get-ScVirtualHardDisk | Where-Object {$_.IsOrphaned -eq "True"}

    #Experimental code for a total progress bar.
    [int]$totalProgress = $VMS.Count+$Clusters.Count+$hvHosts.Count+$StoragePools.Count+$StorageArrays.Count+$ClusterDisks.Count+$ClusterVolumes.Count+$Networks.Count+$Clusters.Count+$ZHVDS.Count

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying virtual machines."

    #Foreach loop for VMs
    Foreach ($VM in $VMS)
    {
            
        <#
        $hash.outputBox.Selectioncolor = "Green"
        Add-OutputBoxLine -Message "`r`nProcessing $VM."
        #>

        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()


        $VHDS = Get-SCVirtualHardDisk -VM $VM
        $storage = 0
        Foreach ($VHD in $VHDS)
        {
            $storage += $VHD.MaximumSize
        }
        $storage = $storage/1GB
        $storage = [math]::Round($storage)
    
        #$fqdn = $VM | Select-Object -ExpandProperty ComputerName
        $fqdn = $VM.ComputerName
        #$status = $VM | Select-Object -ExpandProperty Status
        $status = $VM.Status
        #$vmOS = $VM | Select-Object -ExpandProperty OperatingSystem
        $vmOS = $VM.OperatingSystem
        #$hyperVIntegrationServices = $VM | Select-Object -ExpandProperty HasVMAdditions
        $hyperVIntegrationServices = $VM.HasVMAdditions
        #$checkpoints = $VM | Select-Object -ExpandProperty VMCheckpoints
        [String]$checkpoints = $VM.VMCheckpoints
        #$memory = ($VM | Measure-Object -Property Memory -Sum).Sum -as [int]
        $memory = $VM.memory/1024
        #Round memory.
        $memory = [Math]::Round($memory)

        #$dynamicMemory = ($VM | Measure-Object -Property DynamicMemoryMaximumMB -Sum).Sum -as [int]
        $dynamicMemory = $VM.DynamicMemoryMaximumMB/1024
        #Round dynamic memory
        $dynamicMemory = [Math]::Round($dynamicMemory)

        #Check and see if VM has static or dynamic memory,
        if(!$dynamicMemory)
        {
            $dynamicMemory = "N/a"
        }
        else
        {
            $memory = "N/a"
        }
        #$vCPU = ($VM | Measure-Object -Property CPUCount -Sum).Sum -as [int]
        $vCPU = $VM.CPUCount

        #$vHostname = $VM | Select-Object -ExpandProperty HostName
        $vHostname = $VM.HostName

        #Get VHD location
        $vLocation = $VM.Location
        
        #Get VM network
        $VMnetwork = Get-SCVirtualNetworkAdapter -VM $VM
        #$vmVLAN =  Get-SCVirtualNetworkAdapter -VM $VM | Select-Object -ExpandProperty VLanID
        $vmVLAN = $VMnetwork.VLanID
        #$vmIP = Get-SCVirtualNetworkAdapter -VM $VM | Select-Object -ExpandProperty IPv4Addresses
        [String]$vmIP = $VMnetwork.IPv4Addresses
        #$vmSubnet = Get-SCVirtualNetworkAdapter -VM $VM | Select-Object -ExpandProperty IPv4Subnets
        [String]$vmSubnet = $VMnetwork.IPv4Subnets
        #$vmDefaultGateways = Get-SCVirtualNetworkAdapter -VM $VM | Select-Object -ExpandProperty DefaultIPGateways
        [String]$vmDefaultGateways = $VMnetwork.DefaultIPGateways
    

        <#Experimental new code#>
        $row = $table.NewRow();
        $row["VM Name"] = $VM;
        $row["FQDN"] = $fqdn;
        $row["Status"] = $Status;
        $row["Operating System"] = $vmOS;
        $row["Hyper-V Integration Services"] = $hyperVIntegrationServices;
        $row["Checkpoints"] = $checkpoints;
        $row["Total Storage (GB)"] = "$storage GB";
        $row["Memory"] = "$memory GB";
        $row["Dynamic Memory"] = "$dynamicMemory GB";
        $row["vCPU"] = "$vCPU vCPU";
        $row["Hyper-V Host"] = $vHostname;
        $row["Location"] = $vLocation;
        $row["VLAN"] = $vmVLAN;
        $row["IP Address"] = $vmIP;
        $row["Subnet"] = $vmSubnet;
        $row["Default Gateway"] = $vmDefaultGateways;
        $table.Rows.Add($row)



        #Garbage collection
        $VHDS = $null
        $fqdn = $null
        $status = $null
        $vmOS = $null
        $hyperVIntegrationServices = $null
        $checkpoints = $null
        $storage = $null
        $memory = $null
        $dynamicMemory = $null
        $vCPU = $null
        $vHostname = $null
        $vmVLAN = $null
        $vmIP = $null
        $vmSubnet = $null
        $vmDefaultGateways = $null

    } #Close the foreach loop.
    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nVirtual machine inventory complete."


    
    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying clusters."
    #Start foreach loop for Cluster.
    foreach($cluster in $Clusters)
    {
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        #Variables
        $clusterName = $cluster.Name
        $clusterHostGroup = $cluster.HostGroup
        $clusterClusterReserveDetails = $cluster.ClusterReserveDetails
        $clusterIsVMwareHAEnabled = $cluster.IsVMwareHAEnabled
        $clusterIsVmwareDrsEnabled = $cluster.IsVmwareDrsEnabled
        $clusterAvailableStorageNode = $cluster.AvailableStorageNode
        [String]$clusterNodes = $cluster.Nodes
        [String]$clusterAvailableVolumes = $cluster.AvailableVolumes
        [String]$clusterSharedVolumes = $cluster.SharedVolumes
        [String]$clusterIPAddresses = $cluster.IPAddresses
        $clusterQuorumDiskResourceName = $cluster.QuorumDiskResourceName
        $clusterValidationResult = $cluster.ValidationResult
        $clusterValidationReportPath = $cluster.ValidationReportPath
        [String]$clusterClusterCoreResources = $cluster.ClusterCoreResources
        $clusterIsViewOnly = $cluster.IsViewOnly
        $clusterMarkedForDeletion = $cluster.MarkedForDeletion
        $clusterIsFullyCached = $cluster.IsFullyCached

        #Add to cluster table
        $row2 = $tableCluster.NewRow();
        $row2["Name"] = $clusterName;
        $row2["Host Group"] = $clusterHostGroup;
        $row2["Cluster Reserve Details"] = $clusterClusterReserveDetails;
        $row2["Is VMware HA Enabled"] = $clusterIsVMwareHAEnabled;
        $row2["Is VMware Drs Enabled"] = $clusterIsVmwareDrsEnabled;
        $row2["Available Storage Node"] = $clusterAvailableStorageNode;
        $row2["Nodes"] = $clusterNodes;
        $row2["Available Volumes"] = $clusterAvailableVolumes;
        $row2["Shared Volumes"] = $clusterSharedVolumes;
        $row2["IP Addresses"] = $clusterIPAddresses;
        $row2["Quorum Disk Resource Name"] = $clusterQuorumDiskResourceName;
        $row2["Validation Result"] = $clusterValidationResult;
        $row2["Validation Report Path"] = $clusterValidationReportPath;
        $row2["Cluster Core Resources"] = $clusterClusterCoreResources;
        $row2["Is View Only"] = $clusterIsViewOnly;
        $row2["Marked For Deletion"] = $clusterMarkedForDeletion;
        $row2["Is Fully Cached"] = $clusterIsFullyCached;
        $tableCluster.Rows.Add($row2)

    } #Close foreach loop.
    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nCluster inventory complete."


    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying Hyper-V hosts."
    
    #Hosts Foreach loop.
    Foreach($hvHost in $hvHosts)
    {
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        #Variables
        [String]$hvHostName = $hvHost.Name
        [String]$hvVMPaths = $hvHost.VMPaths
        [String]$hvPhysicalCPUCount = $hvHost.PhysicalCPUCount
        [String]$hvLogicalProcessorCount = $hvHost.LogicalProcessorCount
        [String]$hvProcessorManufacturer = $hvHost.ProcessorManufacturer
        [String]$hvProcessorModel = $hvHost.ProcessorModel
        [String]$hvProcessorSpeed = $hvHost.ProcessorSpeed
        [String]$hvTotalMemory = $hvHost.TotalMemory/1GB
        #Round total memory
        $hvTotalMemory = [math]::Round($hvTotalMemory)
        [String]$hvAvailableMemory = $hvHost.AvailableMemory/1GB
        #Round available Memory
        $hvAvailableMemory = [math]::Round($hvAvailableMemory)
        [String]$hvDiskVolumes = $hvHost.DiskVolumes
        [String]$hvOperatingSystem = $hvHost.OperatingSystem
        [String]$hvVMHostGroup = $hvHost.VMHostGroup
        [String]$hvHostCluster = $hvHost.HostCluster
        [String]$hvClusterNodeStatus = $hvHost.ClusterNodeStatus
        [String]$hvHyperVState = $hvHost.HyperVState
        [String]$hvHyperVVersion = $hvHost.HyperVVersion
        [String]$hvHyperVVersionState = $hvHost.HyperVVersionState
        [String]$hvRemoteConnectEnabled = $hvHost.RemoteConnectEnabled
        [String]$hvRemoteConnectPort = $hvHost.RemoteConnectPort
        [String]$hvSSLTcpPort = $hvHost.SSLTcpPort
        [String]$hvSslCertificateHash = $hvHost.SslCertificateHash
        [String]$hvSshTcpPort = $hvHost.SshTcpPort
        [String]$hvIsRemoteFXRoleInstalled = $hvHost.IsRemoteFXRoleInstalled
        [String]$hvRemoteStorageAvailableCapacity = $hvHost.RemoteStorageAvailableCapacity/1GB
        #Round Remote Storage Available Capacity
        $hvRemoteStorageAvailableCapacity = [math]::Round($hvRemoteStorageAvailableCapacity,1)
        
        [String]$hvRemoteStorageTotalCapacity = $hvHost.RemoteStorageTotalCapacity/1GB
        #Round Remote Storage Total Capacity
        $hvRemoteStorageTotalCapacity = [math]::Round($hvRemoteStorageTotalCapacity,1)
        
        [String]$hvLocalStorageAvailableCapacity = $hvHost.LocalStorageAvailableCapacity/1GB
        #Round Local Storage Available Capacity
        $hvLocalStorageAvailableCapacity = [math]::Round($hvLocalStorageAvailableCapacity,1)
        
        [String]$hvLocalStorageTotalCapacity = $hvHost.LocalStorageTotalCapacity/1GB
        #Eound Local Storage Total Capacity
        $hvLocalStorageTotalCapacity = [math]::Round($hvLocalStorageTotalCapacity,1)

        #Add to cluster table
        $row = $tableHosts.NewRow();
        $row["Name"] = $hvHostName;
        $row["VM Paths"] = $hvVmPaths;
        $row["Physical CPU Count"] = $hvPhysicalCPUCount;
        $row["Logical Processor Count"] = $hvLogicalProcessorCount;
        $row["Processor Manufacturer"] = $hvProcessorManufacturer;
        $row["Processor Model"] = $hvProcessorModel;
        $row["Processor Speed"] = "$hvProcessorSpeed Mhz";
        $row["Total Memory"] = "$hvTotalMemory GB";
        $row["Available Memory"] = "$hvAvailableMemory GB";
        $row["Disk Volumes"] = $hvDiskVolumes;
        $row["Operating System"] = $hvOperatingSystem;
        $row["VM Host Group"] = $hvVMHostGroup;
        $row["Host Cluster"] = $hvHostCluster;
        $row["Cluster Node Status"] = $hvClusterNodeStatus;
        $row["HyperV State"] = $hvHyperVState;
        $row["HyperV Version"] = $hvHyperVVersion;
        $row["HyperV Version State"] = $hvHyperVVersionState;
        $row["Remote Connect Enabled"] = $hvRemoteConnectEnabled;
        $row["Remote Connect Port"] = $hvRemoteConnectPort;
        $row["SSL Tcp Port"] = $hvSSLTcpPort;
        $row["Ssl Certificate Hash"] = $hvSslCertificateHash;
        $row["Ssh Tcp Port"] = $hvSshTcpPort;
        $row["Is RemoteFX Role Installed"] = $hvIsRemoteFXRoleInstalled;
        $row["Remote Storage Available Capacity"] = "$hvRemoteStorageAvailableCapacity GB";
        $row["Remote Storage Total Capacity"] = "$hvRemoteStorageTotalCapacity GB";
        $row["Local Storage Available Capacity"] = "$hvLocalStorageAvailableCapacity GB";
        $row["Local Storage Total Capacity"] = "$hvLocalStorageTotalCapacity GB";
        $tableHosts.Rows.Add($row)

    } #Close foreach loop for SCVMM hosts.

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nHyper-V host inventory complete."


    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying storage pools."

    #Foreach loop for Storage Pools.
    foreach($storagePool in $StoragePools)
    {
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        #Variables
        [String]$storagePoolSMName = $storagePool.SMName
        [String]$storagePoolSMDisplayName = $storagePool.SMDisplayName
        [String]$storagePoolDescription = $storagePool.Description
        [String]$storagePoolPoolId = $storagePool.PoolId
        [String]$storagePoolObjectId = $storagePool.ObjectId
        [String]$storagePoolStorageArray = $storagePool.StorageArray
        [String]$storagePoolRemainingManagedSpace = $storagePool.RemainingManagedSpace/1GB
        #Round StoragePoolRemainingManagedSpace
        $storagePoolRemainingManagedSpace = [math]::Round($storagePoolRemainingManagedSpace,1)

        [String]$storagePoolInUseCapacity = $storagePool.InUseCapacity/1GB
        #Round Storage Pool In Use Capacity
        $storagePoolInUseCapacity = [math]::Round($storagePoolInUseCapacity,1)

        [String]$storagePoolTotalManagedSpace = $storagePool.TotalManagedSpace/1GB
        #Round Storage Pool Total Managed Space
        $storagePoolTotalManagedSpace = [math]::Round($storagePoolTotalManagedSpace,1)

        [String]$storagePoolUsage = $storagePool.Usage
        [String]$storagePoolProvisioningTypeDefault = $storagePool.ProvisioningTypeDefault
        [String]$storagePoolSupportedProvisioningTypes = $storagePool.SupportedProvisioningTypes
        [String]$storagePoolHealthStatus = $storagePool.HealthStatus
        [String]$storagePoolAccessibility = $storagePool.Accessibility
        [String]$storagePoolEnabled = $storagePool.Enabled
        [String]$storagePoolAddedTime = $storagePool.AddedTime
        [String]$storagePoolModifiedTime = $storagePool.ModifiedTime
        [String]$storagePoolMarkedForDeletion = $storagePool.MarkedForDeletion
        [String]$storagePoolIsFullyCached = $storagePool.IsFullyCached


        #Add to Storage Pool table
        $row = $tablestoragePools.NewRow();
        $row["SM Name"] = $storagePoolSMName;
        $row["SM Display Name"] = $storagePoolSMDisplayName;
        $row["Description"] = $storagePoolDescription;
        $row["Pool Id"] = $storagePoolPoolId;
        $row["Object Id"] = $storagePoolObjectId;
        $row["Storage Array"] = $storagePoolStorageArray;
        $row["Remaining Managed Space"] = "$storagePoolRemainingManagedSpace GB";
        $row["In Use Capacity"] = "$storagePoolInUseCapacity GB";
        $row["Total Managed Space"] = "$storagePoolTotalManagedSpace GB";
        $row["Usage"] = $storagePoolUsage;
        $row["Provisioning Type Default"] = $storagePoolProvisioningTypeDefault;
        $row["Supported Provisioning Types"] = $storagePoolSupportedProvisioningTypes;
        $row["Health Status"] = $storagePoolHealthStatus;
        $row["Accessibility"] = $storagePoolAccessibility;
        $row["Enabled"] = $storagePoolEnabled;
        $row["Added Time"] = $storagePoolAddedTime;
        $row["Modified Time"] = $storagePoolModifiedTime;
        $row["Marked For Deletion"] = $storagePoolMarkedForDeletion;
        $row["Is Fully Cached"] = $storagePoolIsFullyCached;
        $tablestoragePools.Rows.Add($row)
    }

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nStorage pool inventory complete."

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying storage arrays."

    #Foreach loop for Storage Arrays.
    foreach($sArray in $StorageArrays)
    {
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        #Variables
        [String]$sArraySMName = $sArray.SMName
        [String]$sArrayDescription = $sArray.Description
        [String]$sArrayManagementServer = $sArray.ManagementServer
        [String]$sArrayManufacturer = $sArray.Manufacturer
        [String]$sArrayModel = $sArray.Model
        [String]$sArrayFirmwareVersion = $sArray.FirmwareVersion
        [String]$sArraySerialNumber = $sArray.SerialNumber
        [String]$sArrayStorageProvider = $sArray.StorageProvider
        [String]$sArrayStoragePools = $sArray.StoragePools
        [String]$sArrayRemainingCapacity = $sArray.RemainingCapacity/1GB
        #Round Array Remaining Capacity
        $sArrayRemainingCapacity = [math]::Round($sArrayRemainingCapacity,1)

        [String]$sArrayInUseCapacity = $sArray.InUseCapacity/1GB
        #Round Array In Use Capacity
        $sArrayInUseCapacity = [math]::Round($sArrayInUseCapacity,1)

        [String]$sArrayTotalCapacity = $sArray.TotalCapacity/1GB
        #Round sArrayTotalCapacity
        $sArrayTotalCapacity = [math]::Round($sArrayTotalCapacity,1)

        [String]$sArrayAddedTime = $sArray.AddedTime
        [String]$sArrayModifiedTime = $sArray.ModifiedTime
        [String]$sArrayEnabled = $sArray.Enabled

        #Add to Storage Arrays table
        $row = $tableStorageArrays.NewRow();
        $row["Name"] = $sArraySMName;
        $row["Description"] = $sArrayDescription;
        $row["Management Server"] = $sArrayManagementServer;
        $row["Manufacturer"] = $sArrayManufacturer;
        $row["Model"] = $sArrayModel;
        $row["Firmware Version"] = $sArrayFirmwareVersion;
        $row["Serial Number"] = $sArraySerialNumber;
        $row["Storage Provider"] = $sArrayStorageProvider;
        $row["Storage Pools"] = $sArrayStoragePools;
        $row["Remaining Capacity"] = "$sArrayRemainingCapacity GB";
        $row["In Use Capacity"] = "$sArrayInUseCapacity GB";
        $row["Total Capacity"] = "$sArrayTotalCapacity GB";
        $row["Added Time"] = $sArrayAddedTime;
        $row["Modified Time"] = $sArrayModifiedTime;
        $row["Enabled"] = $sArrayEnabled;
        $tableStorageArrays.Rows.Add($row)

    }

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nStorage inventory complete."

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying cluster disks."

    #Foreach loop for Cluster Disks.
    foreach($clusterDisk in $ClusterDisks)
    {
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        #Variables
        [String]$clusterDiskName = $clusterDisk.Name
        [String]$clusterDiskVMHostCluster = $clusterDisk.VMHostCluster
        [String]$clusterDiskOwnerNode = $clusterDisk.OwnerNode
        [String]$clusterDiskOnline = $clusterDisk.Online
        [String]$clusterDiskInUse = $clusterDisk.InUse
        [String]$clusterDiskVolumeGuids = $clusterDisk.VolumeGuids
        [String]$clusterDiskID = $clusterDisk.ID
        [String]$clusterDiskUniqueID = $clusterDisk.UniqueID
        [String]$clusterDiskIsViewOnly = $clusterDisk.IsViewOnly
        [String]$clusterDiskObjectType = $clusterDisk.ObjectType
        [String]$clusterDiskMarkedForDeletion = $clusterDisk.MarkedForDeletion
        [String]$clusterDiskIsFullyCached = $clusterDisk.IsFullyCached

        #Add to Storage Cluster Disks table
        $row = $tableClusterDisks.NewRow();
        $row["Name"] = $clusterDiskName;
        $row["VM Host Cluster"] = $clusterDiskVMHostCluster;
        $row["Owner Node"] = $clusterDiskOwnerNode;
        $row["Online"] = $clusterDiskOnline;
        $row["In Use"] = $clusterDiskInUse;
        $row["Volume GUIDs"] = $clusterDiskVolumeGuids;
        $row["ID"] = $clusterDiskID;
        $row["Unique ID"] = $clusterDiskUniqueID;
        $row["Is View Only"] = $clusterDiskIsViewOnly;
        $row["Object Type"] = $clusterDiskObjectType;
        $row["Marked For Deletion"] = $clusterDiskMarkedForDeletion;
        $row["Is Fully Cached"] = $clusterDiskIsFullyCached;
        $tableClusterDisks.Rows.Add($row)

    }

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nCluster disk inventory complete."

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying cluster volumes."

    #Foreach loop for Cluster Volumes.
    Foreach($clusterVolume in $ClusterVolumes)
    {
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        [String]$clusterVolumeVolumeLabel = $clusterVolume.VolumeLabel
        [String]$clusterVolumeVMHost = $clusterVolume.VMHost
        [String]$clusterVolumeName = $clusterVolume.Name
        [String]$clusterVolumeClassification = $clusterVolume.Classification
        [String]$clusterVolumeID = $clusterVolume.ID
        [String]$clusterVolumeStorageVolumeID = $clusterVolume.StorageVolumeID
        [String]$clusterVolumeObjectID = $clusterVolume.ObjectID
        [String]$clusterVolumeHostVolumeID = $clusterVolume.HostVolumeID
        [String]$clusterVolumeMountPoints = $clusterVolume.MountPoints
        [String]$clusterVolumeFreeSpace = $clusterVolume.FreeSpace/1GB
        #Rounding free space
        $clusterVolumeFreeSpace = [Math]::Round($clusterVolumeFreeSpace,1)
        
        [String]$clusterVolumeCapacity = $clusterVolume.Capacity/1GB
        #Rounding capacity
        $clusterVolumeCapacity = [Math]::Round($clusterVolumeCapacity,1)
        [String]$clusterVolumeVolumeLabelFileSystem = $clusterVolume.VolumeLabelFileSystem
        [String]$clusterVolumeIsSANMigrationPossible = $clusterVolume.IsSANMigrationPossible
        [String]$clusterVolumeIsClustered = $clusterVolume.IsClustered
        [String]$clusterVolumeIsClusterSharedVolume = $clusterVolume.IsClusterSharedVolume
        [String]$clusterVolumeInUse = $clusterVolume.InUse
        [String]$clusterVolumeIsAvailableForPlacement = $clusterVolume.IsAvailableForPlacement
        [String]$clusterVolumeHostDisk = $clusterVolume.HostDisk
        [String]$clusterVolumeHostDiskID = $clusterVolume.HostDiskID
        [String]$clusterVolumeStorageDisk = $clusterVolume.StorageDisk
        [String]$clusterVolumeStorageDiskID = $clusterVolume.StorageDiskID
        [String]$clusterVolumeStoragePool = $clusterVolume.StoragePool
        [String]$clusterVolumeIsViewOnly = $clusterVolume.IsViewOnly
        [String]$clusterVolumeObjectType = $clusterVolume.ObjectType
        [String]$clusterVolumeMarkedForDeletion = $clusterVolume.MarkedForDeletion
        [String]$clusterVolumeIsFullyCached = $clusterVolume.IsFullyCached

        #Add to Storage Cluster Volumes table
        $row = $tableClusterVolumes.NewRow();
        $row["Volume Label"] = $clusterVolumeVolumeLabel;
        $row["VM Host"] = $clusterVolumeVMHost;
        $row["Name"] = $clusterVolumeName;
        $row["Classification"] = $clusterVolumeClassification;
        $row["ID"] = $clusterVolumeID;
        $row["Storage Volume ID"] = $clusterVolumeStorageVolumeID;
        $row["Object ID"] = $clusterVolumeObjectID;
        $row["Host Volume ID"] = $clusterVolumeHostVolumeID;
        $row["Mount Points"] = $clusterVolumeMountPoints;
        $row["Free Space"] = "$clusterVolumeFreeSpace GB";
        $row["Capacity"] = "$clusterVolumeCapacity GB";
        $row["Volume Label File System"] = $clusterVolumeVolumeLabelFileSystem;
        $row["Is SAN Migration Possible"] = $clusterVolumeIsSANMigrationPossible;
        $row["Is Clustered"] = $clusterVolumeIsClustered;
        $row["Is Cluster Shared Volume"] = $clusterVolumeIsClusterSharedVolume;
        $row["In Use"] = $clusterVolumeInUse;
        $row["Is Available For Placement"] = $clusterVolumeIsAvailableForPlacement;
        $row["Host Disk"] = $clusterVolumeHostDisk;
        $row["Host Disk ID"] = $clusterVolumeHostDiskID;
        $row["Storage Disk"] = $clusterVolumeStorageDisk;
        $row["Storage Disk ID"] = $clusterVolumeStorageDiskID;
        $row["Storage Pool"] = $clusterVolumeStoragePool;
        $row["Is View Only"] = $clusterVolumeIsViewOnly;
        $row["Object Type"] = $clusterVolumeObjectType;
        $row["Marked For Deletion"] = $clusterVolumeMarkedForDeletion;
        $row["Is Fully Cached"] = $clusterVolumeIsFullyCached;
        $tableClusterVolumes.Rows.Add($row)
    }
    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nCluster volume inventory complete."



    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying networks."

    #Foreach loop for Networks.
    Foreach($network in $Networks)
    {
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        #Variables
        [String]$networkName = $network.Name
        [String]$networkDescription = $network.Description
        [String]$networkLogicalNetwork = $network.LogicalNetwork
        [String]$networkVMSubnet = $network.VMSubnet
        [String]$networkVMNetworkGateways = $network.VMNetworkGateways
        [String]$networkVPNConnections = $network.VPNConnections
        [String]$networkNATConnections = $network.NATConnections
        [String]$networkRoutingDomainId = $network.RoutingDomainId
        [String]$networkIsolationType = $network.IsolationType
        [String]$networkUseGRE = $network.UseGRE
        [String]$networkExternalName = $network.ExternalName
        [String]$networkNetworkEntityAccessType = $network.NetworkEntityAccessType
        [String]$networkIsAssigned = $network.IsAssigned
        [String]$networkIsPrivateVlan = $network.IsPrivateVlan
        [String]$networkHasGatewayConnection = $network.HasGatewayConnection
        [String]$networkNetworkManager = $network.NetworkManager
        [String]$networkPortACL = $network.PortACL
        [String]$networkGrantedToList = $network.GrantedToList
        [String]$networkUserRoleID = $network.UserRoleID
        [String]$networkUserRole = $network.UserRole
        [String]$networkOwner = $network.Owner
        [String]$networkObjectType = $network.ObjectType
        [String]$networkAccessibility = $network.Accessibility
        [String]$networkIsViewOnly = $network.IsViewOnly
        [String]$networkAddedTime = $network.AddedTime
        [String]$networkModifiedTime = $network.ModifiedTime
        [String]$networkEnabled = $network.Enabled
        [String]$networkMostRecentTask = $network.MostRecentTask


        #Add to Networks table
        $row = $tableNetworks.NewRow();
        $row["Name"] = $networkName;
        $row["Description"] = $networkDescription;
        $row["Logical Network"] = $networkLogicalNetwork;
        $row["VM Subnet"] = $networkVMSubnet;
        $row["VM Network Gateways"] = $networkVMNetworkGateways;
        $row["VPN Connections"] = $networkVPNConnections;
        $row["NAT Connections"] = $networkNATConnections;
        $row["Routing Domain Id"] = $networkRoutingDomainId;
        $row["Isolation Type"] = $networkIsolationType;
        $row["Use GRE"] = $networkUseGRE;
        $row["External Name"] = $networkExternalName;
        $row["Network Entity Access Type"] = $networkNetworkEntityAccessType;
        $row["Is Assigned"] = $networkIsAssigned;
        $row["IsPrivateVlan"] = $networkIsPrivateVlan;
        $row["Has Gateway Connection"] = $networkHasGatewayConnection;
        $row["Network Manager"] = $networkNetworkManager;
        $row["Port ACL"] = $networkPortACL;
        $row["Granted To List"] = $networkGrantedToList;
        $row["User Role ID"] = $networkUserRoleID;
        $row["User Role"] = $networkUserRole;
        $row["Owner"] = $networkOwner;
        $row["Object Type"] = $networkObjectType;
        $row["Accessibility"] = $networkAccessibility;
        $row["Is View Only"] = $networkIsViewOnly;
        $row["Added Time"] = $networkAddedTime;
        $row["Modified Time"] = $networkModifiedTime;
        $row["Enabled"] = $networkEnabled;
        $row["Most Recent Task"] = $networkMostRecentTask;
        $tableNetworks.Rows.Add($row)
    }

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nNetwork inventory complete."

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nInventorying cluster networks."


    #Reset counter for cluster networks below.
    $c = 0

    #Foreach loop for Cluster Networks
    Foreach($clusterNetwork in $Clusters)
    {
        #Increment for clusters
        $c++
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        $clusterNetwork = Get-SCClusterVirtualNetwork -VMHostCluster $clusterNetwork[$c]

        #Variables
        [String]$clusterNetworkName = $clusterNetwork.Name
        [String]$clusterNetworkVMHostCluster = $clusterNetwork.VMHostCluster
        [String]$clusterNetworkDescription = $clusterNetwork.Description
        [String]$clusterNetworkBoundToVMHost = $clusterNetwork.BoundToVMHost
        [String]$clusterNetworkHostBoundVlanId = $clusterNetwork.HostBoundVlanId
        [String]$clusterNetworkHasCommonLogicalNetworks = $clusterNetwork.HasCommonLogicalNetworks
        [String]$clusterNetworkLogicalNetworks = $clusterNetwork.LogicalNetworks
        [String]$clusterNetworkHostVirtualNetworks = $clusterNetwork.HostVirtualNetworks
        [String]$clusterNetworkID = $clusterNetwork.ID
        [String]$clusterNetworkIsViewOnly = $clusterNetwork.IsViewOnly
        [String]$clusterNetworkObjectType = $clusterNetwork.ObjectType
        [String]$clusterNetworkMarkedForDeletion = $clusterNetwork.MarkedForDeletion
        [String]$clusterNetworkIsFullyCached = $clusterNetwork.IsFullyCached

        #Add to Cluster Networks table
        $row = $tableClusterNetworks.NewRow();
        $row["Name"] = $clusterNetworkName;
        $row["VM Host Cluster"] = $clusterNetworkVMHostCluster;
        $row["Description"] = $clusterNetworkDescription;
        $row["Bound To VM Host"] = $clusterNetworkBoundToVMHost;
        $row["Host Bound Vlan Id"] = $clusterNetworkHostBoundVlanId;
        $row["Has Common Logical Networks"] = $clusterNetworkHasCommonLogicalNetworks;
        $row["Logical Networks"] = $clusterNetworkLogicalNetworks;
        $row["Host Virtual Networks"] = $clusterNetworkHostVirtualNetworks;
        $row["ID"] = $clusterNetworkID;
        $row["Is View Only"] = $clusterNetworkIsViewOnly;
        $row["Object Type"] = $clusterNetworkObjectType;
        $row["Marked For Deletion"] = $clusterNetworkMarkedForDeletion;
        $row["Is Fully Cached"] = $clusterNetworkIsFullyCached;
        $tableClusterNetworks.Rows.Add($row)

    }

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nCluster network inventory complete."

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nFinding zombie VHDs."

    #Foreach loop for zombie VHDs.
    Foreach($ZVHD in $ZVHDS)
    {
        #Calculate progress
        $i++
        #[int]$progressCount = ($i/$VMS.Count)*100
        [int]$progressCount = ($i/$totalProgress)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()

        [String]$ZVHDName = $ZVHD.Name
        [String]$ZVHDDescription = $ZVHD.Description
        [String]$ZVHDID = $ZVHD.ID
        [String]$ZVHDOperatingSystem = $ZVHD.OperatingSystem
        [String]$ZVHDHostName = $ZVHD.HostName
        [String]$ZVHDVMHost = $ZVHD.VMHost
        [String]$ZVHDVirtualizationPlatform = $ZVHD.VirtualizationPlatform
        [String]$ZVHDNamespace = $ZVHD.Namespace
        [String]$ZVHDVHDFormatType = $ZVHD.VHDFormatType
        [String]$ZVHDVHDType = $ZVHD.VHDType
        [String]$ZVHDObjectType = $ZVHD.ObjectType
        [String]$ZVHDState = $ZVHD.State
        [String]$ZVHDSize = $ZVHD.Size/1GB
        #Round ZVHD Size off.
        $ZVHDSize = [Math]::Round($ZVHDSize)
        
        [String]$ZVHDMaximumSize = $ZVHD.MaximumSize/1GB
        #Round ZVHD Maximum Size
        $ZVHDMaximumSize = [Math]::Round($ZVHDMaximumSize)
        [String]$ZVHDLibraryServer = $ZVHD.LibraryServer
        [String]$ZVHDParentDisk = $ZVHD.ParentDisk
        [String]$ZVHDHostVolume = $ZVHD.HostVolume
        [String]$ZVHDHostVolumeID = $ZVHD.HostVolumeID
        [String]$ZVHDSharePath = $ZVHD.SharePath
        [String]$ZVHDFileShare = $ZVHD.FileShare
        [String]$ZVHDDirectory = $ZVHD.Directory
        [String]$ZVHDFamilyName = $ZVHD.FamilyName
        [String]$ZVHDIsOrphaned = $ZVHD.IsOrphaned
        [String]$ZVHDIsCachedVhd = $ZVHD.IsCachedVhd
        [String]$ZVHDEnabled = $ZVHD.Enabled
        [String]$ZVHDAccessibility = $ZVHD.Accessibility
        [String]$ZVHDIsViewOnly = $ZVHD.IsViewOnly
        [String]$ZVHDAddedTime = $ZVHD.AddedTime
        [String]$ZVHDModifiedTime = $ZVHD.ModifiedTime
        [String]$ZVHDMostRecentTaskIfLocal = $ZVHD.MostRecentTaskIfLocal


        #Add to Zombie VHD Table.
        $row = $tableZVHDs.NewRow();
        $row["Name"] = $ZVHDName;
        $row["Description"] = $ZVHDDescription;
        $row["ID"] = $ZVHDID;
        $row["Operating System"] = $ZVHDOperatingSystem;
        $row["Host Name"] = $ZVHDHostName;
        $row["VM Host"] = $ZVHDVMHost;
        $row["Virtualization Platform"] = $ZVHDVirtualizationPlatform;
        $row["Namespace"] = $ZVHDNamespace;
        $row["VHD Format Type"] = $ZVHDVHDFormatType;
        $row["VHD Type"] = $ZVHDVHDType;
        $row["Object Type"] = $ZVHDObjectType;
        $row["State"] = $ZVHDState;
        $row["Size"] = "$ZVHDSize GB";
        $row["Maximum Size"] = "$ZVHDMaximumSize GB";
        $row["Library Server"] = $ZVHDLibraryServer;
        $row["Parent Disk"] = $ZVHDParentDisk;
        $row["Host Volume"] = $ZVHDHostVolume;
        $row["Host Volume ID"] = $ZVHDHostVolumeID;
        $row["Share Path"] = $ZVHDSharePath;
        $row["File Share"] = $ZVHDFileShare;
        $row["Directory"] = $ZVHDDirectory;
        $row["Family Name"] = $ZVHDFamilyName;
        $row["Is Orphaned"] = $ZVHDIsOrphaned;
        $row["Is Cached Vhd"] = $ZVHDIsCachedVhd;
        $row["Enabled"] = $ZVHDEnabled;
        $row["Accessibility"] = $ZVHDAccessibility;
        $row["Is View Only"] = $ZVHDIsViewOnly;
        $row["Added Time"] = $ZVHDAddedTime;
        $row["Modified Time"] = $ZVHDModifiedTime;
        $row["Most Recent Task If Local"] = $ZVHDMostRecentTaskIfLocal;
        $tableZVHDs.Rows.Add($row)

    }

    $hash.outputBox.Selectioncolor = "Green"
    Add-OutputBoxLine -Message "`r`nZombie VHD search complete."



    #Enable button again
    $hash.buttonRun.enabled = $true
    $hash.buttonBrowse.enabled = $true
    $hash.UsernameBox.ReadOnly = $false
    $hash.PasswordBox.ReadOnly = $false
    $hash.serverBox.ReadOnly = $false

    #Textbox completion output.
    Add-OutputBoxLine -Message "`r`n=============="
    Add-OutputBoxLine -Message "`r`n$VMMServer SCVMM Inventory Complete"
    Add-OutputBoxLine -Message "`r`n=============="
    #Add-OutputBoxLine -Message "`r`nCSV report saved on Desktop.`n$DesktopPath\$filename"

    #Output report - Now deprecated by Save As ZIP feature further below.
    #$table | Export-Csv -Path "$DesktopPath\$filename" -NoTypeInformation
    #$Array | Out-Gridview -Title "$filename" -PassThru
    #$ArrayList | Export-Csv -Path "$DesktopPath\$filename" -NoTypeInformation
    #$ArrayList | Out-Gridview -Title "$filename" -PassThru

    #Display second form.
    $Form2 = New-Object system.Windows.Forms.Form
    $Form2Width = '800'
    $Form2Height = '500'
    $Form2.MinimumSize = "$Form2Width,$Form2Height"
    $Form2.StartPosition = 'CenterScreen'
    $Form2.text = "$VMMServer - HVTools"
    $Form2.Icon = [System.Drawing.SystemIcons]::Shield
    #Autoscaling settings
    $Form2.AutoScale = $true
    $Form2.AutoScaleMode = "Font"
    $ASsize = New-Object System.Drawing.SizeF(7,15)
    $Form2.AutoScaleDimensions = $ASsize
    $Form2.BackColor = 'Moccasin'
    $Form2.Refresh()
    #Disable windows maximize feature.
    #$Form2.MaximizeBox = $False
    #$Form2.FormBorderStyle='FixedDialog'

    #Virtual machines data grid.
    $DataGridView1 = New-Object system.Windows.Forms.DataGridView
    $DataGridView1.DataSource = $table
    $DataGridView1.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridView1.AutoSizeColumnsMode = 'AllCells'
    $DataGridView1.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView1.width = 730
    $DataGridView1.height = 345
    #$DataGridView1.width = 710
    #$DataGridView1.height = 460
    $DataGridView1.location = New-Object System.Drawing.Point(12,14)
    $DataGridView1.AutoSize = $True
    $DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $Datagridview1.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Cluster data grid.
    $DataGridView2 = New-Object system.Windows.Forms.DataGridView
    $DataGridView2.DataSource = $tableCluster
    $DataGridView2.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridView2.AutoSizeColumnsMode = 'AllCells'
    $DataGridView2.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView2.width = 730
    $DataGridView2.height = 345
    $DataGridView2.location = New-Object System.Drawing.Point(12,14)
    $DataGridView2.AutoSize = $True
    $DataGridView2.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $Datagridview2.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Hosts data grid.
    $DataGridView3 = New-Object system.Windows.Forms.DataGridView
    $DataGridView3.DataSource = $tableHosts
    $DataGridView3.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridView3.AutoSizeColumnsMode = 'AllCells'
    $DataGridView3.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView3.width = 730
    $DataGridView3.height = 345
    $DataGridView3.location = New-Object System.Drawing.Point(12,14)
    $DataGridView3.AutoSize = $True
    $DataGridView3.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $Datagridview3.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Storage Pools data grid.
    $DataGridView4 = New-Object system.Windows.Forms.DataGridView
    $DataGridView4.DataSource = $tablestoragePools
    $DataGridView4.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridView4.AutoSizeColumnsMode = 'AllCells'
    $DataGridView4.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView4.width = 730
    $DataGridView4.height = 345
    $DataGridView4.location = New-Object System.Drawing.Point(12,14)
    $DataGridView4.AutoSize = $True
    $DataGridView4.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $Datagridview4.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Storage Arrays data grid.
    $DataGridView5 = New-Object system.Windows.Forms.DataGridView
    $DataGridView5.DataSource = $tableStorageArrays
    $DataGridView5.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridView5.AutoSizeColumnsMode = 'AllCells'
    $DataGridView5.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView5.width = 730
    $DataGridView5.height = 345
    $DataGridView5.location = New-Object System.Drawing.Point(12,14)
    $DataGridView5.AutoSize = $True
    $DataGridView5.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $Datagridview5.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Cluster cluster disks data grid.
    $DataGridView6 = New-Object system.Windows.Forms.DataGridView
    $DataGridView6.DataSource = $tableClusterDisks
    $DataGridView6.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridView6.AutoSizeColumnsMode = 'AllCells'
    $DataGridView6.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView6.width = 730
    $DataGridView6.height = 345
    $DataGridView6.location = New-Object System.Drawing.Point(12,14)
    $DataGridView6.AutoSize = $True
    $DataGridView6.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $Datagridview6.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Cluster cluster storage volumes grid.
    $DataGridViewCSV = New-Object system.Windows.Forms.DataGridView
    $DataGridViewCSV.DataSource = $tableClusterVolumes
    $DataGridViewCSV.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridViewCSV.AutoSizeColumnsMode = 'AllCells'
    $DataGridViewCSV.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridViewCSV.width = 730
    $DataGridViewCSV.height = 345
    $DataGridViewCSV.location = New-Object System.Drawing.Point(12,14)
    $DataGridViewCSV.AutoSize = $True
    $DataGridViewCSV.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $DatagridviewCSV.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Networks data grid.
    $DataGridView7 = New-Object system.Windows.Forms.DataGridView
    $DataGridView7.DataSource = $tableNetworks
    $DataGridView7.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridView7.AutoSizeColumnsMode = 'AllCells'
    $DataGridView7.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView7.width = 730
    $DataGridView7.height = 345
    $DataGridView7.location = New-Object System.Drawing.Point(12,14)
    $DataGridView7.AutoSize = $True
    $DataGridView7.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $Datagridview7.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Cluster Networks data grid.
    $DataGridView8 = New-Object system.Windows.Forms.DataGridView
    $DataGridView8.DataSource = $tableClusterNetworks
    $DataGridView8.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridView8.AutoSizeColumnsMode = 'AllCells'
    $DataGridView8.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView8.width = 730
    $DataGridView8.height = 345
    $DataGridView8.location = New-Object System.Drawing.Point(12,14)
    $DataGridView8.AutoSize = $True
    $DataGridView8.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $Datagridview8.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Zombie VHDs data grid.
    $DataGridViewZVHDS = New-Object system.Windows.Forms.DataGridView
    $DataGridViewZVHDS.DataSource = $tableZVHDs
    $DataGridViewZVHDS.Anchor = 'Top, Bottom, Left, Right'
    #$DataGridViewZVHDS.AutoSizeColumnsMode = 'AllCells'
    $DataGridViewZVHDS.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridViewZVHDS.width = 730
    $DataGridViewZVHDS.height = 345
    $DataGridViewZVHDS.location = New-Object System.Drawing.Point(12,14)
    $DataGridViewZVHDS.AutoSize = $True
    $DataGridViewZVHDS.AlternatingRowsDefaultCellStyle.BackColor = "Moccasin"
    $DataGridViewZVHDS.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    #Search filter experimental
        $textbox1_TextChanged = {
        $dataGridView1.Refresh()
        $dataGridView2.Refresh()
        $dataGridView3.Refresh()
        $dataGridView4.Refresh()
        $dataGridView5.Refresh()
        $dataGridView6.Refresh()
        $dataGridViewCSV.Refresh()
        $dataGridView7.Refresh()
        $dataGridView8.Refresh()
        $DataGridViewZVHDS.Refresh()
        $filter = $textbox1.Text
        If($filter -eq $WatermarkText1)
        {
            $filter = ''
        }

        #$datagridview1.DataSource.DefaultView.RowFilter = "[Computer Name] LIKE '*$($textbox1.Text)*' OR [RDP Test] LIKE '*$($textbox1.Text)*' OR [Ping Test] LIKE '*$($textbox1.Text)*'"
        #$datagridview1.DataSource.DefaultView.RowFilter = "[Computer Name] LIKE '*$($textbox1.Text)*' OR [Operating System] LIKE '*$($textbox1.Text)*' OR [CPU] LIKE '*$($textbox1.Text)*' OR [RAM] LIKE '*$($textbox1.Text)*' OR [DNS Test] LIKE '*$($textbox1.Text)*' OR [RDP Port] LIKE '*$($textbox1.Text)*' OR [RDP Test] LIKE '*$($textbox1.Text)*' OR [Ping Test] LIKE '*$($textbox1.Text)*' OR [OS Drive] LIKE '*$($textbox1.Text)*' OR [Free Space in GB] LIKE '*$($textbox1.Text)*' OR [Total Size in GB] LIKE '*$($textbox1.Text)*' OR [Running Services] LIKE '*$($textbox1.Text)*' OR [IP Address] LIKE '*$($textbox1.Text)*' OR [Subnet Mask] LIKE '*$($textbox1.Text)*' OR [Default Gateway] LIKE '*$($textbox1.Text)*' OR [DNS Server 1] LIKE '*$($textbox1.Text)*' OR [DNS Server 2] LIKE '*$($textbox1.Text)*' OR [SMBv1 Status] LIKE '*$($textbox1.Text)*'"
        $datagridview1.DataSource.DefaultView.RowFilter = "[VM Name] + [FQDN] + [Status] + [Operating System] + [Location] LIKE '*$($filter)*'"
        $datagridview2.DataSource.DefaultView.RowFilter = "[Name] + [Host Group] + [Is VMware HA Enabled] + [Is VMware Drs Enabled] + [Available Storage Node] + [Nodes] + [IP Addresses] + [Validation Result] LIKE '*$($filter)*'"
        $DataGridView3.DataSource.DefaultView.RowFilter = "[Name] + [Logical Processor Count] + [Processor Manufacturer] + [Operating System] + [VM Host Group] + [Host Cluster] + [Cluster Node Status] + [HyperV State] LIKE '*$($filter)*'"
        $datagridview4.DataSource.DefaultView.RowFilter = "[SM Name] + [SM Display Name] + [Description] + [Pool Id] + [Storage Array] + [Usage] + [Provisioning Type Default] + [Health Status] + [Accessibility] + [Enabled] LIKE '*$($filter)*'"
        $datagridview5.DataSource.DefaultView.RowFilter = "[Name] + [Description] + [Management Server] + [Manufacturer] + [Model] + [Firmware Version] + [Serial Number] + [Storage Provider] + [Storage Pools] + [Enabled] LIKE '*$($filter)*'"
        $datagridview6.DataSource.DefaultView.RowFilter = "[Name] + [VM Host Cluster] + [Owner Node] + [Online] + [In Use] + [Volume GUIDs] + [ID] + [Unique ID] + [Is View Only] + [Object Type] + [Marked For Deletion] LIKE '*$($filter)*'"
        $datagridviewCSV.DataSource.DefaultView.RowFilter = "[Volume Label] + [VM Host] + [Name] + [Classification] + [ID] + [Storage Volume ID] + [Object ID] + [Host Volume ID] + [Mount Points] + [Volume Label File System] LIKE '*$($filter)*'"
        $datagridview7.DataSource.DefaultView.RowFilter = "[Name] + [Description] + [Logical Network] + [VM Subnet] + [Routing Domain Id] + [Isolation Type] + [Network Entity Access Type] + [Is Assigned] + [IsPrivateVlan] + [Has Gateway Connection] + [Granted To List] + [User Role ID] + [User Role] + [Owner] + [Accessibility] + [Is View Only] + [Added Time] + [Modified Time] LIKE '*$($filter)*'"
        $datagridview8.DataSource.DefaultView.RowFilter = "[Name] + [ID] + [Is View Only] + [Object Type] + [Marked For Deletion] + [Is Fully Cached] LIKE '*$($filter)*'"
        $DataGridViewZVHDS.DataSource.DefaultView.RowFilter = "[Name] + [Description] + [Operating System] + [Host Name] + [VM Host] + [Virtualization Platform] + [Namespace] + [VHD Format Type] + [VHD Type] + [Host Volume] + [Host Volume ID] + [Share Path] + [Directory] + [Is Cached Vhd] + [State] + [Accessibility] + [Added Time] + [Modified Time] + [Most Recent Task If Local] LIKE '*$($filter)*'"
    }
        

    $MainTab = New-Object System.Windows.Forms.TabControl
    $MainTab.Size = '752,360'
    #$MainTab.Size = '755,390'
    #$MainTab.Location = '15,38'
    $MainTab.Location = '15,53'
    $MainTab.Multiline = $True
    $MainTab.Name = 'Main Tab'
    $MainTab.SelectedIndex = 0
    $MainTab.Anchor = 'Top,Left,Bottom,Right'

    $TabPage1 = New-Object System.Windows.Forms.TabPage
    $Tabpage1.Name = 'Virtual Machines'
    #$Tabpage1.Size = '600, 370'
    #$Tabpage1.Size = '700, 370'
    $Tabpage1.Size = '740, 355'
    $Tabpage1.Padding = '5,5,5,5'
    $Tabpage1.TabIndex = 1
    $Tabpage1.Text = 'Virtual Machines'
    $Tabpage1.UseVisualStyleBackColor = $True
    $Tabpage1.Anchor = 'Top,Left,Bottom,Right'
    #$Tabpage1.AutoScroll = $True
    $TabPage1.HorizontalScroll = $true
    $TabPage1.VerticalScroll = $True
    #$TabPage1.Enabled = $false
    $TabPage1.Controls.AddRange(@($DataGridView1))


    <#
    $Tab2title = New-Object system.Windows.Forms.Label
    $Tab2title.text = "SCVMM cluster information coming soon."
    $Tab2title.AutoSize = $false
    $Tab2title.width = 450
    $Tab2title.height = 20
    $Tab2title.location = New-Object System.Drawing.Point(20,40)
    $Tab2title.Font = 'Verdana,11'
    #>


    #<2nd tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage2 = New-Object System.Windows.Forms.TabPage
    $Tabpage2.Name = 'Clusters'
    $Tabpage2.Size = '740, 355'
    $Tabpage2.Padding = '5,5,5,5'
    $Tabpage2.TabIndex = 2
    $Tabpage2.Text = 'Clusters'
    $Tabpage2.UseVisualStyleBackColor = $True
    $Tabpage2.Anchor = 'Top,Left,Bottom,Right'
    $TabPage2.HorizontalScroll = $true
    $TabPage2.VerticalScroll = $True
    $TabPage2.Enabled = $True
    $TabPage2.Controls.AddRange(@($DataGridView2))

    <#
    #3rd tab
    $Tab3title = New-Object system.Windows.Forms.Label
    $Tab3title.text = "SCVMM host information coming soon."
    $Tab3title.AutoSize = $false
    $Tab3title.width = 450
    $Tab3title.height = 20
    $Tab3title.location = New-Object System.Drawing.Point(20,40)
    $Tab3title.Font = 'Verdana,11'
    #>

    #3rd tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage3 = New-Object System.Windows.Forms.TabPage
    $Tabpage3.Name = 'Hosts'
    $Tabpage3.Size = '740, 355'
    $Tabpage3.Padding = '5,5,5,5'
    $Tabpage3.TabIndex = 2
    $Tabpage3.Text = 'Hosts'
    $Tabpage3.UseVisualStyleBackColor = $True
    $Tabpage3.Anchor = 'Top,Left,Bottom,Right'
    $TabPage3.HorizontalScroll = $true
    $TabPage3.VerticalScroll = $True
    $TabPage3.Enabled = $True
    $TabPage3.Controls.AddRange(@($DataGridView3))

    <#
    #4th tab
    $Tab4title = New-Object system.Windows.Forms.Label
    $Tab4title.text = "SCVMM storage pool information coming soon."
    $Tab4title.AutoSize = $false
    $Tab4title.width = 450
    $Tab4title.height = 20
    $Tab4title.location = New-Object System.Drawing.Point(20,40)
    $Tab4title.Font = 'Verdana,11'
    #>


    #4th tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage4 = New-Object System.Windows.Forms.TabPage
    $Tabpage4.Name = 'Storage Pools'
    $Tabpage4.Size = '740, 355'
    $Tabpage4.Padding = '5,5,5,5'
    $Tabpage4.TabIndex = 2
    $Tabpage4.Text = 'Storage Pools'
    $Tabpage4.UseVisualStyleBackColor = $True
    $Tabpage4.Anchor = 'Top,Left,Bottom,Right'
    $TabPage4.HorizontalScroll = $true
    $TabPage4.VerticalScroll = $True
    $TabPage4.Enabled = $True
    $TabPage4.Controls.AddRange(@($DataGridView4))

    <#5th tab
    $Tab5title = New-Object system.Windows.Forms.Label
    $Tab5title.text = "SCVMM storage arrays page coming soon."
    $Tab5title.AutoSize = $false
    $Tab5title.width = 450
    $Tab5title.height = 20
    $Tab5title.location = New-Object System.Drawing.Point(20,40)
    $Tab5title.Font = 'Verdana,11'
    #>


    #5th tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage5 = New-Object System.Windows.Forms.TabPage
    $Tabpage5.Name = 'Storage Arrays'
    $Tabpage5.Size = '740, 355'
    $Tabpage5.Padding = '5,5,5,5'
    $Tabpage5.TabIndex = 2
    $Tabpage5.Text = 'Storage Arrays'
    $Tabpage5.UseVisualStyleBackColor = $True
    $Tabpage5.Anchor = 'Top,Left,Bottom,Right'
    $TabPage5.HorizontalScroll = $true
    $TabPage5.VerticalScroll = $True
    $TabPage5.Enabled = $True
    $TabPage5.Controls.AddRange(@($DataGridView5))

    #6th tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage6 = New-Object System.Windows.Forms.TabPage
    $Tabpage6.Name = 'Cluster Disks'
    $Tabpage6.Size = '740, 355'
    $Tabpage6.Padding = '5,5,5,5'
    $Tabpage6.TabIndex = 2
    $Tabpage6.Text = 'Cluster Disks'
    $Tabpage6.UseVisualStyleBackColor = $True
    $Tabpage6.Anchor = 'Top,Left,Bottom,Right'
    $TabPage6.HorizontalScroll = $true
    $TabPage6.VerticalScroll = $True
    $TabPage6.Enabled = $True
    $TabPage6.Controls.AddRange(@($DataGridView6))

    #Cluster Storage Volume tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPageCSV = New-Object System.Windows.Forms.TabPage
    $TabpageCSV.Name = 'Cluster Storage Volumes'
    $TabpageCSV.Size = '740, 355'
    $TabpageCSV.Padding = '5,5,5,5'
    $TabpageCSV.TabIndex = 2
    $TabpageCSV.Text = 'Cluster Storage Volumes'
    $TabpageCSV.UseVisualStyleBackColor = $True
    $TabpageCSV.Anchor = 'Top,Left,Bottom,Right'
    $TabPageCSV.HorizontalScroll = $true
    $TabPageCSV.VerticalScroll = $True
    $TabPageCSV.Enabled = $True
    $TabPageCSV.Controls.AddRange(@($DataGridViewCSV))

    #7th tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage7 = New-Object System.Windows.Forms.TabPage
    $Tabpage7.Name = 'Networks'
    $Tabpage7.Size = '740, 355'
    $Tabpage7.Padding = '5,5,5,5'
    $Tabpage7.TabIndex = 2
    $Tabpage7.Text = 'Networks'
    $Tabpage7.UseVisualStyleBackColor = $True
    $Tabpage7.Anchor = 'Top,Left,Bottom,Right'
    $TabPage7.HorizontalScroll = $true
    $TabPage7.VerticalScroll = $True
    $TabPage7.Enabled = $True
    $TabPage7.Controls.AddRange(@($DataGridView7))

    #8th tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage8 = New-Object System.Windows.Forms.TabPage
    $Tabpage8.Name = 'Cluster Network'
    $Tabpage8.Size = '740, 355'
    $Tabpage8.Padding = '5,5,5,5'
    $Tabpage8.TabIndex = 2
    $Tabpage8.Text = 'Cluster Network'
    $Tabpage8.UseVisualStyleBackColor = $True
    $Tabpage8.Anchor = 'Top,Left,Bottom,Right'
    $TabPage8.HorizontalScroll = $true
    $TabPage8.VerticalScroll = $True
    $TabPage8.Enabled = $True
    $TabPage8.Controls.AddRange(@($DataGridView8))

    #ZVHD tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPageZVHDS = New-Object System.Windows.Forms.TabPage
    $TabpageZVHDS.Name = 'Zombie VHDs'
    $TabpageZVHDS.Size = '740, 355'
    $TabpageZVHDS.Padding = '5,5,5,5'
    $TabpageZVHDS.TabIndex = 2
    $TabpageZVHDS.Text = 'Zombie VHDs'
    $TabpageZVHDS.UseVisualStyleBackColor = $True
    $TabpageZVHDS.Anchor = 'Top,Left,Bottom,Right'
    $TabPageZVHDS.HorizontalScroll = $true
    $TabPageZVHDS.VerticalScroll = $True
    $TabPageZVHDS.Enabled = $True
    $TabPageZVHDS.Controls.AddRange(@($DataGridViewZVHDS))


    $Tab9title = New-Object system.Windows.Forms.Label
    $Tab9title.text = "SCVMM health page coming soon."
    $Tab9title.AutoSize = $false
    $Tab9title.width = 450
    $Tab9title.height = 20
    $Tab9title.location = New-Object System.Drawing.Point(20,40)
    $Tab9title.Font = 'Verdana,11'

    #vHealth tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage9 = New-Object System.Windows.Forms.TabPage
    $Tabpage9.Name = 'vHealth'
    $Tabpage9.Size = '740, 355'
    $Tabpage9.Padding = '5,5,5,5'
    $Tabpage9.TabIndex = 2
    $Tabpage9.Text = 'vHealth'
    $Tabpage9.UseVisualStyleBackColor = $True
    $Tabpage9.Anchor = 'Top,Left,Bottom,Right'
    $TabPage9.HorizontalScroll = $true
    $TabPage9.VerticalScroll = $True
    $TabPage9.Enabled = $True
    $TabPage9.Controls.AddRange(@($Tab9title))
    

    #Close Button
    $buttonClose = New-Object System.Windows.Forms.Button
    $buttonClose.text = "Close"
    #$buttonClose.Size = '80,30'
    $buttonClose.Size = '80,30'
    $buttonClose.location = '690, 420'
    $buttonClose.Font = 'Verdana,9'
    $buttonClose.Anchor = 'Bottom,Right'
    $buttonClose.Add_Click({$Form2.Close()})

    #Search filter box
    $textbox1 = New-Object System.Windows.Forms.TextBox
    $textbox1.Location = '585, 31'
    $textbox1.anchor = 'Top,Right'
    $textbox1.Name = 'textbox1'
    $textbox1.Size = '180, 15'
    $textbox1.TabIndex = 1
    $textbox1.add_TextChanged($textbox1_TextChanged)
    #Watermark
    $WatermarkText1 = "Search"
    $textbox1.ForeColor = 'Gray'
    $textbox1.Text = $WatermarkText1
    #If we have focus then clear out the text
    $textbox1.Add_GotFocus(
        {
            If($textbox1.Text -eq $WatermarkText1)
            {
                $textbox1.Text = ''
                $textbox1.ForeColor = 'WindowText'
            }
        }
    )
    #If we have lost focus and the field is empty, reset back to watermark.
    $textbox1.Add_LostFocus(
        {
            If($textbox1.Text -eq '')
            {
                $textbox1.Text = $WatermarkText1
                $textbox1.ForeColor = 'Gray'
            }
        }
    )
    


    #ABOUT FORM
    $FormAbout = New-Object system.Windows.Forms.Form
    $FormAboutWidth = '800'
    $FormAboutHeight = '500'
    $FormAbout.MinimumSize = "$FormAboutWidth,$FormAboutHeight"
    $FormAbout.StartPosition = 'CenterScreen'
    $FormAbout.text = "About"
    $FormAbout.Icon = [System.Drawing.SystemIcons]::Shield
    #Autoscaling settings
    $FormAbout.AutoScale = $true
    $FormAbout.AutoScaleMode = "Font"
    $ASsize = New-Object System.Drawing.SizeF(7,15)
    $FormAbout.AutoScaleDimensions = $ASsize
    $FormAbout.BackColor = 'moccasin'
    $FormAbout.Refresh()
    #Disable windows maximize feature.
    $FormAbout.MaximizeBox = $False
    #$FormAbout.FormBorderStyle ='FixedDialog'

    $AboutHeading = New-Object system.Windows.Forms.Label
    $AboutHeading.text = "HVTools"
    $AboutHeading.AutoSize = $false
    $AboutHeading.width = 150
    $AboutHeading.height = 30
    $AboutHeading.location = New-Object System.Drawing.Point(20,20)
    $AboutHeading.Font = 'Verdana,14'
    $AboutHeading.Anchor = 'top, left'

    $AboutDescription = New-Object system.Windows.Forms.Label
    $AboutDescription.text = "Credits.`r`nAuthor: $Author Build date: $AuthorDate Version: $Version"
    $AboutDescription.AutoSize = $false
    $AboutDescription.width = 500
    $AboutDescription.height = 70
    $AboutDescription.location = New-Object System.Drawing.Point(20,410)
    $AboutDescription.Font = 'Verdana,10'
    $AboutDescription.Anchor = 'bottom, left'

    $Description2Heading = New-Object system.Windows.Forms.Label
    $Description2Heading.text = "Instructions"
    $Description2Heading.AutoSize = $false
    $Description2Heading.width = 150
    $Description2Heading.height = 30
    $Description2Heading.location = New-Object System.Drawing.Point(20,80)
    $Description2Heading.Font = 'Verdana,12'

    $AboutDescription2 = New-Object system.Windows.Forms.Label
    $AboutDescription2.text = "1. Enter SCVMM administrator privileged credentials in the Username and Password fields.`r`n2. Enter SCVMM instance ID address in the third field and click Login.`r`n    Example: scvmminstancefqdn.contoso.com.`r`n3. When inventory is complete, the HVTools report will automatically open.`r`n4. You can then click File -> Save As to save the complete inventory to Excel CSV.`r`n5. By selecting desired rows and performing CTRL + C, you can copy/paste data to Excel (optional).`r`n`r`nDesigned for Microsoft System Centre Virtual Machine Manager 2016."
    $AboutDescription2.AutoSize = $false
    $AboutDescription2.width = 620
    $AboutDescription2.height = 200
    $AboutDescription2.location = New-Object System.Drawing.Point(20,120)
    $AboutDescription2.Font = 'Verdana,10'

    #About Close Button
    $aboutClose = New-Object System.Windows.Forms.Button
    $aboutClose.text = "Close"
    #$buttonClose.Size = '80,30'
    $aboutClose.Size = '80,30'
    $aboutClose.location = '350, 360'
    $aboutClose.Font = 'Verdana,9'
    $aboutClose.Anchor = 'Bottom,Left'
    $aboutClose.Add_Click({$FormAbout.Close()})

    $FormAbout.Controls.AddRange(@($AboutHeading, $AboutDescription, $Description2Heading, $AboutDescription2, $aboutClose))

    #Help Menu
    $helpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
    $helpAbout.Name = "About"
    $helpAbout.Text = "About"
    $helpAbout.Add_Click({$FormAbout.ShowDialog()})

    $menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuHelp.Name = "Help"
    $menuHelp.Text = "Help"
    $menuHelp.DropDownItems.AddRange(@($helpAbout))
    
    
    #Main Menu Toolbar

    #Adding type assembly to create compressed ZIPs.
    Add-Type -assembly "system.io.compression.filesystem"
    #File Menu
    $menuSaveAs = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuSaveAs.Name = "Save As"
    $menuSaveAs.Text = "Save As"
    $menuSaveAs.Add_Click({
        try {
            #Save dialogue
            $saveDlg = New-Object System.Windows.Forms.SaveFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
            $saveDlg.ShowHelp = $true
            $saveDlg.CreatePrompt = $false
            $saveDlg.OverwritePrompt = $false
            $saveDlg.RestoreDirectory = $true
            #$saveDlg.filter = "Csv (*.csv)| *.csv|Txt (*.txt)| *.txt"
            $saveDlg.filter = "Zip (*.zip)| *.zip"

            if($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                #Gets file save path including file name.
                $SaveFilePath = $saveDlg.FileName

                #Get file name only without path.
                $SaveFileName = [System.IO.Path]::GetFileName($saveDlg.FileName)
                
                #Get temp folder
                $tempFolder = [System.IO.Path]::GetTempPath()
                #Create a temp folder
                [system.io.directory]::CreateDirectory("$tempFolder\$SaveFileName")
                
                #Export CSVs to temp directory
                $table | Export-CSV -Path "$tempFolder\$SaveFileName\Virtual Machines.csv" -NoTypeInformation
                $tableCluster | Export-CSV -Path "$tempFolder\$SaveFileName\Cluster.csv" -NoTypeInformation
                $tableHosts | Export-CSV -Path "$tempFolder\$SaveFileName\Hosts.csv" -NoTypeInformation
                $tablestoragePools | Export-CSV -Path "$tempFolder\$SaveFileName\Storage Pools.csv" -NoTypeInformation
                $tableStorageArrays | Export-CSV -Path "$tempFolder\$SaveFileName\Storage Arrays.csv" -NoTypeInformation
                $tableClusterDisks | Export-CSV -Path "$tempFolder\$SaveFileName\Cluster Disks.csv" -NoTypeInformation
                $tableClusterVolumes | Export-CSV -Path "$tempFolder\$SaveFileName\Cluster Volumes.csv" -NoTypeInformation
                $tableNetworks | Export-CSV -Path "$tempFolder\$SaveFileName\Networks.csv" -NoTypeInformation
                $tableClusterNetworks | Export-CSV -Path "$tempFolder\$SaveFileName\Cluster Networks.csv" -NoTypeInformation
                $tableZVHDs | Export-CSV -Path "$tempFolder\$SaveFileName\Zombie VHDs.csv" -NoTypeInformation
                
                #Save to ZIP in desired path.
                $source = "$tempFolder\$SaveFileName"
                [io.compression.zipfile]::CreateFromDirectory($Source, $SaveFilePath) 
                
                #Remove folder
                If(Test-Path "$tempFolder\$SaveFileName")
                {
                Remove-Item -Path "$tempFolder\$SaveFileName" -Recurse -Force
                    #[system.io.directory]::Delete("$tempFolder\$SaveFileName", $true)
                }
                
            }
        }
        catch {
            $hash.outputBox.Selectioncolor = "Red"
            Add-OutputBoxLine -Message "`r`nOoops. An error occurred when saving your report $SaveFileName. Please try again."
    
            }
    })

    $menuClose2 = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuClose2.Name = "Close"
    $menuClose2.Text = "Close"
    $menuClose2.Add_Click({$Form2.Close()})

    #File Menu continued
    $menuFile2 = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuFile2.Name = "File"
    $menuFile2.Text = "File"
    $menuFile2.DropDownItems.AddRange(@($menuSaveAs, $menuClose2))



    #Main menu
    $menuMain2 = New-Object System.Windows.Forms.MenuStrip
    $menuMain2.Items.AddRange(@($menuFile2, $menuHelp))
    #Display the main tab GUI.
    $MainTab.Controls.AddRange(@($TabPage1, $TabPage2, $TabPage3, $TabPage4, $TabPage5, $TabPage6, $TabPageCSV, $TabPage7, $TabPage8, $TabPageZVHDS, $TabPage9))

    #Display the 2nd GUI once the command finishes executing. Contains above tab.
    #$script:Form2.Controls.AddRange(@($MainTab, $buttonClose))
    $script:Form2.Controls.AddRange(@($menuMain2, $mainTab, $textbox1, $buttonClose))
    $Form2.ShowDialog()



    #Garbage collection
    #$Array = $null
    $VMMServer = $null
    #$filename = $null
    $VMS = $null



    } #Close the $scriptRun brackets for the runspace
    
    #RUNSPACE
    #$script:runspace = [runspacefactory]::CreateRunspace()
    $maxthreads = [int]$env:NUMBER_OF_PROCESSORS
    #$maxthreads = 3
    #$session = [initialsessionstate]::CreateDefault2()
    #$session.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('hash', $hash,''))
    $hashVars = New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'hash',$hash,$Null
    $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    #Add the variable to the sessionstate
    $InitialSessionState.Variables.Add($hashVars)
    $script:runspace = [runspacefactory]::CreateRunspacePool(1,$maxthreads,$InitialSessionState, $Host)
    $script:runspace.ApartmentState = "STA"
    #$script:runspace = [runspacefactory]::CreateRunspacePool(1,$maxthreads,$session, $Host)

    
        #$script:runspace.Variables.Add
        $script:powershell = [powershell]::Create()
        $script:runspace.Open()
    
        #$script:powershell = [powershell]::Create()
        #$hash.outputBox.AppendText("`r`nCreating Powershell session in loop")

        #Begin our main code within the runspace
        $script:powershell.AddScript($scriptRun)
        #$hash.outputBox.AppendText("`r`nAdding script")
        $script:powershell.RunspacePool = $script:runspace
        
        #$script:handle = $script:powershell.BeginInvoke()
        $script:handle = $script:powershell.BeginInvoke()
        if ($script:handle.IsCompleted)
        {
            $script:powershell.EndInvoke($script:handle)
            #$script:powershell.Close()
            $script:powershell.Dispose()
            $script:runspace.Dispose()
            $script:runspace.Close()
            [System.GC]::Collect()
        }     

 } #Closing the function.


#Menu GUI begins.

#GUI

# Install .Net Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Main Form
$script:hash.Form = New-Object system.Windows.Forms.Form
$FormWidth = '500'
$FormHeight = '690'
$hash.Form.Size = "$FormWidth,$FormHeight"
$hash.Form.StartPosition = 'CenterScreen'
$hash.Form.text = "HVTools $Version"
$hash.Form.Icon = [System.Drawing.SystemIcons]::Shield
$hash.Form.BackColor = 'orange'
$hash.Form.Refresh()
#Disable windows maximize feature.
$hash.Form.MaximizeBox = $False
$hash.Form.FormBorderStyle='FixedDialog'

#Xmas easter egg
$date = (get-date)
$year = (get-date).Year
$newYear = (get-date).AddYears(1).Year
$xmas = Get-Date -Month 12 -Day 25 -Year $year
$nye = Get-Date -Month 01 -Day 02 -Year $newYear
If(($date -ge $xmas) -and ($date -le $nye))
{
    $hash.Form.text = "HVTools $Version - Merry Christmas and a Happy New Year!"
    $hash.Form.BackColor = 'FireBrick'
    $hash.Form.Refresh()
}


$Description = New-Object system.Windows.Forms.Label
$Description.text = "HVTools"
$Description.AutoSize = $false
$Description.width = 450
$Description.height = 20
$Description.location = New-Object System.Drawing.Point(20,40)
$Description.Font = 'Verdana,11'

$moto = New-Object system.Windows.Forms.Label
$moto.text = "Get information from your SCVMM virtual infrastructure environment."
$moto.AutoSize = $false
$moto.width = 450
$moto.height = 50
$moto.location = New-Object System.Drawing.Point(20,70)
$moto.Font = 'Verdana,10'


$UsernameTitle = New-Object system.Windows.Forms.Label
$UsernameTitle.text = "Username: "
$UsernameTitle.AutoSize = $false
$UsernameTitle.width = 85
$UsernameTitle.height = 20
$UsernameTitle.location = New-Object System.Drawing.Point(90,115)
$UsernameTitle.Font = 'Verdana,8'


#Username Box
$script:hash.userNameBox = New-Object System.Windows.Forms.TextBox 
$hash.userNameBox.Location = New-Object System.Drawing.Size(190,115) 
$hash.userNameBox.Size = New-Object System.Drawing.Size(200,40)
$hash.userNameBox.Font = 'Verdana,9'
$hash.userNameBox.ReadOnly = $False
$hash.usernameWatermarkText = "CONTOSO\Administrator"
$hash.userNameBox.ForeColor = 'Gray'
$hash.userNameBox.Text = $hash.usernameWatermarkText
#If we have focus then clear out the text
$hash.userNameBox.Add_GotFocus(
    {
        If($hash.userNameBox.Text -eq $hash.usernameWatermarkText)
        {
            $hash.userNameBox.Text = ''
            $hash.userNameBox.ForeColor = 'WindowText'
        }
    }
)
#If we have lost focus and the field is empty, reset back to watermark.
$hash.userNameBox.Add_LostFocus(
    {
        If($hash.userNameBox.Text -eq '')
        {
            $hash.userNameBox.Text = $hash.usernameWatermarkText
            $hash.userNameBox.ForeColor = 'Gray'
        }
    }
)

$PasswordTitle = New-Object system.Windows.Forms.Label
$PasswordTitle.text = "Password: "
$PasswordTitle.AutoSize = $false
$PasswordTitle.width = 85
$PasswordTitle.height = 20
$PasswordTitle.location = New-Object System.Drawing.Point(90,145)
$PasswordTitle.Font = 'Verdana,8'

#Password Box
$script:hash.passwordBox = New-Object System.Windows.Forms.MaskedTextBox 
$hash.passwordBox.PasswordChar = '*'
$hash.passwordBox.Location = New-Object System.Drawing.Size(190,145) 
$hash.passwordBox.Size = New-Object System.Drawing.Size(200,40)
$hash.passwordBox.Font = 'Verdana,9'
$hash.passwordBox.ReadOnly = $False
$hash.passwordWatermarkText = ""
$hash.passwordBox.ForeColor = 'Gray'
$hash.passwordBox.Text = $hash.passwordWatermarkText
#If we have focus then clear out the text
$hash.passwordBox.Add_GotFocus(
    {
        If($hash.passwordBox.Text -eq $hash.passwordWatermarkText)
        {
            $hash.passwordBox.Text = ''
            $hash.passwordBox.ForeColor = 'WindowText'
        }
    }
)
#If we have lost focus and the field is empty, reset back to watermark.
$hash.passwordBox.Add_LostFocus(
    {
        If($hash.passwordBox.Text -eq '')
        {
            $hash.passwordBox.Text = $hash.passwordWatermarkText
            $hash.passwordBox.ForeColor = 'Gray'
        }
    }
)



#Input single computer or FQDN textbox
$script:hash.serverBox = New-Object System.Windows.Forms.TextBox 
$hash.serverBox.Location = New-Object System.Drawing.Size(90,175) 
$hash.serverBox.Size = New-Object System.Drawing.Size(300,40)
$hash.serverBox.Font = 'Verdana,9'
$hash.serverBox.ReadOnly = $False
$hash.WatermarkText = "Enter SCVMM server IP or FQDN."
$hash.serverBox.ForeColor = 'Gray'
$hash.serverBox.Text = $hash.WatermarkText
#If we have focus then clear out the text
$hash.serverBox.Add_GotFocus(
    {
        If($hash.serverBox.Text -eq $hash.WatermarkText)
        {
            $hash.serverBox.Text = ''
            $hash.serverBox.ForeColor = 'WindowText'
        }
    }
)
#If we have lost focus and the field is empty, reset back to watermark.
$hash.serverBox.Add_LostFocus(
    {
        If($hash.serverBox.Text -eq '')
        {
            $hash.serverBox.Text = $hash.WatermarkText
            $hash.serverBox.ForeColor = 'Gray'
        }
    }
)

$hash.serverBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        runAppCode
    }
})


<# Output Box which is below all other buttons and displays PS Output #>
$hash.outputBox = New-Object System.Windows.Forms.RichTextBox 
$hash.outputBox.Location = New-Object System.Drawing.Size(65,360) 
$hash.outputBox.Size = New-Object System.Drawing.Size(350,180)
$hash.outputBox.Font = "Verdana, 8"
$hash.outputBox.ReadOnly = $True
$hash.outputBox.MultiLine = $True 
$hash.outputBox.ScrollBars = "Vertical"
#$hash.outputBox.AppendText("PVT Tool ready.")
#$hash.Form.Controls.Add($hash.outputBox)

#PowerShell version check.
If($getPowerShellVersion -ge "4.0")
{
        $hash.outputBox.AppendText("Your PowerShell version is $getPowerShellVersion. `nHVTools is ready.")
}
else
{
        $hash.outputBox.Selectioncolor = "Red"
        $hash.outputBox.AppendText("Your PowerShell version is $getPowerShellVersion. `nVMM Tool may not run correctly on your computer.")
}


$hash.buttonRun = New-Object System.Windows.Forms.Button
$hash.buttonRun.text = "Login"
$hash.buttonRun.Size = '300,40'
$hash.buttonRun.location = '90, 210'
$hash.buttonRun.Font = 'Verdana,11'
$hash.buttonRun.BackColor = "CornflowerBlue"
$hash.buttonRun.Cursor = [System.Windows.Forms.Cursors]::Hand
$hash.buttonRun.Add_Click({RunAppCode})
#$hash.buttonRun.DialogResult = [System.Windows.Forms.DialogResult]::Ok

#Close button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = '190,280'
$cancelButton.Size = '100,40'
$cancelButton.FlatStyle = 'Flat'
$cancelButton.BackColor = 'Brown'
$cancelButton.Font = 'Verdana, 9'
#$cancelFont = New-Object System.Drawing.Font('Calibri',11,[System.Drawing.FontStyle]::Bold)

# Font styles are: Regular, Bold, Italic, Underline, Strikeout
#$cancelButton.Font = $cancelFont
$hash.Form.Controls.Add($CancelButton)
$cancelButton.Text = 'Exit'
$cancelButton.tabindex = 0
$cancelButton.Add_Click({$hash.Form.Tag = $hash.Form.close()
$script:powershell.EndInvoke($script:handle)
#$script:powershell.Close()
$script:powershell.Dispose()
$script:runspace.Close()
$script:runspace.Dispose()
[System.GC]::Collect()})
$hash.Form.CancelButton = $CancelButton


#Progress status Heading
$StatusHeading = New-Object system.Windows.Forms.Label
$StatusHeading.text = "Progress: "
$StatusHeading.AutoSize = $false
$StatusHeading.width = 80
$StatusHeading.height = 20
$StatusHeading.location = New-Object System.Drawing.Point(70,575)
$StatusHeading.Anchor = 'bottom, left'
$StatusHeading.Font = 'Verdana,10'

#Progress bar.
$hash.progressBar1 = New-Object System.Windows.Forms.ProgressBar
$hash.progressBar1.Name = 'progressBar1'
$hash.progressBar1.Value = 0
$hash.progressBar1.Style="Blocks"
$hash.progressBar1.Size = "254, 30"
$hash.progressBar1.location = '160, 570'
$hash.progressBar1.Anchor = 'bottom, left'


#File Menu
$menuClose = New-Object System.Windows.Forms.ToolStripMenuItem
$menuClose.Name = "Close"
$menuClose.Text = "Close"
$menuClose.Add_Click({$hash.Form.Close()
    $script:powershell.EndInvoke($script:handle)
    #$script:powershell.Close()
    $script:powershell.Dispose()
    $script:runspace.Dispose()
    $script:runspace.Close()
    [System.GC]::Collect()})

#File Menu continued
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$menuFile.Name = "File"
$menuFile.Text = "File"
$menuFile.DropDownItems.AddRange(@($menuClose))


#ABOUT FORM
 $FormAbout = New-Object system.Windows.Forms.Form
 $FormAboutWidth = '800'
 $FormAboutHeight = '500'
 $FormAbout.MinimumSize = "$FormAboutWidth,$FormAboutHeight"
 $FormAbout.StartPosition = 'CenterScreen'
 $FormAbout.text = "About"
 $FormAbout.Icon = [System.Drawing.SystemIcons]::Shield
 #Autoscaling settings
 $FormAbout.AutoScale = $true
 $FormAbout.AutoScaleMode = "Font"
 $ASsize = New-Object System.Drawing.SizeF(7,15)
 $FormAbout.AutoScaleDimensions = $ASsize
 $FormAbout.BackColor = 'orange'
 $FormAbout.Refresh()
 #Disable windows maximize feature.
 $FormAbout.MaximizeBox = $False
 #$FormAbout.FormBorderStyle ='FixedDialog'

 $AboutHeading = New-Object system.Windows.Forms.Label
 $AboutHeading.text = "HVTools"
 $AboutHeading.AutoSize = $false
 $AboutHeading.width = 150
 $AboutHeading.height = 30
 $AboutHeading.location = New-Object System.Drawing.Point(20,20)
 $AboutHeading.Font = 'Verdana,14'
 $AboutHeading.Anchor = 'top, left'

 $AboutDescription = New-Object system.Windows.Forms.Label
 $AboutDescription.text = "Credits.`r`nAuthor: $Author Build date: $AuthorDate Version: $Version"
 $AboutDescription.AutoSize = $false
 $AboutDescription.width = 500
 $AboutDescription.height = 70
 $AboutDescription.location = New-Object System.Drawing.Point(20,410)
 $AboutDescription.Font = 'Verdana,10'
 $AboutDescription.Anchor = 'bottom, left'

 $Description2Heading = New-Object system.Windows.Forms.Label
 $Description2Heading.text = "Instructions"
 $Description2Heading.AutoSize = $false
 $Description2Heading.width = 150
 $Description2Heading.height = 30
 $Description2Heading.location = New-Object System.Drawing.Point(20,80)
 $Description2Heading.Font = 'Verdana,12'

$AboutDescription2 = New-Object system.Windows.Forms.Label
$AboutDescription2.text = "1. Enter SCVMM administrator privileged credentials in the Username and Password fields.`r`n2. Enter SCVMM instance ID address in the third field and click Login.`r`n    Example: scvmminstancefqdn.contoso.com.`r`n3. When inventory is complete, the HVTools report will automatically open.`r`n4. You can then click File -> Save As to save the complete inventory to Excel CSV.`r`n5. By selecting desired rows and performing CTRL + C, you can copy/paste data to Excel (optional).`r`n`r`nDesigned for Microsoft System Centre Virtual Machine Manager 2016."
$AboutDescription2.AutoSize = $false
$AboutDescription2.width = 620
$AboutDescription2.height = 200
$AboutDescription2.location = New-Object System.Drawing.Point(20,120)
$AboutDescription2.Font = 'Verdana,10'

#About Close Button
$aboutClose = New-Object System.Windows.Forms.Button
$aboutClose.text = "Close"
#$buttonClose.Size = '80,30'
$aboutClose.Size = '80,30'
$aboutClose.location = '350, 360'
$aboutClose.Font = 'Verdana,9'
$aboutClose.Anchor = 'Bottom,Left'
$aboutClose.Add_Click({$FormAbout.Close()})

$FormAbout.Controls.AddRange(@($AboutHeading, $AboutDescription, $Description2Heading, $AboutDescription2, $aboutClose))

#Help Menu
$helpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$helpAbout.Name = "About"
$helpAbout.Text = "About"
$helpAbout.Add_Click({$FormAbout.ShowDialog()})

$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelp.Name = "Help"
$menuHelp.Text = "Help"
$menuHelp.DropDownItems.AddRange(@($helpAbout))

$menuMain = New-Object System.Windows.Forms.MenuStrip
$menuMain.Items.AddRange(@($menuFile, $menuHelp))



#Display form
$script:hash.Form.Controls.AddRange(@($menuMain, $Description, $UsernameTitle, $hash.usernameBox, $PasswordTitle, $hash.passwordBox, $hash.serverBox, $hash.buttonRun, $moto, $hash.outputBox, $StatusHeading, $hash.progressBar1))
$result = $hash.Form.ShowDialog()

if($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        $hash.Form.Close()
        $script:powershell.EndInvoke($script:handle)
        #$script:powershell.Close()
        $script:powershell.Dispose()
        $script:runspace.Dispose()
        $script:runspace.Close()
        [System.GC]::Collect()
        Exit
    }


# SIG # Begin signature block
# MIIm6AYJKoZIhvcNAQcCoIIm2TCCJtUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU9DHSgDt/qPw6czYfFoIe4FRG
# ncCggh/4MIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5n
# IFJvb3QgUjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIE
# JHQu/xYjApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7
# fbu2ir29BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGr
# YbNzszwLDO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTH
# qi0Eq8Nq6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv
# 64IplXCN/7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2J
# mRCxrds+LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0P
# OM1nqFOI+rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXy
# bGWfv1VbHJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyhe
# Be6QTHrnxvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXyc
# uu7D1fkKdvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7id
# FT/+IAx1yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQY
# MBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJw
# IDaRXBeF5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUE
# DDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmlj
# YXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3Sa
# mES4aUa1qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+
# BtlcY2fUQBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8
# ZsBRNraJAlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx
# 2jLsFeSmTD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyo
# XZ3JHFuu2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p
# 1FiAhORFe1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG
# 9w0BAQwFADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS0wKwYDVQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIB
# igKCAYEAmyudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxD
# eEDIArCS2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk
# 9vT0k2oWJMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7Xw
# iunD7mBxNtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ
# 0arWZVeffvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZX
# nYvZQgWx/SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+t
# AfiWu01TPhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvr
# n35XGf2RPaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn
# 3UayWW9bAgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaR
# XBeF5jAdBgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYD
# VR0gBBQwEjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RS
# NDYuY3JsMHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEM
# BQADggIBAAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXK
# ZDk8+Y1LoNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWk
# vfPkKaAQsiqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3d
# MapandPfYgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwF
# kvjFV3jS49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZa
# PATHvNIzt+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8b
# kinLrYrKpii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7Ew
# oIJB0kak6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TW
# SenLbjBQUGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg
# 51Tbnio1lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoU
# KD85gnJ+t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGaDCCBNCg
# AwIBAgIRAL/9KI6HeeUhYmELSlorbnEwDQYJKoZIhvcNAQEMBQAwVDELMAkGA1UE
# BhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGln
# byBQdWJsaWMgQ29kZSBTaWduaW5nIENBIFIzNjAeFw0yMTA5MTAwMDAwMDBaFw0y
# MjA5MTAyMzU5NTlaMEYxCzAJBgNVBAYTAkFVMREwDwYDVQQIDAhWaWN0b3JpYTER
# MA8GA1UECgwIQW5nbGUgSVQxETAPBgNVBAMMCEFuZ2xlIElUMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAxR2UqErPaN9XXcwEVPrABibDn9uIA+/GbnZU
# boNwvzTyFoIVaUw5/8QO/6s8j6i8TVC9szyXbxBVGpxMgKg++RZ5t4pkZJWtDU9l
# XrHErNitHDT4WMXJj5ihPM2AgkBSJUsYiSbVDuNmwV9x2nizLY3212rJeMYsVYw5
# qQ/UUNIlDx+CYogVc6esFf6gnhnf7UMlJZqDdxV/AidKtabQLoRrQK6UqGbA9CUQ
# YoJECrDQ7bRsGlByWgdOOQHXDtzDgfA3NYculKbIrm63OdIxdhE5lBSuBfc7XQqB
# l+1rPp4loYipXKFyXVolnFlSovL1LggkevQi4g1yr0aVzqAxC/NNtG9JMw/U9e4F
# rKIYTtvAbVYdG5KIx3sqteHJ56N6F1b4SpLHyB+p5zMmuIwi9gPzhCXrzlKO4GmE
# zPoVKmZhYezd1gkhlMjhlT3DkalZQIVYTYTTNgbsbuCuu7Rv6P63ioRaTfku/fyn
# QDxysZEnN8ZVzjsI14cNwGE5ILESYl831o/q5m/diTOWn3ZrC5GV4hGcVQgnkY6r
# jEiY7uo8J/ybPENDuWpyPUrd4APKHTW3w9jtkaI1HapNBQeJKyEYErTXyoVYiRjx
# NHJmbFjuISPgx9FVrYeMc5KAPDjAl8XL+RfOcjv1COlVupCeuoIhOkjt7mvo8nsk
# THuPvdECAwEAAaOCAcEwggG9MB8GA1UdIwQYMBaAFA8qyyCHKLjsb0iuK1SmKaoX
# pM0MMB0GA1UdDgQWBBS7Ym2QqvITo7sv+Bg8ny5yRnwZLTAOBgNVHQ8BAf8EBAMC
# B4AwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzARBglghkgBhvhC
# AQEEBAMCBBAwSgYDVR0gBEMwQTA1BgwrBgEEAbIxAQIBAwIwJTAjBggrBgEFBQcC
# ARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwCAYGZ4EMAQQBMEkGA1UdHwRCMEAw
# PqA8oDqGOGh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVT
# aWduaW5nQ0FSMzYuY3JsMHkGCCsGAQUFBwEBBG0wazBEBggrBgEFBQcwAoY4aHR0
# cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdDQVIz
# Ni5jcnQwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMCMGA1Ud
# EQQcMBqBGGRldmVsb3BlckBhbmdsZWl0LmNvbS5hdTANBgkqhkiG9w0BAQwFAAOC
# AYEAUBEh0nzKQFwac0oHXEFXPdsnuygRXw7R8Q42vmN2Gney52aRRVyrQgIR7SFt
# YH2vLLDvL6U6YxAQ2PEx6unk4ng/uoS8wlZ41Dv1uHRIMEbfQ3BBPJ97aaot63LF
# +8J3dxD6lZgXsanrDQBX3fkRXo/q8E3RonsHwkMzcGskE6wIgfZj+7Qe9l2cyWBj
# Vt6TAbi31/XUf9R3Xj4CtaTwklXs9XBVkknXKVhV/3aowyVpGILQS/4Ifvu/B0v+
# KmjFP5fvBVQ4CEIFvxWnWDahwpplxyk8ILt43MMomiw306TCCfo3hO6PUDiar7Ew
# JjWP2CFfuc3eYs1bevzxHgG58MCcl5AvP6WFuD6LvMA8D0Uz0HYlXbwXbxMswGmu
# jcUNH6QY8NHws/y4KyyWYVEhme/AXLLyR94sS16EzDeez/BfHuj6LiG46YvAqiAB
# GQXn9P2jZzJ8uDioQA4+BnUTZML7p104jAoeZktK1LYoxx1sJQvxr9vemd79Zai5
# edhmMIIG7DCCBNSgAwIBAgIQMA9vrN1mmHR8qUY2p3gtuTANBgkqhkiG9w0BAQwF
# ADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCk5ldyBKZXJzZXkxFDASBgNVBAcT
# C0plcnNleSBDaXR5MR4wHAYDVQQKExVUaGUgVVNFUlRSVVNUIE5ldHdvcmsxLjAs
# BgNVBAMTJVVTRVJUcnVzdCBSU0EgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkwHhcN
# MTkwNTAyMDAwMDAwWhcNMzgwMTE4MjM1OTU5WjB9MQswCQYDVQQGEwJHQjEbMBkG
# A1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxJTAjBgNVBAMTHFNlY3RpZ28gUlNBIFRpbWUg
# U3RhbXBpbmcgQ0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDIGwGv
# 2Sx+iJl9AZg/IJC9nIAhVJO5z6A+U++zWsB21hoEpc5Hg7XrxMxJNMvzRWW5+adk
# FiYJ+9UyUnkuyWPCE5u2hj8BBZJmbyGr1XEQeYf0RirNxFrJ29ddSU1yVg/cyeNT
# mDoqHvzOWEnTv/M5u7mkI0Ks0BXDf56iXNc48RaycNOjxN+zxXKsLgp3/A2UUrf8
# H5VzJD0BKLwPDU+zkQGObp0ndVXRFzs0IXuXAZSvf4DP0REKV4TJf1bgvUacgr6U
# nb+0ILBgfrhN9Q0/29DqhYyKVnHRLZRMyIw80xSinL0m/9NTIMdgaZtYClT0Bef9
# Maz5yIUXx7gpGaQpL0bj3duRX58/Nj4OMGcrRrc1r5a+2kxgzKi7nw0U1BjEMJh0
# giHPYla1IXMSHv2qyghYh3ekFesZVf/QOVQtJu5FGjpvzdeE8NfwKMVPZIMC1Pvi
# 3vG8Aij0bdonigbSlofe6GsO8Ft96XZpkyAcSpcsdxkrk5WYnJee647BeFbGRCXf
# BhKaBi2fA179g6JTZ8qx+o2hZMmIklnLqEbAyfKm/31X2xJ2+opBJNQb/HKlFKLU
# rUMcpEmLQTkUAx4p+hulIq6lw02C0I3aa7fb9xhAV3PwcaP7Sn1FNsH3jYL6uckN
# U4B9+rY5WDLvbxhQiddPnTO9GrWdod6VQXqngwIDAQABo4IBWjCCAVYwHwYDVR0j
# BBgwFoAUU3m/WqorSs9UgOHYm8Cd8rIDZsswHQYDVR0OBBYEFBqh+GEZIA/DQXdF
# KI7RNV8GEgRVMA4GA1UdDwEB/wQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEAMBMG
# A1UdJQQMMAoGCCsGAQUFBwMIMBEGA1UdIAQKMAgwBgYEVR0gADBQBgNVHR8ESTBH
# MEWgQ6BBhj9odHRwOi8vY3JsLnVzZXJ0cnVzdC5jb20vVVNFUlRydXN0UlNBQ2Vy
# dGlmaWNhdGlvbkF1dGhvcml0eS5jcmwwdgYIKwYBBQUHAQEEajBoMD8GCCsGAQUF
# BzAChjNodHRwOi8vY3J0LnVzZXJ0cnVzdC5jb20vVVNFUlRydXN0UlNBQWRkVHJ1
# c3RDQS5jcnQwJQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20w
# DQYJKoZIhvcNAQEMBQADggIBAG1UgaUzXRbhtVOBkXXfA3oyCy0lhBGysNsqfSoF
# 9bw7J/RaoLlJWZApbGHLtVDb4n35nwDvQMOt0+LkVvlYQc/xQuUQff+wdB+PxlwJ
# +TNe6qAcJlhc87QRD9XVw+K81Vh4v0h24URnbY+wQxAPjeT5OGK/EwHFhaNMxcyy
# UzCVpNb0llYIuM1cfwGWvnJSajtCN3wWeDmTk5SbsdyybUFtZ83Jb5A9f0VywRsj
# 1sJVhGbks8VmBvbz1kteraMrQoohkv6ob1olcGKBc2NeoLvY3NdK0z2vgwY4Eh0k
# hy3k/ALWPncEvAQ2ted3y5wujSMYuaPCRx3wXdahc1cFaJqnyTdlHb7qvNhCg0MF
# pYumCf/RoZSmTqo9CfUFbLfSZFrYKiLCS53xOV5M3kg9mzSWmglfjv33sVKRzj+J
# 9hyhtal1H3G/W0NdZT1QgW6r8NDT/LKzH7aZlib0PHmLXGTMze4nmuWgwAxyh8Fu
# TVrTHurwROYybxzrF06Uw3hlIDsPQaof6aFBnf6xuKBlKjTg3qj5PObBMLvAoGMs
# /FwWAKjQxH/qEZ0eBsambTJdtDgJK0kHqv3sMNrxpy/Pt/360KOE2See+wFmd7lW
# EOEgbsausfm2usg1XTN2jvF8IAwqd661ogKGuinutFoAsYyr4/kKyVRd1LlqdJ69
# SK6YMIIHBzCCBO+gAwIBAgIRAIx3oACP9NGwxj2fOkiDjWswDQYJKoZIhvcNAQEM
# BQAwfTELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQ
# MA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSUwIwYD
# VQQDExxTZWN0aWdvIFJTQSBUaW1lIFN0YW1waW5nIENBMB4XDTIwMTAyMzAwMDAw
# MFoXDTMyMDEyMjIzNTk1OVowgYQxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVh
# dGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEsMCoGA1UEAwwjU2VjdGlnbyBSU0EgVGltZSBTdGFtcGluZyBT
# aWduZXIgIzIwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCRh0ssi8Hx
# HqCe0wfGAcpSsL55eV0JZgYtLzV9u8D7J9pCalkbJUzq70DWmn4yyGqBfbRcPlYQ
# gTU6IjaM+/ggKYesdNAbYrw/ZIcCX+/FgO8GHNxeTpOHuJreTAdOhcxwxQ177MPZ
# 45fpyxnbVkVs7ksgbMk+bP3wm/Eo+JGZqvxawZqCIDq37+fWuCVJwjkbh4E5y8O3
# Os2fUAQfGpmkgAJNHQWoVdNtUoCD5m5IpV/BiVhgiu/xrM2HYxiOdMuEh0FpY4G8
# 9h+qfNfBQc6tq3aLIIDULZUHjcf1CxcemuXWmWlRx06mnSlv53mTDTJjU67MximK
# IMFgxvICLMT5yCLf+SeCoYNRwrzJghohhLKXvNSvRByWgiKVKoVUrvH9Pkl0dPyO
# rj+lcvTDWgGqUKWLdpUbZuvv2t+ULtka60wnfUwF9/gjXcRXyCYFevyBI19UCTgq
# YtWqyt/tz1OrH/ZEnNWZWcVWZFv3jlIPZvyYP0QGE2Ru6eEVYFClsezPuOjJC77F
# hPfdCp3avClsPVbtv3hntlvIXhQcua+ELXei9zmVN29OfxzGPATWMcV+7z3oUX5x
# rSR0Gyzc+Xyq78J2SWhi1Yv1A9++fY4PNnVGW5N2xIPugr4srjcS8bxWw+StQ8O3
# ZpZelDL6oPariVD6zqDzCIEa0USnzPe4MQIDAQABo4IBeDCCAXQwHwYDVR0jBBgw
# FoAUGqH4YRkgD8NBd0UojtE1XwYSBFUwHQYDVR0OBBYEFGl1N3u7nTVCTr9X05rb
# nwHRrt7QMA4GA1UdDwEB/wQEAwIGwDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMEAGA1UdIAQ5MDcwNQYMKwYBBAGyMQECAQMIMCUwIwYIKwYB
# BQUHAgEWF2h0dHBzOi8vc2VjdGlnby5jb20vQ1BTMEQGA1UdHwQ9MDswOaA3oDWG
# M2h0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1JTQVRpbWVTdGFtcGluZ0NB
# LmNybDB0BggrBgEFBQcBAQRoMGYwPwYIKwYBBQUHMAKGM2h0dHA6Ly9jcnQuc2Vj
# dGlnby5jb20vU2VjdGlnb1JTQVRpbWVTdGFtcGluZ0NBLmNydDAjBggrBgEFBQcw
# AYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBAEoD
# eJBCM+x7GoMJNjOYVbudQAYwa0Vq8ZQOGVD/WyVeO+E5xFu66ZWQNze93/tk7OWC
# t5XMV1VwS070qIfdIoWmV7u4ISfUoCoxlIoHIZ6Kvaca9QIVy0RQmYzsProDd6aC
# ApDCLpOpviE0dWO54C0PzwE3y42i+rhamq6hep4TkxlVjwmQLt/qiBcW62nW4SW9
# RQiXgNdUIChPynuzs6XSALBgNGXE48XDpeS6hap6adt1pD55aJo2i0OuNtRhcjwO
# hWINoF5w22QvAcfBoccklKOyPG6yXqLQ+qjRuCUcFubA1X9oGsRlKTUqLYi86q50
# 1oLnwIi44U948FzKwEBcwp/VMhws2jysNvcGUpqjQDAXsCkWmcmqt4hJ9+gLJTO1
# P22vn18KVt8SscPuzpF36CAT6Vwkx+pEC0rmE4QcTesNtbiGoDCni6GftCzMwBYj
# yZHlQgNLgM7kTeYqAT7AXoWgJKEXQNXb2+eYEKTx6hkbgFT6R4nomIGpdcAO39Bo
# lHmhoJ6OtrdCZsvZ2WsvTdjePjIeIOTsnE1CjZ3HM5mCN0TUJikmQI54L7nu+i/x
# 8Y/+ULh43RSW3hwOcLAqhWqxbGjpKuQQK24h/dN8nTfkKgbWw/HXaONPB3mBCBP+
# smRe6bE85tB4I7IJLOImYr87qZdRzMdEMoGyr8/fMYIGWjCCBlYCAQEwaTBUMQsw
# CQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJT
# ZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgQ0EgUjM2AhEAv/0ojod55SFiYQtK
# WitucTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUpvFTGXI+9egSxhO/0ARcifiBpYAwDQYJKoZI
# hvcNAQEBBQAEggIAdqFCigjSPJR3wObf/5QIsDQnpsh1gTQUaYZXVrgTjOA5ntyj
# 9DjNs8Od7OcfeFSFmCGdUz79SMkWJEnfAzH6UWWw5sfbZOe600Oohizz8dVLiNRC
# 0dMsPUHYBKlvRq7M68xzXsFyDvlDbfezTcQYjTPqeTN13ToxULww+uXoNfCgaiJo
# GFwkxboPQ1iBK9vb06Tso3sLe6Ct41nH743QuZ/ge6YfI5mW3gUFNIzTyAXvqcNS
# TQHmIQkN08ao6SYEmBSjtBu0lhGFUkLAJm8Y2Y/QCirI5SDJhbPlQDyRgNO6Ru8/
# nmleZ1DJRD1QakK3LJPQX6gPQBuECHbFORWyXkUTHBlJOdmrL35U65CdCYKXhySC
# awA33rT9XIGe8EF1bzNeMI/0e2VVGbqhvhAeLURdswgcCATFTpefuT6wlsU5Alct
# tdBvmW3mmKzlykA3dLl/Uf9eETOb47xlIvCQrwY3VsUto2oAxkBmavlvFnSg+eYg
# ugysyqs535QoF93oGUpU1fo87qovwzUWftqirmSxirXWwDtSKRSdw49mVimKHU6W
# om0JP1GL1E7jNmyNVRZ/8dCfUYfyT5UTirBzOj8z2xyxkN6kNWdAcWyHIMlEARML
# LofHUaAzOuT1CZx9EeuwZjIxeMOXARYfR7fRgnBbSBMHsvUlzK77wk8NMNOhggNM
# MIIDSAYJKoZIhvcNAQkGMYIDOTCCAzUCAQEwgZIwfTELMAkGA1UEBhMCR0IxGzAZ
# BgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYG
# A1UEChMPU2VjdGlnbyBMaW1pdGVkMSUwIwYDVQQDExxTZWN0aWdvIFJTQSBUaW1l
# IFN0YW1waW5nIENBAhEAjHegAI/00bDGPZ86SIONazANBglghkgBZQMEAgIFAKB5
# MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIxMTEw
# MTAzNDc0NlowPwYJKoZIhvcNAQkEMTIEMGJ9rIkThH1IMN7ygHkeHmAm42FvAeDn
# AXE7MbIr8iEG8QkIXTY79uM0G4DeUqFCoTANBgkqhkiG9w0BAQEFAASCAgBn8X1X
# e4MvGSVlUoPbBEX8XluKgf2i2KXJva3ALw2RBzxhcCohwkTcTKZlzQAh+6u/15t5
# DxYc7lI3O8KyHYTap+egZqJN4LhSo0Ps0y1FE8VtdmyAvRXLkA06NykpqpfMhMjp
# khruZqrZo82wywIBQbgjnQWHrpMa21vr+7deI1mrF5RmfCFY3+/oyQWN8klfC7Yk
# BPYHxEWetuBsI+2iRGrBdHdFEEhJMeqim9AWn0vB60BAbCk74XMuxUt3aweyVaOD
# Isip/IPsUssWXWgB6my1Zok3mgccMO0/805Tu4vxHdr4RBKSLG0o6MZB+wRmrCD7
# QFgsfipUp8dd8Vwe8+xAe7q2H7o5pHZRGevIpSsw4MUDqPyBEXZQpGvTnZyJtV+k
# C5DQhGSzjpxR2WIZ49+8v9gXQdPGS24HkAvjSalUjXTFJOyly5D47npS1sjACzJ2
# rYNT8Pp/rFam38tDZPkLkbG3QhqjDM1LsGzPbxIeWvHyjVHM2b/LKwQhYRMVKbUb
# 46Z9BX22bH3uZRJeTs0yqGV0rsnZzCZuzF//gIW7AVDqHKyVa2ORLiwP0eUs6BBO
# KQYO5WLJdWPr+e6RGwfnqf9MlHXLpByKYuS9v132X+YCahYn/xAV7Xkcn8uR0Ud2
# QSR8FtSDXDOcgu2g1xihDXdSF/UzR/faTGlykg==
# SIG # End signature block
