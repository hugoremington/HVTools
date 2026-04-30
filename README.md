# HVTools

**HVTools** (formerly *VMM Tools*) is the swiss-army knife for SCVMM administrators. It is a lightweight application designed to perform a complete SCVMM inventory.

It reports crucial information about your SCVMM environment, including:
* Clusters and Hosts
* Virtual Machines
* Storage Arrays and Storage Pools
* Networks and VLANs
* Workload health and much more

The elegant table views make it easy to use, find gaps, and identify improvements in your SCVMM environment. It is an excellent tool for gaining clearer insight, general troubleshooting, capacity planning, and reporting.

> [!NOTE]  
> This is the Public Edition (PE) of HVTools, as it was a minimum viable product (MvP) back in the day.

---

## Instructions

1. Enter SCVMM administrator privileged credentials in the **Username** and **Password** fields.
2. Enter the SCVMM instance ID address in the third field (e.g., `scvmminstancefqdn.contoso.com`) and click **Login**.
3. When the inventory is complete, the HVTools report will automatically open.
4. You can click `File` $\rightarrow$ `Save As` to save the complete inventory to an **Excel CSV**.
5. **Optional:** By selecting desired rows and performing `CTRL + C`, you can copy/paste data directly into Excel.

> [!NOTE]  
> Designed for **Microsoft System Center Virtual Machine Manager 2016**.

---

---

## To Do

1. Change table column creation to using arrays and foreach loops for elegant creation.
2. HTML reporting.
3. Colourisation (Red-Amber-Green)
4. (Optional) ServiceNow integration using service hooks.

> [!NOTE]  
> There additional features were incorporated into the propietary edition v2.0.0 of HVTools back in 2024.

---

## Changelog

### v1.7.3
- Codesigned compiled file using Sectigo certificate.
- Added VMM server name to title bar form of report.
- Removed empty `var filename`.
- Fixed tab name to **Cluster Storage Volumes**.
- Clipboard copy/paste will now include headers.

### v1.7.2
- Enabled datagrid clipboard `CTRL + C` of selected rows (configured `RunspacePool` apartmentState as `STA`).
- Removed redundant 2nd form `runspace.close` calls.

### v1.7.1
- Fixed exit bug where runspaces would keep the process open; app now exits cleanly.
- Fixed bug where the **Login** button would remain disabled post-inventory completion.
- **New Feature:** Added **Zombie VHDs** tab to report orphaned VHD/VHDX files in SCVMM.

### v1.7.0
- Appended new column `Location` in Virtual Machines table (contains VHD/VHDX path).
- Removed automatic export of Virtual Machines table.

### v1.6.9
- ZIP multi-report saving finalized; added cleanup feature.
- Rounded off all numbers to a single decimal point.
- Added experimental **Save to ZIP** feature. Exports all tables into a temp path before ZIPing to the desired save path via dialog.

### v1.6.8
- **New Feature:** Added **Cluster storage volumes** tab (position 7).

### v1.6.7
- Removed `$script:powershell = [powershell]::Create()` from line 2322.

### v1.6.6
- Cosmetic update: Rounded large capacity numbers to 2 decimal places using `[math]::Round($var,2)`.

### v1.6.5
- Cosmetic update: Added **About/Help** menu in 2nd results form.
- Added reset counter for clusters `$c` in `ClusterNetworks` foreach loop.

### v1.6.4
- Added tab filter for all 8 datagrids.

### v1.6.3
- **Major Update:** Added Clusters, Hosts, Storage Pools, Storage Arrays, Cluster Disks, Networks, and Cluster Networks!

### v1.6.2
- Added **Hosts** feature (including data table, tab, foreach loop, and data grid).
- Re-enabled `maxthreads` to all available processor count on system.
- **New Feature:** Added **Cluster Info** (2nd table for Clusters).

### v1.6.1
- Fixed unprotected memory exceptions by removing `Add-OutputBoxLine` calls within `ForEach` loops (preventing RichTextBox overload).

### v1.6.0
- Added more tables and tabs for comprehensive information including clusters, hosts, storage, and networks.

### v1.5.9
- Attempted to resolve unprotected memory leak by calling `$script:powershell.EndInvoke($script:handle)` at every exit function.
- Changed `$maxthreads` to `3` to resolve unprotected memory exceptions.
- Added extra tabs in preparation for future releases.
- Re-enabled `maxthreads`.
- Updated color scheme.

### v1.5.6
- Attempted to resolve memory leak by casting `$VMS` to an array.
- Fixed VLAN display issue by changing column type from `Int32` to `String`.
- Reduced `$maxthreads` to static `3` to prevent crashes during multi-user RDP sessions.
- Enhanced GUI, streamlined tab view, and changed datagridview color to `moccasin`.
- Fixed 2nd form color.
- Fixed minor bugs regarding table column types (Memory, Dynamic Memory, and vCPU).
- **Major GUI updates:** Converted array to Data Grid View and code-signed the application using Sectigo certificate.

### v1.5.0
- Implemented multi-threading via **Runspace Pools**.
- Added **Out-Grid** GUI view.
- Performance improvements.
- Switched to using Arrays instead of flat memory.
- Added static/dynamic memory optimizations.

---

## Author
**Hugo Remington**