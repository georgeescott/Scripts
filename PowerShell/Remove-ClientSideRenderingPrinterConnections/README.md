# Remove-ClientSideRenderingPrinterConnections.ps1
A script to remove cached client side rendering printer connections from the registry to fix 'Group Policy Printers' Event ID 4098 error code `0x80070057`.


## More Information
'Group Policy Printers' Event ID 4098 error: `The user '<printer name>' preference item in the '<GPO name> {00000000-0000-0000-0000-000000000000}' Group Policy Object did not apply because it failed with error code '0x80070057 The parameter is incorrect.' This error was suppressed.` occurs when mapping a shared printer via GPO where the '**Render print jobs on client computers**' (otherwise known as Client Side Rendering) setting has been configured on the shared printer.

You can check for the presence of these errors on a computer by running the following PowerShell command:
``` PowerShell
Get-EventLog -LogName Application -Source "Group Policy Printers" -Newest 5 | Where-Object {$_.EntryType -eq "Warning" -or $_.EntryType -eq "Error"} | FL TimeGenerated,EventID,Source,EntryType,Message
```

This is caused when domain user profiles are cleared by either the '[Delete user profiles older than a specified number of days on system restart](https://admx.help/?Category=Windows_10_2016&Policy=Microsoft.Policies.UserProfiles::CleanupProfiles)' GPO setting, or by other means such as [DelProf2](https://helgeklein.com/free-tools/delprof2-user-profile-deletion-tool/) or scripts.

These user profile cleanup methods do not clear out the cached user SID's from the `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider` key properly and the Client Side Rendering Print Provider does not clean these up by default either. This leaves a cached/orphaned user SID key which the 'Client Side Rendering Print Provider' can't correctly handle during the print mapping process, resulting in the `0x80070057 The parameter is incorrect` error.

#### Related reading
- https://www.edugeek.net/forums/windows/160941-group-policy-preferences-printer-mapping-issues.html
- https://www.edugeek.net/forums/windows-10/191185-disappearing-printers.html
- https://www.edugeek.net/forums/windows-10/183649-windows-10-printers-keep-coming-back.html
- https://serverfault.com/questions/1082240/where-are-these-printers-coming-from-in-devices-and-printers
- https://social.technet.microsoft.com/Forums/lync/en-US/71d06204-3735-4473-8bc9-20be9e19090e/problem-with-multiple-instances-of-shared-printers-being-installed-on-client-computers-when-the

### Partial workaround - `RemovePrintersAtLogoff`
You can set the following registry key to partially workaround this issue, which will instruct the Client Side Rendering Print Provider to clean up the necessary cached registry keys at logoff:
``` Windows Registry Entries
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider]
"RemovePrintersAtLogoff"=dword:00000001
```

As long as the Client Side Rendering Print Provider correctly cleaned-up the registry on the last logoff when `RemovePrintersAtLogoff` is set, the printer seems to successfuly map the next time the user initiates a printer mapping e.g. during login.
However, this key doesn't seem to cleanup existing cached user SID keys in the registry. So in the event of failed logoffs, power cuts, system crashes etc., these keys are not cleaned-up by the `RemovePrintersAtLogoff` setting.

I've observed the effects of setting the `RemovePrintersAtLogoff` registry key on `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider`. The following keys are removed at logoff:
- `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\S-1-5-21-*`
- `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers\<name>\Printers\*`
- `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers\<name>\Monitors\Client Side Port\*`

### The fix
To resolve this bug with the Client Side Rendering Print Provider, **Remove-ClientSideRenderingPrinterConnections.ps1** replicates the actions of the `RemovePrintersAtLogoff` key to cleanup existing cached user SID's from the registry.

You will still need to have the `RemovePrintersAtLogoff` key set to cleanup the registry on logoff, which this script sets by default.


## Installation
This script will need to be run on any computer that will be mapping a shared printer that has Client Side Rendering enabled, and also utilises any user profile cleanup method e.g. GPO, Delprof2 etc.

It should be run with administrator privileges as it sets/removes registry keys. If you run it as a GPO Computer Policy startup/shutdown script, it will run as SYSTEM.

It should be run when no users are logged in. Ideally at either startup, shutdown or as a scheduled task.

It *shouldn't* require the restart of the 'Print Spooler' service as it's replicating existing functionality that also doesn't require this.

### Group Policy
#### 1. Run the script at startup or shutdown
1. Open **Group Policy Management Editor** and navigate to **Computer Configuration\Policies\Windows Settings\Scripts**
2. Navigate to **Startup** or **Shutdown**
3. Go to the **PowerShell Scripts** tab
4. Click Add...
    - **Script Name**: \<path to script\>
    - **Script Parameters**: \<none\>
5. Click OK

#### 2. RemovePrintersAtLogoff Registry Key (optional)
This script sets the `RemovePrintersAtLogoff` registry key by default.

But if you'd prefer to set it via GPO instead, you can comment out the `Set-RemovePrintersAtLogoff` function call near the bottom of the script and set it via Group Policy Preferences:

1. Open **Group Policy Management Editor** and navigate to **Computer Configuration\Preferences\Windows Settings\Registry**
2. Right-click -> New -> Registry Item
    - **Action**: Update
    - **Hive**: `HKEY_LOCAL_MACHINE`
    - **Key Path**: `SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider`
    - **Value name**: `RemovePrintersAtLogoff`
    - **Value type**: `REG_DWORD`
    - **Value data**: `1`
    - **Base**: Decimal
3. Click OK

## Troubleshooting
This script outputs a transcript log to `C:\Windows\Logs\`.


## Issues
If you have issues using the script, open an issue on the repository.

You can do this by clicking "Issues" at the top and clicking "New Issue" on the following page.

Be sure to mention the script name in the issue report.