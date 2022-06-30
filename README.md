# TeamsManagement
This PowerShell binary module contains the `Search-TeamsMessage` cmdlet. The cmdlet uses the EWS Managed API to search for Microsoft Teams messages. There are two pre-configuration steps that need to be done before this cmdlet will work.  

    1. Create an Application ID in your tenant and give it `EWS.AccessAsUser.All` permissions for delegated access. 
    2. Copy your ApplicationId and your TenantId and add them to the App.config file.  

Now with the pre-configuration out of the way, you can build the project and import the Microsoft.Exchange.Management.Teams.dll.  

Module Import
```powershell
Import-Module Microsoft.Exchange.Management.Teams.dll
```

In this example, we are looking for a message sent by jsmith@contoso.com to a group chat with cstep@contoso.com where the message contains "MFCMapi"  

Example:
```powershell
Search-TeamsMessage -Sender jsmith@contoso.com -Recipient cstep@contoso.com -MessageContains "MFCMapi"
```

Output:
```powershell
Sender              : jsmith@contoso.com
Recipient           : dstep@contoso.com; ebrown@contoso.com; cjohnson@contoso.com; rmoore@contoso.com; cstep@contoso.com
Message             : I think we can use MFCMapi to view the message and associated properties
CreatedDateTime     : 6/23/2022 6:31:09 PM
HasCardAttachments  : False
ThreadId            : 19:78f1808b81e84cad99c4c057d3a2f0d8@thread.v2
ThreadType          : chat
GroupChatName       : Troubleshooting Chat
```
* _**Note:**  By default we show a subset of the most useful properties. If you want to see all properties, pipe the command to Format-List *_.  

<br />
For the next example, we will be searching by Topic or Group Chat Name.  

Example:
```powershell
Search-TeamsMessage -Topic "Troubleshooting Chat" | FT Sender, Message, CreatedDateTime
```
Output:
```powershell
Sender             Message                                                                   CreatedDateTime
------             -------                                                                   ---------------
JSmith@contoso.com I think we can use MFCMapi to view the message and associated properti... 6/23/2022 6:31:09 PM
cstep@contoso.com  We need to see what messages look like inside of a mailbox....            6/23/2022 5:14:55 PM
cstep@contoso.com  We will be troubleshooting Teams Messaging...                             6/23/2022 5:14:39 PM


```

Our final example will be searching by ThreadId. This is a powerful search because you can get the full thread history even if the group chat name has been changed.  

Example:
```powershell
Search-TeamsMessage -ThreadId 19:78f1808b81e84cad99c4c057d3a2f0d8@thread.v2 | FT Sender, Message, CreatedDateTime
```
Output:
```powershell
Sender             Message                                                                   CreatedDateTime
------             -------                                                                   ---------------
JSmith@contoso.com I think we can use MFCMapi to view the message and associated properti... 6/23/2022 6:31:09 PM
cstep@contoso.com  We need to see what messages look like inside of a mailbox....            6/23/2022 5:14:55 PM
cstep@contoso.com  We will be troubleshooting Teams Messaging...                             6/23/2022 5:14:39 PM
cstep@contoso.com  This is a group chat for troubleshooting....                              6/23/2022 5:13:42 PM


```
You can see in the output above that the ThreadId search found one extra message compaired to the Topic Search. This is because we got the first message before the group chat was named.
