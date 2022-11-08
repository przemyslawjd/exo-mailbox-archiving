# Exchange Online - Mailbox in-place archiving setup

A simple powershell script to enable and check status the in-place archiving of the selected exchange online mailbox.

The script allows:</br>
	• Enabling archive for selected mailbox</br>
	• List and assign existing retention policy</br>
	• Disable retention hold settings</br>
	• Enable Managed Folder Assistant</br>
	• Show information about mailbox (total item size, archiving status and progress)</br>
  
Exchange Online PowerShell module is required.  https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.0.0 </br>
You must be assigned the Mail Recipients role in Exchange Online to enable or disable archive mailboxes. 
