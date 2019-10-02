# SentFolderByFrom
Outlook addin that prompts you to choose a "Sent Items" folder for each "From" address you use.

# Background
[*G Suite Sync for Microsoft Outlook* (GSSMO)](https://support.google.com/a/users/answer/153866?hl=en) is Google's tool for connecting your G Suite email account to Outlook.

GSSMO provides support for accessing delegated email accounts via the ["Add Account for Delegation" option](https://support.google.com/a/users/answer/170961?hl=en).

However, items sent from a delegated account will appear in the delegate's primary Sent Items folder, rather than the Sent Items folder for the delegated account.

This add-in works around that issue by moving items to the Sent Items folder you designated for that "From" address.

# Installation
See the [Releases](https://github.com/browlry/SentFolderByFrom/releases) page to donwload and install the latest version.

# Usage
1. Compose a new email message in Outlook.
1. If needed, go to `Options` > `Show Fields` > `From` to unhide the "From" field.
1. Select a From address *other* than your primary email address.
1. Send the email.
1. After the item syncs to your Sent Items folder, a message appears: "Click OK to select the 'Sent Items' folder for items sent from test@example.com". Click "OK".
1. The Select Folder dialog appears. Click the folder where you want items sent from this address to be saved. Click OK.
1. The item is moved to the folder you selected. In the future, items sent from test@example.com will be saved to this folder automatically, with no confirmation messages.

# Oh no, I clicked the wrong folder! How do I get it to ask me for a folder again?
1. Close Outlook.
1. Hit `Windows Key` + `R` on your keyboard.
1. In the Run box, enter `%appata%\browlry\SentFolderByFrom`.
1. Delete `appconfig.bin`.
1. Repeat the steps in the [Usage](#Usage) section.

# Uninstall
Go to "Add or remove programs" in Windows, find "SentFolderByFrom", and click "Uninstall".
