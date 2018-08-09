# Set-MailboxPermissions
A script to manage mailbox permissions when using security groups

# Description
This script manages mailbox delegation for group mailboxes by granting the group
"Send As"and"Send on Behalf" rights to the mailbox.  It then takes the members of
the group and add them individually to the mailbox so the "AutoMapping" property
automatically assigns the mailbox to the user in Outlook.

# Usage
    Set-MailboxPermissions -Group <Security Group> -Mailbox <Group Mailbox>
