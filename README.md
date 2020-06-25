# EmailAutomation1
Create a workflow that reads the Inbox and sorts the emails by moving them to various folders, 
based on “rules” defined in an Excel similar to the one in the attachment.

If the sender address contains the String defined in the Sender column, move the mail to the 
email folder defined in the Folder column.

V.Imp. Note: Make sure you Run your Outlook As an Administrator.
-----------------------------------------------------------------

In UiPath Studio, create a new project and perform following steps:

-- Depending on what mail server / client you are using, the properties due to be set might be different, 
however the logic and activities would be similar. When using in this walkthrough Get Outlook Mail Mail 
Messages, it refers mostly to Get XXX Mail Messages, which XXX could be POP3, IMAP, Exchange, Outlook.

    -- We will start by dragging an activity Get Outlook Mail Messages, we configure which account to read 
    from (if we need one which is not the default account in Outlook), the folder to read from and we create 
    a variable in which the emails will be output. Let’s call this variable, emails.

    -- We will also use a Read Range activity to read the “rules” from the email rules.xlsx file, provided 
    as input. We will Output the result into a DataTable which we will call mailRules.

    -- Notice that the emails variable, created by us in the output property of Get Outlook Mail Message, 
    is of type List<MailMessage> .

    -- As we are trying to apply the rules to each email, we will create an iteration on this emails list 
    by using a For each activity. Considering the list we are iterating consists of MailMessage objects, 
    each item from the iteration will be of type MailMessage and we need to specify this explicitly.

    -- We click on the For each activity and in the property called TypeArgument, we choose from the options 
    in the list, the value System.Net.MailMessage.

    -- If this option is not on the list, please click on Browse for Types and search for the System.Net.Mail.MailMessage 
    in the appearing dialog window.

    -- Now we need to iterate on each of the rules and see if the current mail we are analysing meets any of the 
    criteria. This time we have a DataTable object so we will iterate through it much more comfortably with a 
    For each row activity. The only thing you need to specify is the object to iterate on, which in our case is mailRules.

    -- We now need to check if the rules on the subject apply. As we have access to the current email in the item 
    (we renamed it to mail) variable created automatically and this variable is of type MailMessage, we can access all 
    specific mail properties. We have to see if the Sender.Address property contains the column “From” from the current 
    rule being checked.

    -- If that is true, we use a Move Outlook Mail Message and we move the current mail to the folder that is specified 
    in the current row on column “Folder”.

    -- We also need to add a break to this sequence, in order not to verify the other rules anymore, if one was already 
    applied and the mail was moved to another folder. 

