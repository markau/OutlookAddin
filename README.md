## Overview

This is a code example that forwards the selected Outlook message to the specified address, as an attachment, and then deletes the selected message. 

It has been developed as a learning exercise for better understanding development of Office applications using Visual Studio 2013 and .NET (Outlook Primary Interop assembly), while providing the practical benefit of a tool to handle spam email in a corporate environment.

## What it does

The specific process the application implements is:

* Create a 'report spam' button within the ribbon toolbar
* When clicked, iterate through the messages (in the currently selected folder) and find the selected message
* Build a new Outlook message, and attach the selected message as an object
* Send the email
* Delete the selected message

The code sample includes a function to ensure the email sent is from the default mailbox, where users manage multiple mailboxes.

