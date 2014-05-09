## Overview

This is a demo of using the [Outlook Primary Interop assembly](http://msdn.microsoft.com/en-us/library/office/bb652780(v=office.15).aspx) to create an Add-in for MS Outlook.  It is built from a Visual Studio 2013 template. 

The application creates a new button in the toolbar which, when clicked, forwards the selected message (as an attachment) to a pre-defined address and then deletes the message.  This automates and abstracts the process of reporting spam in our corporate environment.  

With it's simplistic functionality it will serve as a useful quick start for developing future add-ins.

## What it does

Specifically, the Add-in:

* Creates a 'report spam' button within the ribbon toolbar
* When clicked, iterates through the messages (in the currently selected folder) and finds the message that is selected 
* Instantiates a new Outlook message object, and attaches the selected message as an object
* Sends the email
* Deletes the selected message

The code sample includes a function to ensure the email sent is from the default mailbox (in case the user manage multiple mailboxes).

For privacy reasons, it could probably do with a prompt to confirm that the user wants to submit the message they have selected.

