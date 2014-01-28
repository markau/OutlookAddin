## Overview

This is a code example that forwards the selected Outlook message to the specified address, as an attachment, and then deletes the selected message.

It has been developed as a learning exercise with a practical benefit of handling spam in a corporate environment.  Some of the code is from existing examples out in the wild.

## What it does

The Outlook Primary Interop assembly allows the addin to act as native Outlook code. So, properties such as the 'from' address and the outgoing mail server all use the Outlook defaults.

The specific steps the application implements:

* Creates a 'report spam' button to the ribbon toolbar
* When clicked, iterates through the messages and finds the selected message
* Builds a new Outlook message, and attaches the selected message as an object
* Sends the email
* Deletes the selected message

The code sample includes (but does not use) a function to ensure the email sent is from the default mailbox.

