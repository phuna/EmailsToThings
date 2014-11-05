EmailsToThings
==============

Mac OSX Apple scripts for converting Outlook and Mail's messages to Things' tasks

How to use
==========
Use Mail app
------------
Just run `Mail To Things.sept` script, it will show a dialog which allows you to create a new Things task from the selected message.

*Note:* When using conversation mode, if you click on a conversation, it most of the time selects the latest message, but in some of my tests it selects a random message in that conversation. So you should expand the conversation and select the message you want to create task for, or another way is to not use conversation mode.

Use Outlook 2011
----------------
Firstly, we need to be able to open the Outlook message when clicking on the link embeded in Things' task. Because Outlook 2011 doesn't support custom URL scheme for its message, we need to implement one by our own. Steps to do are:

1. Create an Automator application and make it run script `Outlook To Things Helper.scpt` (paste script content into Automaton's `Run AppleScript` action).
2. After creating the application, please move it to `/Applications/` folder (must move it from somewhere to `/Applications/` folder to make Mac OS X recognizes our custom URL scheme

After above steps, you can run 'Outlook To Things.sept' script and it will show a dialog which allows you to create a new Things task from the selected message. Also clicking on the message link in task will open corresponding message in Outlook.

