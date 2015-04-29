#EmailsToThings
Mac OSX Apple scripts and helper application for converting Outlook and Mail.app's messages to Things' tasks, also link Things' tasks back to the corresponding message

##How to use

### Use Mail.app
Just run `Mail To Things.scpt` script, it will show a dialog which allows you to create a new Things task from the selected message.

*Note:* When using conversation mode, if you click on a conversation, it most of the time selects the latest message, but in some of my tests it selects a random message in that conversation. So you should expand the conversation and select the message you want to create task for, or another way is to not use conversation mode.

###Use Outlook 2011
**New Instruction**

1. Extract `Outlook To Things Helper.app.zip` then move the `Outlook To Things Helper.app` to `Applications` folder. This application is used for showing Outlook message when clicking on the link inside Things' task
2. Run `Outlook To Things.scpt`, it will show a dialog which allows you to create a new Things task from the selected message.

**Original instruction, for those want to create `Outlook To Things Helper.app` by hand**

*Note*

For Mac OSX 10.10, the instructions below does not work, I don't know why yet, although the helper application seems to be invoked because a gear is shortly shown on menu, but it Outlook is not activated and message is not shown. On Mac OSX 10.9 the instructions below should work. Infact  my `Outlook To Things Helper.app` is built using below instructions on Mac OSX 10.9 and it can work on both 10.9 and 10.10

*Instructions*

Because Outlook 2011 doesn't support custom URL scheme for its message, we need to implement one by our own so we can click on the link embeded in Things' task to open corresponding message in Outlook. Steps to do are:

1. Create an Automator application and make it run script `Outlook To Things Helper.scpt` (paste script content into Automaton's `Run AppleScript` action) & save it as an application
2. Open application's `info.plist` file (right-click on the app and choose "Show Package Contents". In the `<dic>` section add below information to it:

  ```xml
<key>CFBundleURLTypes</key>
<array>
    <dict>
        <key>CFBundleURLName</key>
        <string>Outlook To Things Helper</string>
        <key>CFBundleURLSchemes</key>
        <array>
            <string>outlooktothings</string>
        </array>
    </dict>
</array>
  ```

3. Move your application to `/Applications/` folder (must move it from somewhere to `/Applications/` folder to make Mac OS X recognizes our custom URL scheme)

After above steps, you can run `Outlook To Things.scpt` script and it will show a dialog which allows you to create a new Things task from the selected message. Also clicking on the message link in task will open the corresponding message in Outlook.

###Recommendation
In my setup, I use Keyboard Maestro to define a shortcut key to invoke `Mail To Things.scpt` and `Outlook To Things.scpt`. Therefore I only need to press `Cmd+E` in Outlook or `Cmd+T` in Mail then the script will run and show the dialog to add the selected message to Things. Very handy.

Credits
=======
Inspired by this post by LuMe96 on Cultured Code forum:

https://culturedcode.com/forums/read.php?7,48866,page=1

License
=======
[MIT](https://github.com/phuna/EmailsToThings/blob/master/LICENSE)
