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

1. Create an Automator application and make it run script `Outlook To Things Helper.scpt` (paste script content into Automaton's `Run AppleScript` action) & save it as an application
2. Open application's `info.plist` file (right-click on the app and choose "Show Package Contents". In the `<dic>` section add below information to it:

  ```xml
  <key>CFBundleURLTypes</key>
  <array>
      <dict>
          <key>CFBundleURLName</key>
          <string>Local File</string>
          <key>CFBundleURLSchemes</key>
          <array>
              <string>local</string>
          </array>
      </dict>
  </array>
  ```

3. Move your application to `/Applications/` folder (must move it from somewhere to `/Applications/` folder to make Mac OS X recognizes our custom URL scheme)

After above steps, you can run `Outlook To Things.scpt` script and it will show a dialog which allows you to create a new Things task from the selected message. Also clicking on the message link in task will open corresponding message in Outlook.

Recommendation
--------------
In my setup, I use Keyboard Maestro to define a shortcut key to invoke corresponding script. 

Credits
=======
Inspired by this post by LuMe96 on Cultured Code forum:

https://culturedcode.com/forums/read.php?7,48866,page=1

License
=======
[MIT](https://github.com/phuna/EmailsToThings/blob/master/LICENSE)
