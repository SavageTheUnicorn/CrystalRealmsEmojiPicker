This script allows Crystal Realms users to easily copy emojis to their clipboard without the need to open a seperate browser window. Ofcourse, this is just the source code; I will provide the exe via dm's to friends if they want it.
Original EXE file hashes "CrystalRealms-EmojiPicker.exe";
MD5: 81247a288f077b5510cf0308659307b8
SHA1: d2935040c97df422e83706f7d774e7e52b6d59e6
SHA256: 6d7814c63faf407962b74b2d836a89a729cf8d3fd99f677be98815d6f6d3ae1a
SHA512: e8bafb1fd3ed446a94c61018a7554d520593e8e9ce7076a9b29922a296cd17ba2f4e30dafbcf166661b768b14ccedbe3c7422e96bd2d08185fcf97d8442840fd 
Use these hashes to verify that you have the application created by me, this way you can ensure safety. I will not provide support to anyone who obtains the executable/source code from an unknown source as anyone can add malware to a project, even the creator.
The command I used to build this from the source code into executable format is "pyinstaller --onefile --add-data "NotoEmoji-Regular.ttf;." --add-data "app.ico;." --icon=app.ico --noconsole --name CrystalRealms-EmojiPicker crystalrealmsemojipicker.py". If you intend to build from source code,
you will need to install the proper dependencies and you will also need to include the app.ico and the NotoEmoji-Regular.ttf font file in the same directory as the source code python file. Your safety is your own responsibility.
NOTE: This application has basic Windows Tray Icon functionality, so you need to close that too to fully close the application.
