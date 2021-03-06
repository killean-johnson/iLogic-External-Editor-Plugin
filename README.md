# iLogic-External-Editor-Plugin
This is used to bridge your iLogic rules so that they can be accessed through external editors.

It's been tested with NeoVim through WSL, which has worked very well, and is what it's primarily written for. It has also been tested with pure Vim through WSL, as well as Visual Studio Code, all of which work, but haven't been tested extensively.

<b>MAKE SURE YOUR PROJECT IS CHECKED OUT, OTHERWISE YOU MAY RUN INTO ISSUES</b>

# Setup
Make sure that you have the latest Inventor SDK installed. If you haven't done that, refer to <a href="https://help.autodesk.com/view/INVNTOR/2021/ENU/?guid=GUID-6FD7AA08-1E43-43FC-971B-5F20E56C8846">this link</a>.

Once that's done, make sure you have Visual Studio 2019 installed with both C# and VB ready.

Clone the project into your desired folder

You should be able to build and run the program from the project.

# Using
Make sure the assembly you want to work on is the active window in Inventor, then run the binary.
After that, there should be a new folder with your rules as vb files located at C:\iLogicTransfer.
From there, you can open, edit, create, delete, and rename files in that folder and they will reflect that in your assembly.

Only .vb files are detected, anything else is ignored. The .vb is parsed out when you rename, so only the actual file name is used for the rules.

# Wishlist
- Folder tree for sub assembly rules, so you can edit a whole project without reloading the app every time you change documents
- Snippets! (Vim plugin for snippets, maybe? Add to already existing coc-snippets plugin?)
- Make iLogic Bridge not one way! (Look into ThisApplication.ActiveDocument.DocumentEvents.OnChange)
