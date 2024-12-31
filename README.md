The Data Folder Selector Utility
See Also: The SetupDataFolders.chm file in the Help folder.

The Data Folder Selector Utility helps users select a Data folder (Client Company) to use when a program starts.
Two parts are needed for this to work;

1. The SetupDataFolder.exe administrator program where Data folders are added, which later can be selected by end-users.
2. The DataFolderSelector.dg dialog. The file should be added to your program's Client_Area and will popup automatically when the program starts. The user can there select which Data folder/Company to use.

What is it?
It is for allowing users to select which Data folder to use when a program starts.
This is handy when a workspace is "multi-tenancy" or client based, aka it contains several Data folders for various "Companies" or such and there is a need to allow for selection of which Data folder (Company) to use when the program starts.

•An administration program (SetupDataFolder.exe) that is used to specify which Data folders and descriptions (aka Company Names) should be presented to the user when the program is started. This work should be made by an administrator.
•The DataFolderSelector.dg has a property named pbAutoPopup that is True by default. The dialog should be added like any other component to your program's Client_Area. When pbAutoPopup is True, the dialog will automatically popup up, for the user to select which Data folder/Company to use. All data path's etc. will be updated automatically when the OK button is clicked, tables re-opened and any SQL connections re-established..
•An Order Entry Sample application and several Data folders to test the concept. After selecting a Data folder with the DataFolderSelector dialog and the Order program has started, you can use the About dialog to check the Data folder changes. Click Help - About - System Info. The dialog contains info about the selected Data Path and Filelist.cfg path and also the Full DFPath.
