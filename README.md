**Data Folder Selector Utility**

**See Also:** Refer to the SetupDataFolders.chm file located in the Help folder.

The Data Folder Selector Utility enables users to choose a Data folder (Client Company) when starting a program.

This utility consists of two key components:

1. **SetupDataFolder.exe**: This administrator program allows the addition of Data folders, which can later be selected by end-users.
2. **DataFolderSelector.dg dialog**: This file should be included in your program's Client_Area and will automatically pop up when the program starts, allowing users to select which Data folder (Company) to utilize.

### What is it?
The Data Folder Selector Utility provides a way for users to select the desired Data folder at program startup. This feature is particularly useful in environments that are "multi-tenancy," where a workspace contains multiple Data folders for various companies, allowing for easy selection of the appropriate Data folder (Company) at launch.

- An **administration program** (SetupDataFolder.exe) must be used by an administrator to specify which Data folders and their corresponding descriptions (Company Names) should be presented to the user upon the program's start.
- The **DataFolderSelector.dg** includes a property called pbAutoPopup, which is set to True by default. This dialog should be added to your program's Client_Area like any other component. When pbAutoPopup is set to True, the dialog will automatically appear, enabling the user to select the desired Data folder/Company. All relevant data paths will be updated automatically once the OK button is clicked, with tables re-opening and any SQL connections being re-established.
- An **Order Entry Sample application** is provided alongside several Data folders to help test the concept. After selecting a Data folder with the DataFolderSelector dialog and launching the Order program, you can verify the changes in the About dialog. To do this, click Help -> About -> System Info. The dialog will display information about the selected Data Path, the Filelist.cfg path, and the complete DFPath.
