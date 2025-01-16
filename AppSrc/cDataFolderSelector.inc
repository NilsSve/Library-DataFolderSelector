//****************************************************************************
// $Module type: Include File
// $Module name: cDataFolderSelector.inc
// $Author     : Nils Svedmyr, RDC Tools International, <mailto:support@rdctools.com>
// Web-site    : http://www.rdctools.com
// Created     : 2019-10-27 @ 09:33 (Military date format: YY-MM-DD)
//
// Purpose     : Include file with declarations used by cDataFolderSelector.pkg
// See Also    : Check out the SetupDataFolders.chm help file in the Help folder.
//
Use cApplication.pkg
Use cIniFile.pkg
Use seq_chnl.pkg
Use Winkern.pkg  
Use WinShell.pkg
Use tWinStructs.pkg
Use GlobalFunctionsProcedures.pkg
Use cDataFolderLanguageConstants.inc

#IFNDEF ghoDFSFolderSelector
    Global_Variable Handle ghoDFSFolderSelector 
    Move 0 to ghoDFSFolderSelector
#ENDIF                                                

#IF (!@ > 190)
    // If DataFlex 19 or later is used; proper encryption logic is used.
    Use cLoginEncryption.pkg
#ELSE
    // Warning! If a pre-DF 19 version is used a proper ENCRYPTION logic IS NOT USED.
    // Instead the values are base64 encoded (!), so you should construct
    // your own encryption logic if you need better password security for the SetupDataFolders program!
    Use Base64Functions.pkg
#ENDIF

#IFNDEF ghoConnection
    Global_Variable Handle ghoConnection
    Define C_ConnectionIniFileName for "DFConnId.ini" // default name of connections INI file.
    Register_Procedure AutoDisconnect
    Register_Procedure Set psConnectionIni String sConnectionIni
    Register_Function AddAllConnections Returns Boolean
    Register_Function LoginAllConnections Returns Boolean
#ENDIF

#IFNDEF GENERIC_READ
    Define GENERIC_READ for            |CI$80000000
#ENDIF
#IFNDEF GENERIC_WRITE
    Define GENERIC_WRITE for           |CI$40000000
#ENDIF
#IFNDEF FILE_SHARE_READ
    Define FILE_SHARE_READ for         |CI$1
#ENDIF
#IFNDEF FILE_SHARE_WRITE
    Define FILE_SHARE_WRITE for        |CI$2
#ENDIF
#IFNDEF FILE_SHARE_RANDOM
    Define FILE_SHARE_RANDOM for       (FILE_SHARE_READ + FILE_SHARE_WRITE)
#ENDIF
#IFNDEF OPEN_EXISTING
    Define OPEN_EXISTING for           |CI$00000003
#ENDIF
#IFNDEF DFS_FNULL
    Define DFS_FNULL for               |CI$0
#ENDIF

#IFNDEF DFS_Createfile
External_Function DFS_Createfile "CreateFileA" Kernel32.dll ;
    Pointer lpFile ;                // Filename
    DWord dwDesAccess ;             // Access mode
    DWord dwShare ;                 // Share mode
    Pointer lpSecAtt ;              // SD
    DWord dwCrDisp ;                // How to create
    DWord dwFlags ;                 // File attributes
    Handle hTempFile ;              // Handle to template fil
    Returns Integer                 // Returns handle that can be used to access the object
#ENDIF

#IFNDEF DFS_CloseHandle
External_Function DFS_CloseHandle "CloseHandle" Kernel32.dll ;
    Handle hObject ;                // Handle to object
    Returns Integer
#ENDIF

#IFNDEF DFS_GetFileTime
External_Function DFS_GetFileTime "GetFileTime" Kernel32.dll ;
    Handle  hFile  ;                // Handle For a file from createfile with read access
    Pointer lpCreationtime ;        // Pointer to a tWinFileTime structure containing the file creation time
    Pointer lpLastAccesstime ;      // Pointer to a tWinFileTime structure containing the file last access time
    Pointer lpLastWritetime ;       // Pointer to a tWinFileTime structure containing the file last write time
    Returns Integer
#ENDIF

#IFDEF Is$Win64
    Define CS_DFSSetupFolderProgram for "SetupDataFolders64.exe"
#ELSE
    Define CS_DFSSetupFolderProgram for "SetupDataFolders.exe"
#ENDIF
Define CS_DFSIniFileName            for "DataFolderSelector.ini"
Define CS_DFSIniFileDataGroup       for "Data Group"
Define CS_DFSIniFileDataFolderName  for "Data"
Define CS_DFSIniFileDescription     for "Description"
Define CS_DFSIniFileActive          for "Active"
Define CS_DFSIniFilePassword        for "Pw"
Define CS_DFSSettings               for "Settings"
Define CS_DFSGridFontSize           for "GridFontSize"
Define CS_DFSLastUsedDataFolder     for "Last Used Data Folder" // For ini-file.

Define CS_DFSCmdLineIniFile  for "/ini"
Define CS_DFSCmdLineLocation for "/loc" 
Define CS_DFSCmdLineWsFile   for "/x"  
Define CS_DFSCmdLineAdmMode  for "/adm"
Define CS_DFSCmdLineSkinFile for "/skinfile"
Define CS_DFSCmdLineSkinIni  for "/skinini"

// Return values for the PopupDFSPasswordDialog function (DFSPassword.dg)
Enum_List
    Define CI_DFSPasswordOK
    Define CI_DFSPasswordErr
    Define CI_DFSPasswordCancel
End_Enum_List

Struct tDFSDataFolders
    String sDataFolder
    String sDescription
    String sPassword
    Boolean bActive 
    Boolean bSelected
    Boolean bIsValid
End_Struct

//  iMode = DF_FILE_ALIAS_DEFAULT if no master or alias
//  iMode = DF_FILE_IS_MASTER if master
//  iMode = DF_FILE_IS_ALIAS if alias
Struct tDFSMasterAlias
    Handle hTable
    Integer iMode
End_Struct    

// These have been moved to the cDataFolderSelector class:
//
//#IFNDEF get_DFSGetApplicationFileName
//Function DFSGetApplicationFileName Global Returns String
//    // Returns the filename from Windows
//    Integer iNumChars
//    String sFilename
//    
//    Move (Repeat(Character(0), 1024)) to sFileName
//    Move (GetModuleFileName(0, AddressOf(sFilename), 1024)) to iNumChars
//    
//    Function_Return (CString(sFilename))
//End_Function
//
//Function DFSGetApplicationPath Global Returns String
//    // Returns the path of the Application (no trailing "\")
//    String sApplicationFileName sPath
//    Boolean bRemoved
//    
//    Get DFSGetApplicationFileName  to sApplicationFileName
//    Move (PathRemoveFileSpec(AddressOf(sApplicationFileName))) to bRemoved
//    Move (CString(sApplicationFileName)) to sPath
//    
//    If (Right(sPath, 1) ="\") ;
//        Move (Left(sPath, Length(sPath) -1)) to sPath
//    Function_Return sPath
//End_Function
//#ENDIF
