//****************************************************************************
// $Module type: View
// $Module name: SetupDataFolders.vw
// $Author     : Nils Svedmyr, RDC Tools International, <mailto:support@rdctools.com>
// Web-site    : http://www.rdctools.com
// Created     : 2019-10-27 @ 09:33 (Military date format: YY-MM-DD)
//
// Purpose     : A utility for adding Data folders to an ini-file
//               (DataFolderSelector.ini)
// See Also    : Check out the SetupDataFolders.chm help file in the Help folder.
//
Use Windows.pkg
Use seq_chnl.pkg
Use Winkern.pkg
Use cRDCView.pkg
Use cRDCCJGrid.pkg
Use cCJGridColumnRowIndicator.pkg
Use cCJGridColumn.pkg
Use File_dlg.pkg
Use DFSYesNoCancel.dg
Use cDataFolderSelector.pkg
Use cDataFolderLanguageConstants.inc
Use DFSFolderBrowse.dg

Register_Procedure Set phoMainView
Register_Function SelectFolder Returns String

Object oDataFolderSelector_vw is a cRDCView
    Set Size to 117 321
    Set Location to 2 2
    Set Border_Style To Border_Thick
    Set View_Mode to Viewmode_Zoom
    Set pbAutoActivate to True
    Set Maximize_Icon to True
    Set pbAcceptDropFiles to True

    Set phoMainView of ghoApplication to Self
    // Keep the originally loaded ini-file data for the grid
    Property tDataSourceRow[] pTheData

    Object oInfo_tb is a TextBox
        Set Auto_Size_State to False
        Set Size to 9 295
        Set Location to 6 13
        Set Label to CS_DFSDragAndDropInfo
        Set Justification_Mode to JMode_Left
        Set peAnchors to anTopLeftRight
        Set FontItalics to True
    End_Object

    Object oSetupDataFolders_grd is a cRDCCJGrid
        Set Size to 91 297
        Set Location to 18 13
        Set piLayoutBuild to 7 
        
        Object oCJGridColumnRowIndicator is a cCJGridColumnRowIndicator
            Set piWidth to 25
        End_Object

        Object oDataFolder_col is a cCJGridColumn
            Set piWidth to 135
            Set psCaption to CS_DFSDataFolder
            Set psCaption to ("  " + CS_DFSDataFolder)
            Set psToolTip to (CS_DFSAddFolder * "(Ctrl+A)")
            Set pbVDFEditControl to False
            Set Prompt_Button_Mode to PB_PromptOn
            Set psImage to "DFSAddFolder.ico"   
            // This is also necessary for displaying the text correctly,
            // when in edit mode as we are not using the VDFEditControl.
            Set TextColor to clBlack
            Set peIconAlignment to xtpAlignmentIconLeft

            Procedure OnEndEdit String sOldValue String sNewValue
                Forward Send OnEndEdit sOldValue sNewValue
                Send Request_Save
            End_Procedure   
            
            Procedure Prompt
                Send ActivateDataFolderSelector
            End_Procedure

            Procedure Prompt_Callback Handle hoPrompt
                Forward Send Prompt_Callback hoPrompt
            End_Procedure

            Procedure OnSetDisplayMetrics Handle hoMetrics Integer iRow String ByRef sValue
                Boolean bIsValid
                
                Get IsDataFolderValid of ghoDFSFolderSelector sValue to bIsValid
                If (bIsValid = False) Begin
                    Set ComForeColor of hoMetrics to clRed
                End
            End_Procedure

        End_Object

        Object oDescription_col is a cCJGridColumn
            Set piWidth to 280
            Set psCaption to CS_DFSDescription
            Set pbVDFEditControl to False
            // This is also necessary for displaying the text correctly,
            // when in edit mode as we are not using the VDFEditControl.
            Set TextColor to clBlack

            // Set pbVisible to false to hide the column.
            // Note: If you change this you _must_ also increase the grid object piLayoutBuild value
            Set pbVisible to True    
            // No point in showing the column in the CodeJock field chooser if the column is hidden.
            Set pbShowInFieldChooser to (pbVisible(Self) = True)
            
            Procedure OnEndEdit String sOldValue String sNewValue
                Forward Send OnEndEdit sOldValue sNewValue
                Send Request_Save
            End_Procedure

            Procedure OnSetDisplayMetrics Handle hoMetrics Integer iRow String ByRef sValue
                Boolean bIsValid      
                String sDataFolder
                
                Get RowDataFolder iRow to sDataFolder
                Get IsDataFolderValid of ghoDFSFolderSelector sDataFolder to bIsValid
                If (bIsValid = False) Begin
                    Set ComForeColor of hoMetrics to clRed
                End
            End_Procedure

        End_Object

        Object oPassword_col is a cCJGridColumn
            Set piWidth to 100
            Set psCaption to CS_DFSPasswordText 
            Set psToolTip to CS_DFSPasswordFolderText
            Set pbVDFEditControl to False
            // This is also necessary for displaying the text correctly,
            // when in edit mode as we are not using the VDFEditControl.
            Set TextColor to clBlack 

            // Set pbVisible to false to hide the column.
            // Note: If you change this you _must_ also increase the grid object piLayoutBuild value
            Set pbVisible to True    
            // No point in showing the column in the CodeJock field chooser if the column is hidden.
            Set pbShowInFieldChooser to (pbVisible(Self) = True)
            
            Procedure OnSetDisplayMetrics Handle hoMetrics Integer iRow String ByRef sValue
                Forward Send OnSetDisplayMetrics hoMetrics iRow sValue
                If (sValue <> "") Begin
                    Move "*****" to sValue
                End
            End_Procedure

            Procedure OnEndEdit String sOldValue String sNewValue
                Forward Send OnEndEdit sOldValue sNewValue
                Send Request_Save
            End_Procedure

        End_Object

        Object oActive_col is a cCJGridColumn
            Set piWidth to 52
            Set psCaption to CS_DFSActiveText
            Set pbCheckbox to True
            Set peHeaderAlignment to xtpAlignmentCenter

            Procedure OnEndEdit String sOldValue String sNewValue
                Forward Send OnEndEdit sOldValue sNewValue
                Send Request_Save
            End_Procedure

        End_Object

        Object oCJContextMenu is a cCJContextMenu
            Set pbShowPopupBarToolTips of ghoCommandBars to True

            Object oAddMenuItem is a cCJMenuItem
                Set psCaption to CS_DFSAddRow
                Set psToolTip to CS_DFSAddRowHelp
                Set psImage to "DFSAdd.ico"
                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Delegate Send Request_AppendRow
                End_Procedure
            End_Object

            Object oInsertMenuItem is a cCJMenuItem
                Set psCaption to CS_DFSInsertRow
                Set psToolTip to CS_DFSInsertRowHelp
                Set psImage to "DFSInsertRow.ico"
                Set psShortcut to "Ctrl+I"
                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send Request_InsertRow
                End_Procedure
            End_Object

            Object oEditMenuItem is a cCJMenuItem
                Set psCaption to CS_DFSEditRow
                Set psToolTip to CS_DFSEditRowHelp
                Set psImage to "DFSEdit.ico"
                Set psShortcut to "Ctrl+F2"
                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Delegate Send ToggleEdit
                End_Procedure
            End_Object

            Object oDeleteMenuItem is a cCJMenuItem
                Set psCaption to CS_DFSDeleteRow
                Set psToolTip to CS_DFSDeleteRowHelp
                Set psImage to "DFSTrash.ico"
                Set psShortcut to "Del"
                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send DeleteItem
                End_Procedure
            End_Object

            Object oToggleSelectMenuItem is a cCJMenuItem
                Set psCaption to CS_DFSToggleRow
                Set psToolTip to CS_DFSToggleRowHelp
                Set psImage to "DFSToggleOn.ico"
                Set psShortcut to "Ctrl+T"

                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send ToggleSelectState
                End_Procedure

            End_Object

            Object oMoveDown_MenuItem is a cCJMenuItem
                Set psCaption to CS_DFSMoveRowDown
                Set psToolTip to CS_DFSMoveRowDownHelp
                Set psShortcut to "Alt+Key-Down"
                Set psImage to "DFSMoveDown.ico"

                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send MoveDown
                End_Procedure

            End_Object

            Object oMoveUp_MenuItem is a cCJMenuItem
                Set pbControlBeginGroup to True
                Set psCaption to CS_DFSMoveRowUp
                Set psToolTip to CS_DFSMoveRowUpHelp
                Set psShortcut to "Alt+Key-Up"
                Set psImage to "DFSMoveUp.ico"

                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send MoveUp
                End_Procedure

            End_Object

            Object oSaveMenuItem is a cCJMenuItem
                Set pbControlBeginGroup to True
                Set psCaption to CS_DFSSaveFile
                Set psToolTip to CS_DFSSaveFileHelp
                Set psImage to "DFSSave.ico"
                Set psShortcut to "Ctrl+S"

                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send SaveIniFile
                End_Procedure

                Function IsEnabled Returns Boolean
                    Boolean bState
                    Get Should_Save to bState
                    Function_Return (bState = True)
                End_Function
            End_Object

            Object oAddFolder_MenuItem is a cCJMenuItem
                Set pbControlBeginGroup to True
                Set psCaption to CS_DFSAddFolder
                Set psToolTip to (CS_DFSAddFolderHelp * "(Ctrl+A)")
                Set psImage to "DFSAddFolder.ico"
                Set psShortcut to "Ctrl+A"
    
                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl  
                    Send ActivateDataFolderSelector
                End_Procedure
    
            End_Object

            Object oOpenContainingFolderMenuItem is a cCJMenuItem
                Set psCaption to CS_DFSOpenFolder
                Set psToolTip to CS_DFSOpenFolderHelp
                Set psImage to "DFSExplorer.ico"
                Set psShortcut to "Ctrl+F"

                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send OpenContainingFolderOfSelectedRow
                End_Procedure

            End_Object

            If (pbAdminMode(ghoApplication) = True) Begin
                Object oOpenDbBuilderMenuItem is a cCJMenuItem
                    Set pbControlBeginGroup to True
                    Set psCaption to CS_DFDbBuilder
                    Set psToolTip to CS_DFDbBuilder
                    Set psImage to "DBBldr.ico"
                    Set psShortcut to "Ctrl+B"
    
                    Procedure OnExecute Variant vCommandBarControl
                        Forward Send OnExecute vCommandBarControl
                        Send LaunchDbBuilder
                    End_Procedure
    
                End_Object
                
                Object oOpenDbExplorerMenuItem is a cCJMenuItem
                    Set psCaption to CS_DFDbExplorer
                    Set psToolTip to CS_DFDbExplorer
                    Set psImage to "DBExplor.ico"
                    Set psShortcut to "Ctrl+D"
    
                    Procedure OnExecute Variant vCommandBarControl
                        Forward Send OnExecute vCommandBarControl
                        Send LaunchDbExplorer
                    End_Procedure
    
                End_Object  
            End

            Object oOpenMenuItem is a cCJMenuItem
                Set pbControlBeginGroup to True
                Set psCaption to CS_DFSOpenHelp
                Set psToolTip to CS_DFSOpenIniFileDesc
                Set psImage to "DFSOpenFolder.ico"
                Set psShortcut to "Ctrl+O"  
                
                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send ActivateOpenDialog
                End_Procedure
                
            End_Object

            Object oClearMenuItem is a cCJMenuItem
                Set psCaption to CS_DFSClearFile
                Set psToolTip to CS_DFSClearFileHelp
                Set psImage to "DFSClear.ico"
                Set psShortcut to "Ctrl+F5"
                
                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send ClearIniFile
                End_Procedure
            
            End_Object

            Object oRefreshMenuItem is a cCJMenuItem
                Set psCaption to CS_DFSRefreshHelp
                Set psToolTip to CS_DFSRefreshDesc
                Set psImage to "DFSRefresh.ico"
                Set psShortcut to "Ctrl+R"

                Procedure OnExecute Variant vCommandBarControl
                    Forward Send OnExecute vCommandBarControl
                    Send RefreshData
                End_Procedure

                Function IsEnabled Returns Boolean
                    Boolean bSave
                    Get Should_Save to bSave
                    Function_Return (bSave = True)
                End_Function

            End_Object  
            
            Set phoContextMenu to Self
        End_Object

        Procedure LoadData
            Handle hoDataSource
            tDFSDataFolders[] aDataFolders
            tDFSDataFolders TheDataFolder
            Integer iSize iCount iData_col iDescription_col iPassword_col iActive_col iRow
            tDataSourceRow[] TheData
            tDataSourceRow TheRow
            Boolean bIsValid bOK 
            
            Move True to bOK 
            Move 0 to iRow
            Send ChangeHeaderText
            Get piColumnId of oDataFolder_col  to iData_col
            Get piColumnId of oDescription_col to iDescription_col
            Get piColumnId of oPassword_col    to iPassword_col
            Get piColumnid of oActive_col      to iActive_col
            Get phoDataSource to hoDataSource
            Send Reset of hoDataSource

            Get ReadDataFoldersFromIniFile of ghoDFSFolderSelector to aDataFolders
            Move (SizeOfArray(aDataFolders)) to iSize
            If (iSize = 0) Begin
                Get ReadWorkspaceFolders of ghoDFSFolderSelector to aDataFolders
                Move (SizeOfArray(aDataFolders)) to iSize
            End
            Decrement iSize
            For iCount from 0 to iSize
                Move aDataFolders[iCount].sDataFolder  to TheRow.sValue[iData_col]
                Move aDataFolders[iCount].sDescription to TheRow.sValue[iDescription_col]
                Move aDataFolders[iCount].sPassword    to TheRow.sValue[iPassword_col]
                Move aDataFolders[iCount].bActive      to TheRow.sValue[iActive_col]
                Move TheRow                            to TheData[iCount]
                Move aDataFolders[iCount].bIsValid     to bIsValid
                If (bIsValid = False and iCount = 0) Begin
                    Move False to bOK
                End
            Loop

            Set pTheData to TheData
            // Initialize Grid with new data
            Send InitializeData TheData
            Set psFooterText of oDescription_col to (CS_DFSNoOfFolders * String(iSize +1))
            If (bOK = False) Begin
                Get FirstValidFolder of ghoDFSFolderSelector to TheDataFolder.sDataFolder
                Move (SearchArray(TheDataFolder, aDataFolders)) to iRow
            End                            
            Send MoveToRow iRow
        End_Procedure

        // Returns the current Data folder.
        Function SelectedDataFolder Returns String
            Integer iCol
            String sDataFolder
            Get piColumnId of oDataFolder_col to iCol
            Get DataValue iCol to sDataFolder
            Function_Return sDataFolder
        End_Function

        Function DataValue Integer iCol Returns String
            tDataSourceRow[] RowData
            Integer iRow
            String sValue
            Handle hoDataSource

            Move "" to sValue
            Get phoDataSource to hoDataSource
            Get SelectedRow   of hoDataSource to iRow
            Get DataSource    of hoDataSource to RowData
            If (iRow >= 0) Begin
                Move RowData[iRow].sValue[iCol] to sValue
            End

            Function_Return sValue
        End_Function

        Function RowDataFolder Integer iRow Returns String
            tDataSourceRow[] RowData
            String sValue
            Handle hoDataSource 
            Integer iCol

            Move "" to sValue
            Get piColumnId of oDataFolder_col to iCol
            Get phoDataSource to hoDataSource
            Get DataSource    of hoDataSource to RowData
            Move RowData[iRow].sValue[iCol] to sValue

            Function_Return sValue
        End_Function

        Procedure RefreshDataFromProperty
            Handle hoDataSource
            tDataSourceRow[] TheData
            Integer iSize

            Send ChangeHeaderText
            Get phoDataSource to hoDataSource
            Send Reset of hoDataSource
            Get pTheData to TheData
            Move (SizeOfArray(TheData)) to iSize
            Send InitializeData TheData
            Set psFooterText of oDescription_col to (CS_DFSNoOfFolders * String(iSize))
            Send MovetoFirstRow
        End_Procedure

        // Adds the passed folder name if the current line contains data.
        // However, it behaves as "insert" of data if the insert key has been 
        // pressed and the current row is empty.
        Procedure AddFolderName String sDataFolder
            Handle hoDataSource
            tDFSDataFolders[] aDataFolders
            Integer iSize iItem iActive_col iData_col iFocusedRow  
            Boolean bInsert
            tDataSourceRow[] TheData
            tDataSourceRow TheRow

            Send ChangeHeaderText
            Get piColumnId of oDataFolder_col  to iData_col
            Get piColumnid of oActive_col      to iActive_col
            Get phoDataSource to hoDataSource
            Get SelectedRow of hoDataSource to iFocusedRow 
            
            Get DataSource of hoDataSource to TheData
            Move (TheData[iFocusedRow].sValue[iData_col] = "") to bInsert
            Move (SizeOfArray(TheData)) to iSize
            If (bInsert = True) Begin
                Move iFocusedRow to iItem
            End                          
            Else Begin
                Move iSize to iItem    
            End
            Move sDataFolder to TheData[iItem].sValue[iData_col]
            Move True        to TheData[iItem].sValue[iActive_col]

            // Initialize Grid with new data
            Send InitializeData TheData
            Send MoveToRow iItem
            If (bInsert = False) Begin
                Send ToggleEdit     
            End
            Get NumberOfRows to iSize
            Set psFooterText of oDescription_col to (CS_DFSNoOfFolders * String(iSize))
        End_Procedure

        // Activate "Add New Folder" dialog:
        Procedure ActivateDataFolderSelector
            String sOrgFolder sFolder sHome
            
            Get SelectedDataFolder of oSetupDataFolders_grd to sOrgFolder
            If (sOrgFolder <> "") Begin  
                // If it is a local Data path we need to make it into a fully qualifying path;                               
                If (not(sOrgFolder contains "\")) Begin   
                    // We assume that the workspace local Data folder (without any path) is located directly under the Home folder;
                    Get psHome of (phoWorkspace(ghoApplication)) to sHome
                    Get AddFolderDelimiter sHome to sHome                                 
                    Move (sHome + sOrgFolder) to sOrgFolder
                End
                Set pbAutoSeed  of ghoDFSFolderSelectorDialog to True
                Set psSeedValue of ghoDFSFolderSelectorDialog to sOrgFolder
            End
            Get SelectFolder of ghoDFSFolderSelectorDialog CS_DFSSelectDataFolder True to sFolder
            If (sFolder <> "") Begin
                Send AddFolderName of oSetupDataFolders_grd sFolder
            End
        End_Procedure

        // Spins through all entered & ACTIVE Data folders and checks that they exists
        Function AllDataFoldersExists Returns Boolean
            Integer iCount iSize iActive_col iData_col iItem
            Handle hoDataSource
            tDataSourceRow[] TheData
            tDataSourceRow TheRow
            Boolean bOK bIsLocalDataFolder bExists bActiveState
            String sHomePath sDataFolder

            Move True to bOK
            Get piColumnId of oDataFolder_col  to iData_col
            Get piColumnId of oActive_col      to iActive_col
            Get phoDataSource to hoDataSource
            Get DataSource of hoDataSource to TheData
            Move (SizeOfArray(TheData)) to iSize
            Decrement iSize
            Move 0 to iItem

            // Load data from the grid datasource array
            For iCount from 0 to iSize
                Move TheData[iCount] to TheRow
                Move TheRow.sValue[iData_col]   to sDataFolder
                Move TheRow.sValue[iActive_col] to bActiveState
                If (sDataFolder <> "" and bActiveState = True) Begin
                    Move (not(sDataFolder contains ":")) to bIsLocalDataFolder
                    If (bIsLocalDataFolder = True) Begin
                        Get psHome of (phoWorkspace(ghoApplication)) to sHomePath
                        Get AddFolderDelimiter sHomePath to sHomePath
                        Move (sHomePath + sDataFolder) to sDataFolder
                    End
                    File_Exist sDataFolder bExists
                    If (bExists = False) Begin
                        Move iSize to iCount // We're done
                        Move False to bOK
                        Send Info_Box (CS_DFSTheDataFolder1 + String(sDataFolder) + String("\n") + CS_DFSTheDataFolder2)
                    End
                End
            Loop

            Function_Return bOK
        End_Function

        Procedure WriteIniFile
            Integer iCount iSize iActive_col iData_col iDescription_col iPassword_col iItem
            Handle hoDataSource
            tDataSourceRow[] TheData
            tDataSourceRow TheRow
            tDFSDataFolders[] aDataFolders
            tDFSDataFolders aDataFolder
            Boolean bOK

            Move ghoDFSFolderSelector to ghoDFSFolderSelector

            Get piColumnId of oDataFolder_col  to iData_col
            Get piColumnId of oDescription_col to iDescription_col
            Get piColumnId of oPassword_col    to iPassword_col
            Get piColumnid of oActive_col      to iActive_col
            Get phoDataSource to hoDataSource
            Get DataSource of hoDataSource to TheData
            Move (SizeOfArray(TheData)) to iSize
            Decrement iSize
            Move 0 to iItem

            // Load data from the grid datasource array
            For iCount from 0 to iSize
                Move TheData[iCount] to TheRow
                If (TheRow.sValue[iData_col] <> "") Begin
                    Move TheRow.sValue[iData_col]        to aDataFolder.sDataFolder
                    Move TheRow.sValue[iDescription_col] to aDataFolder.sDescription
                    Move TheRow.sValue[iPassword_col]    to aDataFolder.sPassword
                    Move TheRow.sValue[iActive_col]      to aDataFolder.bActive
                    Move aDataFolder to aDataFolders[iItem]
                    Increment iItem
                End
            Loop

            Get WriteDataFoldersToIniFile of ghoDFSFolderSelector aDataFolders to bOK
            If (bOK = False) Begin
                Send Info_Box CS_DFSErrorSavingIniFile
                Procedure_Return
            End

            // Update the view property with the newly saved values. (Used to check if anything has changed)
            Set pTheData to TheData
            Send ChangeHeaderText
//            Send Info_Box CS_DFSChangesSaved
            Send ChangeStatusRowText CS_DFSChangesSaved
        End_Procedure

        Procedure MoveUp
            tDataSourceRow[] TheData
            tDataSourceRow TheRow
            Handle hDataSource
            Integer iCurrentRow
            Get phoDataSource to hDataSource
            Get DataSource of hDataSource to TheData
            Get SelectedRow of hDataSource to iCurrentRow
            If (iCurrentRow > 0) Begin
                Move TheData[iCurrentRow - 1] to TheRow
                Move TheData[iCurrentRow] to TheData[iCurrentRow - 1]
                Move TheRow to TheData[iCurrentRow]
                Send ReInitializeData TheData True
                Send MoveToRow (iCurrentRow - 1)
            End
        End_Procedure

        Procedure MoveDown
            tDataSourceRow[] TheData
            tDataSourceRow TheRow
            Handle hDataSource
            Integer iCurrentRow
            Get phoDataSource to hDataSource
            Get DataSource of hDataSource to TheData
            Get SelectedRow of hDataSource to iCurrentRow
            If ((iCurrentRow + 1) < SizeOfArray(TheData)) Begin
                Move TheData[iCurrentRow] to TheRow
                Move TheData[iCurrentRow + 1] to TheData[iCurrentRow]
                Move TheRow to TheData[iCurrentRow + 1]
                Send ReInitializeData TheData True
                Send MoveToRow (iCurrentRow + 1)
            End
        End_Procedure

        // Called when the grid object is created:
        Procedure Activating
            Forward Send Activating
            Send DoChangeFontSize
            Send LoadData
        End_Procedure

        Procedure DoChangeFontSize
            Handle hoFont hoPaintManager
            Variant vFont
            String sFont sFontSize
            Boolean bCreated
            Integer iVal

            Get IsComObjectCreated to bCreated  // When program is started, grid object isn't created yet.
            If (bCreated = False) Begin
                Procedure_Return
            End

            Get phoReportPaintManager to hoPaintManager
            Get Create (RefClass(cComStdFont)) to hoFont  // Create a font object
            Get ComTextFont of hoPaintManager to vFont    // Bind the font object to the Grid's text font
            Set pvComObject of hoFont to vFont            // Connect DataFlex object with com object

            Get ReadString of ghoDFSFolderSelector CS_DFSSettings CS_DFSGridFontSize 8 to iVal
            Set ComSize of hoFont to iVal
            Send ComRedraw
            Send Destroy to hoFont                        // Destroy the font object (releases memory)
        End_Procedure

        Function Should_Save Returns Boolean
            tDataSourceRow[] TheData1 TheData2
            Handle hoDataSource
            Boolean bShouldSave

            Move True to bShouldSave
            Get pTheData to TheData1
            Get phoDataSource to hoDataSource
            Get DataSource    of hoDataSource to TheData2
            #IF (!@ > 180)
                Move (not(IsSameArray(TheData1, TheData2))) to bShouldSave
            #ENDIF

            Function_Return bShouldSave
        End_Function

        Function HasRecord Returns Boolean
            tDataSourceRow[] TheData
            Handle hoDataSource
            Integer iSize

            Get phoDataSource to hoDataSource
            Get DataSource    of hoDataSource to TheData
            Move (SizeOfArray(TheData)) to iSize

            Function_Return (iSize > 0)
        End_Function

        Procedure ChangeHeaderText
            Handle[] hoPanels
            String sFileName
            Boolean bChangesExist

            Send ChangeStatusRowText ""
            Get psFileName of ghoDFSFolderSelector to sFileName
            Get Should_Save to bChangesExist
            If (bChangesExist = True) Begin
                Move (sFileName + "*") to sFileName
            End

            // Not sure why, but if the oStatusPane1 was set to "Set piID to sbpIDIdlePane",
            // it wasn't always updated when this message was send. So instead change the
            // text explicitly:
            If (ghoCommandBars <> 0) Begin
                If ((phoStatusBar(ghoCommandBars))) Begin
                    Get PaneObjects of (phoStatusBar(ghoCommandBars)) to hoPanels
                    Set psText of hoPanels[0] to sFileName
                End
            End
        End_Procedure

        Procedure RemoveCurrentFolder
            Integer iSize iRow iItem
            Handle hoDataSource
            tDataSourceRow[] TheData

            Move 0 to iItem
            Get phoDataSource to hoDataSource
            Get DataSource of hoDataSource to TheData

            Get SelectedRow of hoDataSource to iRow
            If (iRow = -1) Begin
                Procedure_Return
            End

            Move False to Err
            Send Request_Delete

            Get DataSource of hoDataSource to TheData
            Move (SizeOfArray(TheData)) to iSize
            Set psFooterText of oDescription_col to (CS_DFSNoOfFolders * String(iSize))
        End_Procedure

        Procedure RemoveCurrentRow
            Integer iSize iRow iItem
            Handle hoDataSource
            tDataSourceRow[] TheData

            Move 0 to iItem
            Get phoDataSource to hoDataSource
            Get DataSource of hoDataSource to TheData

            Get SelectedRow of hoDataSource to iRow
            If (iRow = -1) Begin
                Procedure_Return
            End

            Move False to Err
            Send Request_Delete

            Get DataSource of hoDataSource to TheData
            Move (SizeOfArray(TheData)) to iSize
            Set psFooterText of oDescription_col to (CS_DFSNoOfFolders * String(iSize))
        End_Procedure

        Procedure ClearData
            tDataSourceRow[] TheData
            Send ChangeHeaderText
            Set pTheData to TheData
            Send InitializeData TheData
        End_Procedure

        Function CurrentRow Returns Integer
            Handle hoDataSource
            Integer iRow

            Get phoDataSource to hoDataSource
            Get SelectedRow   of hoDataSource to iRow
            Function_Return iRow
        End_Function

        Function CurrentRowData Returns tDataSourceRow
            tDataSourceRow[] TheData
            tDataSourceRow TheRow
            Handle ho hoDataSource
            Integer iRow

            Get phoDataSource  to hoDataSource
            Get DataSource     of hoDataSource to TheData
            Get SelectedRow    of hoDataSource to iRow
            Move TheData[iRow] to TheRow

            Function_Return TheRow
        End_Function
        
        Function NumberOfRows Returns Integer
            Handle hoDataSource
            tDataSourceRow[] TheData
            Integer iRows 
        
            Get phoDataSource to hoDataSource
            Get DataSource of hoDataSource to TheData
            Move (SizeOfArray(TheData)) to iRows
            Function_Return iRows
        End_Function

        // Toggles the current row on/off (the checkbox)
        Procedure ToggleSelectState
            Boolean bChecked
            Integer iCol
            Handle hoCol

            Get piColumnId of oActive_col to iCol
            Get ColumnObject iCol   to hoCol
            Get SelectedRowValue    of hoCol to bChecked
            Send UpdateCurrentValue of hoCol (not(bChecked))
            Send Request_Save
        End_Procedure

        Procedure OnComMouseUp Short llButton Short llShift Integer llx Integer lly
            Forward Send OnComMouseUp llButton llShift llx lly
            // Can't have this because it interferes with the Edit menu item action.
//            Send Request_Save
        End_Procedure

        Procedure OnRowChanged Integer iOldRow Integer iNewSelectedRow
            Integer iRow
            Handle hoDataSource
            tDataSourceRow[] RowData

            Forward Send OnRowChanged iOldRow iNewSelectedRow
            Send ChangeHeaderText

            Get phoDataSource to hoDataSource
            Get SelectedRow of hoDataSource to iRow
            If (iRow <> -1) Begin
                Get DataSource of hoDataSource to RowData
            End
        End_Procedure

        Procedure Request_Cancel
            Send None
        End_Procedure

        Procedure Close_Panel
            Send None
        End_Procedure

        Procedure CommitSelectedRow
            Boolean bState
            Get CommitSelectedRow to bState
            Send EndEdit
            Send Request_Save
        End_Procedure

        Procedure OnHeaderClick Integer iCol
            Forward Send OnHeaderClick iCol
            Send ActivateDataFolderSelector
        End_Procedure

        On_Key kCancel                Send CommitSelectedRow
        On_Key Key_Alt+Key_Down_Arrow Send MoveDown
        On_Key Key_Alt+Key_Up_Arrow   Send MoveUp
        On_Key Key_Alt+Key_R          Send RefreshData
        On_Key Key_Ctrl+Key_O         Send ActivateOpenDialog
        On_Key Key_Ctrl+Key_F5        Send ClearIniFile
        On_Key Key_F2                 Send ToggleEdit
        On_Key Key_Ctrl+Key_A         Send ActivateDataFolderSelector
        On_Key Key_Ctrl+Key_I         Send Request_InsertRow
        On_Key Key_Insert             Send Request_InsertRow
        On_Key Key_Ctrl+Key_E         Send ToggleEdit
        On_Key Key_Ctrl+Key_T         Send ToggleSelectState
        On_Key Key_Ctrl+Key_B         send LaunchDbBuilder
        On_Key Key_Ctrl+Key_D         Send LaunchDbExplorer
    End_Object

    Function Should_Save Returns Boolean
        Boolean bShouldSave
        Get Should_Save of oSetupDataFolders_grd to bShouldSave
        Function_Return bShouldSave
    End_Function

    // Public access methods: (used by menu/toolbar system)
    Procedure ActivateOpenDialog
        String sFileName

        Get vSelect_File2 ((CS_DFSDataFolderSetupIni * "(*.ini)|") + CS_DFSIniFileName + "|" + CS_DFSAllIniFiles * "(*.ini)|*.ini|" + CS_DFSAllFiles * "(*.*)|*.*") CS_DFSSelectIniFileFolder "" to sFileName
        If (sFileName <> "") Begin
            Set psFileName of ghoDFSFolderSelector to sFilename
            Send LoadData of oSetupDataFolders_grd
        End
    End_Procedure

    Procedure Close_Panel
    End_Procedure

    Procedure RefreshData
        Handle ho

        Move (oSetupDataFolders_grd(Self)) to ho
        Send ChangeStatusRowText ""
        Send RefreshDataFromProperty of oSetupDataFolders_grd
    End_Procedure 
    
    Function ToolsWorkspaceFile Returns String
        String sWsFile sDataFolder sProgramsFolder
        Integer iPos
        Boolean bExists
        
        Get SelectedDataFolder of oSetupDataFolders_grd to sDataFolder
        Get QualifiedDataPath of ghoDFSFolderSelector sDataFolder to sDataFolder
        Get IsDataFolderValid of ghoDFSFolderSelector sDataFolder to bExists
        
        Move (RightPos("\", sDataFolder)) to iPos
        Move (Left(sDataFolder, (iPos -1))) to sDataFolder
        Move (sDataFolder + "\Programs") to sProgramsFolder
        File_Exist (sProgramsFolder + "\Config.ws") bExists
        If (bExists = True) Begin
            Move (sProgramsFolder + "\Config.ws") to sWsFile
        End
        Else Begin
            Get psWorkspaceWSFile of (phoWorkspace(ghoApplication)) to sWsFile
        End
        Function_Return sWsFile
    End_Function

    // ToDo: How to open the selected Data folder in DbExplorer without making changes to the corresponding 
    //       config.ws file? 
    //       This code will *not* Open the *selected Data folder* If it is not the same as the
    //       "DataPath" setting in the .ws file.
    //       It would be nice if the full Data folder name or Filelist.cfg could be passed on the command 
    //       line to DbExplorer. The same goes for DbBuilder.
    Procedure LaunchDbExplorer
        String sWsFile
        Get ToolsWorkspaceFile to sWsFile
        Move (CS_DFSCmdLineWsFile + String('"' + sWsFile + '"')) to sWsFile
        Runprogram Shell Background "Dbexplor.exe" sWsFile
    End_Procedure

    Procedure LaunchDbBuilder
        String sWsFile
        Get ToolsWorkspaceFile to sWsFile
        Move (CS_DFSCmdLineWsFile + String('"' + sWsFile + '"')) to sWsFile
        Runprogram Shell Background "DBBldr.exe" sWsFile
    End_Procedure
    
    Function Verify_Exit_Application Returns Integer
        Integer iRetval
        Boolean bChanged bStay

        Get Should_Save of oSetupDataFolders_grd to bChanged
        If (bChanged = False) Begin
            Move False to bStay 
        End

        // Uses an improved YesNoCancel dialog
        Else Begin     
            Get DFSYesNoCancel_Box CS_DFSUnsavedChangesInfo to iRetval 
            If (iRetval = MBR_Yes) Begin
                Send SaveIniFile
                Move False to bStay
            End
            If (iRetval = MBR_No) Begin
                Move False to bStay
            End
            If (iRetval = MBR_Cancel) Begin
                Move True to bStay
            End
        End 
        
        Function_Return bStay // Should be False to leave the application
    End_Function

    Procedure SaveIniFile
        Boolean bFoldersExists
        String sFilename sPath

        // We should probably allow for unavailable folders to be saved, as they
        // might be avilable when the data is used.
//        Get AllDataFoldersExists of oSetupDataFolders_grd to bFoldersExists
//        If (bFoldersExists = False) Begin
//            Procedure_Return
//        End
        Send ChangeStatusRowText ""

        Get psFileName of ghoDFSFolderSelector to sFilename
        If (sFilename = "") Begin
            Get psProgramPath of (phoWorkspace(ghoApplication)) to sPath
            Get vSave_File2 ((CS_DFSDataFolderSetupIni * "(*.ini)|") + CS_DFSIniFileName + "|" + CS_DFSAllIniFiles * "(*.ini)|*.ini|" + CS_DFSAllFiles * "(*.*)|*.*") "" sPath to sFileName
            If (sFileName = "") Begin
                Procedure_Return
            End  
            Set psFileName of ghoDFSFolderSelector to sFilename
        End
        Send WriteIniFile of oSetupDataFolders_grd
    End_Procedure

    Procedure ClearIniFile
        Set psFileName of ghoDFSFolderSelector to ""
        Send ClearData of oSetupDataFolders_grd
    End_Procedure

    Procedure DeleteItem
        Send ChangeStatusRowText ""
        Send RemoveCurrentFolder of oSetupDataFolders_grd
    End_Procedure

    Procedure ChangeStatusRowText String sText
        Handle[] hoPanels  
        If (ghoCommandBars <> 0) Begin
            If ((phoStatusBar(ghoCommandBars))) Begin
                Get PaneObjects of (phoStatusBar(ghoCommandBars)) to hoPanels
                If (SizeOfArray(hoPanels)) Begin
                    Set psText of hoPanels[1] to sText
                End
            End
        End
    End_Procedure

    Function AddFolderDelimiter String sPath Returns String
        String sDirSep
        Move (SysConf(SYSCONF_DIR_SEPARATOR)) to sDirSep
        Move (Trim(sPath)) to sPath
        If (Right(sPath, 1) <> sDirSep) Begin
            Move (sPath + sDirSep) to sPath
        End
        Function_Return sPath
    End_Function

    Function IsFileListInFolder String sFolderName Returns Boolean
        Boolean bOK bIsFolder
        String sPath sFileList

        Move False to bOK
        Move (Trim(sFolderName)) to sFolderName
        Get IsFolder of ghoDFSFolderSelector sFolderName to bIsFolder
        If (bIsFolder = True) Begin
            Move sFolderName to sPath
            Get AddFolderDelimiter sPath to sPath
            Get psFileList of (phoWorkspace(ghoApplication)) to sFileList
            Get ExtractFileName sFileList to sFileList
            File_Exist (sPath + sFileList) bOK
        End
        Function_Return bOK
    End_Function

    Function LocalDataFolder String sFileName Returns String
        String sHomePath sTest sDataFolder
        Boolean bIsFolder bIsLocalDataFolder bIsFilelistOK

        Move False to bIsLocalDataFolder
        Move "" to sDataFolder
        Get IsFolder of ghoDFSFolderSelector sFilename to bIsFolder
        If (bIsFolder = False) Begin
            Function_Return ""
        End
        Get IsFileListInFolder sFilename to bIsFilelistOK
        If (bIsFilelistOK = False) Begin
            Function_Return ""
        End
        Get psHome of (phoWorkspace(ghoApplication)) to sHomePath
        Get ExtractFilePath sFilename to sTest
        Move (Lowercase(sHomePath) = Lowercase(sTest)) to bIsLocalDataFolder
        If (bIsLocalDataFolder = True) Begin
            Move (Replace(sHomePath, sFilename, "")) to sDataFolder
        End

        Function_Return sDataFolder
    End_Function

    Procedure OpenContainingFolderOfSelectedRow
        Handle hoGrid
        String sHomePath sDataFolder
        Boolean bIsLocalDataFolder

        Move (oSetupDataFolders_grd(Self)) to hoGrid
        Get SelectedRowValue of (oDataFolder_col(hoGrid)) to sDataFolder
        Move (not(sDataFolder contains ":")) to bIsLocalDataFolder

        If (bIsLocalDataFolder = True) Begin
            Get psHome of (phoWorkspace(ghoApplication)) to sHomePath
            Get AddFolderDelimiter sHomePath to sHomePath
            Move (sHomePath + sDataFolder) to sDataFolder
        End
        Move ("/select," * String(sDataFolder)) to sDataFolder

        Runprogram Shell Background "explorer.exe" sDataFolder
    End_Procedure

    Procedure OnFileDropped String sFilename Boolean bLast
        Boolean bHasChange bIsFolder bIsFilelistOK
        Handle ho hoGrid
        Integer iRetval
        String sFileNameOnly sDataFolder

        Forward Send OnFileDropped sFilename bLast
        // We only react on the last dropped file/folder (if multiple files/folders was dropped)
        If (bLast = False) Begin
            Procedure_Return
        End

        Move (oSetupDataFolders_grd(Self)) to hoGrid
        Get IsFolder of ghoDFSFolderSelector sFilename to bIsFolder
        If (bIsFolder = True) Begin
            Get IsFileListInFolder sFilename to bIsFilelistOK
            If (bIsFilelistOK = False) Begin
                Send Info_Box CS_DFSFolderInvalid
                Procedure_Return
            End
            Get LocalDataFolder sFilename to sDataFolder
            If (sDataFolder = "") Begin
                Move sFilename to sDataFolder
            End
            Send AddFolderName of hoGrid sDataFolder
        End
        Else Begin
            Get ExtractFileName sFilename to sFileNameOnly
            If (Uppercase(sFileNameOnly) <> Uppercase(CS_DFSIniFileName)) Begin
                Send Info_Box (CS_DFSSorryOnly * CS_DFSIniFileName * CS_DFSFilesCanBeDropped)
                Procedure_Return
            End
            Get Should_Save of hoGrid to bHasChange
            If (bHasChange = True) Begin
                Get YesNo_Box CS_DFSLoadNewFile to iRetval
                If (iRetval <> MBR_Yes) Begin
                    Procedure_Return
                End
            End
            Set psFileName of ghoDFSFolderSelector to sFilename
            Send LoadData of hoGrid
        End
    End_Procedure

    Function vSelect_File2 String sFilterString String sCaptionText String sInitalFolder Returns String
        String sFileName
        Boolean bOpen
        String[] asFiles

        Move "" to sFileName

        Object oOpenDialog is a OpenDialog
            Set Filter_String  to sFilterString
            Set Dialog_Caption to sCaptionText
            Set Initial_Folder to sInitalFolder
            Get Show_Dialog to bOpen
            If (bOpen = True) Begin
                Get Selected_Files to asFiles
                Move asFiles[0] to sFileName
            End
        End_Object

        Function_Return sFileName
    End_Function

    Function vSave_File2 String sFilterString String sCaptionText String sInitalFolder Returns String
        String sFileName
        Boolean bSave
        String[] asFiles

        Move "" to sFileName

        Object oSaveAsDialog is a SaveAsDialog
            Set Filter_String  to sFilterString
            Set Dialog_Caption to sCaptionText
            Set Initial_Folder to sInitalFolder
            Get Show_Dialog to bSave
            If (bSave = True) Begin
                Get File_Name to sFileName
            End
        End_Object

        Function_Return sFileName
    End_Function

    // On idle handling:
    Object oIdle is a cIdleHandler
        Procedure OnIdle
          Delegate Send OnIdle
        End_Procedure
    End_Object

    Procedure OnIdle
        Handle ho

        Move (oSetupDataFolders_grd(Self)) to ho
        Send ChangeHeaderText of ho
    End_Procedure

    Procedure Activating
        Handle ho

        Set Maximize_Icon to True
        Set Minimize_Icon to False
        Set Border_Style to Border_Thick
        Set View_Mode to Viewmode_Zoom

        // Note: The following line is essential for the resizing logic
        // to work when starting the program.
        If (ghoCommandBars <> 0) Begin
            Move (Client_Id(ghoCommandBars)) to ho
            Set Border_Style of ho to Border_ClientEdge
        End

        Set pbEnabled of oIdle to True
    End_Procedure

    Procedure Page Integer iPageObject
        Forward Send Page iPageObject
        Set Icon to "DFSSetupDataFolders.ico"
    End_Procedure

    On_Key Key_F2            Send ToggleEdit of oSetupDataFolders_grd
    On_Key kCancel           Send CommitSelectedRow of oSetupDataFolders_grd
    On_Key Key_Ctrl+Key_S    Send SaveIniFile
    On_Key kDelete_Character Send DeleteItem
    On_Key Key_Ctrl+Key_X    Send DeleteItem
    On_Key Key_Ctrl+Key_R    Send RefreshData
    On_Key Key_Ctrl+Key_F    Send OpenContainingFolderOfSelectedRow
End_Object
