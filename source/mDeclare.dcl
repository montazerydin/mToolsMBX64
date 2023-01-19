'===============================================================================
'**
'**		Author				: Montazery Hasibuan
'**		Email				: --insert your email here--
'**
'**		Filename			: mDeclare.dcl
'**		Date Created		: 17 Nov 2010
'**		Last Modified		: 2014-06-14 10:11 AM
'**

'===============================================================================

'Pendeklarasian prosedur-prosedur yang digunakan dalam project mToolsMBX
'(terpisah berdasarkan file MB yang menggunakannya)

Include "mapbasic.def"
Include "icons.def"
Include "menu.def"
Include "mType.typ"
Include "mDef.def"
Include "mDialogID.dig"

'mTools.mb
Declare Sub Main
Declare Sub MainTools
Declare Sub LoginUI
Declare Sub Bye 
Declare Sub cancel_handler
Declare Sub LoadApp

'sub_mTools.mb
Declare Sub GetTables	
Declare Sub GetDirectory
Declare Sub GetColList 
Declare Sub MLBOKButton
Declare Sub ValueParse(sVals() As String, ByVal sInPut As String,ByVal sSeparator As String)

'upcol_mTools.mb
Declare Sub ListDialog
Declare Sub RedoColList 			
Declare Sub EnableButton 			
Declare Sub ProsesUpdate 	

'split_mTools.mb
Declare Sub SplitTable
Declare Sub sRedoColList 
Declare Sub	EnableSortDialogPart2	

'sort_mTools.mb
Declare Sub SortDlg
Declare Sub mSortColList
Declare Sub SortDlgPart2
Declare Sub SortDlgPart3
Declare Sub SortDlgPart4
Declare Sub SortDlgPart5

'worpath_mTools.mb
Declare Sub wToolHandler

'label_mTools.mb
Declare Sub LabelDialog			
Declare Sub lRedoColList 		
Declare Sub LabelDialogPart2
Declare Sub LabelDialogPart3
Declare Sub LabelDialogPart4

'syncfrm_mTools.mb
Declare Sub mSyncFrame

'mpack_mTools.mb
Declare Sub PackDlg

'selstyle_mTools.mb
Declare Sub Buttonhandler
Declare Sub HowTo()
Declare Sub Handle_Lines(ByVal table_name As String, ByVal pstyle As Pen) 
Declare Sub Handle_Areas(ByVal table_name As String, ByVal bstyle As Brush, ByVal pstyle As Pen) 
Declare Sub Handle_Points(ByVal table_name As String, ByVal sstyle As Symbol)
Declare Sub Handle_bitmap_Points(ByVal table_name As String, ByVal sstyle As Symbol)
Declare Sub Handle_truetype_Points(ByVal table_name As String, ByVal sstyle As Symbol)
Declare Sub Handle_mi3_Points(ByVal table_name As String, ByVal sstyle As Symbol)
Declare Sub Handle_Text(ByVal table_name As String, ByVal fstyle As Font) 
Declare Sub UserCanceled 

'qcleaner_mTools.mb
Declare Sub QClean
Declare Sub QCleaner

'shortcuts_mTools.mb
Declare Sub mShort
Declare Sub ExitShort

'uptext_mTools.mb
Declare Sub upText
Declare Sub tRedoColList

'about_mTools.mb
Declare Sub aboutprog

'mreproj_mTools.mb
Declare Sub ReProjTab
Declare Sub ChangeProj
Declare Sub rpDlgLst
Declare Sub rpDlgLst2
Declare Sub rpDlgLst3

'autoedit_mTools.mb
Declare Sub tbHandler
Declare Sub SetTBStat (ByVal stat As Logical)
Declare Sub SelChangedHandler
Declare Sub LayerEdit

'func_mTools.mb
Declare Function GetFolderPath(sOutPath as string) as integer
Declare Function TableOpen(ByVal TabName As String) As Logical
Declare Function Trim(ByVal theStr As String) As String
Declare Function CountChar(ByVal sStr, ByVal sChar As String,ByVal lCaseSens As Logical) As SmallInt
Declare Function ReverseString(ByVal InputString As String) As String
Declare Function Asc2Char() As String
Declare Function getHDD() As Integer

'reg_mTools.mb
Declare Function GetRegistryString(ByVal hKey As Integer, ByVal sEntry As String) As String
Declare Function GetRegistryInt(ByVal hKey As Integer, ByVal sEntry As String) As Integer
Declare Sub SetRegistryString(ByVal hKey As Integer, ByVal sEntry As String, ByVal sEntryString As String)
Declare Sub SetRegistryInt(ByVal hKey As Integer, ByVal sEntry As String, ByVal iEntryInt As Integer)
Declare Sub WriteRegistryStr(ByVal Group As Integer, ByVal Section As String, ByVal Key As String, ByVal ValType As SmallInt, ByVal sValue As String)
Declare Sub WriteRegistryInt(ByVal Group As Integer, ByVal Section As String, ByVal Key As String, ByVal ValType As SmallInt, ByVal iValue As Integer)
Declare Sub DeleteRegKey(ByVal Group As Integer, ByVal Section As String, ByVal Key As String)

' graticule.mb
Declare Sub grtbhandler
Declare Sub grTBStat (ByVal stat As Logical)
Declare Sub maingrat
Declare Sub WinChangedHandler
Declare Sub draw_graticule(graticule_index as integer)
Declare Sub graticule_M_MAP_OPTIONS_handler
Declare Sub graticule_debug(ByVal debug_level as integer, ByVal debug_message as string)
Declare Sub check_graticule_coordsys(target_graticule as integer)
Declare Sub update_graticule_coordsys(target_graticule as integer)
Declare Sub update_graticule_bounds(index_no as integer)
Declare Sub set_coordsys_GraticuleCoordsys(index_no as integer)
Declare Sub add_graticule_handler
Declare Sub remove_graticule_handler
Declare Sub remove_graticule(index_no as integer)
' Declare Sub exit_graticule_handler
' Declare Sub about_graticule_handler
Declare Sub EndHandler
Declare Sub WinClosedHandler
Declare Sub global_options_menu_handler
Declare Sub local_graticule_options_menuhandler
Declare Sub GridOptions_OKButton_Handler
Declare Function local_graticule_options_dialog(ByVal index_no as integer) as smallint
Declare Function add_new_graticule_layer(target_window as integer, index_no as integer) as logical
Declare Function get_graticule_index(ByVal mapwinid as integer) as integer
Declare function create_straight_autogrid_line(byval from_x as float, byval from_y as float, byval to_x as float, byval to_y as float, byval line_type as integer, byval graticule_index as integer, byval label_positioning as integer) as logical
Declare sub create_straight_autogrid_line_nocheck(byval from_x as float, byval from_y as float, byval to_x as float, byval to_y as float, byval line_type as integer, byval graticule_index as integer)
Declare function create_circle_autogrid_line(byval lat as float, graticule_index as integer, from_x as float) as logical
Declare FUNCTION create_warped_autogrid_line(byval from_x as float, 
            byval from_y as float, 
            byval to_x as float, 
            byval to_y as float, 
            byval line_type as integer, 
            byval grid_interval as float,
            byval labelling_object as object,
            byval graticule_index as integer) AS LOGICAL
Declare function format_label(line_value as float, byval xy as integer, byval units as string) as string
Declare function decdeg2d_dm(byval decdeg as float, byval latlong as integer) as string


' misc_tools.mb
Declare Function TableAvailable(ByVal tablename as string) as logical
Declare Function get_temporary_tablename() as string
Declare Function get_coordsys_projection(ByVal coordsys_string as string) as integer
Declare Function get_coordsys_gridunits(ByVal coordsys_string as string) as string
Declare Function get_free_fileno() as integer
Declare Function get_layerno_named(ByVal win_id as integer, ByVal layername as string) as smallint
Declare Function FrontMapperID() as integer
Declare Function extract_delimitedlist_item_no(ByVal list_str As String, ByVal delimiter_str as string, ByVal target_item_no as integer) as string
Declare Function find_nth_occurrence(ByVal search_string as string, ByVal target_string as string, ByVal n as integer) as integer
Declare Function AntiCW_bearing_dec_ll(	ByVal d_lat1 as float, ByVal d_lon1 as float, 
				ByVal d_lat2 as float, ByVal d_lon2 as float, bearing as float) as logical
Declare Function get_coordsys_datum(ByVal coordsys_string as string) as integer
Declare Function parse_coordsys_projection_clause(ByVal	coordsys_string as string,
				projection_no as integer, datum_no as integer, units as string) as logical
Declare Function DatumNo2DatumName(ByVal datum_no as integer) as string
Declare Sub Debug(ByVal this_debug_level as integer, ByVal debug_message as string, ByVal current_active_debug_level as integer)

' projections.mb
Declare Sub update_default_mapper_gridlines(index_no as integer)
Declare Sub update_PolarAzimuthal_gridlines(index_no as integer)
Declare Sub update_latlongNZMG_gridlines(graticule_index as integer)
Declare Sub determine_critical_intersections(graticule_index as integer, cmig as logical, npig as logical, spig as logical, eqig as logical)
Declare Function get_first_grid(win_boundary as float, grid_interval as float) as float
'Declare Function get_grid_interval(ByVal win_siz as float, ByVal interval_units as string) as float
Declare Function get_grid_interval(ByVal win_siz as float, ByVal interval_units as string, ByVal grid_interval_denominator as smallint) as float
Declare Function max_noof_intervals(interval_units as string) as integer
Declare Function get_deg_interval_indexed(i as integer) as float
Declare Function get_m_interval_indexed(i as integer) as float
Declare Function get_interval_indexed(interval_index as integer, interval_units as string) as float

'mCAD_mTools.mb
Declare Sub mCAD
Declare Sub SelFileDlg
Declare Sub cadDlg
Declare Sub cadDlg2
Declare Sub cadDlg3
Declare Sub CADProj
Declare Sub ScoreNonAlpha(sStr As String)
Declare Sub write_out

'mCAD_mTools.mb
Declare Sub mGDB
Declare Sub gdbFileDlg
Declare Sub gdbOut

' .Net call
Declare Method FILEBrowseForFolder
        Class "mLibrary.winLib"  Lib "mLibrary.dll" Alias "BrowseForFolder"
        (ByVal sDescription As String, 'Text to display in the dialog
        ByVal sFolder As String) 'Start folder to use in the dialog
        As String 'Return the folder selected, or "" if the dialog was cancelled
Declare Method FILEOpenFilesDlg
        Class "mLibrary.winLib"  Lib "mLibrary.dll" Alias "OpenFilesDlg"
        (ByVal sTitle As String 'Title of FileOpen dialog
        , ByVal sStartFolder As String 'Folder the dialog show on startup
        , ByVal sFileTypes As String 'List of filetypes that are available in the FileType dropdown list
        'Should be in this form:
        ' description | filter [; filter ...] [| description | filter  [; filter ...]]
        'Single filter: "Text Files (*.txt)|*.txt"
        'Multiple filters: "Tables (*.tab)|*.tab|Workspaces (*.wor)|*.wor"
        'Multiple filetypes at a time: "MapInfo files (*.tab,*.wor)|*.tab;*.wor"
        ) As Integer 'Returns the number of files selected by the user
        ' use GetOpenFilesDlgFileName to get each file item
Declare Method FILEGetOpenFilesDlgFileNames
        Class "mLibrary.winLib"  Lib "mLibrary.dll" Alias "GetOpenFilesDlgFileNames"
        (  arrFileNames() As String 'Array to hold all the filenames selected
        ) As Integer 'Returns full the number of files returned
Declare Method worReg Class "mLibrary.winLib" Lib "mLibrary.dll" (ByVal sKey As String, ByVal sVal As String) As String
Declare Method getReg Class "mLibrary.winLib" Lib "mLibrary.dll" (ByVal sKey As String, ByVal sVal As String) As String
Declare Method GetVersion Class "mLibrary.winLib" Lib "mLibrary.dll" As String

'Windows API
Declare Function SetCurrentDirectoryA Lib "Kernel32.dll" (dir As String) As SmallInt
Declare Function ExecuteAndWait(ByVal cmdLine as string) as integer
Declare Function CloseHandle Lib "kernel32" (hObject As Integer) As smallint
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer,ByVal dwMilliseconds As Integer) As Integer
Declare Function CreateProcessA Lib "kernel32"(ByVal lpApplicationName As Integer,ByVal lpCommandLine As String,
        ByVal lpProcessAttributes As Integer,ByVal lpThreadAttributes As Integer,ByVal bInheritHandles As Integer,
        ByVal dwCreationFlags As Integer,ByVal lpEnvironment As Integer,ByVal lpCurrentDirectory As Integer,
        lpStartupInfo As STARTUPINFO,lpProcessInformation As PROCESS_INFORMATION) As Integer
Declare Function CreateDirectory32 lib "Kernel32.dll" alias "CreateDirectoryA"
        (ByVal sPath as string, tSecurity as SECURITY_ATTRIBUTES) as integer
Declare Function RemoveDirectory32 lib "Kernel32.dll" alias "RemoveDirectoryA"
        (ByVal sPath as string) as integer
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String,
        ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Integer,
        lpMaximumComponentLength As Integer, lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As String,
        ByVal nFileSystemNameSize As Integer) As Integer 
        
'registry API
Declare Function RegOpenKey Lib "ADVAPI32.DLL" Alias "RegOpenKeyA"
        (ByVal hKey As Integer, ByVal SubKey As String, pResult As Integer) As Integer
Declare Function RegCreateKey Lib "ADVAPI32.DLL" Alias "RegCreateKeyA"
        (ByVal hKey As Integer, ByVal SubKey As String, pResult As Integer) As Integer
Declare Function RegCloseKey Lib "ADVAPI32.DLL"
        (ByVal hKey As Integer) As Integer
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" 
        (ByVal hKey As Integer, ByVal dwIndex As Integer, lpName As String, lpcbName As Integer,
        ByVal lpReserved As Integer, ByVal lpClass As String, lpcbClass As Integer, lpftLastWriteTime As FILETIME) As Integer
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" 
        (ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpValueName As String, lpcbValueName As Integer,
        ByVal lpReserved As Integer, lpType As Integer, lpData As SmallInt, lpcbData As Integer) As Integer
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA"
        (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer,
        phkResult As Integer) As Integer 
Declare Function RegQueryValueEx Lib "ADVAPI32.DLL" Alias "RegQueryValueExA"
        (ByVal hKey As Integer, ByVal ValueName As String,ByVal Res1 As Integer, EntryType As Integer, EntryVal As String,
        lpcbData As Integer) As Integer
Declare Function RegQueryNumberEx Lib "ADVAPI32.DLL" Alias "RegQueryValueExA"
        (ByVal hKey As Integer, ByVal ValueName As String, ByVal Res1 As Integer, EntryType As Integer,
        NumVal As Integer, lpcbData As Integer) As Integer
Declare Function RegQueryValueType Lib "ADVAPI32.DLL" Alias "RegQueryValueExA"
        (ByVal hKey As Integer, ByVal ValueName As String, ByVal Res1 As Integer, EntryType As Integer,
        ByVal EntryVal As Integer, lpcbData As Integer) As Integer
Declare Function RegSetValueEx Lib "ADVAPI32.DLL" Alias "RegSetValueExA"
        (ByVal hKey As Integer, ByVal ValueName As String, ByVal Res1 As Integer, ByVal EntryType As Integer,
        ByVal EntryVal As String, ByVal lpcbData As Integer) As Integer
Declare Function RegSetNumberEx Lib "ADVAPI32.DLL" Alias "RegSetValueExA"
        (ByVal hKey As Integer, ByVal ValueName As String, ByVal Res1 As Integer, ByVal EntryType As Integer,
        EntryVal As Integer, ByVal lpcbData As Integer) As Integer
Declare Function RegDeleteKey Lib "ADVAPI32.dll" Alias "RegDeleteKeyA" 
        (ByVal hKey As Integer, ByVal lpSubKey As String) As Integer

' Help File declaration
Declare sub MnuHelp
Declare sub MnuHelpTS
Declare sub HelpHTML
Declare Sub HelpShow (
	byval sHelpFile as string,
	byval sTopic as string)

Declare Function WinExec Lib "kernel32"  Alias "WinExec" (
	ByVal lpCmdLine As String, 
	ByVal nCmdShow As Integer)
	As Integer
Declare Function FindExecutable Lib "shell32.dll" 
	Alias "FindExecutableA" (
	ByVal lpFile As String, 
	ByVal lpDirectory As String, 
	lpResult As String) As integer	
	
Include "mGlobal.glb"
