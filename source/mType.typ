'===============================================================================
'**
'**		Author				: Montazery Hasibuan
'**		Email				: --insert your email here--
'**
'**		Filename			: mType.typ
'**		Date Created		: 17 Nov 2010
'**		Last Modified		: 2014-06-14 10:11 AM
'**

'===============================================================================

'Pendeklarasian variabel custom type

Type FILETIME
     dwLowDateTime As Integer
     dwHighDateTime As Integer
End Type

Type BrowseInfo
		 hWndOwner As integer
		 pIDLRoot As integer
		 pszDisplayName As integer
		 lpszTitle As integer
		 ulFlags As integer
		 lpfnCallback As integer
		 lParam As integer
		 iImage As integer
End Type

'Type Tmap_graticule_ref
'	map_win_id as integer
'	win_minx as float
'	win_miny as float
'	win_maxx as float
'	win_maxy as float
'	tablename as string
'	coordsys_string as string
'	projection as integer
'	gridunits as string
'	warped as logical
'	mapper_bounds_obj as object
'	gridinterval_denom as smallint
'	gridpen as pen
'	mapper_bounds_obj_latlong as object
'End Type

Type STARTUPINFO
	cb As Integer
	lpReserved As String
	lpDesktop As String
	lpTitle As String
	dwX As Integer
	dwY As Integer
	dwXSize As Integer
	dwYSize As Integer
	dwXCountChars As Integer
	dwYCountChars As Integer
	dwFillAttribute As Integer
	dwFlags As Integer
	wShowWindow As Smallint
	cbReserved2 As Smallint
	lpReserved2 As Integer
	hStdInput As Integer
	hStdOutput As Integer
	hStdError As Integer
End Type

Type PROCESS_INFORMATION
hProcess As Integer
hThread As Integer
dwProcessID As Integer
dwThreadID As Integer
End Type

Type SECURITY_ATTRIBUTES
  nLength As Integer              'DWORD in Windows
  lpSecurityDescriptor As Integer 'LPVOID in Windows
  bInheritHandle As Integer       'BOOL in Windows
End Type
