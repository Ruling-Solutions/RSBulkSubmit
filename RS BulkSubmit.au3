#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=RS BulkSubmit.ico
#AutoIt3Wrapper_Outfile=RS BulkSubmit.exe
#AutoIt3Wrapper_Outfile_x64=RS BulkSubmit64.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Comment=Report suspicious malware or false positives to multiple antivirus vendors.
#AutoIt3Wrapper_Res_Description=Report suspicious malware or false positives to multiple antivirus vendors.
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=RS BulkSubmit
#AutoIt3Wrapper_Res_ProductVersion=1.0
#AutoIt3Wrapper_Res_CompanyName=Ruling Solutions
#AutoIt3Wrapper_Res_LegalCopyright=© 2022, Ruling Solutions
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /so
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt('MustDeclareVars', 1)
Opt("TrayIconHide", 1)

#include <Crypt.au3>
#include <GuiButton.au3>
#include <GuiComboBox.au3>
#include <GuiListView.au3>
#include <Misc.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include '..\Shared\RS_Encoding.au3'
#include '..\Shared\RS_Environ.au3'
#include '..\Shared\RS_INI.au3'
#include '..\Shared\RS_SMTP.au3'

; Global constants and variables
Global Const $MinWidth = 600   ; Minimum window width
Global Const $MinHeight = 380  ; Minimum window height
Global $gDblClick = False      ; Boolean state of double clicks in ListView.
Global $gGetEmails = True      ; Boolean state to activate/deactivate retrieve of selected emails.
Global $hLstVendors            ; Vendors ListView handle

Local $hGUI                    ; Main GUI handle
Local $hItmVendors             ; Individual vendor item
Local $hTxtTo                  ; Main recipient address textbox
Local $hTxtBCC                 ; BCC addresses textbox
Local $hChkAll                 ; Check all vendor handle

; Local constants and variables
Local Const $AppName = _fileNameInfo(@ScriptName)
Local Const $AppPath = _fileNameInfo(@ScriptFullPath, 12)
Local Const $PortableMode = _portableMode()
Local Const $UserLocalPath = _addSlash(ENVIRON_specialFolder('LocalAppData')) & 'Ruling Solutions\RS BulkSubmit\'
Local Const $Key = 'RS_SMTP_key'

Local $intFontSize
Local $intVendorCount
Local $strWorkingPath
Local $strSample
Local $strLNGFile

Local $boolSelectAll = False
Local $intHomePageIcon
Local $strFontName
Local $strSelected
Local $strDefaultSource
Local $strListGuide
Local $strListBefore
Local $strListAfter
Local $strMalwareEmailMark
Local $strMalwareSubmitMark
Local $strFPEmailMark
Local $strFPSubmitMark
Local $strHomePageMark
Local $str7ZEXE
Local $str7ZSample
Local $str7ZAdd
Local $str7ZCheck
Local $str7ZVerify
Local $str7ZEncrypted
Local $str7ZNotEncrypted
Local $str7ZSuccess

; Email variables
Local $boolSSL
Local $boolTLS
Local $intPort
Local $intPriority
Local $intSampleType
Local $strSMTPServer
Local $strUser
Local $strPassword
Local $strName
Local $strEmail
Local $strMalwareSubject
Local $strFPSubject
Local $strBody

; Set default escape sequence replacement of INI values to False
INI_replacedEscaped(False)

; Check INI and LST files according to portable mode
Local Const $strINIFile = $PortableMode ? $AppPath & $AppName & '.ini' : $UserLocalPath & $AppName & '.ini'
Local Const $strLSTFile = $PortableMode ? $AppPath & $AppName & '.lst' : $UserLocalPath & $AppName & '.lst'
If FileExists($strINIFile) Then
	$strLNGFile = _removeExt(ENVIRON_replace(INI_valueLoad($strINIFile, 'General', 'Language', 'English'))) & '.lng'
	Opt('GUICloseOnESC', INI_valueLoad($strINIFile, 'General', 'EscToExit', '1') = 1 ? 1 : 0)
	$intHomePageIcon = INI_valueLoad($strINIFile, 'General', 'HomePageIcon', '0')
	$strFontName = INI_valueLoad($strINIFile, 'General', 'FontName', 'Segoe UI')
	$intFontSize = Number(INI_valueLoad($strINIFile, 'General', 'FontSize', '8.5'))
	$str7ZEXE = ENVIRON_replace(INI_valueLoad($strINIFile, 'Packer', 'EXE', ''))
	$str7ZSample = _fileNameInfo(ENVIRON_replace(INI_valueLoad($strINIFile, 'Packer', 'ArchiveName', 'FileSample')), 13)
	$str7ZAdd = INI_valueLoad($strINIFile, 'Packer', 'Add', 'a -pinfected "%ENCRYPTED_ARCHIVE%" "%FILE_SAMPLE%"')
	$str7ZCheck = INI_valueLoad($strINIFile, 'Packer', 'CheckEncryption', 'l -slt "%ENCRYPTED_ARCHIVE%"')
	$str7ZVerify = INI_valueLoad($strINIFile, 'Packer', 'VerifyEncryption', 't -pinfected "%ENCRYPTED_ARCHIVE%"')
	$str7ZEncrypted = INI_valueLoad($strINIFile, 'Packer', 'Encrypted', 'Encrypted = +')
	$str7ZNotEncrypted = INI_valueLoad($strINIFile, 'Packer', 'NotEncrypted', 'Encrypted = -')
	$str7ZSuccess = INI_valueLoad($strINIFile, 'Packer', 'Success', 'Everything is Ok')
	$strSelected = INI_valueLoad($strINIFile, 'Vendors', 'Selected', '')
	$strDefaultSource = INI_valueLoad($strINIFile, 'Vendors', 'DefaultSource', 'https://www.techsupportalert.com/how-to-report-malware-or-false-positives-to-multiple-antivirus-vendors/')
	$strListGuide = INI_valueLoad($strINIFile, 'Vendors', 'ListGuide', '<th width="40%">Submit False Positives</th>')
	$strListBefore = INI_valueLoad($strINIFile, 'Vendors', 'ListBefore', '<strong>')
	$strListAfter = INI_valueLoad($strINIFile, 'Vendors', 'ListAfter', '</table>')
	$strMalwareEmailMark = INI_valueLoad($strINIFile, 'Vendors', 'MalwareEmailMark', 'Report Malware via Email')
	$strMalwareSubmitMark = INI_valueLoad($strINIFile, 'Vendors', 'MalwareSubmitMark', 'malware')
	$strFPEmailMark = INI_valueLoad($strINIFile, 'Vendors', 'FalsePositiveEmailMark', 'Report False Positive via Email')
	$strFPSubmitMark = INI_valueLoad($strINIFile, 'Vendors', 'FalsePositiveSubmitMark', 'false positive')
	$strHomePageMark = INI_valueLoad($strINIFile, 'Vendors', 'HomePageMark', '>HomePage<')
	$strSMTPServer = _decrypt(_fromHex(INI_valueLoad($strINIFile, 'SMTP', 'Server', '')), $Key)
	$strUser = _decrypt(_fromHex(INI_valueLoad($strINIFile, 'SMTP', 'User', '')), $Key)
	$strPassword = _decrypt(_fromHex(INI_valueLoad($strINIFile, 'SMTP', 'Password', '')), $Key)
	$intPort = Number(INI_valueLoad($strINIFile, 'SMTP', 'Port', '25'))
	$boolSSL = INI_valueLoad($strINIFile, 'SMTP', 'SSL', '0') = 1 ? True : False
	$boolTLS = INI_valueLoad($strINIFile, 'SMTP', 'TLS', '0') = 1 ? True : False
	$intSampleType = (INI_valueLoad($strINIFile, 'Message', 'SampleType', '0') = 0 ? 0 : 1)
	$strName = INI_valueLoad($strINIFile, 'Message', 'SenderName', '')
	$strEmail = INI_valueLoad($strINIFile, 'Message', 'SenderEmail', '')
	$intPriority = Number(INI_valueLoad($strINIFile, 'Message', 'Priority', '1'))
	$strMalwareSubject = INI_valueLoad($strINIFile, 'Message', 'MalwareSubject', 'Suspicious File Submission')
	$strFPSubject = INI_valueLoad($strINIFile, 'Message', 'FalsePositiveSubject', 'False Positive Submission')
	$strBody = INI_valueLoad($strINIFile, 'Message', 'Body', 'Sample is in a password protected zip file.\nPassword for attachment is infected.', True)
Else
	$strLNGFile = INI_valueWrite($strINIFile, 'General', 'Language', 'English') & '.lng'
	Opt('GUICloseOnESC', INI_valueWrite($strINIFile, 'General', 'EscToExit', '1'))
	$intHomePageIcon = INI_valueWrite($strINIFile, 'General', 'HomePageIcon', '0')
	$strFontName = INI_valueWrite($strINIFile, 'General', 'FontName', 'Segoe UI')
	$intFontSize = Number(INI_valueWrite($strINIFile, 'General', 'FontSize', '8.5'))
	$str7ZEXE = ENVIRON_replace(INI_valueWrite($strINIFile, 'Packer', 'EXE', FileExists($AppPath & '7z.exe') ? '%APP_DIR%\7z.exe' : '', False))
	$str7ZSample = INI_valueWrite($strINIFile, 'Packer', 'ArchiveName', 'FileSample')
	$str7ZAdd = INI_valueWrite($strINIFile, 'Packer', 'Add', 'a -pinfected "%ENCRYPTED_ARCHIVE%" "%FILE_SAMPLE%"')
	$str7ZCheck = INI_valueWrite($strINIFile, 'Packer', 'CheckEncryption', 'l -slt "%ENCRYPTED_ARCHIVE%"')
	$str7ZVerify = INI_valueWrite($strINIFile, 'Packer', 'VerifyEncryption', 't -pinfected "%ENCRYPTED_ARCHIVE%"')
	$str7ZEncrypted = INI_valueWrite($strINIFile, 'Packer', 'Encrypted', 'Encrypted = +')
	$str7ZNotEncrypted = INI_valueWrite($strINIFile, 'Packer', 'NotEncrypted', 'Encrypted = -')
	$str7ZSuccess = INI_valueWrite($strINIFile, 'Packer', 'Success', 'Everything is Ok')
	$strSelected = INI_valueWrite($strINIFile, 'Vendors', 'Selected', '')
	$strDefaultSource = INI_valueWrite($strINIFile, 'Vendors', 'DefaultSource', 'https://www.techsupportalert.com/how-to-report-malware-or-false-positives-to-multiple-antivirus-vendors/')
	$strListGuide = INI_valueWrite($strINIFile, 'Vendors', 'ListGuide', '<th width="40%">Submit False Positives</th>')
	$strListBefore = INI_valueWrite($strINIFile, 'Vendors', 'ListBefore', '<strong>')
	$strListAfter = INI_valueWrite($strINIFile, 'Vendors', 'ListAfter', '</table>')
	$strMalwareEmailMark = INI_valueWrite($strINIFile, 'Vendors', 'MalwareEmailMark', 'Report Malware via Email')
	$strMalwareSubmitMark = INI_valueWrite($strINIFile, 'Vendors', 'MalwareSubmitMark', 'malware')
	$strFPEmailMark = INI_valueWrite($strINIFile, 'Vendors', 'FalsePositiveEmailMark', 'Report False Positive via Email')
	$strFPSubmitMark = INI_valueWrite($strINIFile, 'Vendors', 'FalsePositiveSubmitMark', 'false positive')
	$strHomePageMark = INI_valueWrite($strINIFile, 'Vendors', 'HomePageMark', '>HomePage<')
	$strSMTPServer = INI_valueWrite($strINIFile, 'SMTP', 'Server', '')
	$strUser = INI_valueWrite($strINIFile, 'SMTP', 'User', '')
	$strPassword = INI_valueWrite($strINIFile, 'SMTP', 'Password', '')
	$intPort = Number(INI_valueWrite($strINIFile, 'SMTP', 'Port', '25'))
	$boolSSL = INI_valueWrite($strINIFile, 'SMTP', 'SSL', '0') = 1 ? True : False
	$boolTLS = INI_valueWrite($strINIFile, 'SMTP', 'TLS', '0') = 1 ? True : False
	$intSampleType = (INI_valueWrite($strINIFile, 'Message', 'SampleType', '0') = 0 ? 0 : 1)
	$strName = INI_valueWrite($strINIFile, 'Message', 'SenderName', '')
	$strEmail = INI_valueWrite($strINIFile, 'Message', 'SenderEmail', '')
	$intPriority = Number(INI_valueWrite($strINIFile, 'Message', 'Priority', '1'))
	$strMalwareSubject = INI_valueWrite($strINIFile, 'Message', 'MalwareSubject', 'Suspicious File Submission')
	$strFPSubject = INI_valueWrite($strINIFile, 'Message', 'FalsePositiveSubject', 'False Positive Submission')
	$strBody = INI_valueWrite($strINIFile, 'Message', 'Body', 'Sample is in a password protected zip file.' & @CRLF & 'Password for attachment is infected.', True)
EndIf

; Check language and database files
If StringInStr($strLNGFile, '\') = 0 Then $strLNGFile = $AppPath & $strLNGFile
If Not FileExists($strLNGFile) Then
	$strLNGFile = $AppPath & 'English.lng'
	INI_valueWrite($strLNGFile, 'Main', '001', 'File')
	INI_valueWrite($strLNGFile, 'Main', '002', 'Open')
	INI_valueWrite($strLNGFile, 'Main', '003', 'Vendors')
	INI_valueWrite($strLNGFile, 'Main', '004', 'Search')
	INI_valueWrite($strLNGFile, 'Main', '005', 'Active')
	INI_valueWrite($strLNGFile, 'Main', '006', 'Discontinued')
	INI_valueWrite($strLNGFile, 'Main', '007', 'Name')
	INI_valueWrite($strLNGFile, 'Main', '008', 'Malware Email')
	INI_valueWrite($strLNGFile, 'Main', '009', 'Malware Submission URL')
	INI_valueWrite($strLNGFile, 'Main', '010', 'False Positive Email')
	INI_valueWrite($strLNGFile, 'Main', '011', 'False Positive Submission URL')
	INI_valueWrite($strLNGFile, 'Main', '012', 'HomePage')
	INI_valueWrite($strLNGFile, 'Main', '013', 'Comments')
	INI_valueWrite($strLNGFile, 'Main', '014', 'Load default')
	INI_valueWrite($strLNGFile, 'Main', '015', 'SMTP options')
	INI_valueWrite($strLNGFile, 'Main', '016', 'Server')
	INI_valueWrite($strLNGFile, 'Main', '017', 'Port')
	INI_valueWrite($strLNGFile, 'Main', '018', 'Use SSL')
	INI_valueWrite($strLNGFile, 'Main', '019', 'Use TLS')
	INI_valueWrite($strLNGFile, 'Main', '020', 'User')
	INI_valueWrite($strLNGFile, 'Main', '021', 'Password')
	INI_valueWrite($strLNGFile, 'Main', '022', 'Sender')
	INI_valueWrite($strLNGFile, 'Main', '023', 'Name')
	INI_valueWrite($strLNGFile, 'Main', '024', 'Email')
	INI_valueWrite($strLNGFile, 'Main', '025', 'Recipients')
	INI_valueWrite($strLNGFile, 'Main', '026', 'To')
	INI_valueWrite($strLNGFile, 'Main', '027', 'BCC')
	INI_valueWrite($strLNGFile, 'Main', '028', 'Message')
	INI_valueWrite($strLNGFile, 'Main', '029', 'Subject')
	INI_valueWrite($strLNGFile, 'Main', '030', 'Body')
	INI_valueWrite($strLNGFile, 'Main', '031', 'Priority')
	INI_valueWrite($strLNGFile, 'Main', '032', 'Normal')
	INI_valueWrite($strLNGFile, 'Main', '033', 'High')
	INI_valueWrite($strLNGFile, 'Main', '034', 'Low')
	INI_valueWrite($strLNGFile, 'Main', '035', 'Sample type')
	INI_valueWrite($strLNGFile, 'Main', '036', 'Malware')
	INI_valueWrite($strLNGFile, 'Main', '037', 'False positive')
	INI_valueWrite($strLNGFile, 'Main', '038', 'Online submission')
	INI_valueWrite($strLNGFile, 'Main', '039', 'Send emails')
	INI_valueWrite($strLNGFile, 'Main', '040', 'Confirm reload')
	INI_valueWrite($strLNGFile, 'Main', '041', 'Reload default vendors?')
	INI_valueWrite($strLNGFile, 'Main', '042', 'Creating default vendor list')
	INI_valueWrite($strLNGFile, 'Open', '001', 'Select file')
	INI_valueWrite($strLNGFile, 'Open', '002', 'Executables')
	INI_valueWrite($strLNGFile, 'Open', '003', 'AutoIt scripts')
	INI_valueWrite($strLNGFile, 'Open', '004', 'Zip archives')
	INI_valueWrite($strLNGFile, 'Open', '005', 'All files')
	INI_valueWrite($strLNGFile, 'Save', '001', 'Archive sample exists')
	INI_valueWrite($strLNGFile, 'Save', '002', 'Archive sample already exists.\nOverwrite?')
	INI_valueWrite($strLNGFile, 'Save', '003', 'Save archive as')
	INI_valueWrite($strLNGFile, 'Save', '004', 'Zip archives')
	INI_valueWrite($strLNGFile, 'Save', '005', 'All files')
	INI_valueWrite($strLNGFile, 'Vendor', '001', 'Add vendor')
	INI_valueWrite($strLNGFile, 'Vendor', '002', 'Edit vendor')
	INI_valueWrite($strLNGFile, 'Vendor', '003', 'Delete vendor')
	INI_valueWrite($strLNGFile, 'Vendor', '004', 'Vendor name')
	INI_valueWrite($strLNGFile, 'Vendor', '005', 'Email for malware')
	INI_valueWrite($strLNGFile, 'Vendor', '006', 'Malware submission URL')
	INI_valueWrite($strLNGFile, 'Vendor', '007', 'Email for false positives')
	INI_valueWrite($strLNGFile, 'Vendor', '008', 'False positive submission')
	INI_valueWrite($strLNGFile, 'Vendor', '009', 'HomePage')
	INI_valueWrite($strLNGFile, 'Vendor', '010', 'Comments')
	INI_valueWrite($strLNGFile, 'Vendor', '011', 'Discontinued')
	INI_valueWrite($strLNGFile, 'Vendor', '012', 'Cancel')
	INI_valueWrite($strLNGFile, 'Vendor', '013', 'Save')
	INI_valueWrite($strLNGFile, 'Vendor', '014', 'Cancel operation?')
	INI_valueWrite($strLNGFile, 'Vendor', '015', 'Are you sure?')
	INI_valueWrite($strLNGFile, 'Error', '001', 'Error')
	INI_valueWrite($strLNGFile, 'Error', '002', 'Archiver "%1" not found.')
	INI_valueWrite($strLNGFile, 'Error', '003', 'Error creating list file %1.')
	INI_valueWrite($strLNGFile, 'Error', '004', 'No data retrieved from URL: %1.')
	INI_valueWrite($strLNGFile, 'Error', '005', 'No valid data found in URL: %1.')
	INI_valueWrite($strLNGFile, 'Error', '006', 'Error updating list file %1.')
	INI_valueWrite($strLNGFile, 'Error', '007', 'Field "%1" missing.')
	INI_valueWrite($strLNGFile, 'Error', '008', 'File sample not defined.\nSelect file or zip archive to proceed.')
	INI_valueWrite($strLNGFile, 'Error', '009', 'Select a valid encrypted archive or\nany other file for automatic creation.')
	INI_valueWrite($strLNGFile, 'Error', '010', 'File %1 is not an archive.')
	INI_valueWrite($strLNGFile, 'Error', '011', 'Archive %1 not encrypted.')
	INI_valueWrite($strLNGFile, 'Error', '012', 'Archive %1 with incorrect password.')
	INI_valueWrite($strLNGFile, 'Error', '013', 'Error deleting archive %1.')
	INI_valueWrite($strLNGFile, 'Error', '014', 'Error creating archive %1.')
	INI_valueWrite($strLNGFile, 'Error', '015', 'Error sending email.')
	INI_valueWrite($strLNGFile, 'Error', '016', 'Error code')
	INI_valueWrite($strLNGFile, 'Error', '017', 'Description')
	INI_valueWrite($strLNGFile, 'Error', '018', 'Error resolving metalink %1.')
	INI_valueWrite($strLNGFile, 'Error', '019', 'Term not found')
EndIf

; Tooltip headers
Local Const $VendorName = INI_valueLoad($strLNGFile, 'Main', '007', 'Name')
Local Const $VendorMalwareEmail = INI_valueLoad($strLNGFile, 'Main', '008', 'Email for malware')
Local Const $VendorMalwareSubmit = INI_valueLoad($strLNGFile, 'Main', '009', 'Malware Submission URL')
Local Const $VendorFPEmail = INI_valueLoad($strLNGFile, 'Main', '010', 'Email for false positives')
Local Const $VendorFPSubmit = INI_valueLoad($strLNGFile, 'Main', '011', 'False positive submission')
Local Const $VendorHomePage = INI_valueLoad($strLNGFile, 'Main', '012', 'HomePage')
Local Const $VendorComments = INI_valueLoad($strLNGFile, 'Main', '013', 'Comments')
Local Const $VendorAdd = INI_valueLoad($strLNGFile, 'Vendor', '001', 'Add vendor')
Local Const $VendorEdit = INI_valueLoad($strLNGFile, 'Vendor', '002', 'Edit vendor')
Local Const $VendorDel = INI_valueLoad($strLNGFile, 'Vendor', '003', 'Delete vendor')
Local Const $VendorWarning = INI_valueLoad($strLNGFile, 'Vendor', '014', 'Cancel operation?')
Local Const $VendorConfirm = INI_valueLoad($strLNGFile, 'Vendor', '015', 'Are you sure?')

; Check 7Z executable
If Not FileExists($str7ZEXE) Then
	MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
			StringReplace(INI_valueLoad($strLNGFile, 'Error', '002', 'Archiver "%1" not found.', True), '%1', $str7ZEXE), 3)
	Exit (1)
EndIf

; If vendors list doesn't exist, create a new one with default values
If Not FileExists($strLSTFile) Then
	Local $strDefault = getDefaultList($strDefaultSource)
	If IsArray($strDefault) Then
		saveList($strDefault)
	Else
		Exit (1)
	EndIf
EndIf

; Check variables
If Not IsArray(_WinAPI_EnumFontFamilies(0, $strFontName)) Then $strFontName = INI_valueWrite($strINIFile, 'General', 'FontName', 'Segoe UI')
If $intFontSize < 1 Or $intFontSize > 14 Then $intFontSize = Number(INI_valueWrite($strINIFile, 'General', 'FontSize', '8.5'))
If $intPriority < 0 Or $intPriority > 2 Then $intPriority = 0

; Main GUI
mainGUI()

#Region GUI PROCEDURES
; <== GUI PROCEDURES ==============================================================================
; <=== mainGUI ====================================================================================
; mainGUI()
; ; Main GUI events.
; ;
; ; @param  NONE
; ; @return NONE
Func mainGUI()
	; Create form and controls: $WS_SIZEBOX 0x00040000, $WS_MINIMIZEBOX 0x00020000, $WS_MAXIMIZEBOX 0x00010000
	Local Const $intWidth = 640
	Local Const $intHeight = 480
	$hGUI = GUICreate(StringReplace($AppName, '64', ' (x64)') & ' ' & _
			FileGetVersion(@ScriptFullPath), $intWidth, $intHeight, _
			@DesktopWidth / 2 - $intWidth / 2, @DesktopHeight / 2 - $intHeight / 2, _
			BitOR(0x00040000, 0x00020000, 0x00010000))

	; File field
	Local $hLblSample = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '001', 'File') & ':', 7, 9, 50, 18)
	Local $hTxtSample = GUICtrlCreateInput('', 60, 7, 485, 20, 0x0800)
	Local $hBtnSample = GUICtrlCreateButton(INI_valueLoad($strLNGFile, 'Main', '002', 'Search'), 545, 7, 85, 20)

	; Vendors list: $LVS_SHOWSELALWAYS 0x0008, ($LVS_EX_GRIDLINES 0x00000001, $LVS_EX_CHECKBOXES 0x00000004, $LVS_EX_FULLROWSELECT 0x00000020)
	Local $hLblVendors = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '003', 'Vendors'), 7, 33, 200, 20)
	Local $hLblSearch = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '004', 'Search') & ':', 7, 58, 50, 20)
	Local $hTxtSearch = GUICtrlCreateInput('', 57, 55, 150, 18)
	$hLstVendors = GUICtrlCreateListView(_StringRepeat(' ', 5) & $VendorName & '|' & _
			$VendorMalwareEmail & '|' & $VendorMalwareSubmit & '|' & _
			$VendorFPEmail & '|' & $VendorFPSubmit & '|' & _
			$VendorHomePage & '|' & $VendorComments, _
			7, 75, 200, 330, 0x0008, BitOR(0x00000001, 0x00000004, 0x00000020))
	$hChkAll = _GUICtrlButton_Create(_GUICtrlListView_GetHeader(GUICtrlGetHandle($hLstVendors)), '', 4, 3, 20, 18, $BS_AUTOCHECKBOX)

	Local $hBtnAdd = GUICtrlCreateButton('+', 7, 408, 20, 20)
	Local $hBtnEdit = GUICtrlCreateButton('!', 27, 408, 20, 20)
	Local $hBtnDel = GUICtrlCreateButton('-', 47, 408, 20, 20)
	Local $hBtnLoadDefault = GUICtrlCreateButton(INI_valueLoad($strLNGFile, 'Main', '014', 'Load default'), 83, 408, 105, 20)
	Local $hBtnDefault = GUICtrlCreateButton('N', 188, 408, 20, 20)

	; SMTP box: $ES_NUMBER 0x2000, $ES_PASSWORD 0x0020
	Local $hBoxSMTP = GUICtrlCreateGroup(INI_valueLoad($strLNGFile, 'Main', '015', 'SMTP options'), 210, 32, 420, 70)
	Local $hLblServer = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '016', 'Server') & ':', 217, 52, 53, 18)
	Local $hTxtSMTPServer = GUICtrlCreateInput($strSMTPServer, 270, 50, 180, 20)
	Local $hLblPort = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '017', 'Port') & ':', 455, 52, 40, 18)
	Local $hTxtPort = GUICtrlCreateInput($intPort, 495, 50, 50, 20, 0x2000)
	Local $hChkSSL = GUICtrlCreateCheckbox(INI_valueLoad($strLNGFile, 'Main', '018', 'Use SSL'), 555, 50, 70, 20)
	Local $hLblUser = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '020', 'User') & ':', 217, 77, 53, 18)
	Local $hTxtUser = GUICtrlCreateInput($strUser, 270, 75, 75, 20)
	Local $hLblPassword = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '021', 'Password') & ':', 350, 77, 80, 18)
	Local $hTxtPassword = GUICtrlCreateInput($strPassword, 430, 75, 115, 20, 0x0020)
	Local $hChkTLS = GUICtrlCreateCheckbox(INI_valueLoad($strLNGFile, 'Main', '019', 'Use TLS'), 555, 75, 70, 20)
	GUICtrlCreateGroup('', -99, -99, 1, 1)

	; Sender box
	Local $hBoxSender = GUICtrlCreateGroup(INI_valueLoad($strLNGFile, 'Main', '022', 'Sender'), 210, 105, 420, 45)
	Local $hLblName = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '023', 'Name') & ':', 217, 122, 55, 18)
	Local $hTxtName = GUICtrlCreateInput($strName, 275, 120, 120, 20)
	Local $hLblEmail = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '024', 'Email') & ':', 405, 122, 60, 18)
	Local $hTxtEmail = GUICtrlCreateInput($strEmail, 470, 120, 150, 20)
	GUICtrlCreateGroup('', -99, -99, 1, 1)

	; Recipients box
	Local $hBoxRecipients = GUICtrlCreateGroup(INI_valueLoad($strLNGFile, 'Main', '025', 'Recipients'), 210, 150, 420, 70)
	Local $hLblTo = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '026', 'To') & ':', 217, 167, 55, 18)
	$hTxtTo = GUICtrlCreateInput('', 275, 165, 348, 20, 0x0800)
	Local $hLblBCC = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '027', 'BCC') & ':', 217, 192, 55, 18)
	$hTxtBCC = GUICtrlCreateInput('', 275, 190, 348, 20, 0x0800)
	GUICtrlCreateGroup('', -99, -99, 1, 1)

	; Message box ($ES_MULTILINE 0x0004, $WS_EX_WINDOWEDGE 0x00000100, $ES_WANTRETURN 0x1000, $WS_VSCROLL 0x00200000),
	;             $CBS_DROPDOWNLIST 0x0003, $WS_EX_RIGHT 0x00001000
	Local $hBoxMessage = GUICtrlCreateGroup(INI_valueLoad($strLNGFile, 'Main', '028', 'Message'), 210, 220, 420, 208)
	Local $hLblSubject = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '029', 'Subject') & ':', 217, 237, 55, 18)
	Local $hTxtSubject = GUICtrlCreateInput('', 275, 235, 348, 20)
	Local $hLblBody = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '030', 'Body') & ':', 217, 265, 55, 18)
	Local $hTxtBody = GUICtrlCreateEdit($strBody, 275, 263, 348, 132, BitOR(0x0004, 0x00000100, 0x1000, 0x00200000))
	Local $hLblPriority = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '031', 'Priority') & ':', 217, 403, 55, 18)
	Local $hLstPriority = GUICtrlCreateCombo('', 275, 400, 68, 20, 0x0003)
	Local $hLblType = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Main', '035', 'Sample type') & ':', 400, 403, 120, 18, -1, 0x00001000)
	Local $hLstType = GUICtrlCreateCombo('', 523, 400, 100, 20, 0x0003)
	GUICtrlCreateGroup('', -99, -99, 1, 1)

	; Bottom buttons
	; HomePage icon: 1 = 'þ' World (America), = 'ü' World (EMEA), 3	= 'ý' World (APAC), Default = 'H' Home icon
	Local $hBtnHomePage = GUICtrlCreateButton($intHomePageIcon = 1 ? 'þ' : ($intHomePageIcon = 2 ? 'ü' : ($intHomePageIcon = 3 ? 'ý' : 'H')), 7, 430, 20, 20)
	Local $hBtnSubmit = GUICtrlCreateButton(INI_valueLoad($strLNGFile, 'Main', '038', 'Online submission'), 28, 430, 180, 20)
	Local $hBtnSend = GUICtrlCreateButton(INI_valueLoad($strLNGFile, 'Main', '039', 'Send emails'), 209, 430, 422, 20)

	; Define controls docking
	GUICtrlSetResizing($hLblSample, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblVendors, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblSearch, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblServer, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblPort, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblUser, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblPassword, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblName, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblEmail, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblTo, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblBCC, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblSubject, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblBody, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblPriority, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLblType, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

	GUICtrlSetResizing($hTxtSample, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
	GUICtrlSetResizing($hTxtSearch, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hTxtSMTPServer, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
	GUICtrlSetResizing($hTxtPort, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hTxtUser, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
	GUICtrlSetResizing($hTxtPassword, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hTxtName, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
	GUICtrlSetResizing($hTxtEmail, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hTxtTo, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
	GUICtrlSetResizing($hTxtBCC, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
	GUICtrlSetResizing($hTxtSubject, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
	GUICtrlSetResizing($hTxtBody, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

	GUICtrlSetResizing($hBoxSMTP, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBoxSender, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBoxRecipients, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBoxMessage, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

	GUICtrlSetResizing($hChkSSL, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hChkTLS, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)

	GUICtrlSetResizing($hLstVendors, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLstPriority, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hLstType, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

	GUICtrlSetResizing($hBtnSample, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBtnAdd, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBtnEdit, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBtnDel, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBtnLoadDefault, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBtnDefault, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBtnHomePage, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBtnSubmit, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
	GUICtrlSetResizing($hBtnSend, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

	; Set controls font
	GUICtrlSetFont($hLblSample, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblServer, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblPort, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblUser, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblPassword, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblName, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblEmail, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblTo, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblBCC, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblSubject, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblBody, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLblPriority, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtSample, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtSMTPServer, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtPort, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtUser, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtPassword, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtName, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtEmail, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtTo, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtBCC, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtSubject, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hTxtBody, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLstPriority, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hLstVendors, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hBtnSample, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hBtnAdd, 10, 400, 0)
	GUICtrlSetFont($hBtnEdit, 10, 400, 0, 'Wingdings')
	GUICtrlSetFont($hBtnDel, 10, 400, 0)
	GUICtrlSetFont($hBtnLoadDefault, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hBtnDefault, 10, 400, 0, 'Webdings')
	GUICtrlSetFont($hBtnHomePage, 10, 400, 0, 'Webdings')
	GUICtrlSetFont($hBtnSubmit, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hBtnSend, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hChkSSL, $intFontSize, 400, 0, $strFontName)
	GUICtrlSetFont($hChkTLS, $intFontSize, 400, 0, $strFontName)

	GUICtrlSetFont($hLblVendors, $intFontSize + 1, 700, 0, $strFontName)
	GUICtrlSetFont($hBoxSMTP, $intFontSize, 700, 2, $strFontName)
	GUICtrlSetFont($hBoxSender, $intFontSize, 700, 2, $strFontName)
	GUICtrlSetFont($hBoxRecipients, $intFontSize, 700, 2, $strFontName)
	GUICtrlSetFont($hBoxMessage, $intFontSize, 700, 2, $strFontName)

	; Set controls background
	GUICtrlSetBkColor($hTxtSample, 0xFBFBFB)
	GUICtrlSetBkColor($hTxtTo, 0xFBFBFB)
	GUICtrlSetBkColor($hTxtBCC, 0xFBFBFB)
	GUICtrlSetBkColor($hLstPriority, 0xFFFFFF)

	; Set ListView groups
	_GUICtrlListView_EnableGroupView($hLstVendors)
	_GUICtrlListView_InsertGroup($hLstVendors, -1, 0, '')
	_GUICtrlListView_InsertGroup($hLstVendors, -1, 1, '')
	_GUICtrlListView_SetGroupInfo($hLstVendors, 0, INI_valueLoad($strLNGFile, 'Main', '005', 'Active'), 0, $LVGS_COLLAPSIBLE)
	_GUICtrlListView_SetGroupInfo($hLstVendors, 1, INI_valueLoad($strLNGFile, 'Main', '006', 'Discontinued'), 0, $LVGS_COLLAPSIBLE)

	; Add data to ListView
	updateList(FileReadToArray($strLSTFile), False)

	; Update check boxes
	GUICtrlSetState($hChkSSL, $boolSSL ? 1 : 4)
	GUICtrlSetState($hChkTLS, $boolTLS ? 1 : 4)

	; Load priority list
	_GUICtrlComboBox_BeginUpdate($hLstPriority)
	GUICtrlSetData($hLstPriority, INI_valueLoad($strLNGFile, 'Main', '032', 'Low') & '|' & _
			INI_valueLoad($strLNGFile, 'Main', '033', 'Normal') & '|' & _
			INI_valueLoad($strLNGFile, 'Main', '034', 'High'))
	_GUICtrlComboBox_EndUpdate($hLstPriority)
	ControlCommand($hGUI, '', $hLstPriority, 'SetCurrentSelection', $intPriority)

	; Load sample type list
	_GUICtrlComboBox_BeginUpdate($hLstType)
	GUICtrlSetData($hLstType, INI_valueLoad($strLNGFile, 'Main', '036', 'Malware') & '|' & _
			INI_valueLoad($strLNGFile, 'Main', '037', 'False Positive'))
	_GUICtrlComboBox_EndUpdate($hLstType)
	ControlCommand($hGUI, '', $hLstType, 'SetCurrentSelection', $intSampleType)
	GUICtrlSetData($hTxtSubject, $intSampleType = 0 ? $strMalwareSubject : $strFPSubject)

	; Register custom windows messages
	GUIRegisterMsg($WM_GETMINMAXINFO, '_WM_GETMINMAXINFO')  ; Limitis to GUI resizing.
	GUIRegisterMsg($WM_NOTIFY, '_WM_NOTIFY')                ; Manage double click, show tooltips and get checked entries in ListView.

	; Show main GUI
	GUISetState(@SW_SHOW, $hGUI)

	; Other handles and variables
	Local $dlgOpenFile
	Local $dlgSaveFile
	Local $intCursorInfo
	Local $intSearch
	Local $strURL
	Local $strTerm

	; Form strings
	Local Const $ReloadTitle = INI_valueLoad($strLNGFile, 'Main', '040', 'Confirm reload')
	Local Const $ReloadMsg = INI_valueLoad($strLNGFile, 'Main', '041', 'Reload default vendors?')
	Local Const $ReloadTooltip = INI_valueLoad($strLNGFile, 'Main', '042', 'Creating default vendor list') & '...'

	; Main loop
	_GUICtrlListView_RegisterSortCallBack($hLstVendors)
	Local $hDLL = DllOpen('user32.dll')
	While 1
		Sleep(10)

		; Allow ToolTip only in ListView
		If WinActive($hGUI) Then
			$intCursorInfo = GUIGetCursorInfo($hGUI)
			If $intCursorInfo[4] <> $hLstVendors Then ToolTip('')
		Else
			ToolTip('')
		EndIf

		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				; Exit script
				ExitLoop

			Case $hBtnSample
				; Select file sample
				$strSample = getSample($strSample)
				GUICtrlSetData($hTxtSample, $strSample)

			Case $hTxtSearch
				; Avoid pipes in search term
				GUICtrlSetData($hTxtSearch, StringReplace(GUICtrlRead($hTxtSearch), '|', ''))

				;	Save fields changes to INI and update related fields
			Case $hTxtSMTPServer
				$strSMTPServer = updateKey($hTxtSMTPServer, 'SMTP', 'Server', True)
			Case $hTxtUser
				$strUser = updateKey($hTxtUser, 'SMTP', 'User', True)
			Case $hTxtPassword
				$strPassword = updateKey($hTxtPassword, 'SMTP', 'Password', True)
			Case $hTxtPort
				$intPort = updateKey($hTxtPort, 'SMTP', 'Port')
			Case $hTxtName
				$strName = updateKey($hTxtName, 'Message', 'SenderName')
			Case $hTxtEmail
				$strEmail = updateKey($hTxtEmail, 'Message', 'SenderEmail')
			Case $hTxtSubject
				If _GUICtrlComboBox_GetCurSel($hLstType) = 0 Then
					$strMalwareSubject = updateKey($hTxtSubject, 'Message', 'MalwareSubject')
				Else
					$strFPSubject = updateKey($hTxtSubject, 'Message', 'FalsePositiveSubject')
				EndIf
			Case $hTxtBody
				$strBody = updateKey($hTxtBody, 'Message', 'Body', False, True)
			Case $hChkSSL
				$boolSSL = (updateKey($hChkSSL, 'SMTP', 'SSL') = 1 ? True : False)
			Case $hChkTLS
				$boolTLS = (updateKey($hChkTLS, 'SMTP', 'TLS') = 1 ? True : False)
			Case $hLstPriority
				$intPriority = updateKey($hLstPriority, 'Message', 'Priority')
			Case $hLstType
				$intSampleType = updateKey($hLstType, 'Message', 'SampleType')
				GUICtrlSetData($hTxtSubject, $intSampleType = 0 ? $strMalwareSubject : $strFPSubject)

			Case $hLstVendors
				; Sort vendors list by columns
				_GUICtrlListView_SortItems($hLstVendors, GUICtrlGetState($hLstVendors))

			Case $hBtnAdd
				addVendor()

			Case $hBtnEdit
				editVendor()

			Case $hBtnDel
				delVendor()

			Case $hBtnLoadDefault
				; Load default vendors list
				If MsgBox(36, $ReloadTitle, $ReloadMsg, 0, $hGUI) = 6 Then
					updateList(getDefaultList($strDefaultSource, $ReloadTooltip))
					saveList()
				EndIf

			Case $hBtnDefault
				; Open default vendor list in default browser
				If StringLen($strDefaultSource) > 0 Then ShellExecute($strDefaultSource)

			Case $hBtnHomePage
				; Open selected vendor homepage in default browser
				$strURL = selectedEntry($hLstVendors, 5)
				If StringLen($strURL) > 0 Then ShellExecute($strURL)

			Case $hBtnSubmit
				; Open selected vendor submission page in default browser, according to sample type
				Local $intType = _GUICtrlComboBox_GetCurSel($hLstType) = 0 ? 2 : 4
				Local $strURL = selectedEntry($hLstVendors, $intType)
				If StringLen($strURL) > 0 Then
					If StringLeft($strURL, 1) = '#' Then
						; Resolve metalinks in URL
						Local $boolFound
						Local $intLen = StringLen($strURL) - 1
						Local $strOriginalURL = $strURL

						While StringLeft($strURL, 1) = '#'
							$boolFound = False
							For $i = 0 To $intVendorCount
								If StringInStr(_trimLeft(_GUICtrlListView_GetItemText($hLstVendors, $i, 0), '*'), StringMid($strURL, 2)) > 0 Then
									$strURL = _GUICtrlListView_GetItemText($hLstVendors, $i, $intType)
									$boolFound = True
									ExitLoop
								EndIf
							Next
							If Not $boolFound And $strURL = $strOriginalURL Then ExitLoop
						WEnd

						If $boolFound Then
							If StringLen($strURL) > 0 Then ShellExecute($strURL)
						Else
							MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
									StringReplace(INI_valueLoad($strLNGFile, 'Error', '018', 'Error resolving metalink %1.', True), '%1', $strURL), 0, $hGUI)
						EndIf
					Else
						ShellExecute($strURL)
					EndIf
				EndIf

			Case $hBtnSend
				; Send email
				sendEmail($hTxtSample, $hTxtSubject)
		EndSwitch

		Select
			Case $gDblClick
				; Edit a vendor with double click in ListView
				$gDblClick = False
				editVendor()

			Case _IsPressed('0D', $hDLL)
				; Check if Enter key was pressed in Search box to search term or in ListView to edit selected vendor
				If WinActive($hGUI) Then
					Switch ControlGetFocus($hGUI)
						Case 'Edit2'
							$strTerm = StringReplace(GUICtrlRead($hTxtSearch), '|', '')
							$intSearch = searchTerm($strTerm)
							While _IsPressed('0D', $hDLL)
							WEnd
						Case 'SysListView321'
							editVendor()
					EndSwitch
				EndIf

			Case _IsPressed('2D', $hDLL)
				; Check if Ins key was pressed in ListView to add a new vendor
				While _IsPressed('2D', $hDLL)
				WEnd
				If WinActive($hGUI) And ControlGetFocus($hGUI) = 'SysListView321' Then
					addVendor()
				EndIf

			Case _IsPressed('2E', $hDLL)
				; Check if Del key was pressed in ListView to delete selected vendor
				While _IsPressed('2E', $hDLL)
				WEnd
				If WinActive($hGUI) And ControlGetFocus($hGUI) = 'SysListView321' Then
					delVendor()
				EndIf

			Case _IsPressed('72', $hDLL)
				; F3 to search next
				While _IsPressed('72', $hDLL)
				WEnd
				If WinActive($hGUI) And ControlGetFocus($hGUI) = 'SysListView321' Then
					$intSearch = searchTerm($strTerm, $intSearch)
				EndIf

			Case BitAND(_GUICtrlButton_GetState($hChkAll), $BST_CHECKED) = $BST_CHECKED
				; Select all vendors
				If Not $boolSelectAll Then
					$boolSelectAll = True
					$gGetEmails = False
					For $i = 0 To $intVendorCount
						_GUICtrlListView_SetItemChecked($hLstVendors, $i, True)
					Next
					$gGetEmails = True
					updateRecipients()
				EndIf

			Case BitAND(_GUICtrlButton_GetState($hChkAll), $BST_UNCHECKED) = $BST_UNCHECKED
				; Unselect all vendors
				If $boolSelectAll Then
					$boolSelectAll = False
					$gGetEmails = False
					For $i = 0 To $intVendorCount
						_GUICtrlListView_SetItemChecked($hLstVendors, $i, False)
					Next
					$gGetEmails = True
					updateRecipients()
				EndIf
		EndSelect
	WEnd

	; Delete GUI and controls.
	_GUICtrlListView_UnRegisterSortCallBack($hLstVendors)
	DllClose($hDLL)
	GUIDelete($hGUI)
EndFunc   ;==>mainGUI
; <=== vendorGUI ==================================================================================
; vendorGUI(Integer, String, String, [String])
; ; Show child window to create or modiy vendor info.
; ;
; ; @param  Integer					Parent GUI handle.
; ; @param  String          Child window title.
; ; @param  String          Message to show at exit without saving.
; ; @param  [String]        Current vendor. Default: Empty.
; ; @return String        	Modified vnedor info.
Func vendorGUI($pParent, $pTitle, $pCancelWarning, $pCurrentVendor = '')
	; Disable main GUI
	GUISetState(@SW_DISABLE, $hGUI)

	; Get vendor fixing entries count
	StringReplace($pCurrentVendor, '|', '')
	If @extended < 6 Then $pCurrentVendor &= _StringRepeat('|', 6 - @extended)
	$pCurrentVendor = StringSplit($pCurrentVendor, '|', 2)

	Local $strVendorCancel = INI_valueLoad($strLNGFile, 'Vendor', '012', 'Cancel')
	Local $strVendorName = _trim($pCurrentVendor[0])

	; Create vendor form
	; EditBox: ($ES_MULTILINE 0x0004, $WS_EX_WINDOWEDGE 0x00000100, $ES_WANTRETURN 0x1000, $WS_VSCROLL 0x00200000)
	; CheckBox: ($BS_AUTOCHECKBOX 0x0003, $BS_RIGHTBUTTON 0x0020, $BS_RIGHT 0x0200)
	Local $hVendorGUI = GUICreate($pTitle, 400, 230, -1, -1, -1, -1, $pParent)
	Local $hLblVendorName = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Vendor', '004', 'Vendor name') & ':', 7, 7, 133, 20)
	Local $hTxtVendorName = GUICtrlCreateInput($strVendorName, 140, 5, 155, 20)
	Local $hChkVendorDiscontinued = GUICtrlCreateCheckbox(INI_valueLoad($strLNGFile, 'Vendor', '011', 'Discontinued'), 300, 5, 95, 20, BitOR(0x0003, 0x0020, 0x0200))
	Local $hLblVendorMalwareEmail = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Vendor', '005', 'Malware email') & ':', 7, 30, 133, 20)
	Local $hTxtVendorMalwareEmail = GUICtrlCreateInput(_trim($pCurrentVendor[1]), 140, 28, 255, 20)
	Local $hLblVendorMalwareSubmit = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Vendor', '006', 'Malware submission') & ':', 7, 53, 133, 20)
	Local $hTxtVendorMalwareSubmit = GUICtrlCreateInput(_trim($pCurrentVendor[2]), 140, 51, 255, 20)
	Local $hLblVendorFPEmail = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Vendor', '007', 'False positive email') & ':', 7, 76, 133, 20)
	Local $hTxtVendorFPEmail = GUICtrlCreateInput(_trim($pCurrentVendor[3]), 140, 74, 255, 20)
	Local $hLblVendorFPSubmit = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Vendor', '008', 'FP submission URL') & ':', 7, 99, 133, 20)
	Local $hTxtVendorFPSubmit = GUICtrlCreateInput(_trim($pCurrentVendor[4]), 140, 97, 255, 20)
	Local $hLblVendorHomePage = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Vendor', '009', 'HomePage') & ':', 7, 122, 133, 20)
	Local $hTxtVendorHomePage = GUICtrlCreateInput(_trim($pCurrentVendor[5]), 140, 120, 255, 20)
	Local $hLblVendorComment = GUICtrlCreateLabel(INI_valueLoad($strLNGFile, 'Vendor', '010', 'Comments') & ':', 7, 145, 133, 20)
	Local $hTxtVendorComment = GUICtrlCreateEdit(_trim($pCurrentVendor[6]), 140, 143, 255, 80, BitOR(0x0004, 0x00000100, 0x1000, 0x00200000))
	Local $hBtnVendorSave = GUICtrlCreateButton(INI_valueLoad($strLNGFile, 'Vendor', '013', 'Save'), 5, 184, 130, 20)
	Local $hBtnVendorCancel = GUICtrlCreateButton($strVendorCancel, 5, 204, 130, 20)

	; Update check boxes
	GUICtrlSetState($hChkVendorDiscontinued, StringLeft($strVendorName, 1) = '*' ? 1 : 4)
	GUISetState()

	; Child loop
	Local $strCurrentVendor
	While 1
		Switch GUIGetMsg()
			Case $hChkVendorDiscontinued, $hTxtVendorName
				$strVendorName = _trimLeft(GUICtrlRead($hTxtVendorName), '*')
				If BitAND(GUICtrlRead($hChkVendorDiscontinued), 1) = 1 Then $strVendorName = '*' & $strVendorName
				GUICtrlSetData($hTxtVendorName, $strVendorName)

			Case $hTxtVendorComment
				Local $strComment = GUICtrlRead($hTxtVendorComment)
				If StringLen($strComment) > 0 Then
					$strComment = StringReplace($strComment, '\l1', ChrW(9679) & ' ', 0, 1)
					$strComment = StringReplace($strComment, '\l2', '   ' & ChrW(9633) & ' ', 0, 1)
					$strComment = StringReplace($strComment, '\l3', '      ' & ChrW(9642) & ' ', 0, 1)
					GUICtrlSetData($hTxtVendorComment, $strComment)
				EndIf

			Case $GUI_EVENT_CLOSE, $hBtnVendorCancel
				If StringLen(GUICtrlRead($hTxtVendorName)) > 0 Or _
						StringLen(GUICtrlRead($hTxtVendorMalwareEmail)) > 0 Or StringLen(GUICtrlRead($hTxtVendorMalwareSubmit)) > 0 Or _
						StringLen(GUICtrlRead($hTxtVendorFPEmail)) > 0 Or StringLen(GUICtrlRead($hTxtVendorFPSubmit)) > 0 Or _
						StringLen(GUICtrlRead($hTxtVendorHomePage)) > 0 Or StringLen(GUICtrlRead($hTxtVendorComment)) > 0 Then
					If MsgBox(52, $strVendorCancel, $pCancelWarning, 0, $hGUI) = 6 Then ExitLoop
				Else
					ExitLoop
				EndIf
			Case $hBtnVendorSave
				$strCurrentVendor = GUICtrlRead($hTxtVendorName) & '|' & _
						GUICtrlRead($hTxtVendorMalwareEmail) & '|' & _
						GUICtrlRead($hTxtVendorMalwareSubmit) & '|' & _
						GUICtrlRead($hTxtVendorFPEmail) & '|' & _
						GUICtrlRead($hTxtVendorFPSubmit) & '|' & _
						GUICtrlRead($hTxtVendorHomePage) & '|' & _
						_fixCR(GUICtrlRead($hTxtVendorComment))
				ExitLoop
		EndSwitch
	WEnd
	GUIDelete()

	; Enable main GUI
	GUISetState(@SW_ENABLE, $hGUI)
	WinActivate($hGUI)

	Return $strCurrentVendor
EndFunc   ;==>vendorGUI
; <=== _WM_NOTIFY =================================================================================
; _WM_NOTIFY(Integer, Integer, Integer, Integer)
; ; UDF for WM_NOTIFY Windows Message to get selected emails.
; ;
; ; @param  Integer         Window handle.
; ; @param  Integer         Message.
; ; @param  Integer         wParam. Common control sending the message.
; ; @param  Integer         lParam. Pointer to NMHDR structure.
; ; @return Integer         $GUI_RUNDEFMSG.
; ; @author Nine            https://www.autoitscript.com/forum/topic/207636-read-value-from-checked-ListView-item/?do=findComment&comment=1497487
Func _WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam, $lParam
	If Number($wParam) = $hLstVendors Then
		Local $tagNM = DllStructCreate($tagNMHDR, $lParam)
		Switch DllStructGetData($tagNM, "Code")
			Case $NM_DBLCLK
				$gDblClick = True

			Case $LVN_HOTTRACK
				; Update ToolTip of ListView items
				Local $intHotItem = _GUICtrlListView_GetHotItem($hLstVendors)
				If $intHotItem <> -1 Then
					Local $strVendorName = _trim(_GUICtrlListView_GetItemText($hLstVendors, $intHotItem, 0))
					If StringLen($strVendorName) = 0 Then $strVendorName = '---'
					Local $strVendorMalwareEmail = _trim(_GUICtrlListView_GetItemText($hLstVendors, $intHotItem, 1))
					$strVendorMalwareEmail = $VendorMalwareEmail & ': ' & (StringLen($strVendorMalwareEmail) = 0 ? '---' : $strVendorMalwareEmail)
					Local $strVendorMalwareSubmit = _trim(_GUICtrlListView_GetItemText($hLstVendors, $intHotItem, 2))
					$strVendorMalwareSubmit = $VendorMalwareSubmit & ': ' & (StringLen($strVendorMalwareSubmit) = 0 ? '---' : $strVendorMalwareSubmit)
					Local $strVendorFPEmail = _trim(_GUICtrlListView_GetItemText($hLstVendors, $intHotItem, 3))
					$strVendorFPEmail = $VendorFPEmail & ': ' & (StringLen($strVendorFPEmail) = 0 ? '---' : $strVendorFPEmail)
					Local $strVendorFPSubmit = _trim(_GUICtrlListView_GetItemText($hLstVendors, $intHotItem, 4))
					$strVendorFPSubmit = $VendorFPSubmit & ': ' & (StringLen($strVendorFPSubmit) = 0 ? '---' : $strVendorFPSubmit)
					Local $strVendorHomePage = _trim(_GUICtrlListView_GetItemText($hLstVendors, $intHotItem, 5))
					$strVendorHomePage = $VendorHomePage & ': ' & (StringLen($strVendorHomePage) = 0 ? '---' : $strVendorHomePage)
					Local $strVendorComment = _trim(_GUICtrlListView_GetItemText($hLstVendors, $intHotItem, 6))
					If StringInStr($strVendorComment, @CRLF) > 0 Then $strVendorComment = @CRLF & $strVendorComment
					$strVendorComment = $VendorComments & ': ' & (StringLen($strVendorComment) = 0 ? '---' : $strVendorComment)
					ToolTip($strVendorMalwareEmail & @CRLF & $strVendorMalwareSubmit & @CRLF & _
							$strVendorFPEmail & @CRLF & $strVendorFPSubmit & @CRLF & _
							$strVendorHomePage & @CRLF & $strVendorComment, Default, Default, $strVendorName, 1)
				EndIf

			Case $LVN_ITEMCHANGED
				; Item was changed in ListView control
				$tagNM = DllStructCreate($tagNMListView, $lParam)
				; If $gGetEmails is true, verify if checked state was changed using bitwise operation: State image index * 0x1000
				If $gGetEmails And BitAND(DllStructGetData($tagNM, "NewState"), $LVIS_STATEIMAGEMASK) Then
					updateRecipients()
				EndIf

		EndSwitch
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_NOTIFY
; <=== _WM_GETMINMAXINFO ==========================================================================
; _WM_GETMINMAXINFO(Integer, Integer, Integer, Integer)
; ; UDF for WM_GETMINMAXINFO Windows Message to limit minimum size of GUI.
; ; Needs Global declaration of minimum width and height.
; ;
; ; @param  Integer         Window handle.
; ; @param  Integer         Message.
; ; @param  Integer         wParam. Not used.
; ; @param  Integer         lParam. Pointer to a MINMAXINFO structure.
; ; @return Integer         $GUI_RUNDEFMSG.
; ; @author BrewManNH       https://www.autoitscript.com/forum/topic/148307-lock-gui-width-when-resize/
Func _WM_GETMINMAXINFO($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam, $lParam
	If $hWnd = $hGUI Then
		Local $tagMaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)
		DllStructSetData($tagMaxinfo, 7, $MinWidth)   ; Min width
		DllStructSetData($tagMaxinfo, 8, $MinHeight)  ; Min height
		DllStructSetData($tagMaxinfo, 9, 99999)       ; Max width
		DllStructSetData($tagMaxinfo, 10, 99999)      ; Max height
		Return $GUI_RUNDEFMSG
	EndIf
EndFunc   ;==>_WM_GETMINMAXINFO
; ============================================================================== GUI PROCEDURES ==>
#EndRegion GUI PROCEDURES
#Region MAIN PROCEDURES
; <== MAIN PROCEDURES =============================================================================
; <=== addVendor ==================================================================================
; addVendor()
; ; Add a new vendor to ListView and save to file.
; ;
; ; @param  NONE
; ; @return NONE
Func addVendor()
	ToolTip('')
	Local $strStringVendor = vendorGUI($hGUI, $VendorAdd, $VendorWarning)
	If StringLen($strStringVendor) > 0 Then
		$intVendorCount = UBound($hItmVendors)
		Local $intCount = $intVendorCount + 1
		ReDim $hItmVendors[$intCount]
		$hItmVendors[$intCount - 1] = GUICtrlCreateListViewItem($strStringVendor, $hLstVendors)
		_GUICtrlListView_SetItemGroupID($hLstVendors, $intCount - 1, StringLeft($strStringVendor, 1) = '*' ? 1 : 0)
		saveList()
		ControlFocus($hGUI, '', $hLstVendors)
;~ 		_GUICtrlListView_ClickItem($hLstVendors, $intCount)
	EndIf
EndFunc   ;==>addVendor
; <=== delVendor ==================================================================================
; delVendor()
; ; Delete selected vendor in ListView and save to file.
; ;
; ; @param  NONE
; ; @return NONE
Func delVendor()
	ToolTip('')

	; If an item is selected, get selected index using $hLstVendors array
	Local $intSelected = GUICtrlRead($hLstVendors)
	If $intSelected > 0 Then
		Local $intIndex = selectedIndex($hLstVendors)
		Local $strVendorName = _GUICtrlListView_GetItemText($hLstVendors, $intIndex, 0)
		If MsgBox(52, $VendorDel, $VendorDel & ': ' & $strVendorName & @LF & $VendorConfirm, 0, $hGUI) = 6 Then
			_GUICtrlListView_DeleteItem($hLstVendors, $intIndex)
			_ArrayDelete($hItmVendors, $intIndex)
			$intVendorCount -= 1
			updateRecipients()
			saveList()
		EndIf
	EndIf
EndFunc   ;==>delVendor
; <=== editVendor =================================================================================
; editVendor()
; ; Edit selected vendor in ListView and save to file.
; ;
; ; @param  NONE
; ; @return NONE
Func editVendor()
	ToolTip('')

	; If an item is selected, edit it
	Local $intSelected = GUICtrlRead($hLstVendors)
	If $intSelected > 0 Then
		Local $strStringVendor = vendorGUI($hGUI, $VendorEdit, $VendorWarning, GUICtrlRead($intSelected))
		If StringLen($strStringVendor) > 0 Then
			GUICtrlSetData($intSelected, ' | | | | | |')
			GUICtrlSetData($intSelected, $strStringVendor)
			_GUICtrlListView_SetItemGroupID($hLstVendors, selectedIndex($hLstVendors), StringLeft($strStringVendor, 1) = '*' ? 1 : 0)
			updateRecipients()
			saveList()

			ControlFocus($hGUI, '', $hLstVendors)
;~ 			_GUICtrlListView_ClickItem($hLstVendors, selectedIndex($hLstVendors))
		EndIf
	EndIf
EndFunc   ;==>editVendor
; <=== getDefaultList() ===========================================================================
; getDefaultList(String, [String])
; ; Get vendors info from default URL.
; ;
; ; @param  String          Default URL. Source: https://www.techsupportalert.com/how-to-report-malware-or-false-positives-to-multiple-antivirus-vendors/
; ; @param  [String]        Message to display in tooltip. Default: Nothing.
; ; @return String()        2D array with vendors' name, email and submit URL.
Func getDefaultList($pURL, $pTooltip = Default)
	Local $intCountTotal
	Local $intCountEmail
	Local $intCountSubmit
	Local $intCountHomePage
	Local $intPos
	Local $strContact[5]
	Local $strLine
	Local $strLines
	Local $strVendors

	; Default parameters
	If $pTooltip = Default Or StringLen($pTooltip) = 0 Then $pTooltip = ''

	; Get source code from URL
	ToolTip($pTooltip)
	$strLines = BinaryToString(InetRead($pURL))
	If StringLen($strLines) = 0 Then
		ToolTip('')
		MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
				StringReplace(INI_valueLoad($strLNGFile, 'Error', '004', 'No data retrieved from URL "%1".', True), '%1', $pURL), 0, $hGUI)
		Return SetError(1)
	EndIf

	; Remove irrelevant data
	$strLines = StringMid($strLines, StringInStr($strLines, $strListGuide))
	$strLines = StringMid($strLines, StringInStr($strLines, $strListBefore) + 8)
	$strLines = StringLeft($strLines, StringInStr($strLines, $strListAfter) - 1)
	If StringLen($strLines) = 0 Then
		ToolTip('')
		MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
				StringReplace(INI_valueLoad($strLNGFile, 'Error', '005', 'No valid data found in URL "%1".', True), '%1', $pURL), 0, $hGUI)
		Return SetError(1)
	EndIf

	; Get vendors
	$strLines = StringSplit($strLines, @LF, 2)
	For $i = 0 To UBound($strLines) - 1
		$strLine = _removeTags($strLines[$i])
		If StringLen($strLine) > 0 Then
			If $strLine <> $strLines[$i] Then
				$intCountTotal += 1
				$strVendors &= $strContact[0] & '|' & $strContact[1] & '|' & $strContact[2] & '|' & $strContact[3] & '|' & _
						$strContact[4] & @LF & DECODE_URL($strLine) & '|'
				$strContact[0] = ''
				$strContact[1] = ''
				$strContact[2] = ''
				$strContact[3] = ''
				$strContact[4] = ''
			Else
				; Get email for malware
				$intPos = StringInStr($strLine, '<a href="mailto:')
				If $intPos > 0 And StringInStr($strLine, $strMalwareEmailMark) > 0 Then
					$intCountEmail += 1
					$strContact[0] = StringMid($strLine, $intPos + 16, StringInStr($strLine, '?', 2, 1, $intPos + 16) - $intPos - 16)
				EndIf
				; Get page for malware submission
				$intPos = StringInStr($strLine, '<a href="')
				If $intPos > 0 And StringInStr($strLine, $strMalwareSubmitMark) > 0 And StringInStr($strLine, '<a href="mailto:') = 0 Then
					$intCountSubmit += 1
					$strContact[1] = StringMid($strLine, $intPos + 9, StringInStr($strLine, '"', 2, 1, $intPos + 9) - $intPos - 9)
				EndIf
				; Get email for false positive
				$intPos = StringInStr($strLine, '<a href="mailto:')
				If $intPos > 0 And StringInStr($strLine, $strFPEmailMark) > 0 Then
					$intCountEmail += 1
					$strContact[2] = StringMid($strLine, $intPos + 16, StringInStr($strLine, '?', 2, 1, $intPos + 16) - $intPos - 16)
				EndIf
				; Get page for False positive submission
				$intPos = StringInStr($strLine, '<a href="')
				If $intPos > 0 And StringInStr($strLine, $strFPSubmitMark) > 0 And StringInStr($strLine, '<a href="mailto:') = 0 Then
					$intCountSubmit += 1
					$strContact[3] = StringMid($strLine, $intPos + 9, StringInStr($strLine, '"', 2, 1, $intPos + 9) - $intPos - 9)
				EndIf
				; Get HomePage
				$intPos = StringInStr($strLine, '<a href="')
				If $intPos > 0 And StringInStr($strLine, $strHomePageMark) > 0 And StringInStr($strLine, '<a href="mailto:') = 0 Then
					$intCountHomePage += 1
					$strContact[4] = StringMid($strLine, $intPos + 9, StringInStr($strLine, '"', 2, 1, $intPos + 9) - $intPos - 9)
				EndIf
			EndIf
		EndIf
	Next

	; Remove empty lines, add entry counters and last entry contacts
	While StringInStr($strVendors, @LF & @LF)
		$strVendors = StringReplace($strVendors, @LF & @LF, @LF)
	WEnd
	$strVendors = $intCountTotal & '|' & $intCountEmail & '|' & $intCountSubmit & '|' & $intCountHomePage & _
			StringTrimLeft($strVendors, 1) & _
			$strContact[0] & '|' & $strContact[1] & '|' & $strContact[2] & '|' & $strContact[3] & '|' & $strContact[4]

	; Backup previous list
	Local $intCount = 0
	Local $strBackup
	Do
		$intCount += 1
		$strBackup = _fileNameInfo($strLSTFile, 13) & '_backup_' & $intCount & _fileNameInfo($strLSTFile, 2)
	Until Not FileExists($strBackup)
	FileCopy($strLSTFile, $strBackup)

	; Return 2D array
	ToolTip('')
	Return _split2D($strVendors)
EndFunc   ;==>getDefaultList
; <=== getFileName ================================================================================
; getFileName()
; ; Returns new filename usign Save CommonDialog.
; ;
; ; @param  NONE
; ; @return String        	Filename.
Func getFileName()
	Local $dlgSaveFile = FileSaveDialog(INI_valueLoad($strLNGFile, 'Save', '003', 'Save archive as'), $strWorkingPath, _
			INI_valueLoad($strLNGFile, 'Save', '004', 'Zip archives') & ' (*.zip)|' & _
			INI_valueLoad($strLNGFile, 'Save', '005', 'All files') & ' (*.*)', $FD_PATHMUSTEXIST, '', $hGUI)
	If @error Then
		Return ''
	Else
		$strWorkingPath = _fileNameInfo($dlgSaveFile, 12)
		Return $dlgSaveFile
	EndIf
EndFunc   ;==>getFileName
; <=== getSample ==================================================================================
; getSample(String)
; ; Get attached filename usign Open CommonDialog.
; ;
; ; @param  String          Previous value of attached filename.
; ; @return String        	New attached filename.
Func getSample($pPrevious)
	; Get file to be attached
	$strWorkingPath = _fileNameInfo($pPrevious, 12)
	Local $dlgOpenFile = FileOpenDialog(INI_valueLoad($strLNGFile, 'Open', '001', 'Select file'), $strWorkingPath, _
			INI_valueLoad($strLNGFile, 'Open', '002', 'Executables') & ' (*.exe)|' & _
			INI_valueLoad($strLNGFile, 'Open', '003', 'AutoIt scripts') & ' (*.au3)|' & _
			INI_valueLoad($strLNGFile, 'Open', '004', 'Zip archives') & ' (*.zip)|' & _
			INI_valueLoad($strLNGFile, 'Open', '005', 'All files') & ' (*.*)', $FD_FILEMUSTEXIST, $pPrevious, $hGUI)
	If @error Then
		Return $pPrevious
	Else
		$strWorkingPath = _fileNameInfo($dlgOpenFile, 12)
		Return $dlgOpenFile
	EndIf
EndFunc   ;==>getSample
; <=== saveList ===================================================================================
; saveList([String()])
; ; Save ListView data to file.
; ;
; ; @param  [String()]			Array with default values. If empty, use ListView data.
; ; @return NONE
Func saveList($pDefault = '')
	; Open file to write in UTF-8 with BOM ($FO_OVERWRITE 2, $FO_UTF8 128)
	Local $hFileOpen = FileOpen($strLSTFile, BitOR(2, 128))
	If $hFileOpen = -1 Then
		Local $strErrMsg
		If IsArray($pDefault) Then
			$strErrMsg = StringReplace(INI_valueLoad($strLNGFile, 'Error', '003', 'Error creating list file %1.', True), '%1', $strLSTFile)
		Else
			$strErrMsg = StringReplace(INI_valueLoad($strLNGFile, 'Error', '006', 'Error updating list file %1.', True), '%1', $strLSTFile)
		EndIf
		MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), $strErrMsg, 0, $hGUI)
		Exit (1)
	EndIf
	FileWriteLine($hFileOpen, '; ======================= VENDORS LIST ======================')
	FileWriteLine($hFileOpen, '; Info of each vendor goes in one line, pipe separated (|):')
	FileWriteLine($hFileOpen, '; Name|Email for malware|Malware submission URL|Email for false positives|FP submission URL|HomePage|Comments')
	FileWriteLine($hFileOpen, '; -----------------------------------------------------------')
	FileWriteLine($hFileOpen, '; NOTE: More columns can be added, but they will be ignored.')
	FileWriteLine($hFileOpen, '; ===========================================================')

	If IsArray($pDefault) Then
		; Save default list
		For $i = 1 To $pDefault[0][0]
			FileWriteLine($hFileOpen, _trim($pDefault[$i][0] & '|' & $pDefault[$i][1] & '|' & _
					$pDefault[$i][2] & '|' & $pDefault[$i][3] & '|' & _
					$pDefault[$i][4] & '|' & $pDefault[$i][5], '|') & @CRLF)
		Next
	Else
		; Save ListView data
		For $i = 0 To $intVendorCount
			FileWriteLine($hFileOpen, _trim(_trim(_GUICtrlListView_GetItemText($hLstVendors, $i, 0)) & '|' & _
					_trim(_GUICtrlListView_GetItemText($hLstVendors, $i, 1)) & '|' & _
					_trim(_GUICtrlListView_GetItemText($hLstVendors, $i, 2)) & '|' & _
					_trim(_GUICtrlListView_GetItemText($hLstVendors, $i, 3)) & '|' & _
					_trim(_GUICtrlListView_GetItemText($hLstVendors, $i, 4)) & '|' & _
					_trim(_GUICtrlListView_GetItemText($hLstVendors, $i, 5)) & '|' & _
					_escapeSeq(_trim(_GUICtrlListView_GetItemText($hLstVendors, $i, 6)), False), '|') & @CRLF)
		Next
	EndIf

	FileClose($hFileOpen)
EndFunc   ;==>saveList
; <=== searchTerm =================================================================================
; searchTerm(String, [Integer])
; ; Search a term in vendor list.
; ;
; ; @param  String					Term to search.
; ; @param  [Integer]				Starting index for next search. Default: 0.
; ; @return Integer					Next index.
Func searchTerm($pTerm, $pStart = 0)
	If StringLen($pTerm) = 0 Then Return 0
	If $pStart < 0 Or $pStart > $intVendorCount Then $pStart = 0

	For $pStart = $pStart To $intVendorCount
		If StringInStr(_GUICtrlListView_GetItemTextString($hLstVendors, $pStart), $pTerm) Then
			ControlFocus($hGUI, '', $hLstVendors)
			_GUICtrlListView_ClickItem($hLstVendors, $pStart)
			Return $pStart + 1 > $intVendorCount ? 0 : $pStart + 1
		EndIf
	Next
	MsgBox(48, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
			INI_valueLoad($strLNGFile, 'Error', '019', 'Term not found'), 0, $hGUI)
	Return 0
EndFunc   ;==>searchTerm
; <=== selectedEntry ===========================================================================
; selectedEntry(Integer, [Integer])
; ; Return ListView selected item content using an array of predefined items.
; ;
; ; @param  Integer					ListView handle.
; ; @param  [Integer]				Subitem index. Default: 0.
; ; @return Integer					Zero-based index of selected item. If nothing is found, return -1.
Func selectedEntry($pListView, $pSubItem = Default)
	If $pSubItem = Default Or StringLen($pSubItem) = 0 Then $pSubItem = 0
	Local $intIndex = selectedIndex($pListView)
	If $intIndex = -1 Then Return ''
	Return _trim(_GUICtrlListView_GetItemText($pListView, $intIndex, $pSubItem))
EndFunc   ;==>selectedEntry
; <=== selectedIndex ===========================================================================
; selectedIndex(Integer)
; ; Return zero-based index of selected item in ListView.
; ;
; ; @param  Integer					ListView handle.
; ; @return Integer					Zero-based index of selected item. If nothing is found, returns -1.
Func selectedIndex($pListView)
	Local $strSelected = _GUICtrlListView_GetSelectedIndices($pListView, True)
	If $strSelected[0] > 0 Then
		Return $strSelected[1]
	Else
		Return -1
	EndIf
EndFunc   ;==>selectedIndex
; <=== sendEmail ==================================================================================
; sendEmail(String, String)
; ; Send email message using RS_SMTP UDF. Attachment can be either archive or file, zip files will
; ; be verified as zip encrypted files with password 'infected'. Otherwise, a new encrypted archive
; ; will be created using selected file.
; ;
; ; @param  String          TextBox handle of attached filename.
; ; @param  String          TextBox handle of current subject (malware or false positive).
; ; @return String        	New attached filename.
Func sendEmail($pSample, $pSubject)
	Local $intNext = 1
	Local $strAttachment
	Local $strReturn
	Local $strTo = GUICtrlRead($hTxtTo)
	Local $strSubject = GUICtrlRead($pSubject)

	; Validate fields
	If _emptyField($strSample, INI_valueLoad($strLNGFile, 'Main', '001', 'File')) Then Return
	If _emptyField($strSMTPServer, INI_valueLoad($strLNGFile, 'Main', '016', 'Server')) Then Return
	If _emptyField($intPort, INI_valueLoad($strLNGFile, 'Main', '017', 'Port')) Then Return
	If _emptyField($strName, INI_valueLoad($strLNGFile, 'Main', '023', 'Name')) Then Return
	If _emptyField($strEmail, INI_valueLoad($strLNGFile, 'Main', '024', 'Email')) Then Return
	If _emptyField($strTo, INI_valueLoad($strLNGFile, 'Main', '026', 'To')) Then Return
	If _emptyField($strSubject, INI_valueLoad($strLNGFile, 'Main', '029', 'Subject')) Then Return
	If _emptyField($strBody, INI_valueLoad($strLNGFile, 'Main', '030', 'Body')) Then Return

	; Check if file sample is defined
	$strSample = GUICtrlRead($pSample)
	Do
		If Not FileExists($strSample) Then
			If MsgBox(17, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
					INI_valueLoad($strLNGFile, 'Error', '008', 'File sample not defined.\nSelect file or zip archive to proceed.', True), 0, $hGUI) = 2 Then
				$intNext = 0
				ExitLoop
			EndIf
			$strSample = getSample($strSample)
			GUICtrlSetData($pSample, $strSample)
		EndIf
	Until FileExists($strSample)

	; Existing file sample retrieved, validate it
	If $intNext = 1 Then
		$intNext = 0
		$strAttachment = _fileNameInfo($str7ZSample, 13) & '.zip'
		If StringInStr($strAttachment, '\') = 0 Then $strAttachment = _fileNameInfo($strSample, 12) & $strAttachment

		; Setting environ user variables
		ENVIRON_userAdd('%WORKING_DIR%', _fileNameInfo($strSample, 12))

		; Validate file by extension
		If _fileNameInfo($strSample, 2) = '.zip' Then
			; Archive, checking encryption
			ENVIRON_userAdd('%ENCRYPTED_ARCHIVE%', $strSample)

			$strReturn = _run($str7ZEXE, ENVIRON_replace($str7ZCheck), $AppPath)
			If StringInStr($strReturn, $str7ZEncrypted) > 0 Then
				; Encrypted archive, verifying password
				If StringInStr(_run($str7ZEXE, ENVIRON_replace($str7ZVerify), $AppPath), $str7ZSuccess) > 0 Then
					$intNext = 2
				Else
					MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
							StringReplace(INI_valueLoad($strLNGFile, 'Error', '012', 'Archive %1 with incorrect password.', True), '%1', $strSample) & @LF & _
							INI_valueLoad($strLNGFile, 'Error', '009', 'Select a valid encrypted archive or\nany other file for automatic creation.', True), 0, $hGUI)
				EndIf
			Else
				If StringInStr($strReturn, $str7ZNotEncrypted) > 0 Then
					MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
							StringReplace(INI_valueLoad($strLNGFile, 'Error', '011', 'Archive %1 not encrypted.', True), '%1', $strSample) & @LF & _
							INI_valueLoad($strLNGFile, 'Error', '009', 'Select a valid encrypted archive or\nany other file for automatic creation.', True), 0, $hGUI)
				Else
					MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
							StringReplace(INI_valueLoad($strLNGFile, 'Error', '010', 'File %1 is not an archive.', True), '%1', $strSample) & @LF & _
							INI_valueLoad($strLNGFile, 'Error', '009', 'Select a valid encrypted archive or\nany other file for automatic creation.', True), 0, $hGUI)
				EndIf
			EndIf
		Else
			; Not archive, check if zip file already exists
			If FileExists($strAttachment) Then
				Switch MsgBox(36, INI_valueLoad($strLNGFile, 'Save', '001', 'Archive sample exists'), _
						INI_valueLoad($strLNGFile, 'Save', '002', 'Archive sample already exists.\nOverwrite?', True), 0, $hGUI)
					Case 6
						; Overwrite archive, delete old one
						If FileDelete($strAttachment) = 1 Then
							$intNext = 1
						Else
							MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), INI_valueLoad($strLNGFile, 'Error', '013', 'Error deleting archive %1.', True), 0, $hGUI)
						EndIf

					Case 7
						; Get new filename
						Local $strNewSample = getFileName()
						If StringLen($strNewSample) > 0 Then
							$intNext = 1
							$strAttachment = $strNewSample
						EndIf
				EndSwitch
			Else
				$intNext = 1
			EndIf
		EndIf
	EndIf

	; Create encrypted archive with password: infected.
	If $intNext = 1 Then
		ENVIRON_userAdd('%ENCRYPTED_ARCHIVE%', $strAttachment)
		ENVIRON_userAdd('%FILE_SAMPLE%', $strSample)

		If StringInStr(_run($str7ZEXE, ENVIRON_replace($str7ZAdd), $AppPath), $str7ZSuccess) > 0 Then
			$intNext = 2
		Else
			MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
					INI_valueLoad($strLNGFile, 'Error', '014', 'Error creating archive %1.', True), 0, $hGUI)
		EndIf
	EndIf

	; Send email
	If $intNext = 2 Then
		Local $strReturn = SMTP_send($strSMTPServer, $strName, $strEmail, _
				$strTo, $strSubject, $strBody, $strAttachment, '', GUICtrlRead($hTxtBCC), _
				($intPriority = 0 ? 'Low' : ($intPriority = 1 ? 'Normal' : 'High')), _
				$strUser, $strPassword, $intPort, $boolSSL, $boolTLS)
		If @error Then
			MsgBox(16, INI_valueLoad($strLNGFile, 'Error', '001', 'Error'), _
					INI_valueLoad($strLNGFile, 'Error', '015', 'Error sending email.', True) & @LF & _
					INI_valueLoad($strLNGFile, 'Error', '016', 'Error code', True) & ': ' & @error & @LF & _
					INI_valueLoad($strLNGFile, 'Error', '017', 'Description', True) & ': ' & $strReturn, 0, $hGUI)
		EndIf
	EndIf
EndFunc   ;==>sendEmail
; <=== updateKey ==================================================================================
; updateKey(Integer, String, String, [Boolean], [Boolean])
; ; Update variables from field values and save them to INI file.
; ;
; ; @param  Integer					Control handle.
; ; @param  String     	    Section in INI file.
; ; @param  String     	    Key in INI file.
; ; @param  [Boolean]   	  True to encrypt value. Default: False.
; ; @param  [Boolean]   	  Restore escape sequences. Default: False.
; ; @return Any							Retrieved value.
Func updateKey($pHandle, $pSection, $pKey, $pEncrypt = Default, $pRestoreEscaped = Default)
	If StringLen($pSection) = 0 Or StringLen($pKey) = 0 Then Return ''
	If $pEncrypt = Default Or StringLen($pEncrypt) = 0 Then $pEncrypt = False
	If $pRestoreEscaped = Default Or StringLen($pRestoreEscaped) = 0 Then $pRestoreEscaped = False

	Local $valValue
	Switch _getType($pHandle)
		Case 'Checkbox'
			$valValue = (GUICtrlRead($pHandle) = 1 ? 1 : 0)
		Case 'Combo'
			$valValue = _GUICtrlComboBox_GetCurSel($pHandle)
		Case 'Input', "Edit"
			$valValue = GUICtrlRead($pHandle)
	EndSwitch

	If $pEncrypt Then
		INI_valueWrite($strINIFile, $pSection, $pKey, _toHex(_encrypt($valValue, $Key)), $pRestoreEscaped)
	Else
		INI_valueWrite($strINIFile, $pSection, $pKey, $valValue, $pRestoreEscaped)
	EndIf

	Return $valValue
EndFunc   ;==>updateKey
; <=== updateList =================================================================================
; updateList(String(), [Boolean])
; ; Add array data to vendors ListView control.
; ;
; ; @param  String     	    Array.
; ; @param  [Boolean]   	  Array is 2D. Default: True.
; ; @return NONE
Func updateList($pArray, $p2D = True)
	If Not IsArray($pArray) Then Return

	; Get items count
	Local $intCount
	If $p2D Then
		; Data directly retrieved from URL default list.
		$intCount = $pArray[0][0]
	Else
		; Data retrieved from list file, count only valid lines.
		For $i = 0 To UBound($pArray) - 1
			If StringLen($pArray[$i]) > 0 And StringLeft($pArray[$i], 1) <> ';' Then $intCount += 1
		Next
	EndIf
	Dim $hItmVendors[$intCount]
	$intVendorCount = $intCount

	; Deactivate $LVN_ITEMCHANGED messages in _WM_NOTIFY
	$gGetEmails = False

	; Empty ListView
	Local $hListHandle = GUICtrlGetHandle($hLstVendors)
	_GUICtrlListView_DeleteAllItems($hLstVendors)

	; Update ListView
	_GUICtrlListView_BeginUpdate($hLstVendors)
	If $p2D Then
		For $i = 1 To $pArray[0][0]
			$hItmVendors[$i - 1] = GUICtrlCreateListViewItem(_ArrayToString($pArray, '|', $i, $i), $hLstVendors)
		Next
	Else
		Local $i = 0
		Local $intPos
		Local $strVendor
		For $strVendorLine In $pArray
			If StringLen($strVendorLine) > 0 And StringLeft($strVendorLine, 1) <> ';' Then
				; Fix vendor info to seven fields
				StringReplace($strVendorLine, '|', '')
				Select
					Case @extended < 6
						$strVendorLine &= _StringRepeat('|', 6 - @extended)
					Case @extended > 6
						$strVendorLine = StringLeft($strVendorLine, StringInStr($strVendorLine, '|', 2, 7) - 1)
				EndSelect

				; Allow only escape sequences in comment field
				$strVendor = StringSplit($strVendorLine, '|', 2)
				$strVendor[0] = _removeBreaks($strVendor[0])
				$strVendor[1] = _removeBreaks($strVendor[1])
				$strVendor[2] = _removeBreaks($strVendor[2])
				$strVendor[3] = _removeBreaks($strVendor[3])
				$strVendor[4] = _removeBreaks($strVendor[4])
				$strVendor[5] = _removeBreaks($strVendor[5])
				$strVendor[6] = _escapeSeq($strVendor[6])

				; Add entry to ListView
				$hItmVendors[$i] = GUICtrlCreateListViewItem(_ArrayToString($strVendor), $hLstVendors)
				_GUICtrlListView_SetItemGroupID($hLstVendors, $i, StringLeft($strVendor[0], 1) = '*' ? 1 : 0)
				$i += 1
			EndIf
		Next
	EndIf
	_GUICtrlListView_SetColumnWidth($hLstVendors, 0, 150)
	_GUICtrlListView_SetColumnWidth($hLstVendors, 1, 50)
	_GUICtrlListView_SetColumnWidth($hLstVendors, 2, 50)
	_GUICtrlListView_SetColumnWidth($hLstVendors, 3, 50)
	_GUICtrlListView_SetColumnWidth($hLstVendors, 4, 50)
	_GUICtrlListView_SetColumnWidth($hLstVendors, 5, 50)
	_GUICtrlListView_SetColumnWidth($hLstVendors, 6, 50)
	_GUICtrlListView_EndUpdate($hLstVendors)

	; Reactivate $LVN_ITEMCHANGED messages in _WM_NOTIFY
	$gGetEmails = True
EndFunc   ;==>updateList
; <=== updateRecipients =================================================================================
; updateRecipients()
; ; Update To and BCC TextBoxes with checked ListView items.
; ;
; ; @param  NONE
; ; @return NONE
Func updateRecipients()
	Local $strSelectedEmail
	Local $strURTo
	Local $strURBCC

	GUICtrlSetData($hTxtTo, '')
	GUICtrlSetData($hTxtBCC, '')

	; Begin check all vendor box
	$boolSelectAll = True
	_GUICtrlButton_SetCheck($hChkAll, $BST_CHECKED)

	; Get selected emails
	For $i = 0 To $intVendorCount
		$strSelectedEmail = _GUICtrlListView_GetItemText($hLstVendors, $i, 1)
		If StringLen($strSelectedEmail) > 0 Then
			If _GUICtrlListView_GetItemChecked($hLstVendors, $i) Then
				If StringLen($strSelectedEmail) > 0 Then
					If StringLen($strURTo) = 0 Then
						$strURTo = $strSelectedEmail
					Else
						$strURBCC &= $strSelectedEmail & '; '
					EndIf
				EndIf
			Else
				$boolSelectAll = False
				_GUICtrlButton_SetCheck($hChkAll, $BST_UNCHECKED)
			EndIf
		EndIf
	Next
	GUICtrlSetData($hTxtTo, $strURTo)
	GUICtrlSetData($hTxtBCC, StringTrimRight($strURBCC, 2))
EndFunc   ;==>updateRecipients
; ============================================================================= MAIN PROCEDURES ==>
#EndRegion MAIN PROCEDURES
#Region INTERNAL PROCEDURES
; <== INTERNAL PROCEDURES =========================================================================
; <=== _addSlash ==================================================================================
; _addSlash(String)
; ; Add trailing slash removing leading and trailing spaces.
; ;
; ; @param  String          Path to verify.
; ; @return String          Path with trailing slash.
Func _addSlash($pPath)
	$pPath = StringStripWS($pPath, 3)
	If StringLen($pPath) = 0 Then Return ''
	While StringRight($pPath, 1) = '\'
		$pPath = StringTrimRight($pPath, 1)
	WEnd
	Return $pPath & '\'
EndFunc   ;==>_addSlash
; <=== _case ======================================================================================
; _case(Integer)
; ; Change string case.
; ;
; ; @param  String          String.
; ; @param  Integer         Case. -1: Unchanged, 0: lower, 1: UPPER, 2: Title case, 3: Proper Case,
; ;                         4: iNVERTED. Default: Unchanged
; ; @return String          String with new casing.
Func _case($pString, $pCase = Default)
	If StringLen($pString) = 0 Or $pCase = Default Or $pCase < 0 Or $pCase > 4 Then Return $pString

	Switch $pCase
		Case 1
			$pString = StringUpper($pString)
		Case 2
			$pString = StringUpper(StringLeft($pString, 1)) & StringLower(StringMid($pString, 2))
		Case 3
			$pString = _StringTitleCase($pString)
		Case 4
			$pString = StringLower(StringLeft($pString, 1)) & StringUpper(StringMid($pString, 2))
		Case Else
			$pString = StringLower($pString)
	EndSwitch
	Return $pString
EndFunc   ;==>_case
; <=== _decrypt ===================================================================================
; _decrypt(String, String)
; ; Decrypt text using AES 256 and SHA 512 key.
; ;
; ; @param  String          Text to decrypt.
; ; @param  String          Password.
; ; @return String          Decrypted text.
Func _decrypt($pText, $pPassword)
	If StringLen($pText) = 0 Or StringLen($pPassword) = 0 Then Return ''
	Local $strKey = _Crypt_DeriveKey($pPassword, $CALG_AES_256, $CALG_SHA_512)
	Local $strText = BinaryToString(_Crypt_DecryptData($pText, $strKey, $CALG_USERKEY))
	_Crypt_DestroyKey($strKey)
	Return $strText
EndFunc   ;==>_decrypt
; <=== _emptyField ================================================================================
; _emptyField(String, String)
; ; Check if a data field is empty.
; ;
; ; @param  String          Field data.
; ; @param  String          Field name.
; ; @return Boolean					True if data field is empty.
Func _emptyField($pField, $pFieldName)
	Local $strErrorTitle = INI_valueLoad($strLNGFile, 'Error', '001', 'Error')
	Local $strErrorMsg = StringReplace(INI_valueLoad($strLNGFile, 'Error', '007', 'Field "%1" missing.'), '%1', $pFieldName)

	If StringLen($pField) = 0 Then
		MsgBox(16, $strErrorTitle, $strErrorMsg, 0, $hGUI)
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>_emptyField
; <=== _encrypt ===================================================================================
; _encrypt(String, String)
; ; Encrypt text using AES 256 and SHA 512 key.
; ;
; ; @param  String          Text to encrypt.
; ; @param  String          Password.
; ; @return String          Encrypted text.
Func _encrypt($pText, $pPassword)
	If StringLen($pText) = 0 Or StringLen($pPassword) = 0 Then Return ''
	Local $strKey = _Crypt_DeriveKey($pPassword, $CALG_AES_256, $CALG_SHA_512)
	Local $strText = BinaryToString(_Crypt_EncryptData($pText, $strKey, $CALG_USERKEY))
	_Crypt_DestroyKey($strKey)
	Return $strText
EndFunc   ;==>_encrypt
; <=== _escapeSeq =================================================================================
; _escapeSeq(String, [Boolean], [Boolean])
; ; Replaces or restores escape sequences in a string.
; ;
; ; @param  String           Text string.
; ; @param  [Boolean]        Direction. True: Resolve escapes. False: Restore them. Default: True.
; ; @param  [Boolean]        True: Remove breaks. False: Keep them. Default: False.
; ; @return String           Processed text string.
Func _escapeSeq($pText, $pDirection = True, $pRemoveBreaks = False)
	If StringLen($pText) = 0 Then Return ''

	Local $intPos
	Local $strChar

	If $pDirection Then
		; Resolve default escape sequences
		$pText = StringReplace($pText, '\\', Chr(1), 0, 2)
		$pText = StringReplace($pText, '\r', @CR, 0, 1)
		$pText = StringReplace($pText, '\n', @CRLF, 0, 1)
		$pText = StringReplace($pText, '\t', Chr(9), 0, 1)

		; Resolve custom bullet sequences
		$pText = StringReplace($pText, '\l1', ChrW(9679) & ' ', 0, 1)
		$pText = StringReplace($pText, '\l2', '   ' & ChrW(9633) & ' ', 0, 1)
		$pText = StringReplace($pText, '\l3', '      ' & ChrW(9642) & ' ', 0, 1)

		Do
			$intPos = StringInStr($pText, '\x', 1)
			If $intPos > 0 Then
				$strChar = StringMid($pText, $intPos + 2, 2)
				$pText = StringReplace($pText, '\x' & $strChar, Chr(Dec($strChar)), 0, 1)
			EndIf
		Until $intPos = 0
		$pText = StringReplace($pText, Chr(1), '\')
	Else
		; Restore default escape sequences
		$pText = StringReplace($pText, '\', '\\', 0, 2)
		$pText = StringReplace($pText, @CRLF, '\n', 0, 2)
		$pText = StringReplace($pText, @LF, '\n', 0, 2)
		$pText = StringReplace($pText, @CR, '\r', 0, 2)
		$pText = StringReplace($pText, Chr(9), '\t', 0, 2)

		; Resolve custom bullet sequences
		$pText = StringReplace($pText, ChrW(9679) & ' ', '\l1', 0, 1)
		$pText = StringReplace($pText, '   ' & ChrW(9633) & ' ', '\l2', 0, 1)
		$pText = StringReplace($pText, '      ' & ChrW(9642) & ' ', '\l3', 0, 1)

		For $i = 1 To StringLen($pText)
			$strChar = StringMid($pText, $i, 1)
			If StringInStr($strASCIIchars, $strChar, 1) > 0 Then $pText = StringReplace($pText, $strChar, '\x' & Hex(Asc($strChar)), 0, 1)
		Next
	EndIf

	Return $pText
EndFunc   ;==>_escapeSeq
; <=== _fileNameInfo ==============================================================================
; _fileNameInfo(String, Integer)
; ; Returns filename info: drive, directory, filename and/or extension and changing case.
; ;
; ; @param  String          Filename with full path.
; ; @param  Integer         Filename flags. 1: Name, 2: Extension, 4: Dir, 8: Drive,
; ;                         16: Change to uppercase, 32: Change to lowercase . Default: Name (1).
; ; @return String          Filename data: Drive, directory, filename and/or extension.
Func _fileNameInfo($pFullPath, $pFlags = Default)
	Local $iCase = 0
	Local $sDrive = ''
	Local $sDir = ''
	Local $sFileName = ''
	Local $sExtension = ''

	If $pFlags = Default Or StringLen($pFlags) = 0 Then $pFlags = 1
	If BitAND($pFlags, 16) Then $iCase = 1
	If BitAND($pFlags, 32) Then $iCase = 2

	_PathSplit($pFullPath, $sDrive, $sDir, $sFileName, $sExtension)
	Switch $iCase
		Case 1
			$sDrive = StringUpper($sDrive)
			$sDir = StringUpper($sDir)
			$sFileName = StringUpper($sFileName)
			$sExtension = StringUpper($sExtension)
		Case 2
			$sDrive = StringLower($sDrive)
			$sDir = StringLower($sDir)
			$sFileName = StringLower($sFileName)
			$sExtension = StringLower($sExtension)
	EndSwitch

	$pFullPath = ''
	If BitAND($pFlags, 8) Then $pFullPath = $sDrive
	If BitAND($pFlags, 4) Then $pFullPath &= $sDir
	If BitAND($pFlags, 1) Then $pFullPath &= $sFileName
	If BitAND($pFlags, 2) Then
		$pFullPath &= $sExtension
	ElseIf $pFlags = 2 Then
		$pFullPath = $sExtension
	EndIf

	Return $pFullPath
EndFunc   ;==>_fileNameInfo
; <=== _fixCR =====================================================================================
; _fixCR(String)
; ; Replace isolated @CR and @LF characters with @CRLF to avoid incorrect line breaks in list.
; ;
; ; @param  String					String text.
; ; @return String					Processed text.
; ; @author Melba23					https://www.autoitscript.com/forum/topic/155849-regexp-replace-lf-with-crlf-if-line-on-has-an-lf/?do=findComment&comment=1126319
Func _fixCR($pText)
	If StringLen($pText) = 0 Then Return ''
	Return StringRegExpReplace($pText, "((?<!\x0d)\x0a|\x0d(?!\x0a))", @CRLF)
EndFunc   ;==>_fixCR
; <=== _fromHex ===================================================================================
; _fromHex(String)
; ; Returns a character string from a HEX sequence.
; ;
; ; @param  String          HEX ASCII values of characters.
; ; @return String          Characters string.
Func _fromHex($pHEX)
	If StringLen($pHEX) = 0 Then Return ''
	Local $strText
	For $i = 1 To StringLen($pHEX) Step 2
		$strText &= Chr(Dec(StringMid($pHEX, $i, 2)))
	Next
	Return $strText
EndFunc   ;==>_fromHex
; <=== _getType ===================================================================================
; _getType(Integer)
; ; Return control type.
; ;
; ; @param  Integer					Control handle.
; ; @return String        	Control type.
; ; @author	guiness					https://www.autoitscript.com/forum/topic/129129-how-to-obtain-the-type-of-gui-control/?do=findComment&comment=896780
Func _getType($pControlHandle)
	Local Const $GWL_STYLE = -16
	Local $intLong
	Local $strClass

	If IsHWnd($pControlHandle) = 0 Then
		$pControlHandle = GUICtrlGetHandle($pControlHandle)
		If IsHWnd($pControlHandle) = 0 Then Return SetError(1, 0, "Unknown")
	EndIf

	$strClass = _WinAPI_GetClassName($pControlHandle)
	If @error Then Return "Unknown"

	$intLong = _WinAPI_GetWindowLong($pControlHandle, $GWL_STYLE)
	If @error Then Return SetError(2, 0, 0)

	Switch $strClass
		Case "Button"
			Select
				Case BitAND($intLong, $BS_GROUPBOX) = $BS_GROUPBOX
					Return "Group"
				Case BitAND($intLong, $BS_CHECKBOX) = $BS_CHECKBOX
					Return "Checkbox"
				Case BitAND($intLong, $BS_AUTOCHECKBOX) = $BS_AUTOCHECKBOX
					Return "Checkbox"
				Case BitAND($intLong, $BS_RADIOBUTTON) = $BS_RADIOBUTTON
					Return "Radio"
				Case BitAND($intLong, $BS_AUTORADIOBUTTON) = $BS_AUTORADIOBUTTON
					Return "Radio"
			EndSelect

		Case "Edit"
			Select
				Case BitAND($intLong, $ES_WANTRETURN) = $ES_WANTRETURN
					Return "Edit"
				Case Else
					Return "Input"
			EndSelect

		Case "Static"
			Select
				Case BitAND($intLong, $SS_BITMAP) = $SS_BITMAP
					Return "Pic"
				Case BitAND($intLong, $SS_ICON) = $SS_ICON
					Return "Icon"
				Case BitAND($intLong, $SS_LEFT) = $SS_LEFT
					If BitAND($intLong, $SS_NOTIFY) = $SS_NOTIFY Then Return "Label"
					Return "Graphic"
			EndSelect

		Case "ComboBox"
			Return "Combo"
		Case "ListBox"
			Return "ListBox"
		Case "msctls_progress32"
			Return "Progress"
		Case "msctls_trackbar32"
			Return "Slider"
		Case "SysDateTimePick32"
			Return "Date"
		Case "SysListView32"
			Return "ListView"
		Case "SysMonthCal32"
			Return "MonthCal"
		Case "SysTabControl32"
			Return "Tab"
		Case "SysTreeView32"
			Return "TreeView"
	EndSwitch

	Return $strClass
EndFunc   ;==>_getType
; <=== _portableMode =================================================================================
; _portableMode(String)
; ; Get or set starting portable mode to read and write configuration files.
; ;
; ; @param  NONE
; ; @return Boolean					True: Configuration files in @ScriptDir. False: In %UserAppData%
Func _portableMode()
	Local Const $PortableFile = StringReplace(@ScriptDir & '\portable.ini', '\\', '\')
	If FileExists($PortableFile) Then
		Return (INI_valueLoad($PortableFile, 'Portable', 'PortableMode', '0') = '1') ? True : False
	Else
		INI_valueWrite($PortableFile, 'Portable', 'PortableMode', '0')
		Return False
	EndIf
EndFunc   ;==>_portableMode
; <=== _removeBreaks =================================================================================
; _removeBreaks(String)
; ; Remove line breaks from a string.
; ;
; ; @param  String			Original string.
; ; @return String  		Processed string.
Func _removeBreaks($pText)
	$pText = StringReplace($pText, '\\', '\', 0, 2)
	$pText = StringReplace($pText, '\r', '', 0, 1)
	$pText = StringReplace($pText, '\n', '', 0, 1)
	$pText = StringReplace($pText, '\t', Chr(9), 0, 1)

	; Replace underscores with spaces in metalinks
	If StringLeft($pText, 1) = '#' Then $pText = StringReplace($pText, '_', ' ', 0, 2)

	Return $pText
EndFunc   ;==>_removeBreaks
; <=== _removeExt =================================================================================
; _removeExt(String)
; ; Remove extension from a filename.
; ;
; ; @param  String			Filename.
; ; @return String  		Filename without extension.
; ; @author guinness
Func _removeExt($pFileName)
	Return StringRegExpReplace($pFileName, '\.[^.\\/]*$', '')
EndFunc   ;==>_removeExt
; <=== _removeTags ================================================================================
; _removeTags(String)
; ; Removes style tags from lines to get vendors name.
; ;
; ; @param  String          Original line from source code.
; ; @return String          Processed string line.
Func _removeTags($pLine)
;~ 	$pLine = DECODE_URL($pLine)
	If StringInStr($pLine, '<td><strong>') Then
		$pLine = StringStripWS(StringReplace($pLine, '<td><strong>', ''), 3)
	EndIf
	If StringInStr($pLine, '<strong>') And StringInStr($pLine, '<td id="') = 0 Then
		$pLine = StringStripWS(StringReplace($pLine, '<strong>', ''), 3)
	EndIf
	If StringInStr($pLine, '</strong></p>') Then
		$pLine = StringStripWS(StringReplace($pLine, '</strong></p>', ''), 3)
	EndIf
	If StringInStr($pLine, '</strong>') Then
		$pLine = StringStripWS(StringReplace($pLine, '</strong>', ''), 3)
	EndIf

	Return $pLine
EndFunc   ;==>_removeTags
; <=== _run =======================================================================================
; _run(String, [String], [String])
; ; Run a program and return StdOut.
; ;
; ; @param  String					Program
; ; @param  [String]				Parameters.
; ; @param  [String]				Path.
; ; @return String					StdOut.
; ; @access Public
Func _run($pProgram, $pParameters = Default, $pPath = Default)
	If $pParameters = Default Or StringLen($pParameters) = 0 Then $pParameters = ''
	If $pPath = Default Or StringLen($pPath) = 0 Then $pPath = @ScriptDir

	$pParameters = _trim($pParameters)
	If StringLen($pParameters) > 0 Then $pParameters = ' ' & $pParameters

	Local $hID = Run(@ComSpec & ' /C ""' & $pProgram & '"' & $pParameters & '"', $pPath, @SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($hID)

	Return StdoutRead($hID)
EndFunc   ;==>_run
; <=== _simplifyName ==============================================================================
; _simplifyName(String, [String])
; ; Remove path if filename is in same folder as script, also removes extension.
; ;
; ; @param  String			Filename.
; ; @param  String			Path to remove. Default: Script dir.
; ; @return String  		Filename without extension.
; ; @author guinness
Func _simplifyName($pFileName, $pPath = Default)
	If $pPath = Default Or StringLen($pPath) = 0 Then $pPath = $AppPath
	$pPath = _addSlash($pPath)

	Local $intLen = StringLen($pPath)
	If StringLeft($pFileName, $intLen) = $pPath Then
		$pPath = _fileNameInfo(StringMid($pFileName, $intLen + 1))
	EndIf
	Return $pPath
EndFunc   ;==>_simplifyName
; <=== _split2D() =================================================================================
; _split2D(String, [String], [String], , [Boolean])
; ; Returns a two dimensional array from a string.
; ;
; ; @param  String   	      String.
; ; @param  [String]   	    Entry separator. Default: Pipe ('|').
; ; @param  [String] 	      Row separator. Default: @LF.
; ; @param  [Boolean]       Remove trailing separators. Default: True.
; ; @return String   	      StdOut.
; ; @access Public
Func _split2D($pText, $pEntrySeparator = '|', $pRowSeparator = @LF, $pFixLast = True)
	Local $intDim1
	Local $intDim2 = 0
	Local $intColCount

	; Fix last item or row.
	If $pFixLast <> False Then
		$pText = _trim($pText, $pRowSeparator)
	EndIf

	Local $strArray1 = StringSplit($pText, $pRowSeparator, 3)
	Local $strArray2
	$intDim1 = UBound($strArray1, 1)
	Local $strTmp[$intDim1][0]

	For $i = 0 To $intDim1 - 1
		$strArray2 = StringSplit($strArray1[$i], $pEntrySeparator, 3)
		$intColCount = UBound($strArray2)
		If $intColCount > $intDim2 Then
			$intDim2 = $intColCount
			ReDim $strTmp[$intDim1][$intDim2]
		EndIf
		For $j = 0 To $intColCount - 1
			$strTmp[$i][$j] = $strArray2[$j]
		Next
	Next
	Return $strTmp
EndFunc   ;==>_split2D
; <=== _subString =================================================================================
; _subString(String, String, [Boolean])
; ; Returns substring from a string.
; ;
; ; @param  String   	      Original string.
; ; @param  String   	      Key text to search.
; ; @param  [Boolean]   	  True to return all text found. False to return one line.
; ; @return String   	      Text found.
; ; @access Public
Func _subString($pText, $pKey, $pFullReturn = True)
	Local $intPos
	If StringLen($pKey) = 0 Then
		Return ''
	Else
		$intPos = StringInStr($pText, $pKey, 2)
		If $intPos = 0 Then
			Return ''
		Else
			Return ($pFullReturn ? _
					_trim(StringMid($pText, $intPos), @LF) : _
					_trim(StringMid($pText, $intPos + StringLen($pKey), StringInStr($pText, @LF, 2, 1, $intPos) - $intPos - StringLen($pKey))))
		EndIf
	EndIf
EndFunc   ;==>_subString
; <=== _toHex =====================================================================================
; _toHex(String)
; ; Returns a HEX ASCII values of characters from a character string.
; ;
; ; @param  String          Characters string.
; ; @return String          HEX ASCII values of characters..
Func _toHex($pText)
	If StringLen($pText) = 0 Then Return ''
	Local $strHex
	For $strChar In StringSplit($pText, '', 2)
		$strHex &= Hex(Asc($strChar), 2)
	Next
	Return $strHex
EndFunc   ;==>_toHex
; <=== _trim ======================================================================================
; _trim(String, [String])
; ; Removes defined leading and trailing characters at both ends.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @return String     	    Processed text.
Func _trim($pText, $pCharacters = Default)
	Return _trimRight(_trimLeft($pText, $pCharacters), $pCharacters)
EndFunc   ;==>_trim
; <=== _trimLeft ==================================================================================
; _trimLeft(String, [String])
; ; Removes defined leading characters.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @return String     	    Processed text.
Func _trimLeft($pText, $pCharacters = Default)
	If StringLen($pText) = 0 Then Return ''
	If $pCharacters = Default Or StringLen($pCharacters) = 0 Then $pCharacters = ' '

	Local $intLen = StringLen($pCharacters)
	While StringLeft($pText, $intLen) = $pCharacters
		$pText = StringTrimLeft($pText, $intLen)
	WEnd
	Return $pText
EndFunc   ;==>_trimLeft
; <=== _trimRight =================================================================================
; _trimRight(String, [String])
; ; Removes defined trailing characters.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @return String     	    Processed text.
Func _trimRight($pText, $pCharacters = Default)
	If StringLen($pText) = 0 Then Return ''
	If $pCharacters = Default Or StringLen($pCharacters) = 0 Then $pCharacters = ' '

	Local $intLen = StringLen($pCharacters)
	While StringRight($pText, $intLen) = $pCharacters
		$pText = StringTrimRight($pText, $intLen)
	WEnd
	Return $pText
EndFunc   ;==>_trimRight
; ========================================================================= INTERNAL PROCEDURES ==>
#EndRegion INTERNAL PROCEDURES
