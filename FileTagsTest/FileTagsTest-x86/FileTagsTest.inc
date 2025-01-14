include windows.inc
include user32.inc
include kernel32.inc
include gdi32.inc
include shell32.inc
include comctl32.inc
include comdlg32.inc
include masm32.inc

includelib user32.lib
includelib kernel32.lib
includelib gdi32.lib
includelib shell32.lib
includelib comctl32.lib
includelib comdlg32.lib
includelib masm32.lib

include FileDialog.inc
includelib FileDialog.lib

include FileTags.inc
includelib FileTags.lib


;------------------------------------------------------------------------------
; FileTagsTest Prototypes
;------------------------------------------------------------------------------
WinMain                 PROTO :DWORD,:DWORD,:DWORD,:DWORD
WndProc                 PROTO :DWORD,:DWORD,:DWORD,:DWORD
GUIInit                 PROTO hWin:DWORD
GUITagOpStatus          PROTO hWin:DWORD, lpszStatusText:DWORD, bSuccess:DWORD
FileTagsBrowseForFile   PROTO hWin:DWORD



.CONST
;------------------------------------------------------------------------------
; FileTagsTest Constants
;------------------------------------------------------------------------------
ICO_MAIN                EQU 100

;FileTagsTest.dlg
IDD_DIALOG				EQU 1000
IDC_BtnOpenFile			EQU 1001
IDC_BtnExit				EQU 1002
IDC_EdtFileTags			EQU 1003
IDC_LblFileTags			EQU 1004
IDC_EdtFilename			EQU 1005
IDC_LblFilename			EQU 1006
IDC_BtnClearTags		EQU 1007
IDC_BtnSaveTags			EQU 1008
IDC_LblTagInfo			EQU 1009
IDC_TxtTagOpStatus      EQU 1010

; Main Menu 
IDM_MENU                EQU 10000
IDM_FILE_EXIT           EQU 10001
IDM_HELP_ABOUT          EQU 10101



.DATA
;------------------------------------------------------------------------------
; FileTagsTest Initialized Data
;------------------------------------------------------------------------------
IFDEF __UNICODE__
ClassName               DB 'D',0,'L',0,'G',0,'C',0,'L',0,'A',0,'S',0,'S',0
                        DB 0,0,0,0
AppName                 DB 'F',0,'i',0,'l',0,'e',0,'T',0,'a',0,'g',0,'s',0,' ',0,'T',0,'e',0,'s',0,'t',0,' ',0,'D',0,'i',0,'a',0,'l',0,'o',0,'g',0
                        DB 0,0,0,0
AboutMsg                DB 'g',0,'i',0,'t',0,'h',0,'u',0,'b',0,'.',0,'c',0,'o',0,'m',0,'/',0,'m',0,'r',0,'f',0,'e',0,'a',0,'r',0,'l',0,'e',0,'s',0,'s',0
                        DB 0,0,0,0
szZeroTags              DB 'F',0,'i',0,'l',0,'e',0,' ',0,'h',0,'a',0,'s',0,' ',0,'n',0,'o',0,' ',0,'t',0,'a',0,'g',0,'s',0
                        DB 0,0,0,0

szTagsReadSuccess       DB 'S',0,'u',0,'c',0,'c',0,'e',0,'s',0,'s',0,' ',0,'r',0,'e',0,'a',0,'d',0,'i',0,'n',0,'g',0,' ',0,'t',0,'a',0,'g',0,'s',0
                        DB ' ',0,'f',0,'r',0,'o',0,'m',0,' ',0,'f',0,'i',0,'l',0,'e',0,'.',0
                        DB 0,0,0,0
szTagsRead0TagsSuccess  DB 'S',0,'u',0,'c',0,'c',0,'e',0,'s',0,'s',0,',',0,' ',0,'b',0,'u',0,'t',0,' ',0,'f',0,'i',0,'l',0,'e',0,' ',0,'h',0,'a',0,'s',0
                        DB ' ',0,'n',0,'o',0,' ',0,'t',0,'a',0,'g',0,'s',0,'.',0
                        DB 0,0,0,0 
szTagsReadError         DB 'E',0,'r',0,'r',0,'o',0,'r',0,' ',0,'r',0,'e',0,'a',0,'d',0,'i',0,'n',0,'g',0,' ',0,'t',0,'a',0,'g',0,'s',0
                        DB ' ',0,'f',0,'r',0,'o',0,'m',0,' ',0,'f',0,'i',0,'l',0,'e',0,'.',0
                        DB 0,0,0,0 
szTagsWriteSuccess      DB 'S',0,'u',0,'c',0,'c',0,'e',0,'s',0,'s',0,' ',0,'w',0,'r',0,'i',0,'t',0,'i',0,'n',0,'g',0,' ',0,'t',0,'a',0,'g',0,'s',0
                        DB ' ',0,'t',0,'o',0,' ',0,'f',0,'i',0,'l',0,'e',0,'.',0
                        DB 0,0,0,0 
szTagsWriteError        DB 'E',0,'r',0,'r',0,'o',0,'r',0,' ',0,'w',0,'r',0,'i',0,'t',0,'i',0,'n',0,'g',0,' ',0,'t',0,'a',0,'g',0,'s',0
                        DB ' ',0,'t',0,'o',0,' ',0,'f',0,'i',0,'l',0,'e',0,'.',0
                        DB 0,0,0,0 
szTagsClearSuccess      DB 'S',0,'u',0,'c',0,'c',0,'e',0,'s',0,'s',0,' ',0,'d',0,'e',0,'l',0,'e',0,'t',0,'i',0,'n',0,'g',0,' ',0,'t',0,'a',0,'g',0,'s',0
                        DB ' ',0,'f',0,'r',0,'o',0,'m',0,' ',0,'f',0,'i',0,'l',0,'e',0,'.',0
                        DB 0,0,0,0 
szTagsClearError        DB 'E',0,'r',0,'r',0,'o',0,'r',0,' ',0,'d',0,'e',0,'l',0,'e',0,'t',0,'i',0,'n',0,'g',0,' ',0,'t',0,'a',0,'g',0,'s',0
                        DB ' ',0,'f',0,'r',0,'o',0,'m',0,' ',0,'f',0,'i',0,'l',0,'e',0,'.',0
                        DB 0,0,0,0 

ELSE
ClassName               DB 'DLGCLASS',0
AppName                 DB 'FileTags Test Dialog',0
AboutMsg                DB 'github.com/mrfearless',0
szZeroTags              DB 'File has no tags',0 

szTagsReadSuccess       DB 'Success reading tags from file.',0 
szTagsRead0TagsSuccess  DB 'Success, but file has no tags.',0 
szTagsReadError         DB 'Error reading tags from file.',0 
szTagsWriteSuccess      DB 'Success writing tags to file.',0 
szTagsWriteError        DB 'Error writing tags to file.',0 
szTagsClearSuccess      DB 'Success deleting tags from file.',0 
szTagsClearError        DB 'Error deleting tags from file.',0 

ENDIF

TagOpSuccess            DD TRUE

; FileTags:
dwTagsCount             DD 0
lpszTags                DD 0

; FileDialog:
dwFiles                 DD 0
lpszTagsFileName        DD 0

IFDEF __UNICODE__
szImageFiles            DB 'I',0,'m',0,'a',0,'g',0,'e',0,' ',0,'F',0,'i',0,'l',0,'e',0,'s',0
                        DB 0,0,0,0
szImageFilesSpec        DB '*',0,'.',0,'j',0,'p',0,'g',0,';',0,'*',0,'.',0,'p',0,'n',0,'g',0,';',0
                        DB '*',0,'.',0,'g',0,'i',0,'f',0,';',0,'*',0,'.',0,'j',0,'p',0,'e',0,'g',0,';',0
                        DB '*',0,'.',0,'b',0,'m',0,'p',0,';',0,'*',0,'.',0,'i',0,'c',0,'o',0,';',0
                        DB '*',0,'.',0,'t',0,'i',0,'f',0,';',0,'*',0,'.',0,'w',0,'e',0,'b',0,'p',0,';',0
                        DB 0,0,0,0

szAllFiles              DB 'A',0,'l',0,'l',0,' ',0,'F',0,'i',0,'l',0,'e',0,'s',0,' ',0,'(',0,'*',0,'.',0,'*',0,')',0
                        DB 0,0,0,0
szAllFilesSpec          DB '*',0,'.',0,'*',0
                        DB 0,0,0,0
ELSE
szImageFiles            DB 'Image Files',0
szImageFilesSpec        DB '*.jpg;*.png;*.gif;*.jpeg;'
                        DB '*.bmp;*.ico;*.tif;*.webp;'
                        DB 0
szAllFiles              DB 'All Files (*.*)',0
szAllFilesSpec          DB '*.*',0
ENDIF

FileSpecs \   
COMDLG_FILTERSPEC <Offset szImageFiles, Offset szImageFilesSpec>
COMDLG_FILTERSPEC <Offset szAllFiles, Offset szAllFilesSpec>

EdtTagsBuffer           DB 512 DUP (0)

.DATA?
;------------------------------------------------------------------------------
; FileTagsTest Uninitialized Data
;------------------------------------------------------------------------------
icc                     INITCOMMONCONTROLSEX <>
hInstance               DD ?
CommandLine             DD ?
hWnd                    DD ?
hIcoMain                DD ?
hEdtFilename            DD ?
hEdtFileTags            DD ?
hTxtTagOpStatus         DD ?
hBtnSaveTags            DD ?
hBtnClearTags           DD ?

