.686
.MMX
.XMM
.x64

option casemap : none
option win64 : 11
option frame : auto
option stackbase : rsp

_WIN64 EQU 1
WINVER equ 0501h

__UNICODE__ EQU 1
IFDEF __UNICODE__
UNICODE EQU 1 ; WinInc definition
ENDIF

;DEBUG64 EQU 1

IFDEF DEBUG64
    PRESERVEXMMREGS equ 1
    includelib \UASM\lib\x64\Debug64.lib
    DBG64LIB equ 1
    DEBUGEXE textequ <'\UASM\bin\DbgWin.exe'>
    include \UASM\include\debug64.inc
    .DATA
    RDBG_DbgWin	DB DEBUGEXE,0
    .CODE
ENDIF

include FileTagsTest.inc

.CODE

;------------------------------------------------------------------------------
; Startup
;------------------------------------------------------------------------------
WinMainCRTStartup proc FRAME
	Invoke GetModuleHandle, NULL
	mov hInstance, rax
	Invoke GetCommandLine
	mov CommandLine, rax
	Invoke InitCommonControls
	mov icc.dwSize, sizeof INITCOMMONCONTROLSEX
    mov icc.dwICC, ICC_COOL_CLASSES or ICC_STANDARD_CLASSES or ICC_WIN95_CLASSES
    Invoke InitCommonControlsEx, offset icc
    
    Invoke FileTagsInit
    
	Invoke WinMain, hInstance, NULL, CommandLine, SW_SHOWDEFAULT
	
	Invoke FileTagsFree
	
	Invoke ExitProcess, eax
    ret
WinMainCRTStartup endp
	

;------------------------------------------------------------------------------
; WinMain
;------------------------------------------------------------------------------
WinMain proc FRAME hInst:HINSTANCE, hPrev:HINSTANCE, CmdLine:LPSTR, iShow:DWORD
	LOCAL msg:MSG
	LOCAL wcex:WNDCLASSEX
	
	mov wcex.cbSize, sizeof WNDCLASSEX
	mov wcex.style, CS_HREDRAW or CS_VREDRAW
	lea rax, WndProc
	mov wcex.lpfnWndProc, rax
	mov wcex.cbClsExtra, 0
	mov wcex.cbWndExtra, DLGWINDOWEXTRA
	mov rax, hInst
	mov wcex.hInstance, rax
	mov wcex.hbrBackground, COLOR_BTNFACE+1 ; COLOR_WINDOW+1 ; 
	mov wcex.lpszMenuName, IDM_MENU ;NULL 
	lea rax, ClassName
	mov wcex.lpszClassName, rax
	Invoke LoadIcon, hInst, ICO_MAIN ; resource icon for main application icon
	mov hIcoMain, rax ; main application icon	
	mov wcex.hIcon, rax
	mov wcex.hIconSm, rax
	Invoke LoadCursor, NULL, IDC_ARROW
	mov wcex.hCursor, rax
	Invoke RegisterClassEx, addr wcex
	
	;Invoke CreateWindowEx, 0, addr ClassName, addr szAppName, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL
	Invoke CreateDialogParam, hInstance, IDD_DIALOG, 0, Addr WndProc, 0
	mov hWnd, rax
	
	Invoke ShowWindow, hWnd, SW_SHOWNORMAL
	Invoke UpdateWindow, hWnd
	
	.WHILE (TRUE)
		Invoke GetMessage, addr msg, NULL, 0, 0
		.BREAK .IF (!rax)		
		
        Invoke IsDialogMessage, hWnd, addr msg
        .IF rax == 0
            Invoke TranslateMessage, addr msg
            Invoke DispatchMessage, addr msg
        .ENDIF
	.ENDW
	
	mov rax, msg.wParam
	ret	
WinMain endp


;------------------------------------------------------------------------------
; WndProc - Main Window Message Loop
;------------------------------------------------------------------------------
WndProc proc FRAME hWin:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
    LOCAL wNotifyCode:DWORD
    
    mov eax, uMsg
	.IF eax == WM_INITDIALOG
		; Init Stuff Here
		Invoke GUIInit, hWin
		
	.ELSEIF eax == WM_COMMAND
        mov rax, wParam
        shr rax, 16
        mov wNotifyCode, eax
        mov rax, wParam
        and eax, 0FFFFh
		.IF rax == IDM_FILE_EXIT || eax == IDC_BtnExit
			Invoke SendMessage, hWin, WM_CLOSE, 0, 0
			
        ;----------------------------------------------------------------------
        ; Browse for a file, and read the file's tags/keywords if it has any.
        ;----------------------------------------------------------------------
        .ELSEIF rax == IDC_BtnOpenFile
            Invoke FileTagsBrowseForFile, hWin
            .IF rax == TRUE
                Invoke SendMessage, hEdtFilename, WM_SETTEXT, 0, 0
                Invoke SendMessage, hEdtFilename, WM_SETTEXT, 0, lpszTagsFileName
                Invoke SendMessage, hEdtFileTags, WM_SETTEXT, 0, 0
                Invoke SendMessage, hTxtTagOpStatus, WM_SETTEXT, 0, 0
                Invoke EnableWindow, hBtnSaveTags, FALSE
                Invoke EnableWindow, hBtnClearTags, FALSE
                
                .IF lpszTags != 0
                    Invoke GlobalFree, lpszTags
                    mov lpszTags, 0
                .ENDIF
                
                Invoke FileTagsRead, lpszTagsFileName, Addr lpszTags, Addr qwTagsCount
                .IF rax == TRUE
                    .IF qwTagsCount == 0
                        Invoke SendMessage, hEdtFileTags, WM_SETTEXT, 0, Addr szZeroTags
                        Invoke GUITagOpStatus, hWin, Addr szTagsRead0TagsSuccess, TRUE
                    .ELSE
                        Invoke SendMessage, hEdtFileTags, WM_SETTEXT, 0, lpszTags
                        Invoke GUITagOpStatus, hWin, Addr szTagsReadSuccess, TRUE
                    .ENDIF
                    Invoke EnableWindow, hBtnClearTags, TRUE
                .ELSE
                    Invoke GUITagOpStatus, hWin, Addr szTagsReadError, FALSE
                .ENDIF
                
            .ENDIF
        
        .ELSEIF rax == IDC_BtnClearTags
            .IF lpszTagsFileName == 0
                ; do nothing, no file 
            .ELSE
                Invoke FileTagsClear, lpszTagsFileName
                .IF rax == TRUE
                    Invoke SendMessage, hEdtFileTags, WM_SETTEXT, 0, 0
                    Invoke GUITagOpStatus, hWin, Addr szTagsClearSuccess, TRUE
                    Invoke EnableWindow, hBtnSaveTags, FALSE
                    Invoke EnableWindow, hBtnClearTags, FALSE
                .ELSE
                    Invoke GUITagOpStatus, hWin, Addr szTagsClearError, FALSE
                .ENDIF
                
            .ENDIF
        
        .ELSEIF rax == IDC_BtnSaveTags
            .IF lpszTagsFileName == 0
                ; do nothing, no file 
            .ELSE
                Invoke GetWindowText, hEdtFileTags, Addr EdtTagsBuffer, 512
                .IF rax == 0
                    ; do nothing, no tags
                .ELSE
                    Invoke FileTagsWrite, lpszTagsFileName, Addr EdtTagsBuffer
                    .IF eax == TRUE
                        Invoke GUITagOpStatus, hWin, Addr szTagsWriteSuccess, TRUE
                        Invoke EnableWindow, hBtnSaveTags, FALSE
                        Invoke EnableWindow, hBtnClearTags, TRUE
                    .ELSE
                        Invoke GUITagOpStatus, hWin, Addr szTagsWriteSuccess, FALSE
                    .ENDIF
                .ENDIF
            .ENDIF
        
        .ELSEIF rax == IDC_EdtFileTags
            .IF wNotifyCode == EN_CHANGE
                .IF lpszTagsFileName == 0
                    ; do nothing, no file
                .ELSE
                    Invoke EnableWindow, hBtnSaveTags, TRUE
                .ENDIF
            .ENDIF
			
		.ELSEIF rax == IDM_HELP_ABOUT
			Invoke ShellAbout, hWin, Addr AppName, Addr AboutMsg, NULL
			
		.ENDIF

    ;--------------------------------------------------------------------------
    ; Color the Tag Operation Status Label
    ;--------------------------------------------------------------------------
    .ELSEIF rax == WM_CTLCOLORSTATIC
        mov rax, lParam
        .IF rax == hTxtTagOpStatus
            .IF TagOpSuccess == TRUE
                Invoke SetTextColor, wParam, 043B008h
            .ELSE
                Invoke SetTextColor, wParam, 02B12C9h
            .ENDIF
            Invoke SetBkMode, wParam, OPAQUE
            Invoke SetBkColor, wParam, 0F0F0F0h
            Invoke GetStockObject, NULL_BRUSH
        .ELSE
            Invoke SetBkMode, wParam, OPAQUE
            Invoke SetBkColor, wParam, 0F0F0F0h
            Invoke GetStockObject, NULL_BRUSH
        .ENDIF
        ret

	.ELSEIF eax == WM_CLOSE
		Invoke DestroyWindow, hWin
		
	.ELSEIF eax == WM_DESTROY
		Invoke PostQuitMessage, NULL
		
	.ELSE
		Invoke DefWindowProc, hWin, uMsg, wParam, lParam ; rcx, edx, r8, r9
		ret
	.ENDIF
	xor rax, rax
	ret
WndProc endp

;------------------------------------------------------------------------------
; GUIInit
;------------------------------------------------------------------------------
GUIInit PROC FRAME hWin:QWORD
    Invoke SendMessage, hWin, WM_SETICON, ICON_BIG, hIcoMain
    Invoke SendMessage, hWin, WM_SETICON, ICON_SMALL, hIcoMain

    Invoke GetDlgItem, hWin, IDC_EdtFilename
    mov hEdtFilename, rax
    Invoke GetDlgItem, hWin, IDC_EdtFileTags
    mov hEdtFileTags, rax
    Invoke GetDlgItem, hWin, IDC_TxtTagOpStatus
    mov hTxtTagOpStatus, rax
    Invoke GetDlgItem, hWin, IDC_BtnSaveTags
    mov hBtnSaveTags, rax
    Invoke GetDlgItem, hWin, IDC_BtnClearTags
    mov hBtnClearTags, rax
    
    Invoke EnableWindow, hBtnSaveTags, FALSE
    Invoke EnableWindow, hBtnClearTags, FALSE
    
    ret
GUIInit ENDP

;------------------------------------------------------------------------------
; GUITagOpStatus
;------------------------------------------------------------------------------
GUITagOpStatus PROC FRAME hWin:QWORD, lpszStatusText:QWORD, bSuccess:QWORD
    Invoke SendMessage, hTxtTagOpStatus, WM_SETTEXT, 0, lpszStatusText
    mov rax, bSuccess
    mov TagOpSuccess, rax
    Invoke InvalidateRect, hTxtTagOpStatus, NULL, TRUE
    ret
GUITagOpStatus ENDP

;------------------------------------------------------------------------------
; FileTagsBrowseForFile
;------------------------------------------------------------------------------
FileTagsBrowseForFile PROC FRAME hWin:QWORD

    .IF qwFiles != 0 && lpszTagsFileName != 0
        Invoke GlobalFree, lpszTagsFileName
        mov lpszTagsFileName, 0
        mov qwFiles, 0
    .ENDIF
    
    Invoke FileOpenDialog, NULL, NULL, NULL, NULL, 2, Addr FileSpecs, hWin, FALSE, Addr qwFiles, Addr lpszTagsFileName
    .IF rax == TRUE
        .IF qwFiles != 0 && lpszTagsFileName != 0
            mov rax, TRUE
        .ELSE
            mov rax, FALSE
        .ENDIF
    .ENDIF
    ret
FileTagsBrowseForFile ENDP



end WinMainCRTStartup

