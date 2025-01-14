.686
.MMX
.XMM
.model flat,stdcall
option casemap:none

__UNICODE__ EQU 1

;DEBUG32 EQU 1

;IFDEF DEBUG32
;    PRESERVEXMMREGS equ 1
;    includelib M:\Masm32\lib\Debug32.lib
;    DBG32LIB equ 1
;    DEBUGEXE textequ <'M:\Masm32\DbgWin.exe'>
;    include M:\Masm32\include\debug32.inc
;    include msvcrt.inc
;    includelib ucrt.lib
;    includelib vcruntime.lib
;ENDIF


include FileTagsTest.inc

.code

start:

    Invoke GetModuleHandle, NULL
    mov hInstance, eax
    Invoke GetCommandLine
    mov CommandLine, eax
    Invoke InitCommonControls
    mov icc.dwSize, sizeof INITCOMMONCONTROLSEX
    mov icc.dwICC, ICC_COOL_CLASSES or ICC_STANDARD_CLASSES or ICC_WIN95_CLASSES
    Invoke InitCommonControlsEx, Offset icc
    
    Invoke FileTagsInit
    
    Invoke WinMain, hInstance, NULL, CommandLine, SW_SHOWDEFAULT
    
    Invoke FileTagsFree
    
    Invoke ExitProcess, eax

;------------------------------------------------------------------------------
; WinMain
;------------------------------------------------------------------------------
WinMain PROC hInst:HINSTANCE, hPrevInst:HINSTANCE, CmdLine:LPSTR, CmdShow:DWORD
    LOCAL wc:WNDCLASSEX
    LOCAL msg:MSG

    mov wc.cbSize, SIZEOF WNDCLASSEX
    mov wc.style, CS_HREDRAW or CS_VREDRAW
    mov wc.lpfnWndProc, Offset WndProc
    mov wc.cbClsExtra, NULL
    mov wc.cbWndExtra, DLGWINDOWEXTRA
    push hInst
    pop wc.hInstance
    mov wc.hbrBackground, COLOR_BTNFACE+1 ; COLOR_WINDOW+1
    mov wc.lpszMenuName, IDM_MENU
    mov wc.lpszClassName, Offset ClassName
    Invoke LoadIcon, hInstance, ICO_MAIN ; resource icon for main application icon
    mov hIcoMain, eax ; main application icon
    mov wc.hIcon, eax
    mov wc.hIconSm, eax
    Invoke LoadCursor, NULL, IDC_ARROW
    mov wc.hCursor,eax
    Invoke RegisterClassEx, Addr wc
    Invoke CreateDialogParam, hInstance, IDD_DIALOG, NULL, Addr WndProc, NULL
    mov hWnd, eax
    Invoke ShowWindow, hWnd, SW_SHOWNORMAL
    Invoke UpdateWindow, hWnd
    .WHILE TRUE
        Invoke GetMessage, Addr msg, NULL, 0, 0
        .BREAK .if !eax
        Invoke IsDialogMessage, hWnd, Addr msg
        .IF eax == 0
            Invoke TranslateMessage, Addr msg
            Invoke DispatchMessage, Addr msg
        .ENDIF
    .ENDW
    mov eax, msg.wParam
    ret
WinMain ENDP

;------------------------------------------------------------------------------
; WndProc - Main Window Message Loop
;------------------------------------------------------------------------------
WndProc PROC hWin:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
    LOCAL wNotifyCode:DWORD
    
    mov eax, uMsg
    .IF eax == WM_INITDIALOG
        ; Init Stuff Here
        Invoke GUIInit, hWin
        
    .ELSEIF eax == WM_COMMAND
        mov eax, wParam
        shr eax, 16
        mov wNotifyCode, eax
        mov eax, wParam
        and eax, 0FFFFh
        
        .IF eax == IDM_FILE_EXIT || eax == IDC_BtnExit
            Invoke SendMessage, hWin, WM_CLOSE, 0, 0
            
        ;----------------------------------------------------------------------
        ; Browse for a file, and read the file's tags/keywords if it has any.
        ;----------------------------------------------------------------------
        .ELSEIF eax == IDC_BtnOpenFile
            Invoke FileTagsBrowseForFile, hWin
            .IF eax == TRUE
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
                
                Invoke FileTagsRead, lpszTagsFileName, Addr lpszTags, Addr dwTagsCount
                .IF eax == TRUE
                    .IF dwTagsCount == 0
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
        
        .ELSEIF eax == IDC_BtnClearTags
            .IF lpszTagsFileName == 0
                ; do nothing, no file 
            .ELSE
                Invoke FileTagsClear, lpszTagsFileName
                .IF eax == TRUE
                    Invoke SendMessage, hEdtFileTags, WM_SETTEXT, 0, 0
                    Invoke GUITagOpStatus, hWin, Addr szTagsClearSuccess, TRUE
                    Invoke EnableWindow, hBtnSaveTags, FALSE
                    Invoke EnableWindow, hBtnClearTags, FALSE
                .ELSE
                    Invoke GUITagOpStatus, hWin, Addr szTagsClearError, FALSE
                .ENDIF
                
            .ENDIF
        
        .ELSEIF eax == IDC_BtnSaveTags
            .IF lpszTagsFileName == 0
                ; do nothing, no file 
            .ELSE
                Invoke GetWindowText, hEdtFileTags, Addr EdtTagsBuffer, 512
                .IF eax == 0
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
        
        .ELSEIF eax == IDC_EdtFileTags
            .IF wNotifyCode == EN_CHANGE
                .IF lpszTagsFileName == 0
                    ; do nothing, no file
                .ELSE
                    Invoke EnableWindow, hBtnSaveTags, TRUE
                .ENDIF
            .ENDIF
        
        .ELSEIF eax == IDM_HELP_ABOUT
            Invoke ShellAbout, hWin, Addr AppName, Addr AboutMsg,NULL
            
        .ENDIF

    ;--------------------------------------------------------------------------
    ; Color the Tag Operation Status Label
    ;--------------------------------------------------------------------------
    .ELSEIF eax == WM_CTLCOLORSTATIC
        mov eax, lParam
        .IF eax == hTxtTagOpStatus
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
        Invoke DefWindowProc, hWin, uMsg, wParam, lParam
        ret
    .ENDIF
    xor eax, eax
    ret
WndProc ENDP

;------------------------------------------------------------------------------
; GUIInit
;------------------------------------------------------------------------------
GUIInit PROC hWin:DWORD
    Invoke SendMessage, hWin, WM_SETICON, ICON_BIG, hIcoMain
    Invoke SendMessage, hWin, WM_SETICON, ICON_SMALL, hIcoMain

    Invoke GetDlgItem, hWin, IDC_EdtFilename
    mov hEdtFilename, eax
    Invoke GetDlgItem, hWin, IDC_EdtFileTags
    mov hEdtFileTags, eax
    Invoke GetDlgItem, hWin, IDC_TxtTagOpStatus
    mov hTxtTagOpStatus, eax
    Invoke GetDlgItem, hWin, IDC_BtnSaveTags
    mov hBtnSaveTags, eax
    Invoke GetDlgItem, hWin, IDC_BtnClearTags
    mov hBtnClearTags, eax
    
    Invoke EnableWindow, hBtnSaveTags, FALSE
    Invoke EnableWindow, hBtnClearTags, FALSE
    
    ret
GUIInit ENDP

;------------------------------------------------------------------------------
; GUITagOpStatus
;------------------------------------------------------------------------------
GUITagOpStatus PROC hWin:DWORD, lpszStatusText:DWORD, bSuccess:DWORD
    Invoke SendMessage, hTxtTagOpStatus, WM_SETTEXT, 0, lpszStatusText
    mov eax, bSuccess
    mov TagOpSuccess, eax
    Invoke InvalidateRect, hTxtTagOpStatus, NULL, TRUE
    ret
GUITagOpStatus ENDP

;------------------------------------------------------------------------------
; FileTagsBrowseForFile
;------------------------------------------------------------------------------
FileTagsBrowseForFile PROC hWin:DWORD

    .IF dwFiles != 0 && lpszTagsFileName != 0
        Invoke GlobalFree, lpszTagsFileName
        mov lpszTagsFileName, 0
        mov dwFiles, 0
    .ENDIF
    
    Invoke FileOpenDialog, NULL, NULL, NULL, NULL, 2, Addr FileSpecs, hWin, FALSE, Addr dwFiles, Addr lpszTagsFileName
    .IF eax == TRUE
        .IF dwFiles != 0 && lpszTagsFileName != 0
            mov eax, TRUE
        .ELSE
            mov eax, FALSE
        .ENDIF
    .ENDIF
    ret
FileTagsBrowseForFile ENDP



end start
