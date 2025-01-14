;==============================================================================
;
; FileTags x64 Library
;
; http://github.com/mrfearless/FileDialog-Library
;
; This software is provided 'as-is', without any express or implied warranty. 
; In no event will the author be held liable for any damages arising from the 
; use of this software.
;
;==============================================================================
;
; FileTags Library consists of functions that wrap the COM implementation of 
; the IPropertyStore object, for accessing the PKEY_Keywords property metadata
; of a file.
; 
; The IPropertyStore object and PKEY_Keywords property is used for the 'Tags' 
; feature in Windows Explorer, when viewing the properties of a file, or when 
; the Details Pane of Windows Explorer is open and a file is selected.
; 
; Thus the FileTags Library functions hide the complexities of interacting with 
; the IPropertyStore COM object, allowing the user to read, write and clear the
; keywords / tags for a file. 

; For more details on file tagging or the file property metadata:
;
; - https://karl-voit.at/2019/11/26/Tagging-Files-With-Windows-10/
;
; - https://github.com/Dijji/FileMeta/wiki/XP,-Vista-and-File-Metadata
;
;------------------------------------------------------------------------------
; 
; References:
;
; https://stackoverflow.com/a/8663182
; https://www.autohotkey.com/boards/viewtopic.php?style=7&t=109629
; https://stackoverflow.com/questions/50686095/adding-editing-new-extended-properties-to-the-details-tab-in-existing-files
; https://learn.microsoft.com/en-us/answers/questions/591953/jpeg-(-jpg)-adding-user-tags-values-to-a-photo-fil
; https://karl-voit.at/2019/11/26/Tagging-Files-With-Windows-10/
; https://github.com/Dijji/FileMeta/wiki/XP,-Vista-and-File-Metadata
; https://github.com/Dijji/FileMeta
;
;------------------------------------------------------------------------------
.686
.MMX
.XMM
.x64

option casemap : none
option win64 : 11
option frame : auto

;DEBUG64 EQU 1
;
;IFDEF DEBUG64
;    PRESERVEXMMREGS equ 1
;    includelib \UASM\lib\x64\Debug64.lib
;    DBG64LIB equ 1
;    DEBUGEXE textequ <'\UASM\bin\DbgWin.exe'>
;    include \UASM\include\debug64.inc
;    .DATA
;    RDBG_DbgWin    DB DEBUGEXE,0
;    .CODE
;ENDIF

includelib user32.lib
includelib kernel32.lib
includelib ole32.lib
includelib shell32.lib
includelib propsys.lib

include FileTags.inc

IFNDEF NULL
NULL EQU 0
ENDIF
IFNDEF TRUE
TRUE EQU 1
ENDIF
IFNDEF FALSE
FALSE EQU 0
ENDIF
IFNDEF GlobalAlloc
GlobalAlloc PROTO uFlags:DWORD, dwBytes:QWORD
ENDIF
IFNDEF GlobalFree
GlobalFree PROTO pMem:QWORD
ENDIF
IFNDEF GMEM_FIXED
GMEM_FIXED EQU 0000h
ENDIF
IFNDEF GMEM_ZEROINIT
GMEM_ZEROINIT EQU 0040h
ENDIF
IFNDEF RtlMoveMemory
RtlMoveMemory PROTO Destination:QWORD, Source:QWORD, qwLength:QWORD
ENDIF
IFNDEF lstrlenA
lstrlenA PROTO lpString:QWORD
ENDIF
IFNDEF lstrlenW
lstrlenW PROTO lpString:QWORD
ENDIF
IFNDEF lstrcpynW
lstrcpynW PROTO lpStringDst:QWORD, lpStringSrc:QWORD, iMaxLength:DWORD
ENDIF
IFNDEF lstrcpynA
lstrcpynA PROTO lpStringDst:QWORD, lpStringSrc:QWORD, iMaxLength:DWORD
ENDIF

IFNDEF WideCharToMultiByte 
WideCharToMultiByte PROTO CodePage:DWORD, dwFlags:DWORD, lpWideCharStr:QWORD, ccWideChar:DWORD, lpMultiByteStr:QWORD, cbMultiByte:DWORD, lpDefaultChar:QWORD, lpUsedDefaultChar:QWORD
ENDIF
IFNDEF MultiByteToWideChar 
MultiByteToWideChar PROTO CodePage:DWORD, dwFlags:DWORD, lpMultiByteStr:QWORD, cbMultiByte:DWORD, lpWideCharStr:QWORD, ccWideChar:DWORD
ENDIF
IFNDEF CoInitializeEx
CoInitializeEx PROTO pvReserved:QWORD, dwCoInit:DWORD
ENDIF
IFNDEF CoUninitialize
CoUninitialize PROTO
ENDIF
IFNDEF CoCreateInstance
CoCreateInstance PROTO rclsid:QWORD, pUnkOuter:QWORD, dwClsContext:DWORD, riid:QWORD, ppv:QWORD
ENDIF
IFNDEF CoTaskMemFree
CoTaskMemFree PROTO pv:QWORD
ENDIF
IFNDEF SHCreateItemFromParsingName
SHCreateItemFromParsingName PROTO pszPath:QWORD, pbc:QWORD, riid:QWORD, ppv:QWORD
ENDIF
IFNDEF SHGetPropertyStoreFromParsingName
SHGetPropertyStoreFromParsingName PROTO pszPath:QWORD, pbc:QWORD, flags:QWORD, riid:QWORD, ppv:QWORD
ENDIF
IFNDEF InitPropVariantFromStringAsVector
InitPropVariantFromStringAsVector PROTO psz:QWORD, ppropvar:QWORD
ENDIF
IFNDEF PropVariantClear
PropVariantClear PROTO pvValue:QWORD
ENDIF
IFNDEF PropVariantToString
PropVariantToString PROTO propvar:QWORD, psz:QWORD, cch:QWORD
ENDIF

;------------------------------------------------------------------------------
; Prototypes for internal use
;------------------------------------------------------------------------------
_FT_ConvertStringToAnsi             PROTO lpszWideString:QWORD
_FT_ConvertStringToWide             PROTO lpszAnsiString:QWORD
_FT_ConvertStringFree               PROTO lpString:QWORD

;------------------------------------------------------------------------------
; COM Prototypes
;------------------------------------------------------------------------------
; IUnknown:
IUnknown_QueryInterface_Proto       TYPEDEF PROTO pThis:QWORD, riid:QWORD, ppvObject:QWORD
IUnknown_AddRef_Proto               TYPEDEF PROTO pThis:QWORD
IUnknown_Release_Proto              TYPEDEF PROTO pThis:QWORD

; IPropertyStore:
IPropertyStore_GetCount_Proto       TYPEDEF PROTO pThis:QWORD, cProps:QWORD
IPropertyStore_GetAt_Proto          TYPEDEF PROTO pThis:QWORD, iProp:QWORD, pkey:QWORD
IPropertyStore_GetValue_Proto       TYPEDEF PROTO pThis:QWORD, key:QWORD, pv:QWORD
IPropertyStore_SetValue_Proto       TYPEDEF PROTO pThis:QWORD, key:QWORD, propvar:QWORD
IPropertyStore_Commit_Proto         TYPEDEF PROTO pThis:QWORD

;------------------------------------------------------------------------------
; Pointer To Prototypes
;------------------------------------------------------------------------------
; IUnknown
IUnknown_QueryInterface_Ptr         TYPEDEF PTR IUnknown_QueryInterface_Proto
IUnknown_AddRef_Ptr                 TYPEDEF PTR IUnknown_AddRef_Proto
IUnknown_Release_Ptr                TYPEDEF PTR IUnknown_Release_Proto

; IPropertyStore:
IPropertyStore_QueryInterface_Ptr   TYPEDEF PTR IUnknown_QueryInterface_Proto
IPropertyStore_AddRef_Ptr           TYPEDEF PTR IUnknown_AddRef_Proto
IPropertyStore_Release_Ptr          TYPEDEF PTR IUnknown_Release_Proto
IPropertyStore_GetCount_Ptr         TYPEDEF PTR IPropertyStore_GetCount_Proto
IPropertyStore_GetAt_Ptr            TYPEDEF PTR IPropertyStore_GetAt_Proto
IPropertyStore_GetValue_Ptr         TYPEDEF PTR IPropertyStore_GetValue_Proto
IPropertyStore_SetValue_Ptr         TYPEDEF PTR IPropertyStore_SetValue_Proto
IPropertyStore_Commit_Ptr           TYPEDEF PTR IPropertyStore_Commit_Proto

;------------------------------------------------------------------------------
; COM Structures
;------------------------------------------------------------------------------
IFNDEF IUnknownVtbl
IUnknownVtbl                  STRUCT 8
    QueryInterface            IUnknown_QueryInterface_Ptr 0
    AddRef                    IUnknown_AddRef_Ptr 0
    Release                   IUnknown_Release_Ptr 0
IUnknownVtbl                  ENDS
ENDIF

IFNDEF IPropertyStoreVtbl
IPropertyStoreVtbl            STRUCT 8
    QueryInterface            IUnknown_QueryInterface_Ptr 0
    AddRef                    IUnknown_AddRef_Ptr 0
    Release                   IUnknown_Release_Ptr 0
    GetCount                  IPropertyStore_GetCount_Ptr 0
    GetAt                     IPropertyStore_GetAt_Ptr 0
    GetValue                  IPropertyStore_GetValue_Ptr 0
    SetValue                  IPropertyStore_SetValue_Ptr 0
    Commit                    IPropertyStore_Commit_Ptr 0
IPropertyStoreVtbl            ENDS
ENDIF

IFNDEF GUID
GUID        STRUCT 8
    Data1   DD ?
    Data2   DW ?
    Data3   DW ?
    Data4   DB 8 DUP (?)
GUID        ENDS
ENDIF

IFNDEF PROPERTYKEY
PROPERTYKEY STRUCT 8
    fmtid   GUID <>
    pid     DWORD ?
PROPERTYKEY ENDS
ENDIF

IFNDEF LARGE_INTEGER
LARGE_INTEGER UNION
    STRUCT
      LowPart  DWORD ?
      HighPart DWORD ?
    ENDS
  QuadPart QWORD ?
LARGE_INTEGER ENDS
ENDIF

IFNDEF ULARGE_INTEGER
ULARGE_INTEGER UNION
    STRUCT
      LowPart  DWORD ?
      HighPart DWORD ?
    ENDS
  QuadPart QWORD ?
ULARGE_INTEGER ENDS
ENDIF

IFNDEF CALPWSTR
CALPWSTR        STRUCT 8
    cElems      DWORD ? ; MUST be set to the total number of elements of the array.
    pElems      QWORD ? ; An array of wchar_t* values.
CALPWSTR        ENDS
ENDIF

IFNDEF PROPVARIANT
PROPVARIANT     STRUCT 8
    vt          DW ?
    wReserved1  DW ?
    wReserved2  DW ?
    wReserved3  DW ?
    UNION
        hVal        LARGE_INTEGER <>    ; VT_I8
        uhVal       ULARGE_INTEGER <>   ; VT_UI8
        uint64Val   QWORD ?             
        fltVal      REAL8 ?             ; VT_R8
        uintVal     DWORD ?             ; VT_UINT
        pwszVal     QWORD ?             ; VT_LPWSTR
        pszVal      QWORD ?             ; VT_LPSTR
        boolVal     DWORD ?             ; VT_BOOL
        puuid       QWORD ?             ; CLSID pointer
        calpwstr    CALPWSTR <>         ; VT_VECTOR | VT_LPSTR
        ; etc
    ENDS
PROPVARIANT     ENDS
ENDIF

.CONST
;------------------------------------------------------------------------------
; COM Constants
;------------------------------------------------------------------------------
CP_ACP	EQU	0
CP_UTF7	EQU	65000
CP_UTF8	EQU	65001

COINIT_APARTMENTTHREADED EQU 02h

IFNDEF S_OK
S_OK EQU 0
ENDIF
IFNDEF S_FALSE
S_FALSE EQU 1
ENDIF
IFNDEF HRESULT
HRESULT TYPEDEF DWORD
ENDIF
IFNDEF HRESULT_ERROR_CANCELLED
HRESULT_ERROR_CANCELLED EQU 800704C7h
ENDIF
IFNDEF MF_E_ATTRIBUTENOTFOUND
MF_E_ATTRIBUTENOTFOUND EQU 0C00D36E6h
ENDIF
IFNDEF MF_E_INVALIDREQUEST
MF_E_INVALIDREQUEST EQU 0C00D36B2h
ENDIF
IFNDEF E_NOINTERFACE
E_NOINTERFACE EQU 80004002h
ENDIF
IFNDEF E_INVALIDARG
E_INVALIDARG EQU 080070057h
ENDIF
IFNDEF E_OUTOFMEMORY
E_OUTOFMEMORY EQU 08007000Eh
ENDIF

; PropVariant Types
VT_EMPTY            EQU  0 ; A property with a type indicator of VT_EMPTY has no data associated with it; that is, the size of the value is zero.
VT_NULL             EQU  1 ; This is like a pointer to NULL.
VT_LPWSTR           EQU 31 ; A pointer to a null-terminated Unicode string in the user default locale.
VT_VECTOR           EQU 1000h 

GPS_READWRITE	    EQU 02h ; Get Property Store Flag For SHGetPropertyStoreFromParsingName

.DATA
ALIGN 4
;------------------------------------------------------------------------------
; COM CLSIDs
;------------------------------------------------------------------------------
IID_IUnknown        GUID <000000000h,00000h,00000h,<0C0h,000h,000h,000h,000h,000h,000h,046h>>
IID_IPropertyStore  GUID <0886D8EEBh,08CF2h,04446h,<08Dh,002h,0CDh,0BAh,01Dh,0BDh,0CFh,099h>>

; Property Key IDs:
PKEY_Keywords       PROPERTYKEY <<0F29F85E0h,04FF9h,01068h,<0ABh,091h,008h,000h,02Bh,027h,0B3h,0D9h>>,5> ; The set of keywords (also known as "tags") assigned to the item.


szFT_SemiSpaceW     DB ';',0,' ',0
                    DB 0,0,0,0
szFT_SemiSpaceA     DB "; ",0

.CODE


;------------------------------------------------------------------------------
;
; Remember to add Invoke CoInitializeEx, NULL, COINIT_APARTMENTTHREADED at 
; start of program and Invoke CoUninitialize at end, or just use FileTagsInit 
; and FileTagsFree functions defined here instead.
;
;------------------------------------------------------------------------------


ALIGN 8
;------------------------------------------------------------------------------
; FileTagsReadW
;
; Read tags / keywords from a file's properties and return them in a buffer.
;
; Parameters:
;
; * pszFilename - Pointer to a null terminated UNICODE buffer that contains the 
;   full filepath of the file to read the tags / keywords for.
;
; * pszTagsBuffer - A qword variable that will contain a pointer to a null 
;   terminated UNICODE buffer containing the tags / keywords of the file pointed 
;   to by the pszFilename parameter. Each tag is seperated by a semi-colon 
;   character in the buffer. Use the GlobalFree function when this buffer is no 
;   longer required.
; 
; * lpqwTagsCount - (optional, can be 0), pointer to a qword variable to store 
;   the count of tags / keywords on succesful return of this function.
;
; Returns:
; 
; TRUE if successful, or FALSE otherwise. If successful, the variable pointed to
; by the pszTagsBuffer parameter will contain the tags / keywords and the 
; variable pointed to by the lpqwTagsCount parameter (if specified), will 
; contain the count of the tags / keywords.
;
; Notes:
;
; The buffer returned in the pszTagsBuffer parameter should be freed with the
; GlobalFree function when no longer required.
;
; See Also:
;
; FileTagsWriteW, FileTagsClearW
;
;------------------------------------------------------------------------------
FileTagsReadW PROC FRAME USES RBX pszFilename:QWORD, pszTagsBuffer:QWORD, lpqwTagsCount:QWORD
    LOCAL pIPropertyStore:QWORD
    LOCAL cElems:DWORD
    LOCAL pElems:QWORD
    LOCAL nElem:DWORD
    LOCAL pElem:QWORD
    LOCAL Tags:PROPVARIANT
    LOCAL lpszWideTagString:QWORD
    LOCAL dwWideTagStringLength:DWORD
    LOCAL dwTagsArraySize:DWORD
    LOCAL pszTags:QWORD
    LOCAL nPos:QWORD
    
    IFDEF DEBUG64
    PrintText 'FileTagsReadW'
    ENDIF
    
    mov pszTags, 0
    mov pIPropertyStore, 0
    mov cElems, 0
    
    .IF pszFilename == 0 || pszTagsBuffer == 0
        jmp FileTagsReadWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Get IPropertyStore Object For File
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'SHGetPropertyStoreFromParsingName'
    ENDIF
    
    Invoke SHGetPropertyStoreFromParsingName, pszFilename, 0, GPS_READWRITE, Addr IID_IPropertyStore, Addr pIPropertyStore
    .IF rax != S_OK
        jmp FileTagsReadWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Initialize Tags PROPVARIANT
    ;--------------------------------------------------------------------------
    mov rax, 0
    lea rbx, Tags
    mov qword ptr [rbx+00], rax
    mov qword ptr [rbx+08], rax

    ;--------------------------------------------------------------------------
    ; Get Tags/Keywords Property From IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'IPropertyStoreVtbl.GetValue'
    ENDIF
    mov rax, pIPropertyStore
    mov rbx, [rax]
    Invoke [rbx].IPropertyStoreVtbl.GetValue, pIPropertyStore, Addr PKEY_Keywords, Addr Tags
    .IF rax != S_OK
        Invoke PropVariantClear, Addr Tags
        jmp FileTagsReadWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; See What The Tags PROPVARIANT Type Is And Decide What To Do 
    ;--------------------------------------------------------------------------
    lea rbx, Tags
    movzx eax, word ptr [rbx].PROPVARIANT.vt
    IFDEF DEBUG64
    PrintText 'PROPVARIANT.vt'
    PrintDec rax
    ENDIF
    
    .IF rax == VT_EMPTY
        ;----------------------------------------------------------------------
        ; No tags, but GetValue was successful
        ;----------------------------------------------------------------------
        IFDEF DEBUG64
        PrintText 'VT_EMPTY'
        ENDIF
        jmp FileTagsReadWExit
        
    .ELSEIF rax == VT_LPWSTR
        ;----------------------------------------------------------------------
        ; Get Length Of Tag / Keyword, Allocate Mem & Copy It To That Mem
        ;----------------------------------------------------------------------
        IFDEF DEBUG64
        PrintText 'VT_LPWSTR'
        ENDIF
        lea rbx, Tags
        mov rax, [rbx].PROPVARIANT.pwszVal
        .IF rax == 0
            Invoke PropVariantClear, Addr Tags
            jmp FileTagsReadWError
        .ENDIF
        mov lpszWideTagString, rax
        
        Invoke lstrlenW, lpszWideTagString
        .IF rax == 0
            Invoke PropVariantClear, Addr Tags
            jmp FileTagsReadWError
        .ENDIF
        shl eax, 1 ; x2 for unicode chars to bytes
        add eax, 2 ; for null terminator
        mov dwWideTagStringLength, eax
        add eax, 4
        
        Invoke GlobalAlloc, GMEM_FIXED or GMEM_ZEROINIT, eax
        .IF rax == 0
            Invoke PropVariantClear, Addr Tags
            jmp FileTagsReadWError
        .ENDIF
        mov pszTags, rax
        
        Invoke lstrcpynW, pszTags, lpszWideTagString, dwWideTagStringLength
    
;        Invoke PropVariantToString, pPropVariantTags, pszTagsBuffer, qwTagsBufferSize
;        .IF eax != S_OK
;            Invoke PropVariantClear, pPropVariantTags
;            jmp FileTagsReadWError
;        .ENDIF
        mov cElems, 1
        
    .ELSEIF rax == (VT_VECTOR or VT_LPWSTR)
        ;----------------------------------------------------------------------
        ; Get Total Length Of All Tags / Keywords
        ;----------------------------------------------------------------------
        IFDEF DEBUG64
        PrintText 'VT_VECTOR or VT_LPWSTR'
        ENDIF
        
        mov dwTagsArraySize, 0

        lea rbx, Tags
        mov eax, dword ptr [rbx].PROPVARIANT.calpwstr.cElems
        mov cElems, eax
        mov rax, [rbx].PROPVARIANT.calpwstr.pElems
        mov pElems, rax
        mov pElem, rax
        
        mov nElem, 0
        mov eax, 0
        .WHILE eax < cElems
            mov rbx, pElem
            mov rax, [rbx] ; pointer to wide string
            mov lpszWideTagString, rax
            
            Invoke lstrlenW, lpszWideTagString
            .IF eax != 0
                shl eax, 1 ; x2 for unicode chars to bytes
            .ENDIF
            add dwTagsArraySize, eax
            
            mov eax, nElem
            inc eax
            .IF eax < cElems
                add dwTagsArraySize, 4 ; for wide '; ' and null
            .ENDIF
            
            add pElem, SIZEOF QWORD
            inc nElem
            mov eax, nElem
        .ENDW
        add dwTagsArraySize, 2 ; for null terminator
    
        ;----------------------------------------------------------------------
        ; Allocate Memory For All Tags / Keywords
        ;----------------------------------------------------------------------
        mov eax, dwTagsArraySize
        add eax, 4
        
        Invoke GlobalAlloc, GMEM_FIXED or GMEM_ZEROINIT, eax
        .IF rax == 0
            Invoke PropVariantClear, Addr Tags
            jmp FileTagsReadWError
        .ENDIF
        mov pszTags, rax
    
        ;----------------------------------------------------------------------
        ; Copy Each Tag / Keyword To Our Allocated Memory, Semi-Colon Seperated
        ;----------------------------------------------------------------------
        mov nPos, 0
        mov rax, pElems
        mov pElem, rax
        
        mov nElem, 0
        mov eax, 0
        .WHILE eax < cElems
            mov rbx, pElem
            mov rax, [rbx] ; pointer to wide string
            mov lpszWideTagString, rax
            
            Invoke lstrlenW, lpszWideTagString
            .IF eax != 0
                shl eax, 1 ; x2 for unicode chars to bytes
            .ENDIF
            mov dwWideTagStringLength, eax
            
            mov rbx, pszTags
            add rbx, nPos
            Invoke RtlMoveMemory, rbx, lpszWideTagString, dwWideTagStringLength
            mov eax, dwWideTagStringLength
            add nPos, rax
            
            mov eax, nElem
            inc eax
            .IF eax < cElems
                mov rbx, pszTags
                add rbx, nPos
                Invoke RtlMoveMemory, rbx, Addr szFT_SemiSpaceW, 4
                mov rax, 4
                add nPos, rax
            .ENDIF
            
            add pElem, SIZEOF QWORD
            inc nElem
            mov eax, nElem
        .ENDW
        
    .ELSE
        ;----------------------------------------------------------------------
        ; Some Other PROPVARIANT Type Found Instead
        ;----------------------------------------------------------------------
        IFDEF DEBUG64
        PrintText 'Something else'
        ENDIF
        Invoke PropVariantClear, Addr Tags
        jmp FileTagsReadWError
        
    .ENDIF
    
    Invoke PropVariantClear, Addr Tags
    jmp FileTagsReadWExit
    

FileTagsReadWError:
    ;--------------------------------------------------------------------------
    ; Error Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'Error FileTagsReadWError'
    ENDIF

    .IF pIPropertyStore != 0
        mov rax, pIPropertyStore
        mov rbx, [rax]
        Invoke [rbx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    .IF pszTagsBuffer != 0
        mov rbx, pszTagsBuffer
        mov rax, 0
        mov [rbx], rax
    .ENDIF
    .IF lpqwTagsCount != 0
        mov rbx, lpqwTagsCount
        mov rax, 0
        mov [rbx], rax
    .ENDIF
    mov rax, FALSE
    ret

FileTagsReadWExit:
    ;--------------------------------------------------------------------------
    ; Success Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'Success FileTagsReadWExit'
    ENDIF

    .IF pIPropertyStore != 0
        mov rax, pIPropertyStore
        mov rbx, [rax]
        Invoke [rbx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    .IF pszTagsBuffer != 0
        mov rbx, pszTagsBuffer
        mov rax, pszTags
        mov [rbx], rax
    .ENDIF
    .IF lpqwTagsCount != 0
        mov rbx, lpqwTagsCount
        mov eax, cElems
        mov [rbx], rax
    .ENDIF
    mov rax, TRUE
    ret
    
FileTagsReadW ENDP

ALIGN 8
;------------------------------------------------------------------------------
; FileTagsReadA
;
; Read tags / keywords from a file's properties and return them in a buffer.
;
; Parameters:
;
; * pszFilename - Pointer to a null terminated ANSI buffer that contains the 
;   full filepath of the file to read the tags / keywords for.
;
; * pszTagsBuffer - A qword variable that will contain a pointer to a null 
;   terminated ANSI buffer containing the tags / keywords of the file pointed 
;   to by the pszFilename parameter. Each tag is seperated by a semi-colon 
;   character in the buffer. Use the GlobalFree function when this buffer is no 
;   longer required.
; 
; * lpqwTagsCount - (optional, can be 0), pointer to a qword variable to store 
;   the count of tags / keywords on succesful return of this function.
;
; Returns:
; 
; TRUE if successful, or FALSE otherwise. If successful, the variable pointed to
; by the pszTagsBuffer parameter will contain the tags / keywords and the 
; variable pointed to by the lpqwTagsCount parameter (if specified), will 
; contain the count of the tags / keywords.
;
; Notes:
;
; The buffer returned in the pszTagsBuffer parameter should be freed with the
; GlobalFree function when no longer required.
;
; See Also:
;
; FileTagsWriteA, FileTagsClearA
;
;------------------------------------------------------------------------------
FileTagsReadA PROC FRAME USES RBX pszFilename:QWORD, pszTagsBuffer:QWORD, lpqwTagsCount:QWORD
    LOCAL lpszWideFilename:QWORD
    LOCAL lpszWideTagsBuffer:QWORD
    LOCAL lpszAnsiTagsBuffer:QWORD
    LOCAL bResult:QWORD
    
    IFDEF DEBUG64
    PrintText 'FileTagsReadA'
    ENDIF
    
    .IF pszFilename == 0 || pszTagsBuffer == 0
        mov rax, FALSE
        ret
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Convert Filename to UNICODE And Call FileTagsReadW
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringToWide, pszFilename
    mov lpszWideFilename, rax
    
    Invoke FileTagsReadW, lpszWideFilename, Addr lpszWideTagsBuffer, lpqwTagsCount
    mov bResult, rax ; save result of call
    
    ;--------------------------------------------------------------------------
    ; Free UNICODE Filename
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringFree, lpszWideFilename
    
    ;--------------------------------------------------------------------------
    ; Convert UNICODE Tags Buffer To ANSI And Return Buffer In pszTagsBuffer
    ;--------------------------------------------------------------------------
    .IF bResult == TRUE
        Invoke _FT_ConvertStringToAnsi, lpszWideTagsBuffer
        mov lpszAnsiTagsBuffer, rax
        
        mov rbx, pszTagsBuffer
        mov rax, lpszAnsiTagsBuffer
        mov [rbx], rax
    .ELSE
        mov rbx, pszTagsBuffer
        mov rax, 0
        mov [rbx], rax
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Free UNICODE Tags Buffer.
    ;--------------------------------------------------------------------------
    .IF lpszWideTagsBuffer != 0
        Invoke GlobalFree, lpszWideTagsBuffer
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; ANSI Tags Buffer To Be Freed At Later Time By User With GlobalFree
    ;--------------------------------------------------------------------------
    mov rax, bResult
    ret
FileTagsReadA ENDP

ALIGN 8
;------------------------------------------------------------------------------
; FileTagsWriteW
;
; Write a buffer containing tags / keywords to a file's properties.
;
; Parameters:
;
; * pszFilename - Pointer to a null terminated UNICODE buffer that contains the 
;   full filepath of the file to write the tags / keywords to.
;
; * pszTagsBuffer - Pointer to a null terminated UNICODE buffer that contains 
;   the tags / keywords to be written to the file pointed to by the pszFilename 
;   parameter.
; 
; Returns:
; 
; TRUE if successful, or FALSE otherwise.
;
; See Also:
;
; FileTagsReadW, FileTagsClearW
;
;------------------------------------------------------------------------------
FileTagsWriteW PROC FRAME USES RBX pszFilename:QWORD, pszTagsBuffer:QWORD
    LOCAL pIPropertyStore:QWORD
    LOCAL Tags:PROPVARIANT ; Pointer to a PROPVARIANT structure of type VT_LPWSTR, used for the tags/keywords
    
    IFDEF DEBUG64
    PrintText 'FileTagsWriteW'
    ENDIF
    
    mov pIPropertyStore, 0
    
    .IF pszFilename == 0 || pszTagsBuffer == 0
        jmp FileTagsWriteWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Get IPropertyStore Object For File
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'SHGetPropertyStoreFromParsingName'
    ENDIF
    Invoke SHGetPropertyStoreFromParsingName, pszFilename, 0, GPS_READWRITE, Addr IID_IPropertyStore, Addr pIPropertyStore
    .IF rax != S_OK
        jmp FileTagsWriteWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Initialize Tags PROPVARIANT
    ;--------------------------------------------------------------------------
    mov rax, 0
    lea rbx, Tags
    mov qword ptr [rbx+00], rax
    mov qword ptr [rbx+08], rax
    
    ;--------------------------------------------------------------------------
    ; Split String By Semi-Colon Into Array Of String Pointers
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'InitPropVariantFromStringAsVector'
    ENDIF
    Invoke InitPropVariantFromStringAsVector, pszTagsBuffer, Addr Tags
    .IF rax != S_OK
        jmp FileTagsWriteWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Set Tags/Keywords Property For IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'IPropertyStore.SetValue'
    ENDIF
    mov rax, pIPropertyStore
    mov rbx, [rax]
    Invoke [rbx].IPropertyStoreVtbl.SetValue, pIPropertyStore, Addr PKEY_Keywords, Addr Tags
    .IF rax != S_OK
        Invoke PropVariantClear, Addr Tags
        jmp FileTagsWriteWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Commit Property Changes For IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'IPropertyStore.Commit'
    ENDIF
    mov rax, pIPropertyStore
    mov rbx, [rax]
    Invoke [rbx].IPropertyStoreVtbl.Commit, pIPropertyStore
    .IF rax != S_OK
        Invoke PropVariantClear, Addr Tags
        jmp FileTagsWriteWError
    .ENDIF
    
    Invoke PropVariantClear, Addr Tags
    jmp FileTagsWriteWExit
    
FileTagsWriteWError:
    ;--------------------------------------------------------------------------
    ; Error Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'Error FileTagsWriteWError'
    ENDIF
    .IF pIPropertyStore != 0
        mov rax, pIPropertyStore
        mov rbx, [rax]
        Invoke [rbx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    mov rax, FALSE
    ret

FileTagsWriteWExit:
    ;--------------------------------------------------------------------------
    ; Success Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'Success FileTagsWriteWExit'
    ENDIF
    .IF pIPropertyStore != 0
        mov rax, pIPropertyStore
        mov rbx, [rax]
        Invoke [rbx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    mov rax, TRUE
    ret
    
FileTagsWriteW ENDP

ALIGN 8
;------------------------------------------------------------------------------
; FileTagsWriteA
;
; Write a buffer containing tags / keywords to a file's properties.
;
; Parameters:
;
; * pszFilename - Pointer to a null terminated ANSI buffer that contains the 
;   full filepath of the file to write the tags / keywords to.
;
; * pszTagsBuffer - Pointer to a null terminated ANSI buffer that contains the 
;   tags / keywords to be written to the file pointed to by the pszFilename 
;   parameter.
; 
; Returns:
; 
; TRUE if successful, or FALSE otherwise.
;
; See Also:
;
; FileTagsReadA, FileTagsClearA
;
;------------------------------------------------------------------------------
FileTagsWriteA PROC FRAME USES RBX pszFilename:QWORD, pszTagsBuffer:QWORD
    LOCAL lpszWideFilename:QWORD
    LOCAL lpszWideTagsBuffer:QWORD
    LOCAL bResult:QWORD
    
    IFDEF DEBUG64
    PrintText 'FileTagsWriteA'
    ENDIF
    
    .IF pszFilename == 0 || pszTagsBuffer == 0
        mov rax, FALSE
        ret
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Convert Filename And Tags to UNICODE And Call FileTagsWriteW
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringToWide, pszFilename
    mov lpszWideFilename, rax
    
    Invoke _FT_ConvertStringToWide, pszTagsBuffer
    mov lpszWideTagsBuffer, rax
    
    Invoke FileTagsWriteW, lpszWideFilename, lpszWideTagsBuffer
    mov bResult, rax
    
    ;--------------------------------------------------------------------------
    ; Free UNICODE Filename And Tags
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringFree, lpszWideFilename
    Invoke _FT_ConvertStringFree, lpszWideTagsBuffer
    
    mov rax, bResult
    ret
FileTagsWriteA ENDP

ALIGN 8
;------------------------------------------------------------------------------
; FileTagsClearW
;
; Clear any tags/keywords from a file's properties.
;
; Parameters:
;
; * pszFilename - Pointer to a null terminated UNICODE buffer that contains the 
;   full filepath of the file to clear the tags / keywords for.
;
; Returns:
; 
; TRUE if successful, or FALSE otherwise.
;
; See Also:
;
; FileTagsReadW, FileTagsWriteW
;
;------------------------------------------------------------------------------
FileTagsClearW PROC USES RBX pszFilename:QWORD
    LOCAL pIPropertyStore:QWORD
    LOCAL Tags:PROPVARIANT ; Pointer to a PROPVARIANT structure of type VT_LPWSTR, used for the tags/keywords
    
    IFDEF DEBUG64
    PrintText 'FileTagsClearW'
    ENDIF
    
    mov pIPropertyStore, 0
    
    .IF pszFilename == 0
        jmp FileTagsClearWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Get IPropertyStore Object For File
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'SHGetPropertyStoreFromParsingName'
    ENDIF
    Invoke SHGetPropertyStoreFromParsingName, pszFilename, 0, GPS_READWRITE, Addr IID_IPropertyStore, Addr pIPropertyStore
    .IF rax != S_OK
        jmp FileTagsClearWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Set Tags PROPVARIANT to VT_EMPTY to remove exiting tags
    ;--------------------------------------------------------------------------
    lea rbx, Tags
    mov rax, 0
    mov qword ptr [rbx+00], rax
    mov qword ptr [rbx+08], rax
    
    ;--------------------------------------------------------------------------
    ; Set Tags/Keywords Property For IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'IPropertyStore.SetValue'
    ENDIF
    mov rax, pIPropertyStore
    mov rbx, [rax]
    Invoke [rbx].IPropertyStoreVtbl.SetValue, pIPropertyStore, Addr PKEY_Keywords, Addr Tags
    .IF rax != S_OK
        jmp FileTagsClearWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Commit Property Changes For IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'IPropertyStore.Commit'
    ENDIF
    mov rax, pIPropertyStore
    mov rbx, [rax]
    Invoke [rbx].IPropertyStoreVtbl.Commit, pIPropertyStore
    .IF rax != S_OK
        jmp FileTagsClearWError
    .ENDIF
    
    jmp FileTagsClearWExit
    
FileTagsClearWError:
    ;--------------------------------------------------------------------------
    ; Error Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'Error FileTagsClearWError'
    ENDIF
    .IF pIPropertyStore != 0
        mov rax, pIPropertyStore
        mov rbx, [rax]
        Invoke [rbx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    mov rax, FALSE
    ret

FileTagsClearWExit:
    ;--------------------------------------------------------------------------
    ; Success Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG64
    PrintText 'Success FileTagsClearWExit'
    ENDIF
    .IF pIPropertyStore != 0
        mov rax, pIPropertyStore
        mov rbx, [rax]
        Invoke [rbx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    mov rax, TRUE
    ret

FileTagsClearW ENDP

ALIGN 8
;------------------------------------------------------------------------------
; FileTagsClearA
;
; Clear any tags/keywords from a file's properties.
;
; Parameters:
;
; * pszFilename - Pointer to a null terminated ANSI buffer that contains the 
;   full filepath of the file to clear the tags / keywords for.
;
; Returns:
; 
; TRUE if successful, or FALSE otherwise.
;
; See Also:
;
; FileTagsReadA, FileTagsWriteA
;
;------------------------------------------------------------------------------
FileTagsClearA PROC FRAME USES RBX pszFilename:QWORD
    LOCAL lpszWideFilename:QWORD
    LOCAL bResult:QWORD
    
    IFDEF DEBUG64
    PrintText 'FileTagsClearA'
    ENDIF
    
    .IF pszFilename == 0
        mov rax, FALSE
        ret
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Convert Filename to UNICODE And Call FileTagsClearW
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringToWide, pszFilename
    mov lpszWideFilename, rax
    
    Invoke FileTagsClearW, lpszWideFilename
    mov bResult, rax
    
    ;--------------------------------------------------------------------------
    ; Free UNICODE Filename
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringFree, lpszWideFilename
    
    mov rax, bResult
    ret
FileTagsClearA ENDP

ALIGN 8
;------------------------------------------------------------------------------
; FileTagsInit
;
; Initialize COM for FileTags functions.
;
; Parameters:
;
; None
;
; Returns:
; 
; Return values from CoInitializeEx.
;
; See Also:
;
; FileTagsFree
;
;------------------------------------------------------------------------------
FileTagsInit PROC FRAME
    Invoke CoInitializeEx, NULL, COINIT_APARTMENTTHREADED
    ret
FileTagsInit ENDP

ALIGN 8
;------------------------------------------------------------------------------
; FileTagsInit
;
; Uninitialize COM for FileTags functions.
;
; Parameters:
;
; None
;
; Returns:
; 
; None.
;
; See Also:
;
; FileTagsFree
;
;------------------------------------------------------------------------------
FileTagsFree PROC FRAME
    Invoke CoUninitialize
    ret
FileTagsFree ENDP



ALIGN 8
;------------------------------------------------------------------------------
; _FT_ConvertStringToAnsi 
;
; Converts a Wide/Unicode string to an ANSI/UTF8 string.
;
; Parameters:
; 
; * lpszWideString - pointer to a wide string to convert to an Ansi string.
; 
; Returns:
; 
; A pointer to the Ansi string if successful, or NULL otherwise.
; 
; Notes:
;
; The string that is converted should be freed when it is no longer needed with 
; a call to the _FT_ConvertStringFree function.
;
; See Also:
;
; _FT_ConvertStringToWide, _FT_ConvertStringFree
; 
;------------------------------------------------------------------------------
_FT_ConvertStringToAnsi PROC FRAME lpszWideString:QWORD
    LOCAL dwAnsiStringSize:DWORD
    LOCAL lpszAnsiString:QWORD

    .IF lpszWideString == NULL
        mov rax, NULL
        ret
    .ENDIF
    Invoke WideCharToMultiByte, CP_UTF8, 0, lpszWideString, -1, NULL, 0, NULL, NULL
    .IF rax == 0
        ret
    .ENDIF
    mov dwAnsiStringSize, eax
    ;shl rax, 1 ; x2 to get non wide char count
    add rax, 4 ; add 4 for good luck and nulls
    Invoke GlobalAlloc, GMEM_FIXED or GMEM_ZEROINIT, rax
    .IF rax == NULL
        ret
    .ENDIF
    mov lpszAnsiString, rax    
    Invoke WideCharToMultiByte, CP_UTF8, 0, lpszWideString, -1, lpszAnsiString, dwAnsiStringSize, NULL, NULL
    .IF rax == 0
        ret
    .ENDIF
    mov rax, lpszAnsiString
    ret
_FT_ConvertStringToAnsi ENDP

ALIGN 8
;------------------------------------------------------------------------------
; _FT_ConvertStringToWide
;
; Converts a Ansi string to an Wide/Unicode string.
;
; Parameters:
; 
; * lpszAnsiString - pointer to an Ansi string to convert to a Wide string.
; 
; Returns:
; 
; A pointer to the Wide string if successful, or NULL otherwise.
; 
; Notes:
;
; The string that is converted should be freed when it is no longer needed with 
; a call to the _FT_ConvertStringFree function.
;
; See Also:
;
; _FT_ConvertStringToAnsi, _FT_ConvertStringFree
; 
;------------------------------------------------------------------------------
_FT_ConvertStringToWide PROC FRAME lpszAnsiString:QWORD
    LOCAL dwWideStringSize:DWORD
    LOCAL lpszWideString:QWORD
    
    .IF lpszAnsiString == NULL
        mov rax, NULL
        ret
    .ENDIF
    Invoke MultiByteToWideChar, CP_UTF8, 0, lpszAnsiString, -1, NULL, 0
    .IF rax == 0
        ret
    .ENDIF
    mov dwWideStringSize, eax
    shl rax, 1 ; x2 to get non wide char count
    add rax, 4 ; add 4 for good luck and nulls
    Invoke GlobalAlloc, GMEM_FIXED or GMEM_ZEROINIT, rax
    .IF rax == NULL
        ret
    .ENDIF
    mov lpszWideString, rax
    Invoke MultiByteToWideChar, CP_UTF8, 0, lpszAnsiString, -1, lpszWideString, dwWideStringSize
    .IF rax == 0
        ret
    .ENDIF
    mov rax, lpszWideString
    ret
_FT_ConvertStringToWide ENDP

ALIGN 8
;------------------------------------------------------------------------------
; _FT_ConvertStringFree
;
; Frees a string created by _FT_ConvertStringToWide or _FT_ConvertStringToAnsi
;
; Parameters:
; 
; * lpString - pointer to a converted string to free.
; 
; Returns:
; 
; None.
; 
; See Also:
;
; _FT_ConvertStringToWide, _FT_ConvertStringToAnsi
; 
;------------------------------------------------------------------------------
_FT_ConvertStringFree PROC FRAME lpString:QWORD
    mov rax, lpString
    .IF rax == NULL
        mov rax, FALSE
        ret
    .ENDIF
    Invoke GlobalFree, rax
    mov rax, TRUE
    ret
_FT_ConvertStringFree ENDP


END

