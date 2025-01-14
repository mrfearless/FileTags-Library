;==============================================================================
;
; FileTags x86 Library
;
; http://github.com/mrfearless
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
.model flat,stdcall
option casemap:none
;include \masm32\macros\macros.asm

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

include windows.inc

include user32.inc
includelib user32.lib

include kernel32.inc
includelib kernel32.lib

includelib ole32.lib
includelib shell32.lib
includelib propsys.lib

IFNDEF CoInitializeEx
CoInitializeEx PROTO pvReserved:DWORD, dwCoInit:DWORD
ENDIF
IFNDEF CoUninitialize
CoUninitialize PROTO
ENDIF
IFNDEF CoCreateInstance
CoCreateInstance PROTO rclsid:DWORD, pUnkOuter:DWORD, dwClsContext:DWORD, riid:DWORD, ppv:DWORD
ENDIF
IFNDEF CoTaskMemAlloc
CoTaskMemAlloc PROTO cb:DWORD
ENDIF
IFNDEF CoTaskMemFree
CoTaskMemFree PROTO pv:DWORD
ENDIF
IFNDEF SHCreateItemFromParsingName
SHCreateItemFromParsingName PROTO pszPath:DWORD, pbc:DWORD, riid:DWORD, ppv:DWORD
ENDIF
IFNDEF SHGetPropertyStoreFromParsingName
SHGetPropertyStoreFromParsingName PROTO pszPath:DWORD, pbc:DWORD, flags:DWORD, riid:DWORD, ppv:DWORD
ENDIF
IFNDEF RtlCompareMemory
RtlCompareMemory PROTO Source1:DWORD, Source2:DWORD, dwLength:DWORD
ENDIF
IFNDEF InitPropVariantFromStringAsVector
InitPropVariantFromStringAsVector PROTO psz:DWORD, ppropvar:DWORD
ENDIF
IFNDEF PropVariantClear
PropVariantClear PROTO pvValue:DWORD
ENDIF
IFNDEF PropVariantToString
PropVariantToString PROTO propvar:DWORD, psz:DWORD, cch:DWORD
ENDIF

include FileTags.inc

;------------------------------------------------------------------------------
; Prototypes for internal use
;------------------------------------------------------------------------------
_FT_ConvertStringToAnsi             PROTO lpszWideString:DWORD
_FT_ConvertStringToWide             PROTO lpszAnsiString:DWORD
_FT_ConvertStringFree               PROTO lpString:DWORD

;------------------------------------------------------------------------------
; COM Prototypes
;------------------------------------------------------------------------------
; IUnknown:
IUnknown_QueryInterface_Proto       TYPEDEF PROTO STDCALL pThis:DWORD, riid:DWORD, ppvObject:DWORD
IUnknown_AddRef_Proto               TYPEDEF PROTO STDCALL pThis:DWORD
IUnknown_Release_Proto              TYPEDEF PROTO STDCALL pThis:DWORD

; IPropertyStore:
IPropertyStore_GetCount_Proto       TYPEDEF PROTO STDCALL pThis:DWORD, cProps:DWORD
IPropertyStore_GetAt_Proto          TYPEDEF PROTO STDCALL pThis:DWORD, iProp:DWORD, pkey:DWORD
IPropertyStore_GetValue_Proto       TYPEDEF PROTO STDCALL pThis:DWORD, key:DWORD, pv:DWORD
IPropertyStore_SetValue_Proto       TYPEDEF PROTO STDCALL pThis:DWORD, key:DWORD, propvar:DWORD
IPropertyStore_Commit_Proto         TYPEDEF PROTO STDCALL pThis:DWORD

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
IUnknownVtbl                  STRUCT
    QueryInterface            IUnknown_QueryInterface_Ptr 0
    AddRef                    IUnknown_AddRef_Ptr 0
    Release                   IUnknown_Release_Ptr 0
IUnknownVtbl                  ENDS
ENDIF

IFNDEF IPropertyStoreVtbl
IPropertyStoreVtbl            STRUCT
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
GUID        STRUCT
    Data1   DD ?
    Data2   DW ?
    Data3   DW ?
    Data4   DB 8 DUP (?)
GUID        ENDS
ENDIF

IFNDEF PROPERTYKEY
PROPERTYKEY STRUCT
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
CALPWSTR        STRUCT
    cElems      DWORD ? ; MUST be set to the total number of elements of the array.
    pElems      DWORD ? ; An array of wchar_t* values.
CALPWSTR        ENDS
ENDIF

IFNDEF PROPVARIANT
PROPVARIANT     STRUCT
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
        pwszVal     DWORD ?             ; VT_LPWSTR
        pszVal      DWORD ?             ; VT_LPSTR
        boolVal     DWORD ?             ; VT_BOOL
        puuid       DWORD ?             ; CLSID pointer
        calpwstr    CALPWSTR <>         ; VT_VECTOR | VT_LPSTR
        ; etc
    ENDS
PROPVARIANT     ENDS
ENDIF


.CONST
;------------------------------------------------------------------------------
; COM Constants
;------------------------------------------------------------------------------
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


ALIGN 4
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
; * pszTagsBuffer - A dword variable that will contain a pointer to a null 
;   terminated UNICODE buffer containing the tags / keywords of the file pointed 
;   to by the pszFilename parameter. Each tag is seperated by a semi-colon 
;   character in the buffer. Use the GlobalFree function when this buffer is no 
;   longer required.
; 
; * lpdwTagsCount - (optional, can be 0), pointer to a dword variable to store 
;   the count of tags / keywords on succesful return of this function.
;
; Returns:
; 
; TRUE if successful, or FALSE otherwise. If successful, the variable pointed to
; by the pszTagsBuffer parameter will contain the tags / keywords and the 
; variable pointed to by the lpdwTagsCount parameter (if specified), will 
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
FileTagsReadW PROC USES EBX pszFilename:DWORD, pszTagsBuffer:DWORD, lpdwTagsCount:DWORD
    LOCAL pIPropertyStore:DWORD
    LOCAL cElems:DWORD
    LOCAL pElems:DWORD
    LOCAL nElem:DWORD
    LOCAL pElem:DWORD
    LOCAL Tags:PROPVARIANT
    LOCAL lpszWideTagString:DWORD
    LOCAL dwWideTagStringLength:DWORD
    LOCAL dwTagsArraySize:DWORD
    LOCAL pszTags:DWORD
    LOCAL nPos:DWORD
    
    IFDEF DEBUG32
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
    IFDEF DEBUG32
    PrintText 'SHGetPropertyStoreFromParsingName'
    ENDIF
    
    Invoke SHGetPropertyStoreFromParsingName, pszFilename, 0, GPS_READWRITE, Addr IID_IPropertyStore, Addr pIPropertyStore
    .IF eax != S_OK
        jmp FileTagsReadWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Initialize Tags PROPVARIANT
    ;--------------------------------------------------------------------------
    mov eax, 0
    lea ebx, Tags
    mov dword ptr [ebx+00], eax
    mov dword ptr [ebx+04], eax
    mov dword ptr [ebx+08], eax
    mov dword ptr [ebx+12], eax
    ;mov dword ptr [ebx].PROPVARIANT.vt, VT_EMPTY

    ;--------------------------------------------------------------------------
    ; Get Tags/Keywords Property From IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'IPropertyStoreVtbl.GetValue'
    ENDIF
    mov eax, pIPropertyStore
    mov ebx, [eax]
    Invoke [ebx].IPropertyStoreVtbl.GetValue, pIPropertyStore, Addr PKEY_Keywords, Addr Tags
    .IF eax != S_OK
        Invoke PropVariantClear, Addr Tags
        jmp FileTagsReadWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; See What The Tags PROPVARIANT Type Is And Decide What To Do 
    ;--------------------------------------------------------------------------
    lea ebx, Tags
    movzx eax, word ptr [ebx].PROPVARIANT.vt
    IFDEF DEBUG32
    PrintText 'PROPVARIANT.vt'
    PrintDec eax
    ENDIF
    
    .IF eax == VT_EMPTY
        ;----------------------------------------------------------------------
        ; No tags, but GetValue was successful
        ;----------------------------------------------------------------------
        IFDEF DEBUG32
        PrintText 'VT_EMPTY'
        ENDIF
        jmp FileTagsReadWExit
        
    .ELSEIF eax == VT_LPWSTR
        ;----------------------------------------------------------------------
        ; Get Length Of Tag / Keyword, Allocate Mem & Copy It To That Mem
        ;----------------------------------------------------------------------
        IFDEF DEBUG32
        PrintText 'VT_LPWSTR'
        ENDIF
        lea ebx, Tags
        mov eax, [ebx].PROPVARIANT.pwszVal
        .IF eax == 0
            Invoke PropVariantClear, Addr Tags
            jmp FileTagsReadWError
        .ENDIF
        mov lpszWideTagString, eax
        
        Invoke lstrlenW, lpszWideTagString
        .IF eax == 0
            Invoke PropVariantClear, Addr Tags
            jmp FileTagsReadWError
        .ENDIF
        add eax, 2 ; for null terminator
        mov dwWideTagStringLength, eax
        add eax, 4
        
        Invoke GlobalAlloc, GMEM_FIXED or GMEM_ZEROINIT, eax
        .IF eax == 0
            Invoke PropVariantClear, Addr Tags
            jmp FileTagsReadWError
        .ENDIF
        mov pszTags, eax
        
        Invoke lstrcpynW, pszTags, lpszWideTagString, dwWideTagStringLength

        mov cElems, 1
        
    .ELSEIF eax == (VT_VECTOR or VT_LPWSTR)
        ;----------------------------------------------------------------------
        ; Get Total Length Of All Tags / Keywords
        ;----------------------------------------------------------------------
        IFDEF DEBUG32
        PrintText 'VT_VECTOR or VT_LPWSTR'
        ENDIF
        
        mov dwTagsArraySize, 0

        lea ebx, Tags
        mov eax, [ebx].PROPVARIANT.calpwstr.cElems
        mov cElems, eax
        mov eax, [ebx].PROPVARIANT.calpwstr.pElems
        mov pElems, eax
        mov pElem, eax
        
        mov nElem, 0
        mov eax, 0
        .WHILE eax < cElems
            mov ebx, pElem
            mov eax, [ebx] ; pointer to wide string
            mov lpszWideTagString, eax
            
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
            
            add pElem, SIZEOF DWORD
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
        .IF eax == 0
            Invoke PropVariantClear, Addr Tags
            jmp FileTagsReadWError
        .ENDIF
        mov pszTags, eax
    
        ;----------------------------------------------------------------------
        ; Copy Each Tag / Keyword To Our Allocated Memory, Semi-Colon Seperated
        ;----------------------------------------------------------------------
        mov nPos, 0
        mov eax, pElems
        mov pElem, eax
        
        mov nElem, 0
        mov eax, 0
        .WHILE eax < cElems
            mov ebx, pElem
            mov eax, [ebx] ; pointer to wide string
            mov lpszWideTagString, eax
            
            Invoke lstrlenW, lpszWideTagString
            .IF eax != 0
                shl eax, 1 ; x2 for unicode chars to bytes
            .ENDIF
            mov dwWideTagStringLength, eax
            
            mov ebx, pszTags
            add ebx, nPos
            Invoke RtlMoveMemory, ebx, lpszWideTagString, dwWideTagStringLength
            mov eax, dwWideTagStringLength
            add nPos, eax
            
            mov eax, nElem
            inc eax
            .IF eax < cElems
                mov ebx, pszTags
                add ebx, nPos
                Invoke RtlMoveMemory, ebx, Addr szFT_SemiSpaceW, 4
                mov eax, 4
                add nPos, eax
            .ENDIF
            
            add pElem, SIZEOF DWORD
            inc nElem
            mov eax, nElem
        .ENDW
        
    .ELSE
        ;----------------------------------------------------------------------
        ; Some Other PROPVARIANT Type Found Instead
        ;----------------------------------------------------------------------
        IFDEF DEBUG32
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
    IFDEF DEBUG32
    PrintText 'Error FileTagsReadWError'
    ENDIF

    .IF pIPropertyStore != 0
        mov eax, pIPropertyStore
        mov ebx, [eax]
        Invoke [ebx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    .IF pszTagsBuffer != 0
        mov ebx, pszTagsBuffer
        mov eax, 0
        mov [ebx], eax
    .ENDIF
    .IF lpdwTagsCount != 0
        mov ebx, lpdwTagsCount
        mov eax, 0
        mov [ebx], eax
    .ENDIF
    mov eax, FALSE
    ret

FileTagsReadWExit:
    ;--------------------------------------------------------------------------
    ; Success Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'Success FileTagsReadWExit'
    ENDIF

    .IF pIPropertyStore != 0
        mov eax, pIPropertyStore
        mov ebx, [eax]
        Invoke [ebx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    .IF pszTagsBuffer != 0
        mov ebx, pszTagsBuffer
        mov eax, pszTags
        mov [ebx], eax
    .ENDIF
    .IF lpdwTagsCount != 0
        mov ebx, lpdwTagsCount
        mov eax, cElems
        mov [ebx], eax
    .ENDIF
    mov eax, TRUE
    ret
    
FileTagsReadW ENDP

ALIGN 4
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
; * pszTagsBuffer - A dword variable that will contain a pointer to a null 
;   terminated ANSI buffer containing the tags / keywords of the file pointed 
;   to by the pszFilename parameter. Each tag is seperated by a semi-colon 
;   character in the buffer. Use the GlobalFree function when this buffer is no 
;   longer required.
; 
; * lpdwTagsCount - (optional, can be 0), pointer to a dword variable to store 
;   the count of tags / keywords on succesful return of this function.
;
; Returns:
; 
; TRUE if successful, or FALSE otherwise. If successful, the variable pointed to
; by the pszTagsBuffer parameter will contain the tags / keywords and the 
; variable pointed to by the lpdwTagsCount parameter (if specified), will 
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
FileTagsReadA PROC USES EBX pszFilename:DWORD, pszTagsBuffer:DWORD, lpdwTagsCount:DWORD
    LOCAL lpszWideFilename:DWORD
    LOCAL lpszWideTagsBuffer:DWORD
    LOCAL lpszAnsiTagsBuffer:DWORD
    LOCAL bResult:DWORD
    
    IFDEF DEBUG32
    PrintText 'FileTagsReadA'
    ENDIF
    
    .IF pszFilename == 0 || pszTagsBuffer == 0
        mov eax, FALSE
        ret
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Convert Filename to UNICODE And Call FileTagsReadW
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringToWide, pszFilename
    mov lpszWideFilename, eax
    
    Invoke FileTagsReadW, lpszWideFilename, Addr lpszWideTagsBuffer, lpdwTagsCount
    mov bResult, eax ; save result of call
    
    ;--------------------------------------------------------------------------
    ; Free UNICODE Filename
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringFree, lpszWideFilename
    
    ;--------------------------------------------------------------------------
    ; Convert UNICODE Tags Buffer To ANSI And Return Buffer In pszTagsBuffer
    ;--------------------------------------------------------------------------
    .IF bResult == TRUE
        Invoke _FT_ConvertStringToAnsi, lpszWideTagsBuffer
        mov lpszAnsiTagsBuffer, eax
        
        mov ebx, pszTagsBuffer
        mov eax, lpszAnsiTagsBuffer
        mov [ebx], eax
    .ELSE
        mov ebx, pszTagsBuffer
        mov eax, 0
        mov [ebx], eax
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
    mov eax, bResult
    ret
FileTagsReadA ENDP

ALIGN 4
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
FileTagsWriteW PROC USES EBX pszFilename:DWORD, pszTagsBuffer:DWORD
    LOCAL pIPropertyStore:DWORD
    LOCAL Tags:PROPVARIANT ; Pointer to a PROPVARIANT structure of type VT_LPWSTR, used for the tags/keywords
    
    IFDEF DEBUG32
    PrintText 'FileTagsWriteW'
    ENDIF
    
    mov pIPropertyStore, 0
    
    .IF pszFilename == 0 || pszTagsBuffer == 0
        jmp FileTagsWriteWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Get IPropertyStore Object For File
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'SHGetPropertyStoreFromParsingName'
    ENDIF
    Invoke SHGetPropertyStoreFromParsingName, pszFilename, 0, GPS_READWRITE, Addr IID_IPropertyStore, Addr pIPropertyStore
    .IF eax != S_OK
        jmp FileTagsWriteWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Initialize Tags PROPVARIANT
    ;--------------------------------------------------------------------------
    mov eax, 0
    lea ebx, Tags
    mov dword ptr [ebx+00], eax
    mov dword ptr [ebx+04], eax
    mov dword ptr [ebx+08], eax
    mov dword ptr [ebx+12], eax
    ;mov dword ptr [ebx].PROPVARIANT.vt, VT_EMPTY    
    
    ;--------------------------------------------------------------------------
    ; Split String By Semi-Colon Into Array Of String Pointers
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'InitPropVariantFromStringAsVector'
    ENDIF
    Invoke InitPropVariantFromStringAsVector, pszTagsBuffer, Addr Tags
    .IF eax != S_OK
        jmp FileTagsWriteWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Set Tags/Keywords Property For IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'IPropertyStore.SetValue'
    ENDIF
    mov eax, pIPropertyStore
    mov ebx, [eax]
    Invoke [ebx].IPropertyStoreVtbl.SetValue, pIPropertyStore, Addr PKEY_Keywords, Addr Tags
    .IF eax != S_OK
        Invoke PropVariantClear, Addr Tags
        jmp FileTagsWriteWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Commit Property Changes For IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'IPropertyStore.Commit'
    ENDIF
    mov eax, pIPropertyStore
    mov ebx, [eax]
    Invoke [ebx].IPropertyStoreVtbl.Commit, pIPropertyStore
    .IF eax != S_OK
        Invoke PropVariantClear, Addr Tags
        jmp FileTagsWriteWError
    .ENDIF
    
    Invoke PropVariantClear, Addr Tags
    jmp FileTagsWriteWExit
    
FileTagsWriteWError:
    ;--------------------------------------------------------------------------
    ; Error Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'Error FileTagsWriteWError'
    ENDIF
    .IF pIPropertyStore != 0
        mov eax, pIPropertyStore
        mov ebx, [eax]
        Invoke [ebx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    mov eax, FALSE
    ret

FileTagsWriteWExit:
    ;--------------------------------------------------------------------------
    ; Success Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'Success FileTagsWriteWExit'
    ENDIF
    .IF pIPropertyStore != 0
        mov eax, pIPropertyStore
        mov ebx, [eax]
        Invoke [ebx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    mov eax, TRUE
    ret

FileTagsWriteW ENDP

ALIGN 4
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
FileTagsWriteA PROC USES EBX pszFilename:DWORD, pszTagsBuffer:DWORD
    LOCAL lpszWideFilename:DWORD
    LOCAL lpszWideTagsBuffer:DWORD
    LOCAL bResult:DWORD
    
    IFDEF DEBUG32
    PrintText 'FileTagsWriteA'
    ENDIF
    
    .IF pszFilename == 0 || pszTagsBuffer == 0
        mov eax, FALSE
        ret
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Convert Filename And Tags to UNICODE And Call FileTagsWriteW
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringToWide, pszFilename
    mov lpszWideFilename, eax
    
    Invoke _FT_ConvertStringToWide, pszTagsBuffer
    mov lpszWideTagsBuffer, eax
    
    Invoke FileTagsWriteW, lpszWideFilename, lpszWideTagsBuffer
    mov bResult, eax
    
    ;--------------------------------------------------------------------------
    ; Free UNICODE Filename And Tags
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringFree, lpszWideFilename
    Invoke _FT_ConvertStringFree, lpszWideTagsBuffer
    
    mov eax, bResult
    ret
FileTagsWriteA ENDP

ALIGN 4
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
FileTagsClearW PROC USES EBX pszFilename:DWORD
    LOCAL pIPropertyStore:DWORD
    LOCAL Tags:PROPVARIANT ; Pointer to a PROPVARIANT structure of type VT_LPWSTR, used for the tags/keywords
    
    IFDEF DEBUG32
    PrintText 'FileTagsClearW'
    ENDIF
    
    mov pIPropertyStore, 0
    
    .IF pszFilename == 0
        jmp FileTagsClearWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Get IPropertyStore Object For File
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'SHGetPropertyStoreFromParsingName'
    ENDIF
    Invoke SHGetPropertyStoreFromParsingName, pszFilename, 0, GPS_READWRITE, Addr IID_IPropertyStore, Addr pIPropertyStore
    .IF eax != S_OK
        jmp FileTagsClearWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Set Tags PROPVARIANT to VT_EMPTY to remove exiting tags
    ;--------------------------------------------------------------------------
    lea ebx, Tags
    mov eax, 0
    mov dword ptr [ebx+00], eax
    mov dword ptr [ebx+04], eax
    mov dword ptr [ebx+08], eax
    mov dword ptr [ebx+12], eax
    ;mov dword ptr [ebx].PROPVARIANT.vt, VT_EMPTY
    
    ;--------------------------------------------------------------------------
    ; Set Tags/Keywords Property For IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'IPropertyStore.SetValue'
    ENDIF
    mov eax, pIPropertyStore
    mov ebx, [eax]
    Invoke [ebx].IPropertyStoreVtbl.SetValue, pIPropertyStore, Addr PKEY_Keywords, Addr Tags
    .IF eax != S_OK
        jmp FileTagsClearWError
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Commit Property Changes For IPropertyStore Object
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'IPropertyStore.Commit'
    ENDIF
    mov eax, pIPropertyStore
    mov ebx, [eax]
    Invoke [ebx].IPropertyStoreVtbl.Commit, pIPropertyStore
    .IF eax != S_OK
        jmp FileTagsClearWError
    .ENDIF
    
    jmp FileTagsClearWExit
    
FileTagsClearWError:
    ;--------------------------------------------------------------------------
    ; Error Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'Error FileTagsClearWError'
    ENDIF
    .IF pIPropertyStore != 0
        mov eax, pIPropertyStore
        mov ebx, [eax]
        Invoke [ebx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    mov eax, FALSE
    ret

FileTagsClearWExit:
    ;--------------------------------------------------------------------------
    ; Success Exit
    ;--------------------------------------------------------------------------
    IFDEF DEBUG32
    PrintText 'Success FileTagsClearWExit'
    ENDIF
    .IF pIPropertyStore != 0
        mov eax, pIPropertyStore
        mov ebx, [eax]
        Invoke [ebx].IPropertyStoreVtbl.Release, pIPropertyStore
    .ENDIF
    mov eax, TRUE
    ret

FileTagsClearW ENDP

ALIGN 4
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
FileTagsClearA PROC USES EBX pszFilename:DWORD
    LOCAL lpszWideFilename:DWORD
    LOCAL bResult:DWORD
    
    IFDEF DEBUG32
    PrintText 'FileTagsClearA'
    ENDIF
    
    .IF pszFilename == 0
        mov eax, FALSE
        ret
    .ENDIF
    
    ;--------------------------------------------------------------------------
    ; Convert Filename to UNICODE And Call FileTagsClearW
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringToWide, pszFilename
    mov lpszWideFilename, eax
    
    Invoke FileTagsClearW, lpszWideFilename
    mov bResult, eax
    
    ;--------------------------------------------------------------------------
    ; Free UNICODE Filename
    ;--------------------------------------------------------------------------
    Invoke _FT_ConvertStringFree, lpszWideFilename
    
    mov eax, bResult
    ret
FileTagsClearA ENDP

ALIGN 4
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
FileTagsInit PROC
    Invoke CoInitializeEx, NULL, COINIT_APARTMENTTHREADED
    ret
FileTagsInit ENDP

ALIGN 4
;------------------------------------------------------------------------------
; FileTagsFree
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
; FileTagsInit
;
;------------------------------------------------------------------------------
FileTagsFree PROC
    Invoke CoUninitialize
    ret
FileTagsFree ENDP

ALIGN 4
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
_FT_ConvertStringToAnsi PROC lpszWideString:DWORD
    LOCAL dwAnsiStringSize:DWORD
    LOCAL lpszAnsiString:DWORD

    .IF lpszWideString == NULL
        mov eax, NULL
        ret
    .ENDIF
    Invoke WideCharToMultiByte, CP_UTF8, 0, lpszWideString, -1, NULL, 0, NULL, NULL
    .IF eax == 0
        ret
    .ENDIF
    mov dwAnsiStringSize, eax
    ;shl eax, 1 ; x2 to get non wide char count
    add eax, 4 ; add 4 for good luck and nulls
    Invoke GlobalAlloc, GMEM_FIXED or GMEM_ZEROINIT, eax
    .IF eax == NULL
        ret
    .ENDIF
    mov lpszAnsiString, eax    
    Invoke WideCharToMultiByte, CP_UTF8, 0, lpszWideString, -1, lpszAnsiString, dwAnsiStringSize, NULL, NULL
    .IF eax == 0
        ret
    .ENDIF
    mov eax, lpszAnsiString
    ret
_FT_ConvertStringToAnsi ENDP

ALIGN 4
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
_FT_ConvertStringToWide PROC lpszAnsiString:DWORD
    LOCAL dwWideStringSize:DWORD
    LOCAL lpszWideString:DWORD
    
    .IF lpszAnsiString == NULL
        mov eax, NULL
        ret
    .ENDIF
    Invoke MultiByteToWideChar, CP_UTF8, 0, lpszAnsiString, -1, NULL, 0
    .IF eax == 0
        ret
    .ENDIF
    mov dwWideStringSize, eax
    shl eax, 1 ; x2 to get non wide char count
    add eax, 4 ; add 4 for good luck and nulls
    Invoke GlobalAlloc, GMEM_FIXED or GMEM_ZEROINIT, eax
    .IF eax == NULL
        ret
    .ENDIF
    mov lpszWideString, eax
    Invoke MultiByteToWideChar, CP_UTF8, 0, lpszAnsiString, -1, lpszWideString, dwWideStringSize
    .IF eax == 0
        ret
    .ENDIF
    mov eax, lpszWideString
    ret
_FT_ConvertStringToWide ENDP

ALIGN 4
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
_FT_ConvertStringFree PROC lpString:DWORD
    mov eax, lpString
    .IF eax == NULL
        mov eax, FALSE
        ret
    .ENDIF
    Invoke GlobalFree, eax
    mov eax, TRUE
    ret
_FT_ConvertStringFree ENDP

END
