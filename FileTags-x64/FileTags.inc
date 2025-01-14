;==============================================================================
;
; FileTags x64 Library
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

FileTagsInit    PROTO
FileTagsFree    PROTO

FileTagsReadA   PROTO pszFilename:QWORD, pszTagsBuffer:QWORD, lpqwTagsCount:QWORD
FileTagsReadW   PROTO pszFilename:QWORD, pszTagsBuffer:QWORD, lpqwTagsCount:QWORD

FileTagsWriteA  PROTO pszFilename:QWORD, pszTagsBuffer:QWORD
FileTagsWriteW  PROTO pszFilename:QWORD, pszTagsBuffer:QWORD

FileTagsClearA  PROTO pszFilename:QWORD
FileTagsClearW  PROTO pszFilename:QWORD

IFDEF __UNICODE__
    FileTagsRead        EQU <FileTagsReadW>
    FileTagsWrite       EQU <FileTagsWriteW>
    FileTagsClear       EQU <FileTagsClearW>
ELSE
    IFDEF UNICODE ; WinInc Define For Unicode
        FileTagsRead        EQU <FileTagsReadW>
        FileTagsWrite       EQU <FileTagsWriteW>
        FileTagsClear       EQU <FileTagsClearW>
    ELSE ; ANSI
        FileTagsRead        EQU <FileTagsReadA>
        FileTagsWrite       EQU <FileTagsWriteA>
        FileTagsClear       EQU <FileTagsClearA>
    ENDIF
ENDIF










