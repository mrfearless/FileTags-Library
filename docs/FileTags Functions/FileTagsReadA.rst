.. _FileTagsReadA:

=============
FileTagsReadA
=============

Read tags / keywords from a file's properties and return them in a buffer.

::

   FileTagsReadA PROTO pszFilename:DWORD, pszTagsBuffer:DWORD, lpdwTagsCount:DWORD


**Parameters**

* ``pszFilename`` - Pointer to a null terminated ANSI buffer that contains the full filepath of the file to read the tags / keywords for.

* ``pszTagsBuffer`` - A dword variable that will contain a pointer to a null terminated ANSI buffer containing the tags / keywords of the file pointed to by the pszFilename parameter. Each tag is seperated by a semi-colon character in the buffer. Use the GlobalFree function when this buffer is no longer required.

* ``lpdwTagsCount`` - (optional, can be 0), pointer to a dword variable to store the count of tags / keywords on succesful return of this function.


**Returns**

TRUE if successful, or FALSE otherwise. If successful, the variable pointed to by the pszTagsBuffer parameter will contain the tags / keywords and the variable pointed to by the lpdwTagsCount parameter (if specified), will contain the count of the tags / keywords.


**Notes**

The buffer returned in the pszTagsBuffer parameter should be freed with the GlobalFree function when no longer required.


**See Also**

:ref:`FileTagsWriteA<FileTagsWriteA>`, :ref:`FileTagsClearA<FileTagsClearA>`
