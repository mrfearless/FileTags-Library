.. _FileTagsWriteA:

==============
FileTagsWriteA
==============

Write a buffer containing tags / keywords to a file's properties.

::

   FileTagsWriteA PROTO pszFilename:DWORD, pszTagsBuffer:DWORD


**Parameters**

* ``pszFilename`` - Pointer to a null terminated ANSI buffer that contains the full filepath of the file to write the tags / keywords to.

* ``pszTagsBuffer`` - Pointer to a null terminated ANSI buffer that contains the tags / keywords to be written to the file pointed to by the pszFilename parameter.


**Returns**

TRUE if successful, or FALSE otherwise.


**See Also**

:ref:`FileTagsReadA<FileTagsReadA>`, :ref:`FileTagsClearA<FileTagsClearA>`
