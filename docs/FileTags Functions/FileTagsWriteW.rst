.. _FileTagsWriteW:

==============
FileTagsWriteW
==============

Write a buffer containing tags / keywords to a file's properties.

::

   FileTagsWriteW PROTO pszFilename:DWORD, pszTagsBuffer:DWORD


**Parameters**

* ``pszFilename`` - Pointer to a null terminated UNICODE buffer that contains the full filepath of the file to write the tags / keywords to.

* ``pszTagsBuffer`` - Pointer to a null terminated UNICODE buffer that contains the tags / keywords to be written to the file pointed to by the pszFilename parameter.


**Returns**

TRUE if successful, or FALSE otherwise.


**See Also**

:ref:`FileTagsReadW<FileTagsReadW>`, :ref:`FileTagsClearW<FileTagsClearW>`
