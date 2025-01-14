.. _FileTagsClearW:

==============
FileTagsClearW
==============

Clear any tags/keywords from a file's properties.

::

   FileTagsClearW PROTO pszFilename:DWORD


**Parameters**

* ``pszFilename`` - Pointer to a null terminated UNICODE buffer that contains the full filepath of the file to clear the tags / keywords for.


**Returns**

TRUE if successful, or FALSE otherwise.


**See Also**

:ref:`FileTagsReadW<FileTagsReadW>`, :ref:`FileTagsWriteW<FileTagsWriteW>`
