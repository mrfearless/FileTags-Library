.. _FileTagsClearA:

==============
FileTagsClearA
==============

Clear any tags/keywords from a file's properties.

::

   FileTagsClearA PROTO pszFilename:DWORD


**Parameters**

* ``pszFilename`` - Pointer to a null terminated ANSI buffer that contains the full filepath of the file to clear the tags / keywords for.


**Returns**

TRUE if successful, or FALSE otherwise.


**See Also**

:ref:`FileTagsReadA<FileTagsReadA>`, :ref:`FileTagsWriteA<FileTagsWriteA>`
