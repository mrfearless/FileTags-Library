.. _Overview:

============
Overview
============

FileTags Library consists of functions that wrap the COM implementation of the `IPropertyStore  <https://learn.microsoft.com/en-us/windows/win32/api/propsys/nn-propsys-ipropertystore>`_ object, for accessing the `PKEY_Keywords <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-keywords>`_ property metadata of a file.

The `IPropertyStore  <https://learn.microsoft.com/en-us/windows/win32/api/propsys/nn-propsys-ipropertystore>`_ object and `PKEY_Keywords <https://learn.microsoft.com/en-us/windows/win32/properties/props-system-keywords>`_ property is used for the 'Tags' feature in Windows Explorer, when viewing the properties of a file, or when the Details Pane of Windows Explorer is open and a file is selected.

Thus the FileTags Library functions hide the complexities of interacting with the `IPropertyStore  <https://learn.microsoft.com/en-us/windows/win32/api/propsys/nn-propsys-ipropertystore>`_ COM object, allowing the user to read, write and clear the keywords / tags for a file. 

For more details on file tagging or the file property metadata:

*  `https://karl-voit.at/2019/11/26/Tagging-Files-With-Windows-10/ <https://karl-voit.at/2019/11/26/Tagging-Files-With-Windows-10/>`_
* `https://github.com/Dijji/FileMeta/wiki/XP,-Vista-and-File-Metadata <https://github.com/Dijji/FileMeta/wiki/XP,-Vista-and-File-Metadata>`_

The FileTags library and source code are free to use for anyone, and anyone can contribute to the FileTags Library project.

.. _Download_Overview:

Download
--------

The FileTags Library is available to download from the github repository at `github.com/mrfearless/FileTags-Library <https://github.com/mrfearless/FileTags-Library>`_


.. _Features_Overview:

Features
--------

* Read file tags (Ansi and Wide/Unicode)
* Write file tags (Ansi and Wide/Unicode)
* Clear file tags (Ansi and Wide/Unicode)


.. _Installation_Overview:

Installation
------------

See the :ref:`Installation & Setup<Installation & Setup>` section for more details.


.. _Contribute_Overview:

Contribute
----------

If you wish to contribute, please follow the :ref:`Contributing<Contributing>` guide for details on how to add or edit, the FileTags Library source or documentation.


.. _FAQ_Overview:

Frequently Asked Questions
--------------------------

Please visit the :ref:`Frequently Asked Questions<FAQ>` section for details.

