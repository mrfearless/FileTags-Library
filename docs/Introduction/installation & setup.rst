.. _Installation & Setup:

====================
Installation & Setup
====================


.. _MASM for 32bit assembler:

MASM for 32bit assembler
------------------------

- Download and install the `MASM32 <http://www.masm32.com>`_ package.

- Download the 32bit version of the FileTags Library: `FileTags-x86.zip <https://github.com/mrfearless/FileTags-Library/blob/master/releases/FileTags-x86.zip?raw=true>`_

- Copy the ``FileTags.inc`` file to ``X:\MASM32\Include`` folder overwriting any previous versions.

- Copy the ``FileTags.lib`` file to ``X:\MASM32\Lib`` folder overwriting any previous versions.

.. note:: ``X`` is the drive letter where the `MASM32 <http://www.masm32.com>`_ package has been installed to.

**Adding FileTags Library to your MASM project**

You are now ready to begin using the FileTags library in your Masm projects. Simply add the following lines to your project:

::

   include FileTags.inc
   includelib FileTags.lib


.. _UASM for 64bit assembler:

UASM for 64bit assembler
------------------------

- Download and install the `UASM <http://www.terraspace.co.uk/uasm.html>`_ assembler. Ideally you will have a setup that mimics the MASM32 setup, where you create manually folders for ``bin``, ``include`` and ``lib``

- Download the 64bit version of the FileTags Library: `FileTags-x64.zip <https://github.com/mrfearless/FileTags-Library/blob/master/releases/FileTags-x64.zip?raw=true>`_

- Copy the ``FileTags.inc`` file to ``X:\UASM\Include`` folder overwriting any previous versions.

- Copy the ``FileTags.lib`` file to ``X:\UASM\Lib\x64`` folder overwriting any previous versions.

.. note:: ``X`` is the drive letter where the `UASM <http://www.terraspace.co.uk/uasm.html>`_ package has been installed to.


**Adding FileTags Library to your UASM project**

You are now ready to begin using the FileTags library in your Uasm projects. Simply add the following lines to your project:

::

   include FileTags.inc
   includelib FileTags.lib



.. note:: See the following for more details on setting up UASM to work with RadASM and other details that may be useful in creating a development environment that mimics the MASM32 SDK: `UASM-with-RadASM <https://github.com/mrfearless/UASM-with-RadASM>`_, `UASM-SDK <https://github.com/mrfearless/UASM-SDK>`_

.. tip:: `UASM <http://www.terraspace.co.uk/uasm.html>`_ can be used as a x86 32 bit assembler as well.