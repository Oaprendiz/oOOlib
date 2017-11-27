#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"

CLASS oOOmeta
DATA Cargo
DATA oCreationDate
DATA oGenerator

METHOD Read(cPath)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Read(cPath) CLASS oOOmeta
LOCAL oNode
::Cargo := HXMLDoc():Read(cPath + "\meta.xml")
Return Self

//---------------------------------------------------------------------//
