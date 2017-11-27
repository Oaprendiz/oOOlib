#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"

CLASS oOOManifest
DATA Cargo

METHOD Read(cPath)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Read(cPath) CLASS oOOManifest
LOCAL oNode
::Cargo := HXMLDoc():Read(cPath + "\META-INF\manifest.xml")

Return Self

//---------------------------------------------------------------------//
