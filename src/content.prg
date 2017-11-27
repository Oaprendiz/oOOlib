#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"

CLASS oOOcontent
DATA Cargo
DATA oBody

METHOD Read(cPath)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Read(cPath) CLASS oOOcontent
::Cargo := HXMLDoc():Read(cPath + "\content.xml")
::oBody := ::Cargo:aItems[1]:Find("office:body")
Return Self

//---------------------------------------------------------------------//

METHOD getSheets() CLASS oOOcontent
Return ::oBody:Find("office:spreadsheet")

//---------------------------------------------------------------------//

METHOD getByName(cName) CLASS oOOcontent
Return ::oBody:Find("office:spreadsheet")

//---------------------------------------------------------------------//
