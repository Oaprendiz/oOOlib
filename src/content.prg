#include "hbclass.ch"

CLASS oOOcontent
DATA Cargo
DATA oBody

METHOD Load(cPath)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Load(cPath) CLASS oOOcontent
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
