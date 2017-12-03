#include "hbclass.ch"

CLASS oOOdocument
DATA HasLocation
DATA GetLocation
DATA IsReadOnly

DATA Cargo
DATA oBody

METHOD Load(cPath)
METHOD Store() VIRTUAL
METHOD StoreAsURL(cURL, oProperties) VIRTUAL       //save as...
METHOD StoreToURL(cURL, oProperties) VIRTUAL
METHOD Print() VIRTUAL
METHOD StyleFamilies() VIRTUAL
METHOD Sheets(nPos) VIRTUAL
METHOD GetSheets() VIRTUAL
METHOD CreateInstance() VIRTUAL        //"com.sun.star.sheet.Spreadsheet"
ENDCLASS

//---------------------------------------------------------------------//

METHOD Load(cPath) CLASS oOOdocument
::Cargo := HXMLDoc():Read(cPath + "\content.xml")
::oBody := ::Cargo:aItems[1]:Find("office:body")
Return Self

//---------------------------------------------------------------------//

METHOD getSheets() CLASS oOOdocument
Return ::oBody:Find("office:spreadsheet")

//---------------------------------------------------------------------//

METHOD getByName(cName) CLASS oOOdocument
Return ::oBody:Find("office:spreadsheet")

//---------------------------------------------------------------------//
