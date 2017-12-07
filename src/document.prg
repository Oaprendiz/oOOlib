#include "hbclass.ch"

CLASS oOOdocument
DATA HasLocation
DATA GetLocation
DATA IsReadOnly
DATA cFile

DATA Cargo
DATA oSheet
DATA aSheets INIT {}
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

METHOD Load(cPath, cFile) CLASS oOOdocument
LOCAL n
::cFile := cFile
::Cargo := HXMLDoc():Read(cPath + "\content.xml")
::oSheet := ::Cargo:aItems[1]:Find("office:body"):Find("office:spreadsheet")
FOR n := 1 TO len(::oSheet:aItems)
   if ::oSheet:aItems[n]:Title == "table:table"
   //altd()
      aadd(::aSheets, oOOsheet():Load(::oSheet:aItems[n]))
   endif
NEXT
Return Self

//---------------------------------------------------------------------//

METHOD getSheets() CLASS oOOdocument
Return ::oBody:Find("office:spreadsheet")

//---------------------------------------------------------------------//

METHOD getByName(cName) CLASS oOOdocument
Return ::oBody:Find("office:spreadsheet")

//---------------------------------------------------------------------//
