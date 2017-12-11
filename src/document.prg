#include "hbclass.ch"

CLASS oOOdocument
DATA HasLocation
DATA GetLocation
DATA IsReadOnly
DATA cFile

DATA Cargo
DATA oSheet
DATA aSheets INIT {}
DATA nCorrent
METHOD Load(cPath)
METHOD Store() VIRTUAL
METHOD StoreAsURL(cURL, oProperties) VIRTUAL       //save as...
METHOD StoreToURL(cURL, oProperties) VIRTUAL
METHOD Print() VIRTUAL
METHOD StyleFamilies() VIRTUAL
METHOD Sheets(nPos) VIRTUAL
METHOD GetSheets()
METHOD getByName(cName)
METHOD CreateInstance() VIRTUAL        //"com.sun.star.sheet.Spreadsheet"
ENDCLASS

//---------------------------------------------------------------------//

METHOD Load(cPath, cFile) CLASS oOOdocument
LOCAL n
::cFile := cFile
::Cargo := HXMLDoc():Read(cPath + "\content.xml")
::oSheet := ::Cargo:aItems[1]:Find("office:body"):Find("office:spreadsheet")
::nCorrent := 1
FOR n := 1 TO len(::oSheet:aItems)
   if ::oSheet:aItems[n]:Title == "table:table"
   //altd()
      aadd(::aSheets, oOOsheet():Load(::oSheet:aItems[n], Self))
   endif
NEXT
Return Self

//---------------------------------------------------------------------//

METHOD getSheets(nSheet) CLASS oOOdocument
if HB_ISNIL(nSheet) .and. len(::aSheets) > 0
   Return ::aSheets[::nCorrent]
elseif HB_ISNUMERIC(nSheet)
   nSheet++
   if nSheet < 1 .or. nSheet > len(::aSheets)
      ? "ERRO - Index"
      altd()
   endif
   Return ::aSheets[nSheet]
else
   ? "ERRO o valor tem que ser numérico ou NIL"
   altd()
endif
Return "ERRO - GetSheets"

//---------------------------------------------------------------------//

METHOD getByName(cName) CLASS oOOdocument
Local n
if HB_ISSTRING(cName)
   n := Ascan( ::aSheets,{|o|o:cName == cName} )
   IF n != 0
      Return ::aSheets[n]
   ENDIF
endif
Return "ERRO - getByName"

//---------------------------------------------------------------------//
