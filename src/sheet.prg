#include "hbclass.ch"

CLASS oOOsheet
DATA aRow INIT {}
DATA cName
DATA cStyle

METHOD Load(oSheets)
METHOD getByName(cName)
METHOD getByIndex(nPos)
METHOD Name(cName) SETGET
METHOD getCellByPosition(nRow, nCol) Virtual
METHOD getCellRangeByName(cName) Virtual
METHOD insertCells(cell, nMod) Virtual
METHOD removeRange(cell, nMod) Virtual
METHOD moveRange(cell, cell) Virtual
METHOD Columns(nPos) Virtual
METHOD Rows(nPos) Virtual
ENDCLASS

//---------------------------------------------------------------------//

METHOD Load(oSheet) CLASS oOOsheet
LOCAL n, x, nRow
::cName  := oSheet:GetAttribute( "table:name" )
::cStyle := oSheet:GetAttribute( "table:style-name" )
for n := 1 to len(oSheet:aItems)
   if oSheet:aItems[n]:Title == "table:table-row"
      if !(nRow := oSheet:aItems[n]:GetAttribute("table:number-rows-repeated", "N")) == NIL
         for x := 1 to nRow
            aadd(::aRow, oOOrow():Load(oSheet:aItems[n]))
         next
      else
         aadd(::aRow, oOOrow():Load(oSheet:aItems[n]))
      endif
   endif
next n
Return Self

//---------------------------------------------------------------------//

METHOD Name(cName) CLASS oOOsheet
LOCAL n
for n := 1 to len(::Cargo)
   if ::Cargo[n]:GetAttribute("table:name") == cName
      ? "ERRO o nome já existe " + cName
   endif
next n
Return cName

//---------------------------------------------------------------------//

METHOD getByName(cName) CLASS oOOsheet
LOCAL n
for n := 1 to len(::Cargo)
   if ::Cargo[n]:GetAttribute("table:name") == cName
      RETURN ::Cargo[n]
   endif
next n
Return nil

//---------------------------------------------------------------------//

METHOD getByIndex(nPos) CLASS oOOsheet
LOCAL n
if nPos < 0 .or. nPos >= len(::Cargo)
   ? "ERRO getByIndex (" + alltrim(str(nPos)) + ")"
   wait
   quit
endif
Return ::Cargo[nPos + 1]

//---------------------------------------------------------------------//
