#include "hbclass.ch"

CLASS oOOsheet
DATA Cargo INIT {}
DATA nSheets

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

METHOD Load(oSheets) CLASS oOOsheet
LOCAL n
for n := 1 to len(oSheets:aItems)
   if oSheets:aItems[n]:Title == "table:table"
      aadd(::Cargo, oSheets:aItems[n])
   endif
next n
Return Self

//---------------------------------------------------------------------//

METHOD Name(cName) CLASS oOOsheet
LOCAL n
for n := 1 to len(::Cargo)
   if ::Cargo[n]:GetAttribute("table:name") == cName
      ? "ERRO o nome j� existe " + cName
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
