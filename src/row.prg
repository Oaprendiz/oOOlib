#include "hbclass.ch"

CLASS oOOrow
DATA aCell INIT {}

METHOD Load(oRow)
METHOD getByName(cName)
METHOD getByIndex(nPos)
METHOD Name(cName) SETGET
ENDCLASS

//---------------------------------------------------------------------//

METHOD Load(oRow) CLASS oOOrow
LOCAL n, x, nCol
for n := 1 to len(oRow:aItems)
   if oRow:aItems[n]:Title == "table:table-cell"
      if !(nCol := oRow:aItems[n]:GetAttribute("table:number-columns-repeated", "N")) == NIL
         for x := 1 to nCol
            aadd(::aCell, oOOcell())
         next
      else
         aadd(::aCell, oOOcell():Load(oRow:aItems[n]))
      endif
   endif
next n
Return Self

//---------------------------------------------------------------------//

METHOD Name(cName) CLASS oOOrow
LOCAL n
for n := 1 to len(::Cargo)
   if ::Cargo[n]:GetAttribute("table:name") == cName
      ? "ERRO o nome já existe " + cName
   endif
next n
Return cName

//---------------------------------------------------------------------//

METHOD getByName(cName) CLASS oOOrow
LOCAL n
for n := 1 to len(::Cargo)
   if ::Cargo[n]:GetAttribute("table:name") == cName
      RETURN ::Cargo[n]
   endif
next n
Return nil

//---------------------------------------------------------------------//

METHOD getByIndex(nPos) CLASS oOOrow
LOCAL n
if nPos < 0 .or. nPos >= len(::Cargo)
   ? "ERRO getByIndex (" + alltrim(str(nPos)) + ")"
   wait
   quit
endif
Return ::Cargo[nPos + 1]

//---------------------------------------------------------------------//
