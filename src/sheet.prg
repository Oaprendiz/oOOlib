#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"

CLASS oOOsheet
DATA Cargo

METHOD New(oSheets)
METHOD getByName(cName)
ENDCLASS

//---------------------------------------------------------------------//

METHOD New(oSheets) CLASS oOOsheet
::Cargo := oSheets
Return Self

//---------------------------------------------------------------------//

METHOD getByName(cName) CLASS oOOsheet
LOCAL n
for n := 1 to len(::Cargo:aItems)
   if ::Cargo:aItems[n]:GetAttribute("table:name") == cName
      RETURN ::Cargo:aItems[n]
   endif
next n
Return nil

//---------------------------------------------------------------------//
