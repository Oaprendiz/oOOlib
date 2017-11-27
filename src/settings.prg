#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"

CLASS odsSettings INHERIT HXMLDoc
DATA oGenerator
DATA oDescription
DATA oSubject
DATA oTitle
DATA oInitCreator
DATA oCreationDate
DATA oCreator
DATA oDate

METHOD New()
ENDCLASS

//---------------------------------------------------------------------//

METHOD New() CLASS odsSettings
LOCAL oNode
Return Self

//---------------------------------------------------------------------//
