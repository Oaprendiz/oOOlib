#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"

Function Main()
LOCAL oDoc, oSheet
LOCAL cFile := HB_PathNormalize(hb_DirSepToOS(HB_DirBase() + "novo.ods" ))

SetMode( 40, 100 )
cls
altd()

oDoc := oOOlib():Open(cFile)
oSheet := oDoc:getSheets:getByName("Folha1")
//oSheet:getCellByPosition( 0, nLine ):setValue( oResultado:Fields( 0 ):Value)
//oSheet:getCellByPosition( 1, nLine ):setString( cDoc )
//ATail(aItems):new()
oDoc:Close()
wait
RETURN

//---------------------------------------------------------------------//
