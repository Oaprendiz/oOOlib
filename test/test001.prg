/*
#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"
*/
Function Main()
LOCAL oDoc, oSheet
LOCAL cFile := HB_PathNormalize(hb_DirSepToOS(HB_DirBase() + "../000/novo.ods" ))

SetMode( 40, 100 )
cls
altd()
? cFile
? ConvertToURL(cFile)
? ConvertFromUrl(ConvertToURL(cFile))

//hb_StarDesktop():loadComponentFromURL(cURL, cFrame, nSearchFlags, oProperties)

//oDoc := oOOlib():Load(cFile)
oSheet := oDoc:getSheets:getByName("Folha1")
oSheet := oDoc:getSheets:getByIndex(2)
//oSheet:getCellByPosition( 0, nLine ):setValue( oResultado:Fields( 0 ):Value)
//oSheet:getCellByPosition( 1, nLine ):setString( cDoc )
//ATail(aItems):new()
oDoc:Close()
wait
RETURN

//---------------------------------------------------------------------//

//---------------------------------------------------------------------//

/*
MsgBox ConvertToUrl("C:\doc\test.odt") 
  ' supplies file:///C:/doc/test.odt
MsgBox ConvertFromUrl("file:///C:/doc/test.odt")    
  '  supplies (under Windows) c:\doc\test.odt
*/

