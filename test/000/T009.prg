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

CLASS oOOlib
DATA cFileName
DATA cDirTemp

DATA cMimeType
DATA oMeta
DATA oContent
DATA oSettings
DATA oStyles
DATA oManifest

DATA getSheets

METHOD Open(cFile)
METHOD OpenZip(cFile, bBlock, cPassword, cPath)
METHOD Close()
ENDCLASS

//---------------------------------------------------------------------//

METHOD Open(cFile) CLASS oOOlib
LOCAL cPath := HB_DirTemp() + HB_FNameName(cFile) + "_oOOlibTEMP"
::cFileName := cFile
::cDirTemp  := cPath
::OpenZip(cFile, , , cPath)
::cMimeType := MemoRead(::cDirTemp + "\mimetype")
::oMeta := oOOmeta():Read(cPath)
::oContent := oOOcontent():Read(cPath)
ALTD()
::getSheets := oOOsheet():New(::oContent:oBody:Find("office:spreadsheet"))

Return Self

//---------------------------------------------------------------------//

METHOD OpenZip(cFileName, bBlock, cPassword, cPath) CLASS oOOlib
LOCAL i := 0 , hUnzip , nErr, hHandle
LOCAL cFile, dDate, cTime, nSize, nCompSize, lCrypted, cComment, cStorePath

IF ! Empty( cPath )
   cPath := hb_DirSepAdd( HB_PathNormalize( HB_DirSepToOS( cPath )))
   IF ! hb_DirExists( cPath )
      hb_DirBuild( cPath )
   ENDIF
ENDIF

hUnzip := hb_UnZipOpen( cFileName )
nErr := hb_UnZipFileFirst( hUnzip )

DO WHILE nErr == 0
   HB_UnzipFileInfo( hUnzip, @cFile, @dDate, @cTime,,,, @nSize, @nCompSize, @lCrypted, @cComment )

   cFile := hb_DirSepToOS( cFile )
   IF ! Empty( (cStorePath := hb_FNameDir( cFile )) )
      cStorePath := cPath + hb_DirSepAdd(cStorePath)
      IF ! hb_DirExists( cStorePath )
         hb_DirBuild( cStorePath )
      ENDIF
   ENDIF

   if valtype (bBlock) == 'B'
      Eval ( bBlock , cFile , ++i )
   endif

   IF ( hHandle := FCreate( cPath + cFile ) ) != F_ERROR
      hb_UnzipExtractCurrentFileToHandle( hUnzip, hHandle, iif( lCrypted, cPassword, NIL ) )
      FClose( hHandle )
   ENDIF
   nErr := hb_UnZipFileNext( hUnzip )

ENDDO
hb_UnZipClose( hUnzip )
RETURN Self

//---------------------------------------------------------------------//

METHOD Close() CLASS oOOlib
//hb_processRun( "cmd.exe /c rd " + ::cDirTemp + " /s /q" )
local oShell, RET, cComando := "rd " + ::cDirTemp + " /s /q"
oShell := win_oleCreateObject( "WScript.Shell" ) 
RET := oShell:Run( "%comspec% /c " + cComando, 0, .T. ) 
oShell := NIL 
Return iif( RET = 0, .T., .F. ) 

//---------------------------------------------------------------------//

CLASS oOOmeta
DATA Cargo
DATA oCreationDate
DATA oGenerator

METHOD Read(cPath)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Read(cPath) CLASS oOOmeta
LOCAL oNode
::Cargo := HXMLDoc():Read(cPath + "\meta.xml")
Return Self

//---------------------------------------------------------------------//

CLASS oOOManifest
DATA Cargo

METHOD Read(cPath)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Read(cPath) CLASS oOOManifest
LOCAL oNode
::Cargo := HXMLDoc():Read(cPath + "\META-INF\manifest.xml")

Return Self

//---------------------------------------------------------------------//

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

CLASS oOOcontent
DATA Cargo
DATA oBody

METHOD Read(cPath)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Read(cPath) CLASS oOOcontent
::Cargo := HXMLDoc():Read(cPath + "\content.xml")
::oBody := ::Cargo:aItems[1]:Find("office:body")
Return Self

//---------------------------------------------------------------------//

METHOD getSheets() CLASS oOOcontent
Return ::oBody:Find("office:spreadsheet")

//---------------------------------------------------------------------//

METHOD getByName(cName) CLASS oOOcontent
Return ::oBody:Find("office:spreadsheet")

//---------------------------------------------------------------------//

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


//---------------------------------------------------------------------//


//---------------------------------------------------------------------//


//---------------------------------------------------------------------//



*------------------------------------------------------------------------------*
