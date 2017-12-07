#include "fileio.ch"
#include "hbclass.ch"

CLASS hb_StarDesktop
DATA aItems  INIT {}

METHOD LoadComponentFromURL(cURL, cFrame, nSearchFlags, oProperties)
METHOD Add(oDoc)
ENDCLASS

//---------------------------------------------------------------------//

METHOD loadComponentFromURL(cURL, cFrame, nSearchFlags, oProperties) CLASS hb_StarDesktop
LOCAL cFile, cTemp, oDoc
HB_SYMBOL_UNUSED( cFrame )
HB_SYMBOL_UNUSED( nSearchFlags )
HB_SYMBOL_UNUSED( oProperties )
cFile := ConvertFromUrl(cURL)
cTemp := HB_DirTemp() + HB_FNameName(cFile) + "_StarDesktop"
OpenFile(cFile, , , cTemp)
oDoc := oOOdocument():Load(cTemp, cFile)
::Add(oDoc)
Return oDoc

//---------------------------------------------------------------------//

METHOD Add(oDoc)
Aadd( ::aItems, oDoc )
Return oDoc

//---------------------------------------------------------------------//

STATIC FUNCTION OpenFile(cFileName, bBlock, cPassword, cPath)
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
RETURN NIL

//----------------------------------------------------------------------------//

FUNCTION ConvertToURL( cString )
IF !( Left( cString, 2 ) == "\\" )
   cString := StrTran( cString, ":", "|" )
   cString := "///" + cString
ENDIF
cString := StrTran( cString, "\", "/" )
cString := StrTran( cString, " ", "%20" )
RETURN "file:" + cString

//----------------------------------------------------------------------------//

FUNCTION ConvertFromUrl( cString )
IF ( Left( cString, 5 ) == "file:" )
   cString := StrTran( cString, "file:", "" )
   cString := StrTran( cString, "|", ":" )
   cString := StrTran( cString, "///", "" )
ENDIF
cString := HB_DirSepToOS(cString)
cString := StrTran( cString, "%20", " " )
RETURN cString

//----------------------------------------------------------------------------//

Function oSProperty(oSrvManager,cName,uValue)
Local oOOSProperty := oSrvManager:Bridge_GetStruct("com.sun.star.beans.PropertyValue")
oOOSProperty:Name  := cName
oOOSProperty:Value := uValue
return( oOOSProperty )

//----------------------------------------------------------------------------//

//---------------------------------------------------------------------//

Function CloseTemp()
//hb_processRun( "cmd.exe /c rd " + ::cDirTemp + " /s /q" )
local oShell, RET, cComando := "rd " + ::cDirTemp + " /s /q"
oShell := win_oleCreateObject( "WScript.Shell" ) 
RET := oShell:Run( "%comspec% /c " + cComando, 0, .T. ) 
oShell := NIL 
Return iif( RET = 0, .T., .F. ) 

//---------------------------------------------------------------------//


//---------------------------------------------------------------------//

//---------------------------------------------------------------------//
