#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"

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
