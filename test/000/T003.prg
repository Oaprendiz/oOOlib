// hbmk2 T003 xmlparse.c hxmldoc.prg myxml.prg -b hbmzip.hbc

#include "fileio.ch"

Function Main()
LOCAL oDoc, oXml, oFont, oStyles, oAuto, oMaster
SetMode( 40, 100 )
cls
altd()

//ODT_UnZip( "Novo.ods" , , , "x1/")
//oXml := HXMLDoc():Read( "x1/" + "styles.xml" )
oXml := XMLDoc():Read( "x1/" + "styles.xml" )
if len(oXml:aAttr) == 0 .and. len(oXml:aItems) == 0
   ? "ERRO!!!!"
endif
altd()
oFont   := oXml:aItems[1]:Find("office:font-face-decls")
oStyles := oXml:aItems[1]:Find("office:styles")
oAuto   := oXml:aItems[1]:Find("office:automatic-styles")
oMaster := oXml:aItems[1]:Find("office:master-styles")

altd()
oXml:Save("teste.xml")
wait
? oXml
RETURN


*------------------------------------------------------------------------------*
function ODT_UnZip ( cFileName , bBlock , cPassword, cPath )
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

RETURN
