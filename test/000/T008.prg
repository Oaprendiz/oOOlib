#require "hbwin"

#include "fileio.ch"
#include "hbclass.ch"
#include "hxml.ch"

Function Main()
LOCAL oDoc, oSheet
LOCAL cFile := HB_PathNormalize(hb_DirSepToOS(HB_DirBase() + "Novo.ods" ))

SetMode( 40, 100 )
cls
altd()

? HB_DirTemp()
oDoc := oOOlib():Open(cFile)
oDoc:Close()
wait
RETURN
/*
STATIC FUNCTION DirUnbuild( cDir )
   LOCAL cDirTemp
   LOCAL tmp

   IF hb_DirExists( cDir )

      cDir := DirDelPathSep( cDir )

      cDirTemp := cDir
      DO WHILE ! Empty( cDirTemp )
         IF hb_DirExists( cDirTemp )
            IF hb_DirDelete( cDirTemp ) != 0
               RETURN .F.
            ENDIF
         ENDIF
         IF ( tmp := RAt( hb_osPathSeparator(), cDirTemp ) ) == 0
            EXIT
         ENDIF
         cDirTemp := Left( cDirTemp, tmp - 1 )
         IF ! Empty( hb_osDriveSeparator() ) .AND. ;
            Right( cDirTemp, 1 ) == hb_osDriveSeparator()
            EXIT
         ENDIF
      ENDDO
   ENDIF

   RETURN .T.
*/
/*   
STARTUPINFO si = { sizeof(STARTUPINFO) };
si.dwFlags = STARTF_USESHOWWINDOW;
si.wShowWindow = SW_HIDE;
PROCESS_INFORMATION pi;
CreateProcess(NULL, "cmd /C rmdir /S /Q c:\\MyFiles\\MyTemporaryFolder", NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);


01	////////////////////////////////////////////////////////////////////////////
02	//
03	//         Autor: Jose Carlos da Rocha                                                                                             
04	//          Data: 07/05/2015
05	//         Email: irochinha@hotmail.com.br
06	//     Linguagem: xBase / Fivewin
07	//   Plataformas: DOS, Windows
08	// Requerimentos: Harbour/xHarbour
09	//
10	/////////////////////////////////////////////////////////////////////////////
11	 
12	#include "FiveWin.ch"
13	 
14	FUNCTION MAIN()
15	 
16	   ? ;
17	   'ShellExecute( "CMD.EXE", "RUNAS", "", "C:\WINDOWS\SYSTEM32", 1 )',,;
18	   ShellExecute( "CMD.EXE", "RUNAS", "", "C:\WINDOWS\SYSTEM32", 1 )
19	 
20	RETURN .T.
21	 
22	#pragma BEGINDUMP
23	        #include <windows.h>
24	        #include <hbapi.h>
25	        // ShellExecute( cFile, cOperation, cParams, cDir, nFlag )
26	        HB_FUNC( SHELLEXECUTE )
27	        {
28	        hb_retnl( (LONG) ShellExecute( GetActiveWindow(),
29	                  ISNIL(2) ? NULL : (LPCSTR) hb_parc(2),
30	                  (LPCSTR) hb_parc(1),
31	                  ISNIL(3) ? NULL : (LPCSTR) hb_parc(3),
32	                  ISNIL(4) ? "C:\\" : (LPCSTR) hb_parc(4),
33	                  ISNIL(5) ? 1 : hb_parni(5) ) ) ;
34	        }
35	#pragma ENDDUMP
36	 


REQUEST HB_GT_WIN_DEFAULT
ShellExecute (0, "open", "AcroRd32.exe" , "Armado Cables Fiscales.pdf" )


FUNCTION Abre_arquivo( cHelpFile )
   LOCAL nRet, cPath, cFileName, cFileExt
   HB_FNameSplit( cHelpFile, @cPath, @cFileName, @cFileExt )
   nRet := _OpenHelpFile( cPath, cHelpFile )
RETURN nRet

#pragma BEGINDUMP

   #pragma comment( lib, "shell32.lib" )
   #include "hbapi.h"
   #include 
   HB_FUNC( _OPENHELPFILE )
   {
     HINSTANCE hInst;
     LPCTSTR lpPath = (LPTSTR) hb_parc( 1 );
     LPCTSTR lpHelpFile = (LPTSTR) hb_parc( 2 );
     hInst = ShellExecute( 0, "open", lpHelpFile, 0, lpPath, SW_SHOW );
     hb_retnl( (LONG) hInst );
     return;
   }
#pragma ENDDUMP




****************************************************** 
function MYRUN( cComando ) 
****************************************************** 

local oShell, RET 

oShell := CreateObject( "WScript.Shell" ) 
RET := oShell:Run( "%comspec% /c " + cComando, 0, .T. ) 
oShell := NIL 

return iif( RET = 0, .T., .F. ) 

*/

*------------------------------------------------------------------------------*

CLASS oOOlib
DATA cFileName
DATA cMimeType
DATA oMeta
DATA oContent
DATA oSettings
DATA oStyles
DATA oManifest
DATA cDirTemp

METHOD Open(cFile)
METHOD OpenZip(cFile, bBlock, cPassword, cPath)
METHOD Close()
ENDCLASS

//---------------------------------------------------------------------//

METHOD Open(cFile) CLASS oOOlib
//LOCAL cPath := HB_DirTemp() + HB_FNameName(cFile)
LOCAL cPath := HB_FNameName(cFile) + "_TEMP"
? cPath
::cDirTemp := cPath
::OpenZip(cFile, , , cPath)
ALTD()
//WAPI_SHELLEXECUTE( 0, "open", cFile, , , 1 )

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
//WAPI_SHELLEXECUTE( 0, "", "rd " + ::cDirTemp + " /s /q", , , 1 )
//? WAPI_SHELLEXECUTE( 0, "open", "cmd.exe rd" + ::cDirTemp + " /s /q", , , 1 )
//? hb_processRun( "cmd.exe rd " + ::cDirTemp + " /s /q" )
Return Self

//---------------------------------------------------------------------//


//---------------------------------------------------------------------//

CLASS odsMeta INHERIT HXMLDoc
DATA oCreationDate
DATA oGenerator

METHOD New()
ENDCLASS

//---------------------------------------------------------------------//

METHOD New() CLASS odsMeta
LOCAL oNode
::SUPER:New("UTF-8")
::Add( ElementNode():New("office:document-meta") )
::aItems[1]:setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
::aItems[1]:setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
::aItems[1]:setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
::aItems[1]:setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
::aItems[1]:setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
::aItems[1]:setAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")
::aItems[1]:setAttribute("office:version", "1.2")

oNode := ::Add( ElementNode():New("office:meta") )
::oCreationDate := oNode:Add( ElementNode():New("meta:creation-date", ;
                                                 StrTran( hb_TToC( hb_DateTime(), "YYYY-MM-DD",  ), " ", "T" ) +"000000") )
::oGenerator    := oNode:Add( ElementNode():New("meta:generator", "oOOlib-0.0") )
Return Self

//---------------------------------------------------------------------//

CLASS ElementNode INHERIT HXMLNode
METHOD New(cElement, cData)
METHOD Add(cData)
ENDCLASS

//---------------------------------------------------------------------//

METHOD New(cElement, cData) CLASS ElementNode
::SUPER:New(cElement)
IF HB_ISSTRING(cData)
   ::Add(cData)
ENDIF
RETURN Self

//---------------------------------------------------------------------//

METHOD Add(cData) CLASS ElementNode
IF HB_ISSTRING(cData)
   cData := StrTran( cData, "&", "&amp;" )
   cData := StrTran( cData, '"', "&quot;" )
   cData := StrTran( cData, "'", "&apos;" )
   cData := StrTran( cData, "<", "&lt;" )
   cData := StrTran( cData, ">", "&gt;" )
ENDIF

RETURN ::SUPER:Add(cData)

//---------------------------------------------------------------------//

CLASS odsManifest INHERIT HXMLDoc
METHOD New()
ENDCLASS

//---------------------------------------------------------------------//

METHOD New() CLASS odsManifest
LOCAL oNode
::SUPER:New("UTF-8")
::Add( ElementNode():New("manifest:manifest") )
::aItems[1]:setAttribute("xmlns:manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")
::aItems[1]:setAttribute("manifest:version", "1.2")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "/")
oNode:setAttribute("manifest:version", "1.2")
oNode:setAttribute("manifest:media-type", "application/vnd.oasis.opendocument.spreadsheet")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "Configurations2/accelerator/current.xml")
oNode:setAttribute("manifest:media-type", "")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "Configurations2/")
oNode:setAttribute("manifest:media-type", "application/vnd.sun.xml.ui.configuration")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "styles.xml")
oNode:setAttribute("manifest:media-type", "text/xml")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "manifest.rdf")
oNode:setAttribute("manifest:media-type", "application/rdf+xml")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "meta.xml")
oNode:setAttribute("manifest:media-type", "text/xml")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "Thumbnails/thumbnail.png")
oNode:setAttribute("manifest:media-type", "image/png")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "content.xml")
oNode:setAttribute("manifest:media-type", "text/xml")

oNode := ::aItems[1]:Add( ElementNode():New('manifest:file-entry') )
oNode:setAttribute("manifest:full-path", "settings.xml")
oNode:setAttribute("manifest:media-type", "text/xml")

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
::SUPER:New("UTF-8")
::Add( ElementNode():New("office:document-settings") )
::aItems[1]:setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
::aItems[1]:setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
::aItems[1]:setAttribute("xmlns:config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0")
::aItems[1]:setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
::aItems[1]:setAttribute("office:version", "1.2")

oNode := ::Add( ElementNode():New("office:meta") )
::oCreationDate := oNode:Add( ElementNode():New("meta:creation-date", ;
                                                 StrTran( hb_TToC( hb_DateTime(), "YYYY-MM-DD",  ), " ", "T" ) +"000000") )
::oDate         := oNode:Add( ElementNode():New("dc:Date") )
::oGenerator    := oNode:Add( ElementNode():New("meta:generator", "oOOlib-0.0") )
::oCreator      := oNode:Add( ElementNode():New("dc:creator") )

::oDescription  := oNode:Add( ElementNode():New("dc:description") )
::oSubject      := oNode:Add( ElementNode():New("dc:subject") )
::oTitle        := oNode:Add( ElementNode():New("dc:title") )
::oInitCreator  := oNode:Add( ElementNode():New("meta:initial-creator") )
Return Self

//---------------------------------------------------------------------//


//---------------------------------------------------------------------//


//---------------------------------------------------------------------//


//---------------------------------------------------------------------//


//---------------------------------------------------------------------//


//---------------------------------------------------------------------//



*------------------------------------------------------------------------------*
