// hbmk2 T007 xmlparse.c hxmldoc.prg myxml.prg -b hbmzip.hbc
#include "fileio.ch"

#include "hbclass.ch"
#include "hxml.ch"

Function Main()
LOCAL oDoc, oXml, oFont, oStyles, oAuto, oMaster
SetMode( 40, 100 )
cls
altd()
? StrTran( "&teste", "&", "&amp;" )
oDoc := oOOlib("private:factory/scalc")   // create new document
/*
oDoc := oOOlib(cFile)   // open document
*/
wait
RETURN

FUNCTION oOOlib(xData)
LOCAL oDoc
IF !Empty(xData) .AND. HB_ISSTRING(xData)
   ? "OK"
   IF lower(xData) == "private:factory/scalc"
      oDoc := OOffice():NewODS()
   ELSEIF lower(xData) == "private:factory/swriter"
      ? "New swriter"
   ELSE
      oDoc := OpenFile(xData)
   ENDIF
ELSE
   ? "ERRO"
ENDIF
RETURN oDoc

FUNCTION OpenFile(cFile)
? "OpenFile"
RETURN oDoc
*------------------------------------------------------------------------------*

CLASS OOffice
DATA cMimeType
DATA oMeta
DATA oContent
DATA oSettings
DATA oStyles
DATA oManifest
DATA cFileName

METHOD NewODS()
METHOD NewODT()
//METHOD OpenODS(cFileName)
//METHOD OpenODT(cFileName)
//METHOD Save(cFileName)
ENDCLASS

//---------------------------------------------------------------------//

METHOD NewODS() CLASS OOffice
::cMimeType := "application/vnd.oasis.opendocument.spreadsheet"
::oMeta := odsMeta():New()
::oManifest := odsManifest():New()

::oMeta:SetTitle("TESTE")
? ::oMeta:save()
?
?
? ::oManifest:save()
Return Self

//---------------------------------------------------------------------//

METHOD NewODT() CLASS OOffice
::cMimeType := ""
Return Self

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

