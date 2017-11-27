// hbmk2 T005 xmlparse.c hxmldoc.prg myxml.prg -b hbmzip.hbc

#include "fileio.ch"

#include "hbclass.ch"
#include "hxml.ch"

Function Main()
LOCAL oDoc, oXml, oFont, oStyles, oAuto, oMaster
SetMode( 40, 100 )
cls
altd()
? hb_DateTime()
? hb_TToC( hb_DateTime(), "YYYY-MM-DD", "HH:MM:SSSSSSS" )
? hb_TToC( hb_DateTime(), "YYYY-MM-DD",  )
?
? hb_StrFormat( "%1$s UTC%2$s%3$02d%4$02d %5$s", ;
      hb_TToC( hb_DateTime(), "YYYY-MM-DD", "HH:MM:SS" ), ;
      iif( hb_UTCOffset() < 0, "-", "+" ), ;
      Int( hb_UTCOffset() / 3600 ), ;
      Int( ( ( hb_UTCOffset() / 3600 ) - Int( hb_UTCOffset() / 3600 ) ) * 60 ), ;
      "Jorge Nunes" ) + hb_eol()
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
Return Self

//---------------------------------------------------------------------//

METHOD NewODT() CLASS OOffice
::cMimeType := ""
Return Self

//---------------------------------------------------------------------//

CLASS odsMeta
DATA oMeta
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

METHOD New() CLASS odsMeta
LOCAL oDoc, oMeta
oDoc := xmlDoc():New("UTF-8")
::oMeta := oDoc:Add( ElementNode():New("office:document-meta") )
::oMeta:setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
::oMeta:setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
::oMeta:setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
::oMeta:setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
::oMeta:setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
::oMeta:setAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")
::oMeta:setAttribute("office:version", "1.2")

oMeta := ::oMeta:Add( ElementNode():New('office:meta') )
::oCreationDate := oMeta:Add( ElementNode():New("meta:creation-date", "2016-11-17T11:35:21.265000000") )
::oDate         := oMeta:Add( ElementNode():New("dc:Date") )
::oGenerator    := oMeta:Add( ElementNode():New("meta:generator", "oOOlib-0.0") )
::oCreator      := oMeta:Add( ElementNode():New("dc:creator") )
/*
::oDescription  := oMeta:Add( ElementNode():New("dc:description") )
::oSubject      := oMeta:Add( ElementNode():New("dc:subject") )
::oTitle        := oMeta:Add( ElementNode():New("dc:title") )
::oInitCreator  := oMeta:Add( ElementNode():New("meta:initial-creator") )
*/
? DToC( Date() ) + "T" + Time()
? oDoc:Save()
Return Self

//---------------------------------------------------------------------//

CLASS ElementNode INHERIT HXMLNode
METHOD New(cElement, cData)
ENDCLASS
//---------------------------------------------------------------------//

METHOD New(cElement, cData) CLASS ElementNode
::SUPER:New(cElement)
IF HB_ISSTRING(cData)
   ::Add(cData)
ENDIF
RETURN Self

//---------------------------------------------------------------------//


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

