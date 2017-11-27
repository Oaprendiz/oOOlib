// hbmk2 T001 -b hbziparc.hbc

Function Main()
LOCAL aExtract := hb_GetFilesInZip( "Novo.ods" )  // extract all files in zip

SetMode( 40, 100 )
cls
altd()
? hb_ps()
?
/*
IF hb_UnzipFile( "Novo.ods",,,, "." + hb_ps() + "odt" + hb_ps(), aExtract )
   ? "File was successfully extracted"
ENDIF
*/

aExtract := hb_GetFilesInZip( "Novo.ods" )  // extract all files in zip
IF hb_UnzipFile( "Novo.ods", {| cFile | QOut( cFile ) },.t.,, "x" + hb_ps(), aExtract )
   ? "File was successfully extracted"
ENDIF
RETURN
