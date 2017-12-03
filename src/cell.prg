//https://ask.libreoffice.org/en/question/83622/macro-how-to-get-a-cell-value-set-by-a-formula/
//https://wiki.openoffice.org/wiki/PyUNO_samples#createTable.28.29

#include "hbclass.ch"

CLASS oOOcell
DATA nValue
DATA cString
DATA cFormula
DATA cType     INIT "empty"
DATA aProperty INIT {}

METHOD Value(nValue) SETGET
METHOD String(cValue) SETGET
METHOD Formula(cValue) SETGET
METHOD setPropertyValue(cProperty, xValue)
METHOD getPropertyValue(cProperty)
METHOD delProperty(cProperty)
METHOD NumberFormat() VIRTUAL
METHOD CellBackColor(nColor)
METHOD IsCellBackgroundTransparent(lValue)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Value(nValue) CLASS oOOcell
IF !HB_ISNUMERIC( nValue )
   ? "ERRO o valor tem que ser numérico"
   altd()
ENDIF
::nValue   := nValue
::cString  := alltrim(str(nValue))
::cFormula := nil
::cType    := "float"
Return Self

//---------------------------------------------------------------------//

METHOD String(cValue) CLASS oOOcell
IF !HB_ISSTRING( cValue ) .and. !HB_ISNUMERIC( nValue )
   ? "ERRO o valor tem que ser texto ou numérico"
   altd()
ENDIF
IF HB_ISNUMERIC( nValue )
   cValue := alltrim(str(nValue))
ENDIF
::nValue   := nil
::cString  := cValue
::cFormula := nil
::cType    := "string"
Return cValue

//---------------------------------------------------------------------//

METHOD Formula(cValue) CLASS oOOcell
IF !HB_ISSTRING( cValue )
   ? "ERRO o valor tem que ser texto"
   altd()
ENDIF
::nValue   := nil    //corregir
::cString  := nil    //corregir
::cFormula := cValue
::cType    := "formula"
Return cValue

//---------------------------------------------------------------------//

METHOD getPropertyValue( cName ) CLASS oOOcell
Local i := Ascan( ::aProperty,{|a|a[1]==cName} )
IF i != 0
   Return ::aProperty[ i,2 ]
ENDIF
Return Nil

//---------------------------------------------------------------------//

METHOD setPropertyValue( cName,xValue ) CLASS oOOcell
Local i := Ascan( ::aProperty,{|a|a[1]==cName} )
IF i == 0
   Aadd( ::aProperty,{ cName,xValue } )
ELSE
   ::aProperty[ i,2 ] := xValue
ENDIF
Return xValue

//---------------------------------------------------------------------//

METHOD delProperty( cName ) CLASS oOOcell
Local i := Ascan( ::aProperty,{|a|a[1]==cName} )
IF i != 0
   Adel( ::aProperty, i )
   Asize( ::aProperty, Len( ::aProperty ) - 1 )
   Return .T.
ENDIF
Return .F.

//---------------------------------------------------------------------//

METHOD CellBackColor(nColor) CLASS oOOcell
IF !HB_ISNUMERIC( nColor )
   ? "Erro: CellBackColor"
   altd()
ENDIF
Return ::setPropertyValue("CellBackColor", nColor)

//---------------------------------------------------------------------//

METHOD IsCellBackgroundTransparent(lValue) CLASS oOOcell
IF !HB_ISLOGICAL( lValue )
   ? "Erro: IsCellBackgroundTransparent"
   altd()
ENDIF
Return ::setPropertyValue("IsCellBackgroundTransparent", nColor)

//---------------------------------------------------------------------//

//#define HB_CLS_NOTOBJECT

CLASS ShadowFormat
DATA Location
DATA ShadowWidth
DATA IsTransparent
DATA Color
ENDCLASS

//---------------------------------------------------------------------//

