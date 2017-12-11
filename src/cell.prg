//https://ask.libreoffice.org/en/question/83622/macro-how-to-get-a-cell-value-set-by-a-formula/
//https://wiki.openoffice.org/wiki/PyUNO_samples#createTable.28.29

#include "hbclass.ch"

CLASS oOOcell
DATA nValue
DATA cValue
DATA cFormula
DATA cMoeda
DATA dDate
DATA cType     INIT "empty"   // string, float, currency, date
DATA aProperty INIT {}

METHOD Load(oCell)
METHOD Value(nValue)
METHOD String(cValue)
METHOD Formula(cValue)
METHOD setPropertyValue(cProperty, xValue)
METHOD getPropertyValue(cProperty)
METHOD delProperty(cProperty)
METHOD NumberFormat() VIRTUAL
METHOD CellBackColor(nColor)
METHOD IsCellBackgroundTransparent(lValue)
ENDCLASS

//---------------------------------------------------------------------//

METHOD Load(oCell) CLASS oOOcell
LOCAL n
if oCell:GetAttribute("office:value-type") == "string"
   if !HB_ISNIL(oCell:GetAttribute("table:formula"))
      ::cFormula := oCell:GetAttribute("table:formula")
      ::nValue := nil
      ::cValue := oCell:aItems[1]:aItems[1]
      ::cType := "string"
   else
      ::cValue := oCell:aItems[1]:aItems[1]  //oCell:Find("text:p"):aItems[1]
      ::nValue := nil
      ::cType := "string"
   endif
elseif oCell:GetAttribute("office:value-type") == "float"
   if !HB_ISNIL(oCell:GetAttribute("table:formula"))
      ::cFormula := oCell:GetAttribute("table:formula")
      ::nValue := oCell:GetAttribute("office:value-type", "N")
      ::cValue := oCell:aItems[1]:aItems[1]
      ::cType := "formula"
   else
      ::nValue := oCell:GetAttribute("office:value-type", "N")
      ::cValue := oCell:aItems[1]:aItems[1]
      ::cType  := "float"
   endif
elseif oCell:GetAttribute("office:value-type") == "currency"
   if !HB_ISNIL(oCell:GetAttribute("table:formula"))
      ::cFormula := oCell:GetAttribute("table:formula")
      ::nValue := oCell:GetAttribute("office:value-type", "N")
      ::cValue := oCell:aItems[1]:aItems[1]
      ::cType := "formula"
   else
      ::nValue := oCell:GetAttribute("office:value-type", "N")
      ::cValue := oCell:aItems[1]:aItems[1]
      ::cMoeda := oCell:GetAttribute("office:currency")
      ::cType  := "currency"
   endif
elseif oCell:GetAttribute("office:value-type") == "date"
   if !HB_ISNIL(oCell:GetAttribute("table:formula"))
      ::cFormula := oCell:GetAttribute("table:formula")
      ::nValue := oCell:GetAttribute("office:value-type", "N")
      ::cValue := oCell:aItems[1]:aItems[1]
      ::cType := "formula"
   else
      ::dDate  := oCell:GetAttribute("office:date-value", "D")
      ::cValue := oCell:aItems[1]:aItems[1]
      ::cType  := "date"
   endif
elseif !empty(oCell:GetAttribute("office:value-type"))
   ? oCell:GetAttribute("office:value-type")
   altd()
   wait
   quit
endif
Return Self

//---------------------------------------------------------------------//

METHOD Value(nValue) CLASS oOOcell
IF !HB_ISNIL(nValue)
   IF !HB_ISNUMERIC( nValue )
      ? "ERRO o valor tem que ser numérico"
      altd()
   ENDIF
   ::nValue   := nValue
   ::cValue   := alltrim(str(nValue))
   ::cFormula := nil
   ::cType    := "float"
ENDIF
Return ::nValue

//---------------------------------------------------------------------//

METHOD String(cValue) CLASS oOOcell
IF !HB_ISNIL(cValue)
   IF !HB_ISSTRING( cValue ) .and. !HB_ISNUMERIC( cValue )
      ? "ERRO o valor tem que ser texto ou numérico"
      altd()
   ENDIF
   IF HB_ISNUMERIC( cValue )
      cValue := alltrim(str(cValue))
   ENDIF
   ::nValue   := nil
   ::cValue  := cValue
   ::cFormula := nil
   ::cType    := "string"
ENDIF
Return ::cValue

//---------------------------------------------------------------------//

METHOD Formula(cFormula) CLASS oOOcell
IF !HB_ISNIL(cFormula)
   IF !HB_ISSTRING( cFormula )
      ? "ERRO o valor tem que ser texto"
      altd()
   ENDIF
   ::nValue   := nil
   ::cValue  := nil
   ::cFormula := cFormula
   ::cType    := "formula"
ENDIF
Return ::cFormula

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

