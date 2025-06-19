Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R4484

    ' Tipo de dado do AKB
    Private Enum AKBTypeConst
        cAKBTypeNone = -1
        cAKBTypeNumeric = 1
        cAKBTypeDate = 2
        cAKBTypeString = 3
        cAKBTypeLogical = 4

        ' atv 13.453-107.368
        cAKBTypeUndefined = 0
        cAKBTypeQuery = 5
    End Enum

    Public Const AKB_DecimalPoint = ","
    Public Const DecimalSymbol = "."
    Public Const DateIdentifier = "'"
    Public Const StringIdentifier = "'"
    Public Shared Sub Main ()

    End Sub

    Public Shared Function R4484(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P12201 As Double, ByVal P12202 As Double, ByVal P12203 As Double, ByVal P12204 As Object, ByVal P12205 As Object, ByVal P40621 As Object, ByVal P40622 As Object, ByVal P40894 As Object, ByVal P68671 As Object) As DataTable
    ' 
    ' 4484 - Verifica Comissão do Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R4484
        Dim sCurrComponent as String

        Dim QueryC52364 As New Object
        Dim RsC52364 As New Object
        Dim C52364DataType() As Short
        Dim RsC52364_EOF As Boolean
        Dim C52365 As DataTable
        Dim C52365CurrentRow As Integer
        Dim C52365DataType() As Short
        Dim C52366 As Object
        Dim C52366DataType As Short
        Dim C52368DataType() As Short
        Dim C52369 As Object
        Dim C52369DataType As Short
        Dim C52370 As Boolean
        Dim C52370DataType As Short
        Dim QueryC52371 As New Object
        Dim RsC52371 As New Object
        Dim C52371DataType() As Short
        Dim RsC52371_EOF As Boolean
        Dim C52372 As Object
        Dim C52372DataType As Short
        Dim C52373 As Object
        Dim C52373DataType As Short
        Dim C52374 As Object
        Dim C52374DataType As Short
        Dim C52375 As Object
        Dim C52375DataType As Short
        Dim C52376 As Boolean
        Dim C52376DataType As Short
        Dim C52385 As Object
        Dim C52385DataType As Short
        Dim C52386 As Boolean
        Dim C52386DataType As Short
        Dim C52387 As Short
        Dim C52387DataType As Short
        Dim C52387ReturnDataType() As Short

        Dim C52389DataType() As Short
        Dim C52390 As Object
        Dim C52390DataType As Short
        Dim C52392 As Double
        Dim C52392DataType As Short
        Dim C52393 As Object
        Dim C52393DataType As Short
        Dim C52823 As Boolean
        Dim C52823DataType As Short
        Dim C52824 As Boolean
        Dim C52824DataType As Short
        Dim C52833DataType() As Short
        Dim QueryC52838 As New Object
        Dim RsC52838 As New Object
        Dim C52838DataType() As Short
        Dim RsC52838_EOF As Boolean
        Dim C52839 As Object
        Dim C52839DataType As Short
        Dim C52840 As Boolean
        Dim C52840DataType As Short
        Dim C52843 As Object
        Dim C52843DataType As Short
        Dim C52845 As Object
        Dim C52845DataType As Short
        Dim QueryC53562 As New Object
        Dim RsC53562 As New Object
        Dim C53562DataType() As Short
        Dim RsC53562_EOF As Boolean
        Dim C53563 As Boolean
        Dim C53563DataType As Short
        Dim QueryC55367 As New Object
        Dim RsC55367 As New Object
        Dim C55367DataType() As Short
        Dim RsC55367_EOF As Boolean
        Dim C55368 As Object
        Dim C55368DataType As Short
        Dim C55369 As Boolean
        Dim C55369DataType As Short
        Dim C55370 As Object
        Dim C55370DataType As Short
        Dim C55371 As Object
        Dim C55371DataType As Short
        Dim C55373DataType() As Short
        Dim QueryC233199 As New Object
        Dim RsC233199 As New Object
        Dim C233199DataType() As Short
        Dim RsC233199_EOF As Boolean
        Dim C233201 As Boolean
        Dim C233201DataType As Short
        Dim QueryC233213 As New Object
        Dim C233213 As Integer
        Dim C233213DataType As Short
        Dim QueryC233217 As New Object
        Dim RsC233217 As New Object
        Dim C233217DataType() As Short
        Dim RsC233217_EOF As Boolean
        Dim C233218 As Boolean
        Dim C233218DataType As Short
        Dim C233219 As Object
        Dim C233219DataType As Short
        Dim C233220 As Boolean
        Dim C233220DataType As Short
        Dim C233223 As Boolean
        Dim C233223DataType As Short
        Dim C233224 As Object
        Dim C233224DataType As Short
        Dim C233225 As Short
        Dim C233225DataType As Short
        Dim C233225ReturnDataType() As Short

        Dim C233226 As Boolean
        Dim C233226DataType As Short
        Dim C233227 As Object
        Dim C233227DataType As Short
        Dim QueryC233228 As New Object
        Dim C233228 As Integer
        Dim C233228DataType As Short
        Dim C233229DataType() As Short
        Dim QueryC233256 As New Object
        Dim C233256 As Integer
        Dim C233256DataType As Short
        Dim C233257 As Boolean
        Dim C233257DataType As Short
        Dim QueryC233258 As New Object
        Dim C233258 As Integer
        Dim C233258DataType As Short
        Dim QueryC233259 As New Object
        Dim C233259 As Integer
        Dim C233259DataType As Short
        Dim C233260 As Object
        Dim C233260DataType As Short
        Dim C233261DataType() As Short
        Dim C233262 As Boolean
        Dim C233262DataType As Short
        Dim C233263 As Boolean
        Dim C233263DataType As Short
        Dim C233264 As Boolean
        Dim C233264DataType As Short
        Dim C233265 As Boolean
        Dim C233265DataType As Short
        Dim QueryC233266 As New Object
        Dim RsC233266 As New Object
        Dim C233266DataType() As Short
        Dim RsC233266_EOF As Boolean
        Dim C233267 As Object
        Dim C233267DataType As Short
        Dim C233538 As DataTable
        Dim C233538CurrentRow As Integer
        Dim C233538DataType() As Short
        Dim C233540 As Boolean
        Dim C233540DataType As Short
        Dim C233541 As Object
        Dim C233541DataType As Short
        Dim QueryC238895 As New Object
        Dim RsC238895 As New Object
        Dim C238895DataType() As Short
        Dim RsC238895_EOF As Boolean
        Dim C238906 As Boolean
        Dim C238906DataType As Short
        Dim QueryC238908 As New Object
        Dim RsC238908 As New Object
        Dim C238908DataType() As Short
        Dim RsC238908_EOF As Boolean
        Dim C238909 As Boolean
        Dim C238909DataType As Short
        Dim C238910 As Object
        Dim C238910DataType As Short
        Dim C238911 As Object
        Dim C238911DataType As Short
        Dim C238912 As Object
        Dim C238912DataType As Short
        Dim C238936 As DataTable
        Dim C238936CurrentRow As Integer
        Dim C238936DataType() As Short
        Dim C239002 As DataTable
        Dim C239002CurrentRow As Integer
        Dim C239002DataType() As Short
        Dim C239075 As Double
        Dim C239075DataType As Short
        Dim C239079 As Object
        Dim C239079DataType As Short
        Dim C239080 As Object
        Dim C239080DataType As Short
        Dim QueryC239081 As New Object
        Dim RsC239081 As New Object
        Dim C239081DataType() As Short
        Dim RsC239081_EOF As Boolean
        Dim C239085 As Double
        Dim C239085DataType As Short
        Dim C239087 As Object
        Dim C239087DataType As Short
        Dim C239088 As Boolean
        Dim C239088DataType As Short
        Dim C239089 As Double
        Dim C239089DataType As Short
        Dim C239090 As Object
        Dim C239090DataType As Short
        Dim C296826 As Object
        Dim C296826DataType As Short
        Dim QueryC402182 As New Object
        Dim RsC402182 As New Object
        Dim C402182DataType() As Short
        Dim RsC402182_EOF As Boolean
        Dim C402183 As Boolean
        Dim C402183DataType As Short
        Dim QueryC402184 As New Object
        Dim C402184 As Integer
        Dim C402184DataType As Short
        Dim C402185DataType() As Short
        Dim QueryC402878 As New Object
        Dim RsC402878 As New Object
        Dim C402878DataType() As Short
        Dim RsC402878_EOF As Boolean
        Dim C402879 As Boolean
        Dim C402879DataType As Short
        Dim QueryC531933 As New Object
        Dim RsC531933 As New Object
        Dim C531933DataType() As Short
        Dim RsC531933_EOF As Boolean
        Dim C531934 As Boolean
        Dim C531934DataType As Short
        Dim QueryC531935 As New Object
        Dim RsC531935 As New Object
        Dim C531935DataType() As Short
        Dim RsC531935_EOF As Boolean
        Dim QueryC531936 As New Object
        Dim RsC531936 As New Object
        Dim C531936DataType() As Short
        Dim RsC531936_EOF As Boolean
        Dim C531937 As Double
        Dim C531937DataType As Short
        Dim QueryC531938 As New Object
        Dim RsC531938 As New Object
        Dim C531938DataType() As Short
        Dim RsC531938_EOF As Boolean
        Dim C535150 As Object
        Dim C535150DataType As Short
        Dim QueryC535151 As New Object
        Dim RsC535151 As New Object
        Dim C535151DataType() As Short
        Dim RsC535151_EOF As Boolean
        Dim QueryC535152 As New Object
        Dim RsC535152 As New Object
        Dim C535152DataType() As Short
        Dim RsC535152_EOF As Boolean
        Dim C535153 As Object
        Dim C535153DataType As Short
        Dim C535154 As Object
        Dim C535154DataType As Short
        Dim C535155 As Object
        Dim C535155DataType As Short
        Dim C535156 As Object
        Dim C535156DataType As Short
        Dim QueryC535157 As New Object
        Dim C535157 As Integer
        Dim C535157DataType As Short
        Dim C535158 As Boolean
        Dim C535158DataType As Short
        P40621 = IIf(IsDBNull(P40621), 0, P40621)

        P40894 = IIf(IsDBNull(P40894), 0, P40894)

        ReDim ReturnDatatype(0)

        GoTo Comp_C233224

    Comp_C52364:

        ' SelPolítica
        sCurrComponent = "SelPolítica"
        QueryC52364 = con.CreateCommand()
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "SELECT "
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "WF_POLITICA_COMISSAO.Politica , WF_POLITICA_COMISSAO.Descricao , WF_POLITICA_COMISSAO.Porc_Comissao "
        QueryC52364.CommandText = QueryC52364.CommandText & " " & ""
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "FROM AKBUSER01.WF_POLITICA_COMISSAO , WF_COMISSAO_TIPO_PRODUTO, WF_POL_COMISSAO_REPRES"
        QueryC52364.CommandText = QueryC52364.CommandText & " " & ""
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "WHERE  WF_POLITICA_COMISSAO.Validade_Inicial <= " & _ 
ConvertValue(P12204, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "AND  WF_POLITICA_COMISSAO.Validade_Final +1 > " & _ 
ConvertValue(P12204, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52364.CommandText = QueryC52364.CommandText & " " & ""
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "AND  WF_POLITICA_COMISSAO.Politica = WF_COMISSAO_TIPO_PRODUTO.POLITICA"
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "AND  WF_COMISSAO_TIPO_PRODUTO.TP_PRODUTO = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52364.CommandText = QueryC52364.CommandText & " " & ""
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "AND WF_POLITICA_COMISSAO.Politica = WF_POL_COMISSAO_REPRES.Politica "
        QueryC52364.CommandText = QueryC52364.CommandText & " " & "AND WF_POL_COMISSAO_REPRES.Cod_Repres = " & _ 
ConvertValue(P40621, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52364.CommandText = QueryC52364.CommandText & " " & ""
        QueryC52364.CommandText = QueryC52364.CommandText & " " & ""
        QueryC52364.Transaction = txn
        RsC52364 = QueryC52364.ExecuteReader()
        Dim iC52364 As Short
        ReDim C52364DataType(RsC52364.FieldCount)
        For iC52364 = 0 to RsC52364.FieldCount - 1
            Select Case RsC52364.GetDataTypeName(iC52364).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C52364DataType(iC52364 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C52364DataType(iC52364 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C52364DataType(iC52364 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC52364
        RsC52364_EOF = Not RsC52364.Read()

        GoTo Comp_C52369

    Comp_C52365:

        ' Desconto
        sCurrComponent = "Desconto"
        'C52365 = clsRuleDynamicallyCompiled_R4483.R4483(con, txn, IIf(Not IsDBNull(P12201), P12201, System.DBNull.Value), IIf(Not IsDBNull(P12202), P12202, System.DBNull.Value), fn_ConvertValueCompiled( "1", 1, AKB_DecimalPoint), IIf(Not IsDBNull(C52366), C52366, System.DBNull.Value))
        C52365CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C52365) Then
          iColumns = C52365.Columns.Count
        End If
        ReDim C52365DataType(iColumns)
        For iCol = 1 To iColumns
            'C52365DataType(iCol) = clsRuleDynamicallyCompiled_R4483.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C52371

    Comp_C52366:

        ' Política
        sCurrComponent = "Política"
        C52366DataType = 0
        C52366 = RsC52364(0)
        C52366DataType = C52364Datatype(1)
        If C52366DataType = AKBTypeConst.cAKBTypeString Then
          C52366 = IIF(IsDBNull(C52366), C52366, Strings.RTrim(C52366))
        End If 

        GoTo Comp_C52372

    Comp_C52368:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C52368 As DataTable = New DataTable()
        tb_C52368.Columns.Add("vRet" & "")
        Dim row_C52368 As DataRow = tb_C52368.NewRow()
        row_C52368(0) = C233224
        tb_C52368.Rows.Add(row_C52368)
        R4484 = tb_C52368
        ReDim C52368DataType(1)
        C52368DataType(1) = C233224DataType
        ReturnDataType = C52368DataType

        GoTo Exit_R4484

    Comp_C52369:

        ' Eof
        sCurrComponent = "Eof"
        C52369DataType = 4
        C52369 = RsC52364_EOF
        GoTo Comp_C52370

    Comp_C52370:

        ' Eof=1
        sCurrComponent = "Eof=1"
        C52370 = (fn_ConvertValueCompiled(C52369, C52369DataType, AKB_DecimalPoint, False) = 1)
        C52370DataType = AKBTypeConst.cAKBTypeLogical
        If C52370 Then
            GoTo Comp_C233218
        Else
            GoTo Comp_C52366
        End If

    Comp_C52371:

        ' SelTabela
        sCurrComponent = "SelTabela"
        QueryC52371 = con.CreateCommand()
        QueryC52371.CommandText = QueryC52371.CommandText & " " & "SELECT "
        QueryC52371.CommandText = QueryC52371.CommandText & " " & "NVL( WF_COMISSAO_RED_DESCONTOS.Porc_Reducao , 0) , "
        QueryC52371.CommandText = QueryC52371.CommandText & " " & "NVL( WF_COMISSAO_RED_DESCONTOS.Porc_Comissao , 0)"
        QueryC52371.CommandText = QueryC52371.CommandText & " " & ""
        QueryC52371.CommandText = QueryC52371.CommandText & " " & "FROM  AKBUSER01.WF_COMISSAO_RED_DESCONTOS "
        QueryC52371.CommandText = QueryC52371.CommandText & " " & ""
        QueryC52371.CommandText = QueryC52371.CommandText & " " & "WHERE WF_COMISSAO_RED_DESCONTOS.Politica = " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52371.CommandText = QueryC52371.CommandText & " " & ""
        QueryC52371.CommandText = QueryC52371.CommandText & " " & "AND  WF_COMISSAO_RED_DESCONTOS.Faixa_Inicial <= " & _ 
ConvertValue(C52365.rows(C52365CurrentRow - 1)(0), C52365DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52371.CommandText = QueryC52371.CommandText & " " & "AND  WF_COMISSAO_RED_DESCONTOS.Faixa_Final >= " & _ 
ConvertValue(C52365.rows(C52365CurrentRow - 1)(0), C52365DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52371.CommandText = QueryC52371.CommandText & " " & ""
        QueryC52371.Transaction = txn
        RsC52371 = QueryC52371.ExecuteReader()
        Dim iC52371 As Short
        ReDim C52371DataType(RsC52371.FieldCount)
        For iC52371 = 0 to RsC52371.FieldCount - 1
            Select Case RsC52371.GetDataTypeName(iC52371).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C52371DataType(iC52371 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C52371DataType(iC52371 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C52371DataType(iC52371 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC52371
        RsC52371_EOF = Not RsC52371.Read()

        GoTo Comp_C52375

    Comp_C52372:

        ' %ComissãoMáximo
        sCurrComponent = "%ComissãoMáximo"
        C52372DataType = 0
        C52372DataType = C52364Datatype(3)
        If C52372DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC52364(2)) Then
          C52372 = Strings.RTrim(RsC52364(2))
        Else 
          C52372 = RsC52364(2)
        End If 

        GoTo Comp_C52385

    Comp_C52373:

        ' %ComissãoTabela
        sCurrComponent = "%ComissãoTabela"
        C52373DataType = 0
        C52373DataType = C52371Datatype(2)
        If C52373DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC52371(1)) Then
          C52373 = Strings.RTrim(RsC52371(1))
        Else 
          C52373 = RsC52371(1)
        End If 

        GoTo Comp_C52374

    Comp_C52374:

        ' %ReduçãoTabela
        sCurrComponent = "%ReduçãoTabela"
        C52374DataType = 0
        C52374 = RsC52371(0)
        C52374DataType = C52371Datatype(1)
        If C52374DataType = AKBTypeConst.cAKBTypeString Then
          C52374 = IIF(IsDBNull(C52374), C52374, Strings.RTrim(C52374))
        End If 

        GoTo Comp_C52386

    Comp_C52375:

        ' EofTabela
        sCurrComponent = "EofTabela"
        C52375DataType = 4
        C52375 = RsC52371_EOF
        GoTo Comp_C52376

    Comp_C52376:

        ' EofTabela=1
        sCurrComponent = "EofTabela=1"
        C52376 = (fn_ConvertValueCompiled(C52375, C52375DataType, AKB_DecimalPoint, False) = 1)
        C52376DataType = AKBTypeConst.cAKBTypeLogical
        If C52376 Then
            GoTo Comp_C52823
        Else
            GoTo Comp_C52373
        End If

    Comp_C52385:

        ' vComissãoCalculada
        sCurrComponent = "vComissãoCalculada"
        C52385 = C52372 
        C52385DataType = 1
        GoTo Comp_C238910

    Comp_C52386:

        ' %Red=0
        sCurrComponent = "%Red=0"
        C52386 = (fn_ConvertValueCompiled(C52374, C52374DataType, AKB_DecimalPoint, False) = 0)
        C52386DataType = AKBTypeConst.cAKBTypeLogical
        If C52386 Then
            GoTo Comp_C52390
        Else
            GoTo Comp_C52392
        End If

    Comp_C52387:

        ' Confirma
        sCurrComponent = "Confirma"
        C52387 = 6
        C52387DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C52824

    Comp_C52389:

        ' RetComissaoDif
        sCurrComponent = "RetComissaoDif"
        Dim tb_C52389 As DataTable = New DataTable()
        tb_C52389.Columns.Add("vRet" & "")
        Dim row_C52389 As DataRow = tb_C52389.NewRow()
        row_C52389(0) = C233224
        tb_C52389.Rows.Add(row_C52389)
        R4484 = tb_C52389
        ReDim C52389DataType(1)
        C52389DataType(1) = C233224DataType
        ReturnDataType = C52389DataType

        GoTo Exit_R4484

    Comp_C52390:

        ' Att1
        sCurrComponent = "Att1"
        C52390DataType = 4
        C52385 = fn_ConvertValueCompiled(C52373 , 1, AKB_DecimalPoint)
        C52390 = True
        GoTo Comp_C52823

    Comp_C52392:

        ' CálculoRedução1
        sCurrComponent = "CálculoRedução1"
        C52392 = fn_ConvertValueCompiled(C52373, C52373DataType, AKB_DecimalPoint, True) - ( fn_ConvertValueCompiled(C52373, C52373DataType, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(C52374, C52374DataType, AKB_DecimalPoint, True) /100)
        C52392DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C52393

    Comp_C52393:

        ' Att2
        sCurrComponent = "Att2"
        C52393DataType = 4
        C52385 = fn_ConvertValueCompiled(C52392 , 1, AKB_DecimalPoint)
        C52393 = True
        GoTo Comp_C52823

    Comp_C52823:

        ' Igual
        sCurrComponent = "Igual"
        C52823 = (fn_ConvertValueCompiled(P12205, 1, AKB_DecimalPoint, False)  =  fn_ConvertValueCompiled(C52385, C52385DataType, AKB_DecimalPoint, False))
        C52823DataType = AKBTypeConst.cAKBTypeLogical
        If C52823 Then
            GoTo Comp_C233256
        Else
            GoTo Comp_C233264
        End If

    Comp_C52824:

        ' Confirmado
        sCurrComponent = "Confirmado"
        C52824 = (fn_ConvertValueCompiled(C52387, C52387DataType, AKB_DecimalPoint, False) = 6)
        C52824DataType = AKBTypeConst.cAKBTypeLogical
        If C52824 Then
            GoTo Comp_C233258
        Else
            GoTo Comp_C233260
        End If

    Comp_C52833:

        ' RetComissaoCalc
        sCurrComponent = "RetComissaoCalc"
        Dim tb_C52833 As DataTable = New DataTable()
        tb_C52833.Columns.Add("vRet" & "")
        Dim row_C52833 As DataRow = tb_C52833.NewRow()
        row_C52833(0) = C233224
        tb_C52833.Rows.Add(row_C52833)
        R4484 = tb_C52833
        ReDim C52833DataType(1)
        C52833DataType(1) = C233224DataType
        ReturnDataType = C52833DataType

        GoTo Exit_R4484

    Comp_C52838:

        ' SelRepres
        sCurrComponent = "SelRepres"
        QueryC52838 = con.CreateCommand()
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "SELECT  RF.PORC_COMISSAO"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "FROM WF_POL_COMISSAO_REPRES R , WF_POL_COMIS_REPRES_Faixa RF, WF_PEDIDO P"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "WHERE  R.POLITICA = " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND  P.TP_PRODUTO = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND  P.PEDIDO = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & ""
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND R.POLITICA = RF.POLITICA"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND R.COD_REPRES = RF.COD_REPRES "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND  R.COD_REPRES = P.COD_REPRES"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & ""
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND RF.FAIXA_INICIAL <= " & _ 
ConvertValue(C238910, C238910DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND RF.FAIXA_FINAL >= " & _ 
ConvertValue(C238910, C238910DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & ""
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "Union all"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & ""
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "SELECT  R.PORC_COMISSAO"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "FROM  WF_POL_COMISSAO_REPRES R , WF_PEDIDO P"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "WHERE  R.POLITICA = " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND  P.TP_PRODUTO = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND  P.PEDIDO = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "AND  R.COD_REPRES = P.COD_REPRES"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & ""
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "and not exists ( Select RF.POLITICA"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "                               from WF_POL_COMIS_REPRES_Faixa RF"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "                               where R.POLITICA = RF.POLITICA"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "                               and R.COD_REPRES = RF.COD_REPRES "
        QueryC52838.CommandText = QueryC52838.CommandText & " " & "               )"
        QueryC52838.CommandText = QueryC52838.CommandText & " " & ""
        QueryC52838.Transaction = txn
        RsC52838 = QueryC52838.ExecuteReader()
        Dim iC52838 As Short
        ReDim C52838DataType(RsC52838.FieldCount)
        For iC52838 = 0 to RsC52838.FieldCount - 1
            Select Case RsC52838.GetDataTypeName(iC52838).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C52838DataType(iC52838 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C52838DataType(iC52838 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C52838DataType(iC52838 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC52838
        RsC52838_EOF = Not RsC52838.Read()

        GoTo Comp_C52839

    Comp_C52839:

        ' EofRepres
        sCurrComponent = "EofRepres"
        C52839DataType = 4
        C52839 = RsC52838_EOF
        GoTo Comp_C52840

    Comp_C52840:

        ' EofRepres=1
        sCurrComponent = "EofRepres=1"
        C52840 = (fn_ConvertValueCompiled(C52839, C52839DataType, AKB_DecimalPoint, False) = 1)
        C52840DataType = AKBTypeConst.cAKBTypeLogical
        If C52840 Then
            GoTo Comp_C55367
        Else
            GoTo Comp_C52843
        End If

    Comp_C52843:

        ' %CalculadoRepres
        sCurrComponent = "%CalculadoRepres"
        C52843DataType = 0
        C52843 = RsC52838(0)
        C52843DataType = C52838Datatype(1)
        If C52843DataType = AKBTypeConst.cAKBTypeString Then
          C52843 = IIF(IsDBNull(C52843), C52843, Strings.RTrim(C52843))
        End If 

        GoTo Comp_C52845

    Comp_C52845:

        ' Att3
        sCurrComponent = "Att3"
        C52845DataType = 4
        C52385 = fn_ConvertValueCompiled(C52843 , 1, AKB_DecimalPoint)
        C52845 = True
        GoTo Comp_C52823

    Comp_C53562:

        ' VerCondição
        sCurrComponent = "VerCondição"
        QueryC53562 = con.CreateCommand()
        QueryC53562.CommandText = QueryC53562.CommandText & " " & "SELECT NVL(NAO_GERA_COMISSAO,0)"
        QueryC53562.CommandText = QueryC53562.CommandText & " " & ""
        QueryC53562.CommandText = QueryC53562.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO , AKBUSER01.WF_CONDICAO_PGTO "
        QueryC53562.CommandText = QueryC53562.CommandText & " " & ""
        QueryC53562.CommandText = QueryC53562.CommandText & " " & "WHERE WF_CONDICAO_PGTO.Cod_Pagto = WF_PEDIDO.Cod_Pagto "
        QueryC53562.CommandText = QueryC53562.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC53562.CommandText = QueryC53562.CommandText & " " & "AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC53562.CommandText = QueryC53562.CommandText & " " & ""
        QueryC53562.Transaction = txn
        RsC53562 = QueryC53562.ExecuteReader()
        Dim iC53562 As Short
        ReDim C53562DataType(RsC53562.FieldCount)
        For iC53562 = 0 to RsC53562.FieldCount - 1
            Select Case RsC53562.GetDataTypeName(iC53562).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C53562DataType(iC53562 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C53562DataType(iC53562 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C53562DataType(iC53562 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC53562
        RsC53562_EOF = Not RsC53562.Read()

        GoTo Comp_C53563

    Comp_C53563:

        ' VerCondição=1
        sCurrComponent = "VerCondição=1"
        C53563 = (fn_ConvertValueCompiled(RsC53562(0), C53562DataType(1), AKB_DecimalPoint, False) = 1)
        C53563DataType = AKBTypeConst.cAKBTypeLogical
        If C53563 Then
            GoTo Comp_C52368
        Else
            GoTo Comp_C402878
        End If

    Comp_C55367:

        ' SelCliente
        sCurrComponent = "SelCliente"
        QueryC55367 = con.CreateCommand()
        QueryC55367.CommandText = QueryC55367.CommandText & " " & "SELECT  R.PORC_COMISSAO"
        QueryC55367.CommandText = QueryC55367.CommandText & " " & ""
        QueryC55367.CommandText = QueryC55367.CommandText & " " & "FROM  WF_POL_COMISSAO_CLIENTES R , WF_PEDIDO P"
        QueryC55367.CommandText = QueryC55367.CommandText & " " & ""
        QueryC55367.CommandText = QueryC55367.CommandText & " " & "WHERE  R.POLITICA = " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC55367.CommandText = QueryC55367.CommandText & " " & "AND  P.TP_PRODUTO = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC55367.CommandText = QueryC55367.CommandText & " " & "AND  P.PEDIDO = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC55367.CommandText = QueryC55367.CommandText & " " & "AND  R.COD_CLIENTE = P.COD_CLIENTE"
        QueryC55367.CommandText = QueryC55367.CommandText & " " & ""
        QueryC55367.Transaction = txn
        RsC55367 = QueryC55367.ExecuteReader()
        Dim iC55367 As Short
        ReDim C55367DataType(RsC55367.FieldCount)
        For iC55367 = 0 to RsC55367.FieldCount - 1
            Select Case RsC55367.GetDataTypeName(iC55367).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C55367DataType(iC55367 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C55367DataType(iC55367 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C55367DataType(iC55367 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC55367
        RsC55367_EOF = Not RsC55367.Read()

        GoTo Comp_C55368

    Comp_C55368:

        ' EofCliente
        sCurrComponent = "EofCliente"
        C55368DataType = 4
        C55368 = RsC55367_EOF
        GoTo Comp_C55369

    Comp_C55369:

        ' EofCliente=1
        sCurrComponent = "EofCliente=1"
        C55369 = (fn_ConvertValueCompiled(C55368, C55368DataType, AKB_DecimalPoint, False) = 1)
        C55369DataType = AKBTypeConst.cAKBTypeLogical
        If C55369 Then
            GoTo Comp_C233266
        Else
            GoTo Comp_C55370
        End If

    Comp_C55370:

        ' %CalculadoCliente
        sCurrComponent = "%CalculadoCliente"
        C55370DataType = 0
        C55370 = RsC55367(0)
        C55370DataType = C55367Datatype(1)
        If C55370DataType = AKBTypeConst.cAKBTypeString Then
          C55370 = IIF(IsDBNull(C55370), C55370, Strings.RTrim(C55370))
        End If 

        GoTo Comp_C55371

    Comp_C55371:

        ' Att4
        sCurrComponent = "Att4"
        C55371DataType = 4
        C52385 = fn_ConvertValueCompiled(C55370 , 1, AKB_DecimalPoint)
        C55371 = True
        GoTo Comp_C52823

    Comp_C55373:

        ' RetZona
        sCurrComponent = "RetZona"
        Dim tb_C55373 As DataTable = New DataTable()
        tb_C55373.Columns.Add("vRet" & "")
        Dim row_C55373 As DataRow = tb_C55373.NewRow()
        row_C55373(0) = C233224
        tb_C55373.Rows.Add(row_C55373)
        R4484 = tb_C55373
        ReDim C55373DataType(1)
        C55373DataType(1) = C233224DataType
        ReturnDataType = C55373DataType

        GoTo Exit_R4484

    Comp_C233199:

        ' Sel_PermiteDig?
        sCurrComponent = "Sel_PermiteDig?"
        QueryC233199 = con.CreateCommand()
        QueryC233199.CommandText = QueryC233199.CommandText & " " & "Select NVL(WF_TP_PRODUTO.Permite_Dig_Comissao_Ped ,0)"
        QueryC233199.CommandText = QueryC233199.CommandText & " " & "from AKBUSER01.WF_TP_PRODUTO "
        QueryC233199.CommandText = QueryC233199.CommandText & " " & "where WF_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233199.Transaction = txn
        RsC233199 = QueryC233199.ExecuteReader()
        Dim iC233199 As Short
        ReDim C233199DataType(RsC233199.FieldCount)
        For iC233199 = 0 to RsC233199.FieldCount - 1
            Select Case RsC233199.GetDataTypeName(iC233199).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C233199DataType(iC233199 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C233199DataType(iC233199 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C233199DataType(iC233199 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC233199
        RsC233199_EOF = Not RsC233199.Read()

        GoTo Comp_C233538

    Comp_C233201:

        ' PermiteDig=1
        sCurrComponent = "PermiteDig=1"
        C233201 = (fn_ConvertValueCompiled(RsC233199(0), C233199DataType(1), AKB_DecimalPoint, False) = 1)
        C233201DataType = AKBTypeConst.cAKBTypeLogical
        If C233201 Then
            GoTo Comp_C233225
        Else
            GoTo Comp_C233228
        End If

    Comp_C233213:

        ' UpComissãoInfo
        sCurrComponent = "UpComissãoInfo"
        QueryC233213 = con.CreateCommand()
        QueryC233213.CommandText = QueryC233213.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_SEQ SET WF_PEDIDO_SEQ.Porc_Comissao = " & _ 
ConvertValue(P12205, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233213.CommandText = QueryC233213.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto =  " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233213.CommandText = QueryC233213.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233213.Transaction = txn
        C233213 = QueryC233213.ExecuteNonQuery()
        C233213DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C55373

    Comp_C233217:

        ' ComissaoZona
        sCurrComponent = "ComissaoZona"
        QueryC233217 = con.CreateCommand()
        QueryC233217.CommandText = QueryC233217.CommandText & " " & "SELECT  NVL( WF_REPRES_ZONA.Comissao , 0 )"
        QueryC233217.CommandText = QueryC233217.CommandText & " " & ""
        QueryC233217.CommandText = QueryC233217.CommandText & " " & "FROM AKBUSER01.WF_CLIENTES , AKBUSER01.WF_REPRES_ZONA "
        QueryC233217.CommandText = QueryC233217.CommandText & " " & ""
        QueryC233217.CommandText = QueryC233217.CommandText & " " & "WHERE  WF_CLIENTES.Cod_Cliente = " & _ 
ConvertValue(P40622, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233217.CommandText = QueryC233217.CommandText & " " & "AND  WF_REPRES_ZONA.Cod_Repres = " & _ 
ConvertValue(P40621, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233217.CommandText = QueryC233217.CommandText & " " & "AND  WF_CLIENTES.Cod_Zona =  WF_REPRES_ZONA.Cod_Zona "
        QueryC233217.CommandText = QueryC233217.CommandText & " " & ""
        QueryC233217.Transaction = txn
        RsC233217 = QueryC233217.ExecuteReader()
        Dim iC233217 As Short
        ReDim C233217DataType(RsC233217.FieldCount)
        For iC233217 = 0 to RsC233217.FieldCount - 1
            Select Case RsC233217.GetDataTypeName(iC233217).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C233217DataType(iC233217 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C233217DataType(iC233217 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C233217DataType(iC233217 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC233217
        RsC233217_EOF = Not RsC233217.Read()

        GoTo Comp_C233219

    Comp_C233218:

        ' Cod_Repres=0
        sCurrComponent = "Cod_Repres=0"
        C233218 = (fn_ConvertValueCompiled(P40621, 1, AKB_DecimalPoint, False) = 0)
        C233218DataType = AKBTypeConst.cAKBTypeLogical
        If C233218 Then
            GoTo Comp_C233229
        Else
            GoTo Comp_C233217
        End If

    Comp_C233219:

        ' EOF_Zona
        sCurrComponent = "EOF_Zona"
        C233219DataType = 4
        C233219 = RsC233217_EOF
        GoTo Comp_C233220

    Comp_C233220:

        ' EOF_Zona=1
        sCurrComponent = "EOF_Zona=1"
        C233220 = (fn_ConvertValueCompiled(C233219, C233219DataType, AKB_DecimalPoint, False) = 1)
        C233220DataType = AKBTypeConst.cAKBTypeLogical
        If C233220 Then
            GoTo Comp_C233262
        Else
            GoTo Comp_C233223
        End If

    Comp_C233223:

        ' ComInf=ComZona
        sCurrComponent = "ComInf=ComZona"
        C233223 = (fn_ConvertValueCompiled(P12205, 1, AKB_DecimalPoint, False) = fn_ConvertValueCompiled(RsC233217(0), C233217DataType(1), AKB_DecimalPoint, False))
        C233223DataType = AKBTypeConst.cAKBTypeLogical
        If C233223 Then
            GoTo Comp_C233213
        Else
            GoTo Comp_C296826
        End If

    Comp_C233224:

        ' vRet
        sCurrComponent = "vRet"
        C233224 = 1
        C233224DataType = 4
        GoTo Comp_C53562

    Comp_C233225:

        ' ConfirmaZona
        sCurrComponent = "ConfirmaZona"
        C233225 = 6
        C233225DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C233226

    Comp_C233226:

        ' ConfirmaZona=Sim
        sCurrComponent = "ConfirmaZona=Sim"
        C233226 = (fn_ConvertValueCompiled(C233225, C233225DataType, AKB_DecimalPoint, False) = 6)
        C233226DataType = AKBTypeConst.cAKBTypeLogical
        If C233226 Then
            GoTo Comp_C233213
        Else
            GoTo Comp_C233227
        End If

    Comp_C233227:

        ' atvFalso
        sCurrComponent = "atvFalso"
        C233227DataType = 4
        C233224 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C233227 = True
        GoTo Comp_C55373

    Comp_C233228:

        ' UpComissaoZona
        sCurrComponent = "UpComissaoZona"
        QueryC233228 = con.CreateCommand()
        QueryC233228.CommandText = QueryC233228.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_SEQ SET WF_PEDIDO_SEQ.Porc_Comissao = " & _ 
ConvertValue(RsC233217(0), C233217DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC233228.CommandText = QueryC233228.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto =  " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233228.CommandText = QueryC233228.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233228.Transaction = txn
        C233228 = QueryC233228.ExecuteNonQuery()
        C233228DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C233261

    Comp_C233229:

        ' RetRepres
        sCurrComponent = "RetRepres"
        Dim tb_C233229 As DataTable = New DataTable()
        tb_C233229.Columns.Add("vRet" & "")
        Dim row_C233229 As DataRow = tb_C233229.NewRow()
        row_C233229(0) = C233224
        tb_C233229.Rows.Add(row_C233229)
        R4484 = tb_C233229
        ReDim C233229DataType(1)
        C233229DataType(1) = C233224DataType
        ReturnDataType = C233229DataType

        GoTo Exit_R4484

    Comp_C233256:

        ' UpComissaoCalc
        sCurrComponent = "UpComissaoCalc"
        QueryC233256 = con.CreateCommand()
        QueryC233256.CommandText = QueryC233256.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_SEQ SET WF_PEDIDO_SEQ.Porc_Comissao = " & _ 
ConvertValue(C52385, C52385DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233256.CommandText = QueryC233256.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto =  " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233256.CommandText = QueryC233256.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233256.Transaction = txn
        C233256 = QueryC233256.ExecuteNonQuery()
        C233256DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C52833

    Comp_C233257:

        ' Sel_PermiteDig=1
        sCurrComponent = "Sel_PermiteDig=1"
        C233257 = (fn_ConvertValueCompiled(RsC233199(0), C233199DataType(1), AKB_DecimalPoint, False) = 1)
        C233257DataType = AKBTypeConst.cAKBTypeLogical
        If C233257 Then
            GoTo Comp_C52387
        Else
            GoTo Comp_C233259
        End If

    Comp_C233258:

        ' UpComissaoInfo
        sCurrComponent = "UpComissaoInfo"
        QueryC233258 = con.CreateCommand()
        QueryC233258.CommandText = QueryC233258.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_SEQ SET WF_PEDIDO_SEQ.Porc_Comissao = " & _ 
ConvertValue(P12205, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233258.CommandText = QueryC233258.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto =  " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233258.CommandText = QueryC233258.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233258.Transaction = txn
        C233258 = QueryC233258.ExecuteNonQuery()
        C233258DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C52389

    Comp_C233259:

        ' UpComissCalcDif
        sCurrComponent = "UpComissCalcDif"
        QueryC233259 = con.CreateCommand()
        QueryC233259.CommandText = QueryC233259.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_SEQ SET WF_PEDIDO_SEQ.Porc_Comissao = " & _ 
ConvertValue(C52385, C52385DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233259.CommandText = QueryC233259.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto =  " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233259.CommandText = QueryC233259.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233259.Transaction = txn
        C233259 = QueryC233259.ExecuteNonQuery()
        C233259DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C52389

    Comp_C233260:

        ' atvRet=0
        sCurrComponent = "atvRet=0"
        C233260DataType = 4
        C233224 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C233260 = True
        GoTo Comp_C52389

    Comp_C233261:

        ' RetComisZonaDif
        sCurrComponent = "RetComisZonaDif"
        Dim tb_C233261 As DataTable = New DataTable()
        tb_C233261.Columns.Add("vRet" & "")
        Dim row_C233261 As DataRow = tb_C233261.NewRow()
        row_C233261(0) = C233224
        tb_C233261.Rows.Add(row_C233261)
        R4484 = tb_C233261
        ReDim C233261DataType(1)
        C233261DataType(1) = C233224DataType
        ReturnDataType = C233261DataType

        GoTo Exit_R4484

    Comp_C233262:

        ' Permite=1
        sCurrComponent = "Permite=1"
        C233262 = (fn_ConvertValueCompiled(RsC233199(0), C233199DataType(1), AKB_DecimalPoint, False) = 1)
        C233262DataType = AKBTypeConst.cAKBTypeLogical
        If C233262 Then
            GoTo Comp_C233213
        Else
            GoTo Comp_C55373
        End If

    Comp_C233263:

        ' %ComissãoInformada>0
        sCurrComponent = "%ComissãoInformada>0"
        C233263 = (fn_ConvertValueCompiled(C296826, C296826DataType, AKB_DecimalPoint, False) = 1 or fn_ConvertValueCompiled(P40894, 4, AKB_DecimalPoint, False) = 1)
        C233263DataType = AKBTypeConst.cAKBTypeLogical
        If C233263 Then
            GoTo Comp_C233228
        Else
            GoTo Comp_C233201
        End If

    Comp_C233264:

        ' ComissaoInfo>0
        sCurrComponent = "ComissaoInfo>0"
        C233264 = (( fn_ConvertValueCompiled(P12205, 1, AKB_DecimalPoint, False) Is System.DBNull.Value )  or ( fn_ConvertValueCompiled(P40894, 4, AKB_DecimalPoint, False)=1 ))
        C233264DataType = AKBTypeConst.cAKBTypeLogical
        If C233264 Then
            GoTo Comp_C233259
        Else
            GoTo Comp_C233257
        End If

    Comp_C233265:

        ' EOF_Desc=1
        sCurrComponent = "EOF_Desc=1"
        C233265 = (fn_ConvertValueCompiled(C233267, C233267DataType, AKB_DecimalPoint, False) = 1)
        C233265DataType = AKBTypeConst.cAKBTypeLogical
        If C233265 Then
            GoTo Comp_C52823
        Else
            GoTo Comp_C52365
        End If

    Comp_C233266:

        ' SelDesconto
        sCurrComponent = "SelDesconto"
        QueryC233266 = con.CreateCommand()
        QueryC233266.CommandText = QueryC233266.CommandText & " " & "SELECT WF_PED_SEQ_DESCONTO.Porc_Desconto "
        QueryC233266.CommandText = QueryC233266.CommandText & " " & ""
        QueryC233266.CommandText = QueryC233266.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_DESCONTO ,  WF_COMISSAO_TIPO_DESCONTO D"
        QueryC233266.CommandText = QueryC233266.CommandText & " " & ""
        QueryC233266.CommandText = QueryC233266.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233266.CommandText = QueryC233266.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233266.CommandText = QueryC233266.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Seq = 1"
        QueryC233266.CommandText = QueryC233266.CommandText & " " & ""
        QueryC233266.CommandText = QueryC233266.CommandText & " " & "AND D.TIPO_DESCONTO = WF_PED_SEQ_DESCONTO.Tipo_Desconto "
        QueryC233266.CommandText = QueryC233266.CommandText & " " & ""
        QueryC233266.CommandText = QueryC233266.CommandText & " " & "AND D.POLITICA = " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233266.CommandText = QueryC233266.CommandText & " " & ""
        QueryC233266.CommandText = QueryC233266.CommandText & " " & ""
        QueryC233266.Transaction = txn
        RsC233266 = QueryC233266.ExecuteReader()
        Dim iC233266 As Short
        ReDim C233266DataType(RsC233266.FieldCount)
        For iC233266 = 0 to RsC233266.FieldCount - 1
            Select Case RsC233266.GetDataTypeName(iC233266).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C233266DataType(iC233266 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C233266DataType(iC233266 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C233266DataType(iC233266 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC233266
        RsC233266_EOF = Not RsC233266.Read()

        GoTo Comp_C233267

    Comp_C233267:

        ' EOF_Desc
        sCurrComponent = "EOF_Desc"
        C233267DataType = 4
        C233267 = RsC233266_EOF
        GoTo Comp_C233265

    Comp_C233538:

        ' VerificaZona
        sCurrComponent = "VerificaZona"
        C233538 = clsRuleDynamicallyCompiled_R12221.R12221(con, txn, IIf(Not IsDBNull(P40622), P40622, System.DBNull.Value), IIf(Not IsDBNull(P12201), P12201, System.DBNull.Value), IIf(Not IsDBNull(P12202), P12202, System.DBNull.Value), IIf(Not IsDBNull(P68671), P68671, System.DBNull.Value))
        C233538CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C233538) Then
          iColumns = C233538.Columns.Count
        End If
        ReDim C233538DataType(iColumns)
        For iCol = 1 To iColumns
            C233538DataType(iCol) = clsRuleDynamicallyCompiled_R12221.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C233540

    Comp_C233540:

        ' VerificaZona=1
        sCurrComponent = "VerificaZona=1"
        C233540 = (fn_ConvertValueCompiled(C233538.rows(C233538CurrentRow - 1)(0), C233538DataType(1), AKB_DecimalPoint, False) = 1)
        C233540DataType = AKBTypeConst.cAKBTypeLogical
        If C233540 Then
            GoTo Comp_C531933
        Else
            GoTo Comp_C233541
        End If

    Comp_C233541:

        ' atFalso
        sCurrComponent = "atFalso"
        C233541DataType = 4
        C233224 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C233541 = True
        GoTo Comp_C233229

    Comp_C238895:

        ' SelDescItens
        sCurrComponent = "SelDescItens"
        QueryC238895 = con.CreateCommand()
        QueryC238895.CommandText = QueryC238895.CommandText & " " & "Select  count( distinct WF_PED_ITENS_DESC.seq_item)"
        QueryC238895.CommandText = QueryC238895.CommandText & " " & "from WF_PED_ITENS_DESC, WF_COMISSAO_TIPO_DESCONTO"
        QueryC238895.CommandText = QueryC238895.CommandText & " " & "where WF_COMISSAO_TIPO_DESCONTO.TIPO_DESCONTO = WF_PED_ITENS_DESC.TIPO_DESCONTO"
        QueryC238895.CommandText = QueryC238895.CommandText & " " & "and WF_PED_ITENS_DESC.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC238895.CommandText = QueryC238895.CommandText & " " & "and WF_PED_ITENS_DESC.TP_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC238895.CommandText = QueryC238895.CommandText & " " & "and WF_COMISSAO_TIPO_DESCONTO.POLITICA = " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC238895.Transaction = txn
        RsC238895 = QueryC238895.ExecuteReader()
        Dim iC238895 As Short
        ReDim C238895DataType(RsC238895.FieldCount)
        For iC238895 = 0 to RsC238895.FieldCount - 1
            Select Case RsC238895.GetDataTypeName(iC238895).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C238895DataType(iC238895 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C238895DataType(iC238895 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C238895DataType(iC238895 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC238895
        RsC238895_EOF = Not RsC238895.Read()

        GoTo Comp_C238906

    Comp_C238906:

        ' BOF_DescItens=1
        sCurrComponent = "BOF_DescItens=1"
        C238906 = (fn_ConvertValueCompiled(RsC238895(0), C238895DataType(1), AKB_DecimalPoint, False) = 0)
        C238906DataType = AKBTypeConst.cAKBTypeLogical
        If C238906 Then
            GoTo Comp_C238908
        Else
            GoTo Comp_C239002
        End If

    Comp_C238908:

        ' SelDescPed
        sCurrComponent = "SelDescPed"
        QueryC238908 = con.CreateCommand()
        QueryC238908.CommandText = QueryC238908.CommandText & " " & "SELECT NVL(sum(WF_PED_SEQ_DESCONTO.Porc_Desconto),0)"
        QueryC238908.CommandText = QueryC238908.CommandText & " " & ""
        QueryC238908.CommandText = QueryC238908.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_DESCONTO ,  WF_COMISSAO_TIPO_DESCONTO D"
        QueryC238908.CommandText = QueryC238908.CommandText & " " & ""
        QueryC238908.CommandText = QueryC238908.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC238908.CommandText = QueryC238908.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC238908.CommandText = QueryC238908.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Seq = 1"
        QueryC238908.CommandText = QueryC238908.CommandText & " " & ""
        QueryC238908.CommandText = QueryC238908.CommandText & " " & "AND D.TIPO_DESCONTO = WF_PED_SEQ_DESCONTO.Tipo_Desconto "
        QueryC238908.CommandText = QueryC238908.CommandText & " " & ""
        QueryC238908.CommandText = QueryC238908.CommandText & " " & "AND D.POLITICA = " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC238908.Transaction = txn
        RsC238908 = QueryC238908.ExecuteReader()
        Dim iC238908 As Short
        ReDim C238908DataType(RsC238908.FieldCount)
        For iC238908 = 0 to RsC238908.FieldCount - 1
            Select Case RsC238908.GetDataTypeName(iC238908).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C238908DataType(iC238908 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C238908DataType(iC238908 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C238908DataType(iC238908 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC238908
        RsC238908_EOF = Not RsC238908.Read()

        GoTo Comp_C238909

    Comp_C238909:

        ' DescPed=0
        sCurrComponent = "DescPed=0"
        C238909 = (fn_ConvertValueCompiled(RsC238908(0), C238908DataType(1), AKB_DecimalPoint, False) = 0)
        C238909DataType = AKBTypeConst.cAKBTypeLogical
        If C238909 Then
            GoTo Comp_C239088
        Else
            GoTo Comp_C238936
        End If

    Comp_C238910:

        ' vDesconto
        sCurrComponent = "vDesconto"
        C238910 = 0
        C238910DataType = 1
        GoTo Comp_C239079

    Comp_C238911:

        ' atvDescItens
        sCurrComponent = "atvDescItens"
        C238911DataType = 4
        C238910 = fn_ConvertValueCompiled(C239002.rows(C239002CurrentRow - 1)(0) , 1, AKB_DecimalPoint)
        C238911 = True
        GoTo Comp_C239080

    Comp_C238912:

        ' atvDescPed
        sCurrComponent = "atvDescPed"
        C238912DataType = 4
        C238910 = fn_ConvertValueCompiled(C239075 , 1, AKB_DecimalPoint)
        C238912 = True
        GoTo Comp_C239085

    Comp_C238936:

        ' DescPed
        sCurrComponent = "DescPed"
        'C238936 = clsRuleDynamicallyCompiled_R4483.R4483(con, txn, IIf(Not IsDBNull(P12201), P12201, System.DBNull.Value), IIf(Not IsDBNull(P12202), P12202, System.DBNull.Value), fn_ConvertValueCompiled( "1", 1, AKB_DecimalPoint), IIf(Not IsDBNull(C52366), C52366, System.DBNull.Value))
        C238936CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C238936) Then
          iColumns = C238936.Columns.Count
        End If
        ReDim C238936DataType(iColumns)
        For iCol = 1 To iColumns
            'C238936DataType(iCol) = clsRuleDynamicallyCompiled_R4483.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C239081

    Comp_C239002:

        ' DescItem
        sCurrComponent = "DescItem"
        'C239002 = clsRuleDynamicallyCompiled_R12432.R12432(con, txn, IIf(Not IsDBNull(P12201), P12201, System.DBNull.Value), IIf(Not IsDBNull(P12202), P12202, System.DBNull.Value), IIf(Not IsDBNull(C52366), C52366, System.DBNull.Value))
        C239002CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C239002) Then
          iColumns = C239002.Columns.Count
        End If
        ReDim C239002DataType(iColumns)
        For iCol = 1 To iColumns
            'C239002DataType(iCol) = clsRuleDynamicallyCompiled_R12432.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C238911

    Comp_C239075:

        ' cDescPed
        sCurrComponent = "cDescPed"
        C239075 = Math.Round(fn_ConvertValueCompiled(C238910, C238910DataType, AKB_DecimalPoint, True) + ( fn_ConvertValueCompiled(C238936.rows(C238936CurrentRow - 1)(0), C238936DataType(1), AKB_DecimalPoint, True) * fn_ConvertValueCompiled(RsC239081(0), C239081DataType(1), AKB_DecimalPoint, True) ), 2)
        C239075DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C238912

    Comp_C239079:

        ' vQtdeItens
        sCurrComponent = "vQtdeItens"
        C239079 = 0
        C239079DataType = 1
        GoTo Comp_C238895

    Comp_C239080:

        ' atvQtdeItens
        sCurrComponent = "atvQtdeItens"
        C239080DataType = 4
        C239079 = fn_ConvertValueCompiled(RsC238895(0) , 1, AKB_DecimalPoint)
        C239080 = True
        GoTo Comp_C238908

    Comp_C239081:

        ' SelItens
        sCurrComponent = "SelItens"
        QueryC239081 = con.CreateCommand()
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "Select count(distinct WF_PEDIDO_ITENS.seq_item)"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "from WF_Pedido_Itens, WF_PED_SEQ_DESCONTO, WF_COMISSAO_TIPO_DESCONTO D"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "where WF_Pedido_Itens.Porc_desconto is not null"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and WF_Pedido_Itens.porc_desconto <> 0"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and WF_Pedido_Itens.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and WF_Pedido_Itens.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and D.POLITICA =  " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and WF_PED_SEQ_DESCONTO.Seq = 1"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and WF_Pedido_Itens.TP_Produto = WF_PED_SEQ_DESCONTO.Tp_Produto"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and WF_Pedido_Itens.Pedido = WF_PED_SEQ_DESCONTO.Pedido"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and D.TIPO_DESCONTO = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & ""
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "and not exists  ( Select WF_PED_ITENS_DESC.seq_item"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                  from WF_PED_ITENS_DESC, WF_COMISSAO_TIPO_DESCONTO"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                  where WF_COMISSAO_TIPO_DESCONTO.TIPO_DESCONTO = WF_PED_ITENS_DESC.TIPO_DESCONTO"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                  and WF_PED_ITENS_DESC.Seq_Item = WF_Pedido_Itens.Seq_Item"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                  and WF_PED_ITENS_DESC.Pedido = WF_Pedido_Itens.Pedido"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                  and WF_PED_ITENS_DESC.TP_Produto = WF_Pedido_Itens.TP_Produto"
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                  and WF_PED_ITENS_DESC.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                  and WF_PED_ITENS_DESC.TP_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                  and WF_COMISSAO_TIPO_DESCONTO.POLITICA = " & _ 
ConvertValue(C52366, C52366DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC239081.CommandText = QueryC239081.CommandText & " " & "                )"
        QueryC239081.Transaction = txn
        RsC239081 = QueryC239081.ExecuteReader()
        Dim iC239081 As Short
        ReDim C239081DataType(RsC239081.FieldCount)
        For iC239081 = 0 to RsC239081.FieldCount - 1
            Select Case RsC239081.GetDataTypeName(iC239081).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C239081DataType(iC239081 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C239081DataType(iC239081 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C239081DataType(iC239081 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC239081
        RsC239081_EOF = Not RsC239081.Read()

        GoTo Comp_C239075

    Comp_C239085:

        ' cQtdePed
        sCurrComponent = "cQtdePed"
        C239085 = fn_ConvertValueCompiled(C239079, C239079DataType, AKB_DecimalPoint, True) + fn_ConvertValueCompiled(RsC239081(0), C239081DataType(1), AKB_DecimalPoint, True)
        C239085DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C239087

    Comp_C239087:

        ' atvQtdePed
        sCurrComponent = "atvQtdePed"
        C239087DataType = 4
        C239079 = fn_ConvertValueCompiled(C239085 , 1, AKB_DecimalPoint)
        C239087 = True
        GoTo Comp_C239088

    Comp_C239088:

        ' vQtdeItens=0
        sCurrComponent = "vQtdeItens=0"
        C239088 = (fn_ConvertValueCompiled(C239079, C239079DataType, AKB_DecimalPoint, False) = 0)
        C239088DataType = AKBTypeConst.cAKBTypeLogical
        If C239088 Then
            GoTo Comp_C52838
        Else
            GoTo Comp_C239089
        End If

    Comp_C239089:

        ' cMediaDesc
        sCurrComponent = "cMediaDesc"
        C239089 = Math.Round(fn_ConvertValueCompiled(C238910, C238910DataType, AKB_DecimalPoint, True) / fn_ConvertValueCompiled(C239079, C239079DataType, AKB_DecimalPoint, True), 2)
        C239089DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C239090

    Comp_C239090:

        ' atvvDesc
        sCurrComponent = "atvvDesc"
        C239090DataType = 4
        C238910 = fn_ConvertValueCompiled(C239089 , 1, AKB_DecimalPoint)
        C239090 = True
        GoTo Comp_C52838

    Comp_C296826:

        ' NuloComissaoInformada
        sCurrComponent = "NuloComissaoInformada"
        C296826DataType = 4
        C296826 = IsDBNull (P12205)
        GoTo Comp_C233263

    Comp_C402182:

        ' %ComisTabPrc
        sCurrComponent = "%ComisTabPrc"
        QueryC402182 = con.CreateCommand()
        QueryC402182.CommandText = QueryC402182.CommandText & " " & "SELECT NVL(T.Percentual_Comissao,0)"
        QueryC402182.CommandText = QueryC402182.CommandText & " " & " FROM WF_TABELA_PRECO_VENDA T, WF_PEDIDO P"
        QueryC402182.CommandText = QueryC402182.CommandText & " " & "WHERE T.Tabela_Preco_Venda = P.Tabela_Preco_Venda"
        QueryC402182.CommandText = QueryC402182.CommandText & " " & "AND P.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC402182.CommandText = QueryC402182.CommandText & " " & "AND P.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC402182.Transaction = txn
        RsC402182 = QueryC402182.ExecuteReader()
        Dim iC402182 As Short
        ReDim C402182DataType(RsC402182.FieldCount)
        For iC402182 = 0 to RsC402182.FieldCount - 1
            Select Case RsC402182.GetDataTypeName(iC402182).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C402182DataType(iC402182 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C402182DataType(iC402182 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C402182DataType(iC402182 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC402182
        RsC402182_EOF = Not RsC402182.Read()

        GoTo Comp_C402183

    Comp_C402183:

        ' DESVIO1
        sCurrComponent = "DESVIO1"
        C402183 = (fn_ConvertValueCompiled(RsC402182(0), C402182DataType(1), AKB_DecimalPoint, False) > 0)
        C402183DataType = AKBTypeConst.cAKBTypeLogical
        If C402183 Then
            GoTo Comp_C402184
        Else
            GoTo Comp_C52364
        End If

    Comp_C402184:

        ' upComisTabPrc
        sCurrComponent = "upComisTabPrc"
        QueryC402184 = con.CreateCommand()
        QueryC402184.CommandText = QueryC402184.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_SEQ SET WF_PEDIDO_SEQ.Porc_Comissao = " & _ 
ConvertValue(RsC402182(0), C402182DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC402184.CommandText = QueryC402184.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto =  " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC402184.CommandText = QueryC402184.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC402184.Transaction = txn
        C402184 = QueryC402184.ExecuteNonQuery()
        C402184DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C402185

    Comp_C402185:

        ' Ret11
        sCurrComponent = "Ret11"
        Dim tb_C402185 As DataTable = New DataTable()
        tb_C402185.Columns.Add("vRet" & "")
        Dim row_C402185 As DataRow = tb_C402185.NewRow()
        row_C402185(0) = C233224
        tb_C402185.Rows.Add(row_C402185)
        R4484 = tb_C402185
        ReDim C402185DataType(1)
        C402185DataType(1) = C233224DataType
        ReturnDataType = C402185DataType

        GoTo Exit_R4484

    Comp_C402878:

        ' NaoGeraDupl
        sCurrComponent = "NaoGeraDupl"
        QueryC402878 = con.CreateCommand()
        QueryC402878.CommandText = QueryC402878.CommandText & " " & "SELECT C.Nao_Gera_Duplicata"
        QueryC402878.CommandText = QueryC402878.CommandText & " " & " FROM WF_PEDIDO P, WF_CONDICAO_PGTO C"
        QueryC402878.CommandText = QueryC402878.CommandText & " " & " WHERE P.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC402878.CommandText = QueryC402878.CommandText & " " & " AND P.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC402878.CommandText = QueryC402878.CommandText & " " & " AND P.Cod_Pagto = C.Cod_Pagto"
        QueryC402878.Transaction = txn
        RsC402878 = QueryC402878.ExecuteReader()
        Dim iC402878 As Short
        ReDim C402878DataType(RsC402878.FieldCount)
        For iC402878 = 0 to RsC402878.FieldCount - 1
            Select Case RsC402878.GetDataTypeName(iC402878).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C402878DataType(iC402878 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C402878DataType(iC402878 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C402878DataType(iC402878 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC402878
        RsC402878_EOF = Not RsC402878.Read()

        GoTo Comp_C402879

    Comp_C402879:

        ' DESVIO2
        sCurrComponent = "DESVIO2"
        C402879 = (fn_ConvertValueCompiled(RsC402878(0), C402878DataType(1), AKB_DecimalPoint, False) = 1)
        C402879DataType = AKBTypeConst.cAKBTypeLogical
        If C402879 Then
            GoTo Comp_C52368
        Else
            GoTo Comp_C233199
        End If

    Comp_C531933:

        ' CountTabPromocPE
        sCurrComponent = "CountTabPromocPE"
        QueryC531933 = con.CreateCommand()
        QueryC531933.CommandText = QueryC531933.CommandText & " " & "SELECT NVL(COUNT(*),0)"
        QueryC531933.CommandText = QueryC531933.CommandText & " " & " FROM WF_TABELA_PRECO_VENDA T, WF_PEDIDO P"
        QueryC531933.CommandText = QueryC531933.CommandText & " " & "WHERE T.Tabela_Preco_Venda = P.Tabela_Preco_Venda"
        QueryC531933.CommandText = QueryC531933.CommandText & " " & "AND P.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC531933.CommandText = QueryC531933.CommandText & " " & "AND P.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC531933.CommandText = QueryC531933.CommandText & " " & "AND EXISTS (SELECT * FROM WF_COMIS_TAB_PCO_DESCONTO CT"
        QueryC531933.CommandText = QueryC531933.CommandText & " " & "                               WHERE CT.Tabela_Preco_Venda = T.Tabela_Preco_Venda)"
        QueryC531933.Transaction = txn
        RsC531933 = QueryC531933.ExecuteReader()
        Dim iC531933 As Short
        ReDim C531933DataType(RsC531933.FieldCount)
        For iC531933 = 0 to RsC531933.FieldCount - 1
            Select Case RsC531933.GetDataTypeName(iC531933).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C531933DataType(iC531933 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C531933DataType(iC531933 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C531933DataType(iC531933 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC531933
        RsC531933_EOF = Not RsC531933.Read()

        GoTo Comp_C531934

    Comp_C531934:

        ' CountTabPromocPE=0?
        sCurrComponent = "CountTabPromocPE=0?"
        C531934 = (fn_ConvertValueCompiled(RsC531933(0), C531933DataType(1), AKB_DecimalPoint, False) = 0)
        C531934DataType = AKBTypeConst.cAKBTypeLogical
        If C531934 Then
            GoTo Comp_C402182
        Else
            GoTo Comp_C535152
        End If

    Comp_C531935:

        ' QtdDiasCP
        sCurrComponent = "QtdDiasCP"
        QueryC531935 = con.CreateCommand()
        QueryC531935.CommandText = QueryC531935.CommandText & " " & "SELECT SUM(WF_OCORRENCIAS_COND_PAGTO.Quant_Dias) FROM AKBUSER01.WF_OCORRENCIAS_COND_PAGTO , AKBUSER01.WF_PEDIDO WHERE WF_OCORRENCIAS_COND_PAGTO.Cod_Pagto = WF_PEDIDO.Cod_Pagto AND  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC531935.Transaction = txn
        RsC531935 = QueryC531935.ExecuteReader()
        Dim iC531935 As Short
        ReDim C531935DataType(RsC531935.FieldCount)
        For iC531935 = 0 to RsC531935.FieldCount - 1
            Select Case RsC531935.GetDataTypeName(iC531935).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C531935DataType(iC531935 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C531935DataType(iC531935 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C531935DataType(iC531935 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC531935
        RsC531935_EOF = Not RsC531935.Read()

        GoTo Comp_C531936

    Comp_C531936:

        ' QtdDplCP
        sCurrComponent = "QtdDplCP"
        QueryC531936 = con.CreateCommand()
        QueryC531936.CommandText = QueryC531936.CommandText & " " & "SELECT COUNT(WF_OCORRENCIAS_COND_PAGTO.Seq_Cond_Pagto) FROM AKBUSER01.WF_OCORRENCIAS_COND_PAGTO , AKBUSER01.WF_PEDIDO WHERE WF_OCORRENCIAS_COND_PAGTO.Cod_Pagto = WF_PEDIDO.Cod_Pagto AND  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC531936.Transaction = txn
        RsC531936 = QueryC531936.ExecuteReader()
        Dim iC531936 As Short
        ReDim C531936DataType(RsC531936.FieldCount)
        For iC531936 = 0 to RsC531936.FieldCount - 1
            Select Case RsC531936.GetDataTypeName(iC531936).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C531936DataType(iC531936 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C531936DataType(iC531936 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C531936DataType(iC531936 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC531936
        RsC531936_EOF = Not RsC531936.Read()

        GoTo Comp_C531937

    Comp_C531937:

        ' DiasMedio
        sCurrComponent = "DiasMedio"
        C531937 = Math.Round(fn_ConvertValueCompiled(RsC531935(0), C531935DataType(1), AKB_DecimalPoint, True) / fn_ConvertValueCompiled(RsC531936(0), C531936DataType(1), AKB_DecimalPoint, True), 0)
        C531937DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C531938

    Comp_C531938:

        ' %DescPed
        sCurrComponent = "%DescPed"
        QueryC531938 = con.CreateCommand()
        QueryC531938.CommandText = QueryC531938.CommandText & " " & "SELECT NVL(WF_PEDIDO_SEQ.Porc_Desc_Ped,0) "
        QueryC531938.CommandText = QueryC531938.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_SEQ "
        QueryC531938.CommandText = QueryC531938.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC531938.CommandText = QueryC531938.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC531938.Transaction = txn
        RsC531938 = QueryC531938.ExecuteReader()
        Dim iC531938 As Short
        ReDim C531938DataType(RsC531938.FieldCount)
        For iC531938 = 0 to RsC531938.FieldCount - 1
            Select Case RsC531938.GetDataTypeName(iC531938).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C531938DataType(iC531938 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C531938DataType(iC531938 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C531938DataType(iC531938 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC531938
        RsC531938_EOF = Not RsC531938.Read()

        GoTo Comp_C535150

    Comp_C535150:

        ' v%DescPed
        sCurrComponent = "v%DescPed"
        C535150DataType = 0
        C535150 = RsC531938(0)
        C535150DataType = C531938Datatype(1)
        If C535150DataType = AKBTypeConst.cAKBTypeString Then
          C535150 = IIF(IsDBNull(C535150), C535150, Strings.RTrim(C535150))
        End If 

        GoTo Comp_C535151

    Comp_C535151:

        ' ComisDesc
        sCurrComponent = "ComisDesc"
        QueryC535151 = con.CreateCommand()
        QueryC535151.CommandText = QueryC535151.CommandText & " " & "SELECT NVL(Porc_Comissao,0), Bloqueia_Pedido"
        QueryC535151.CommandText = QueryC535151.CommandText & " " & "FROM WF_COMIS_TAB_PCO_DESCONTO "
        QueryC535151.CommandText = QueryC535151.CommandText & " " & "WHERE Tabela_Preco_Venda = " & _ 
ConvertValue(C535153, C535153DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC535151.CommandText = QueryC535151.CommandText & " " & "AND " & _ 
ConvertValue(C535150, C535150DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " >= Porc_Desc_Inicial AND " & _ 
ConvertValue(C535150, C535150DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " <= Porc_Desc_Final"
        QueryC535151.CommandText = QueryC535151.CommandText & " " & "AND " & _ 
ConvertValue(C531937, C531937DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " >= Dias_Medio_Inicial AND " & _ 
ConvertValue(C531937, C531937DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " <= Dis_Medio_Final "
        QueryC535151.Transaction = txn
        RsC535151 = QueryC535151.ExecuteReader()
        Dim iC535151 As Short
        ReDim C535151DataType(RsC535151.FieldCount)
        For iC535151 = 0 to RsC535151.FieldCount - 1
            Select Case RsC535151.GetDataTypeName(iC535151).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C535151DataType(iC535151 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C535151DataType(iC535151 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C535151DataType(iC535151 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC535151
        RsC535151_EOF = Not RsC535151.Read()

        GoTo Comp_C535154

    Comp_C535152:

        ' TabPreço
        sCurrComponent = "TabPreço"
        QueryC535152 = con.CreateCommand()
        QueryC535152.CommandText = QueryC535152.CommandText & " " & "SELECT T.Tabela_Preco_Venda"
        QueryC535152.CommandText = QueryC535152.CommandText & " " & " FROM WF_TABELA_PRECO_VENDA T, WF_PEDIDO P"
        QueryC535152.CommandText = QueryC535152.CommandText & " " & "WHERE T.Tabela_Preco_Venda = P.Tabela_Preco_Venda"
        QueryC535152.CommandText = QueryC535152.CommandText & " " & "AND P.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC535152.CommandText = QueryC535152.CommandText & " " & "AND P.Tp_Produto = " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC535152.Transaction = txn
        RsC535152 = QueryC535152.ExecuteReader()
        Dim iC535152 As Short
        ReDim C535152DataType(RsC535152.FieldCount)
        For iC535152 = 0 to RsC535152.FieldCount - 1
            Select Case RsC535152.GetDataTypeName(iC535152).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C535152DataType(iC535152 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C535152DataType(iC535152 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C535152DataType(iC535152 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC535152
        RsC535152_EOF = Not RsC535152.Read()

        GoTo Comp_C535153

    Comp_C535153:

        ' vTabPreço
        sCurrComponent = "vTabPreço"
        C535153DataType = 0
        C535153 = RsC535152(0)
        C535153DataType = C535152Datatype(1)
        If C535153DataType = AKBTypeConst.cAKBTypeString Then
          C535153 = IIF(IsDBNull(C535153), C535153, Strings.RTrim(C535153))
        End If 

        GoTo Comp_C531935

    Comp_C535154:

        ' v%Comis
        sCurrComponent = "v%Comis"
        C535154DataType = 0
        C535154 = RsC535151(0)
        C535154DataType = C535151Datatype(1)
        If C535154DataType = AKBTypeConst.cAKBTypeString Then
          C535154 = IIF(IsDBNull(C535154), C535154, Strings.RTrim(C535154))
        End If 

        GoTo Comp_C535155

    Comp_C535155:

        ' vBloqPed
        sCurrComponent = "vBloqPed"
        C535155DataType = 0
        C535155DataType = C535151Datatype(2)
        If C535155DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC535151(1)) Then
          C535155 = Strings.RTrim(RsC535151(1))
        Else 
          C535155 = RsC535151(1)
        End If 

        GoTo Comp_C535158

    Comp_C535156:

        ' ATRIBUIVA1
        sCurrComponent = "ATRIBUIVA1"
        C535156DataType = 4
        C233224 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C535156 = True
        GoTo Comp_C233229

    Comp_C535157:

        ' UpComisDesc
        sCurrComponent = "UpComisDesc"
        QueryC535157 = con.CreateCommand()
        QueryC535157.CommandText = QueryC535157.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_SEQ SET WF_PEDIDO_SEQ.Porc_Comissao = " & _ 
ConvertValue(C535154, C535154DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC535157.CommandText = QueryC535157.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto =  " & _ 
ConvertValue(P12201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC535157.CommandText = QueryC535157.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P12202, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC535157.Transaction = txn
        C535157 = QueryC535157.ExecuteNonQuery()
        C535157DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C233229

    Comp_C535158:

        ' vBloqPed=1?
        sCurrComponent = "vBloqPed=1?"
        C535158 = (fn_ConvertValueCompiled(C535155, C535155DataType, AKB_DecimalPoint, False) = 1)
        C535158DataType = AKBTypeConst.cAKBTypeLogical
        If C535158 Then
            GoTo Comp_C535156
        Else
            GoTo Comp_C535157
        End If

    Exit_R4484:

        Exit Function

    Err_R4484:
        If Not IsNothing(txn) Then
            txn.Rollback()
        End If
    End Function
    '
    ' Converte um valor String para o tipo apropriado
    '
    Public Shared Function fn_ConvertValueCompiled(ByVal vValue As Object, ByVal lDatatype As Integer, ByVal sDecimalPoint As String, _
                                            Optional bExpression As Boolean = false) As Object

        If Not IsDBNull(vValue) Then

            ' Verifica o tipo do dado
            Select Case lDatatype
                Case AKBTypeConst.cAKBTypeNumeric
                    If CStr(vValue) <> "" Then
                        fn_ConvertValueCompiled = CDbl(vValue)
                    End If
                Case AKBTypeConst.cAKBTypeDate
                    fn_ConvertValueCompiled = CDate(vValue)
                Case AKBTypeConst.cAKBTypeString, AKBTypeConst.cAKBTypeQuery
                    fn_ConvertValueCompiled = vValue
                Case AKBTypeConst.cAKBTypeLogical
                    fn_ConvertValueCompiled = IIf(vValue = 1 Or vValue = True, 1, 0)
                Case Else
                    If CStr(vValue) = "True" Or CStr(vValue) = "False" Then
                        fn_ConvertValueCompiled = IIf(vValue = 1 Or vValue = True, 1, 0)
                    ElseIf IsNumeric(vValue) Then
                        fn_ConvertValueCompiled = CDbl(vValue)
                    ElseIf IsDate(vValue) Then
                        fn_ConvertValueCompiled = CDate(vValue)
                    Else
                        fn_ConvertValueCompiled = vValue
                    End If
            End Select
        Else
            If Not bExpression Then
                fn_ConvertValueCompiled = System.DBNull.Value
            Else

                ' Verifica o tipo do dado
                Select Case lDatatype
                    Case AKBTypeConst.cAKBTypeNumeric
                        fn_ConvertValueCompiled = 0
                    Case AKBTypeConst.cAKBTypeDate
                        fn_ConvertValueCompiled = ""
                    Case AKBTypeConst.cAKBTypeString, AKBTypeConst.cAKBTypeQuery
                        fn_ConvertValueCompiled = ""
                    Case AKBTypeConst.cAKBTypeLogical
                        fn_ConvertValueCompiled = 0
                    Case Else
                        fn_ConvertValueCompiled = ""
                End Select
            End If
        End If
    End Function
    '
    ' Retorna o inteiro maior ou igual ao número
    '
    Public Shared Function fn_Ceiling(ByVal dParam As Double) As Integer

        ' Obtém valor
        If dParam <= 0 Then
            fn_Ceiling = Fix(dParam)
        Else
            If Int(dParam) = dParam Then
                fn_Ceiling = dParam
            Else
                fn_Ceiling = Fix(dParam) + 1
            End If
        End If
    End Function
    '
    ' Retorna o inteiro menor ou igual ao número
    '
    Public Shared Function fn_Floor(ByVal dParam As Double) As Integer

        ' Obtém valor
        If dParam >= 0 Then
            fn_Floor = Fix(dParam)
        Else
            If Int(dParam) = dParam Then
                fn_Floor = dParam
            Else
                fn_Floor = Fix(dParam) - 1
            End If
        End If
    End Function
    Public Shared ReturnDatatype() As Short
    Private Shared iColumns As Short
    Private Shared iCol As Short
    Public Shared Function GetReturnDatatype() As Short()
        GetReturnDatatype = ReturnDatatype
    End Function

    '
    'Retorna o valor por extenso
    '
    Public Shared Function Extenso(valor)

       Dim g(10, 1) As String
       Dim CEN(9) As String
       Dim UNI(9) As String
       Dim DEZ(19) As String
       Dim PARTE(9) As String
       Dim vlr As Double
       Dim Numero$
       Dim p
       Dim i
       Dim TEXTO
       Dim c, D, U
       Dim DE
       Dim VALOR_ANT
       Dim VALOR_POS
       Dim X1, X2, X3
       Dim E1, E2, E3
       Dim GRANDEZA

       If Not IsNumeric(valor) Then
           Extenso = "Valor Não Numerico"
           Exit Function
       End If

       g(1, 0) = "CENTAVO"
       g(2, 0) = ""
       g(3, 0) = "MIL"
       g(4, 0) = "MILHÃO"
       g(5, 0) = "BILHÃO"
       g(6, 0) = "TRILHÃO"

       g(1, 1) = "CENTAVOS"
       g(2, 1) = ""
       g(3, 1) = "MIL"
       g(4, 1) = "MILHÕES"
       g(5, 1) = "BILHÕES"
       g(6, 1) = "TRILHÕES"

       CEN(0) = ""
       CEN(1) = "CENTO"
       CEN(2) = "DUZENTOS"
       CEN(3) = "TREZENTOS"
       CEN(4) = "QUATROCENTOS"
       CEN(5) = "QUINHENTOS"
       CEN(6) = "SEISCENTOS"
       CEN(7) = "SETECENTOS"
       CEN(8) = "OITOCENTOS"
       CEN(9) = "NOVECENTOS"

       DEZ(0) = ""
       DEZ(1) = "DEZ"
       DEZ(2) = "VINTE"
       DEZ(3) = "TRINTA"
       DEZ(4) = "QUARENTA"
       DEZ(5) = "CINQUENTA"
       DEZ(6) = "SESSENTA"
       DEZ(7) = "SETENTA"
       DEZ(8) = "OITENTA"
       DEZ(9) = "NOVENTA"
       DEZ(11) = "ONZE"
       DEZ(12) = "DOZE"
       DEZ(13) = "TREZE"
       DEZ(14) = "QUATORZE"
       DEZ(15) = "QUINZE"
       DEZ(16) = "DEZESSEIS"
       DEZ(17) = "DEZESSETE"
       DEZ(18) = "DEZOITO"
       DEZ(19) = "DEZENOVE"

       UNI(0) = ""
       UNI(1) = "UM"
       UNI(2) = "DOIS"
       UNI(3) = "TRES"
       UNI(4) = "QUATRO"
       UNI(5) = "CINCO"
       UNI(6) = "SEIS"
       UNI(7) = "SETE"
       UNI(8) = "OITO"
       UNI(9) = "NOVE"

       vlr = CDbl(valor)
       Numero$ = Format$(valor, "000000000000000.00")
       p = 1
       For i = 6 To 2 Step -1
          PARTE$(i) = Mid$(Numero$, p, 3)
          p = p + 3
       Next
       PARTE$(1) = 0 & Right$(Numero$, 2)

       TEXTO = ""

       For i = 6 To 1 Step -1
           If Val(PARTE$(i)) <> 0 Then
              c = Val(Left$(PARTE$(i), 1))
              D = Val(Mid$(PARTE$(i), 2, 1))
              U = Val(Mid$(PARTE$(i), 3, 1))
              DE = Val(Mid$(PARTE$(i), 2, 2))
              VALOR_ANT = "0"
              VALOR_POS = "0"
              If i < 6 Then
                 VALOR_ANT = Left$(Numero$, (6 - i) * 3)
              End If
              If i > 2 Then
                 VALOR_POS = Mid$(Numero$, (7 - i) * 3 + 1, (i - 2) * 3)
              End If
              X1 = ""
              X2 = ""
              X3 = ""
              If c = 1 And DE = 0 Then
                   X1 = "CEM"
              Else
                   X1 = CEN(c)
              End If
              If DE < 11 Or DE > 19 Then
                   X2 = DEZ(D)
                   X3 = UNI(U)
              Else
                   X2 = DEZ(DE)
              End If
              'DETERMINA E's
              E1 = " "
              E2 = ""
              E3 = ""
              'Entre as centenas, dezenas e unidades
              If c > 0 And DE > 0 Then
                   E2 = " E "
              End If
              If D > 1 And U > 0 Then
                   E3 = " E "
              End If

              'Entre grandezas diferentes
              If i > 1 And c = 0 And DE <> 0 And Val(VALOR_ANT) <> 0 And Val(VALOR_POS) = 0 Then
                   E2 = " E "
              End If
              If i > 1 And c > 0 And DE = 0 And Val(VALOR_ANT) <> 0 And Val(VALOR_POS) = 0 Then
                   E1 = " E "
              End If

              'Entre a parte inteira e os centavos
              If i = 1 And DE <> 0 And Int(valor) <> 0 Then
                   E1 = " E "
              End If
              If Val(PARTE$(i)) > 1 Then
                   GRANDEZA = g(i, 1)
              Else
                   GRANDEZA = g(i, 0)
              End If
              TEXTO = TEXTO & E1 & X1 & E2 & X2 & E3 & X3 & " " & GRANDEZA
           End If
           If i = 2 Then
               If Int(valor) = 1 Then
                   TEXTO = TEXTO & " REAL"
               Else
                   If PARTE$(3) = "000" And PARTE$(2) = "000" Then
                       TEXTO = TEXTO & " DE REAIS"
                   Else
                       TEXTO = TEXTO & " REAIS"
                   End If
               End If
           End If

       Next
       Extenso = Trim(TEXTO)
    End Function
    '
    ' Validar CGC
    '
    Public Shared Function CalculaCGC(Numero As String) As String

    Dim i As Short
    Dim prod As Short
    Dim mult As Short
    Dim digito As Short

    If Not IsNumeric(Numero) Then
      CalculaCGC = """"
      Exit Function
    End If

    mult = 2
    For i = Len(Numero) To 1 Step -1
     prod = prod + Val(Mid(Numero, i, 1)) * mult
     mult = IIf(mult = 9, 2, mult + 1)
    Next

    digito = 11 - Int(prod Mod 11)
    digito = IIf(digito = 10 Or digito = 11, 0, digito)

    CalculaCGC = Trim(Str(digito))

    End Function
    '
    'Validar CGC
    '
    Public Shared Function ValidaCGC(CGC As String) As Boolean

       'Retira os caracteres de formatação
       CGC = Replace(CGC, ".", "")
       CGC = Replace(CGC, "/", "")
       CGC = Replace(CGC, "-", "")

    If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
      ValidaCGC = False
      Exit Function
    End If

    If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
      ValidaCGC = False
      Exit Function
    End If

    ValidaCGC = True

    End Function

    '
    'Validar CPF
    '
    Public Shared Function calculacpf(CPF As String) As Boolean

       'Retira os caracteres de formatação
       CPF = Replace(CPF, ".", "")
       CPF = Replace(CPF, "/", "")
       CPF = Replace(CPF, "-", "")

       'Esta rotina foi adaptada da revista Fórum Access
       On Error GoTo Err_CPF
       Dim i As Short        'utilizada nos FOR... NEXT
       Dim strcampo As String  'armazena do CPF que será utilizada para o cálculo
       Dim strCaracter As String   'armazena os dígitos do CPF da direita para a esquerda
       Dim intNumero As Short    'armazena o digito separado para cálculo (uma a um)
       Dim intMais As Short  'armazena o digito específico multiplicado pela sua base
       Dim lngSoma As Integer     'armazena a soma dos dígitos multiplicados pela sua base(intmais)
       Dim dblDivisao As Double    'armazena a divisão dos dígitos * base por 11
       Dim lngInteiro As Integer  'armazena inteiro da divisão
       Dim intResto As Short     'armazena o resto
       Dim intDig1 As Short  'armazena o 1º digito verificador
       Dim intDig2 As Short  'armazena o 2º digito verificador
       Dim strConf As String   'armazena o digito verificador

       lngSoma = 0
       intNumero = 0
       intMais = 0
       strcampo = Left(CPF, 9)

       'Inicia cálculos do 1º dígito
       For i = 2 To 10
           strCaracter = Right(strcampo, i - 1)
           intNumero = Left(strCaracter, 1)
           intMais = intNumero * i
           lngSoma = lngSoma + intMais
       Next i
       dblDivisao = lngSoma / 11

       lngInteiro = Int(dblDivisao) * 11
       intResto = lngSoma - lngInteiro
       If intResto = 0 Or intResto = 1 Then
           intDig1 = 0
       Else
           intDig1 = 11 - intResto
       End If

       strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
       lngSoma = 0
       intNumero = 0
       intMais = 0
       'Inicia cálculos do 2º dígito
       For i = 2 To 11
           strCaracter = Right(strcampo, i - 1)
           intNumero = Left(strCaracter, 1)
           intMais = intNumero * i
           lngSoma = lngSoma + intMais
       Next i
       dblDivisao = lngSoma / 11
       lngInteiro = Int(dblDivisao) * 11
       intResto = lngSoma - lngInteiro
       If intResto = 0 Or intResto = 1 Then
           intDig2 = 0
       Else
           intDig2 = 11 - intResto
       End If
       strConf = intDig1 & intDig2
       'Caso o CPF esteja errado dispara a mensagem
       If strConf <> Right(CPF, 2) Then
           calculacpf = False
       Else
           calculacpf = True
       End If
       Exit Function

    Exit_CPF:
       Exit Function
    Err_CPF:
       MsgBox(Err.Description)
       Resume Exit_CPF
    End Function
    '
    'Converte um valor para uma String no formato esperado pelo Banco
    '
    Public Shared Function ConvertValue(ByVal Value As Object, ByVal ValueType As Integer, ByVal DecimalSymbol As String, ByVal DateIdentifier As String, ByVal StringIdentifier As String, ByVal DateFormat as String) As String

       On Error GoTo ConvertValue_Error

       Dim sTemp As String

       If Not IsDBNull(Value) Then

           'Verifica o tipo do dado
           Select Case ValueType

               Case AKBTypeConst.cAKBTypeNumeric

                   'Substitui o ponto pelo separador do banco
                   sTemp = Trim(Str(Value))
                   ConvertValue = Replace(sTemp, ".", DecimalSymbol)

               Case AKBTypeConst.cAKBTypeDate
                   ConvertValue = "TO_DATE(" & DateIdentifier & Format(CDate(Value), DateFormat) & DateIdentifier & ",'dd-mm-yyyy hh24:mi:ss')"

               Case AKBTypeConst.cAKBTypeString
                   ConvertValue = StringIdentifier & Replace(Value, StringIdentifier, StrDup(2, StringIdentifier)) & StringIdentifier

               Case AKBTypeConst.cAKBTypeLogical
                   ConvertValue = IIf(Value, 1, 0)

               Case Else
                   ConvertValue = Value

           End Select
       Else

           'Se for nulo retorna a String "Null"
           ConvertValue = "Null"
       End If

       Exit Function

    ConvertValue_Error:
       MsgBox("Erro ao executar a versão compilada GVINCI da função ConvertValue ")
    End Function
    '
    ' Retorna a data atual do banco de dados
    '
    Public Shared Function GetDate(ByRef con As OracleConnection) As Object
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select sysdate from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            GetDate = Rs(0)
        End If
        Rs.Close()
    End Function
    '
    ' Retorna a data atual do banco de dados
    '
    Public Shared Function ObjTable_NewIdentity(ByRef con As OracleConnection, ByVal TableName as String) As Long
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select SEQ_" & TableName & ".nextval from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            ObjTable_NewIdentity = Rs(0)
        End If
        Rs.Close()
    End Function

End Class
