Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R3125

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

    Public Shared Function R3125(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P7290 As Double, ByVal P7291 As Double) As DataTable
    ' 
    ' 3125 - Verifica Itens Duplicados na Digitação do Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R3125
        Dim sCurrComponent as String

        Dim QueryC29457 As New Object
        Dim RsC29457 As New Object
        Dim C29457DataType() As Short
        Dim RsC29457_EOF As Boolean
        Dim C29458 As Boolean
        Dim C29458DataType As Short
        Dim C29459 As Object
        Dim C29459DataType As Short
        Dim C29460DataType() As Short
        Dim C29461 As Short
        Dim C29461DataType As Short
        Dim C29461ReturnDataType() As Short

        Dim C29463 As Object
        Dim C29463DataType As Short
        Dim QueryC82517 As New Object
        Dim RsC82517 As New Object
        Dim C82517DataType() As Short
        Dim RsC82517_EOF As Boolean
        Dim C82519 As Object
        Dim C82519DataType As Short
        Dim C82520 As Object
        Dim C82520DataType As Short
        Dim C82521 As Object
        Dim C82521DataType As Short
        Dim C205141 As Object
        Dim C205141DataType As Short
        Dim C205321 As Object
        Dim C205321DataType As Short
        Dim QueryC256596 As New Object
        Dim RsC256596 As New Object
        Dim C256596DataType() As Short
        Dim RsC256596_EOF As Boolean
        Dim C256597 As Boolean
        Dim C256597DataType As Short
        Dim C256598 As Short
        Dim C256598DataType As Short
        Dim C256598ReturnDataType() As Short

        Dim C296995 As Object
        Dim C296995DataType As Short
        Dim C296996 As Boolean
        Dim C296996DataType As Short
        Dim C411793 As Object
        Dim C411793DataType As Short
        Dim C411794 As Object
        Dim C411794DataType As Short
        Dim C411795 As Boolean
        Dim C411795DataType As Short
        Dim C411796 As Object
        Dim C411796DataType As Short
        Dim C411998 As Object
        Dim C411998DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C29459

    Comp_C29457:

        ' SelItens
        sCurrComponent = "SelItens"
        QueryC29457 = con.CreateCommand()
        QueryC29457.CommandText = QueryC29457.CommandText & " " & "SELECT COUNT(*) "
        QueryC29457.CommandText = QueryC29457.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS "
        QueryC29457.CommandText = QueryC29457.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P7290, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC29457.CommandText = QueryC29457.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P7291, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC29457.CommandText = QueryC29457.CommandText & " " & "AND WF_PEDIDO_ITENS.Flag_Pos_Ped NOT IN (7,8,10,11,12,13,299 )   "
        QueryC29457.CommandText = QueryC29457.CommandText & " " & "" & _ 
ConvertValue(C411794, C411794DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC29457.Transaction = txn
        RsC29457 = QueryC29457.ExecuteReader()
        Dim iC29457 As Short
        ReDim C29457DataType(RsC29457.FieldCount)
        For iC29457 = 0 to RsC29457.FieldCount - 1
            Select Case RsC29457.GetDataTypeName(iC29457).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C29457DataType(iC29457 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C29457DataType(iC29457 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C29457DataType(iC29457 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC29457
        RsC29457_EOF = Not RsC29457.Read()

        GoTo Comp_C296995

    Comp_C29458:

        ' SelItens=0
        sCurrComponent = "SelItens=0"
        C29458 = (fn_ConvertValueCompiled(RsC29457(0), C29457DataType(1), AKB_DecimalPoint, False) =0)
        C29458DataType = AKBTypeConst.cAKBTypeLogical
        If C29458 Then
            GoTo Comp_C29460
        Else
            GoTo Comp_C82517
        End If

    Comp_C29459:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C29459 = 1
        C29459DataType = 4
        GoTo Comp_C411794

    Comp_C29460:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C29460 As DataTable = New DataTable()
        tb_C29460.Columns.Add("vRetorno" & "")
        Dim row_C29460 As DataRow = tb_C29460.NewRow()
        row_C29460(0) = C29459
        tb_C29460.Rows.Add(row_C29460)
        R3125 = tb_C29460
        ReDim C29460DataType(1)
        C29460DataType(1) = C29459DataType
        ReturnDataType = C29460DataType

        GoTo Exit_R3125

    Comp_C29461:

        ' MsgConfirma
        sCurrComponent = "MsgConfirma"
        C29461 = 1
        C29461DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C29463

    Comp_C29463:

        ' AtrFalseParaRetorno
        sCurrComponent = "AtrFalseParaRetorno"
        C29463DataType = 4
        C29459 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C29463 = True
        GoTo Comp_C29460

    Comp_C82517:

        ' Qry_01
        sCurrComponent = "Qry_01"
        QueryC82517 = con.CreateCommand()
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "SELECT WF_PEDIDO_ITENS.Sigla_Prod , "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WF_PEDIDO_ITENS.Acesso , "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WF_PEDIDO_ITENS.Cod_Embalagem , "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "COUNT(*), "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WF_PEDIDO_ITENS.Prazo_Entr_Cliente ,"
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WF_PEDIDO_ITENS.Tipo_Ped"
        QueryC82517.CommandText = QueryC82517.CommandText & " " & ""
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & ""
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P7290, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P7291, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "AND WF_PEDIDO_ITENS.Flag_Pos_Ped NOT IN (7,8,10,11,12,13,299 )   "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "GROUP BY WF_PEDIDO_ITENS.Sigla_Prod , "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WF_PEDIDO_ITENS.Acesso , "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WF_PEDIDO_ITENS.Cod_Embalagem , "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WF_PEDIDO_ITENS.Tipo_Ped, "
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "WF_PEDIDO_ITENS.Prazo_Entr_Cliente"
        QueryC82517.CommandText = QueryC82517.CommandText & " " & "HAVING COUNT(*) > 1"
        QueryC82517.Transaction = txn
        RsC82517 = QueryC82517.ExecuteReader()
        Dim iC82517 As Short
        ReDim C82517DataType(RsC82517.FieldCount)
        For iC82517 = 0 to RsC82517.FieldCount - 1
            Select Case RsC82517.GetDataTypeName(iC82517).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82517DataType(iC82517 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82517DataType(iC82517 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82517DataType(iC82517 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82517
        RsC82517_EOF = Not RsC82517.Read()

        GoTo Comp_C82519

    Comp_C82519:

        ' Sigla_Prod
        sCurrComponent = "Sigla_Prod"
        C82519DataType = 0
        C82519 = RsC82517(0)
        C82519DataType = C82517Datatype(1)
        If C82519DataType = AKBTypeConst.cAKBTypeString Then
          C82519 = IIF(IsDBNull(C82519), C82519, Strings.RTrim(C82519))
        End If 

        GoTo Comp_C82520

    Comp_C82520:

        ' Acesso
        sCurrComponent = "Acesso"
        C82520DataType = 0
        C82520DataType = C82517Datatype(2)
        If C82520DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82517(1)) Then
          C82520 = Strings.RTrim(RsC82517(1))
        Else 
          C82520 = RsC82517(1)
        End If 

        GoTo Comp_C82521

    Comp_C82521:

        ' Cod_Embalagem
        sCurrComponent = "Cod_Embalagem"
        C82521DataType = 0
        C82521DataType = C82517Datatype(3)
        If C82521DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82517(2)) Then
          C82521 = Strings.RTrim(RsC82517(2))
        Else 
          C82521 = RsC82517(2)
        End If 

        GoTo Comp_C205141

    Comp_C205141:

        ' Prazo_entrega
        sCurrComponent = "Prazo_entrega"
        C205141DataType = 0
        C205141DataType = C82517Datatype(5)
        If C205141DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82517(4)) Then
          C205141 = Strings.RTrim(RsC82517(4))
        Else 
          C205141 = RsC82517(4)
        End If 

        GoTo Comp_C205321

    Comp_C205321:

        ' TpPedido
        sCurrComponent = "TpPedido"
        C205321DataType = 0
        C205321DataType = C82517Datatype(6)
        If C205321DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82517(5)) Then
          C205321 = Strings.RTrim(RsC82517(5))
        Else 
          C205321 = RsC82517(5)
        End If 

        GoTo Comp_C29461

    Comp_C256596:

        ' FlagPed
        sCurrComponent = "FlagPed"
        QueryC256596 = con.CreateCommand()
        QueryC256596.CommandText = QueryC256596.CommandText & " " & "SELECT WF_PEDIDO.Flag_Pos_Ped ,"
        QueryC256596.CommandText = QueryC256596.CommandText & " " & "WF_TP_PRODUTO.Tp_Tab_Preco ,"
        QueryC256596.CommandText = QueryC256596.CommandText & " " & "NVL(WF_TP_PRODUTO.Verifica_Prazo_Ent_Cliente, 0) "
        QueryC256596.CommandText = QueryC256596.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO , AKBUSER01.WF_TP_PRODUTO "
        QueryC256596.CommandText = QueryC256596.CommandText & " " & "WHERE  WF_PEDIDO.Tp_Produto = WF_TP_PRODUTO.Tp_Produto "
        QueryC256596.CommandText = QueryC256596.CommandText & " " & "AND WF_PEDIDO.Pedido = " & _ 
ConvertValue(P7291, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC256596.CommandText = QueryC256596.CommandText & " " & "AND WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P7290, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC256596.Transaction = txn
        RsC256596 = QueryC256596.ExecuteReader()
        Dim iC256596 As Short
        ReDim C256596DataType(RsC256596.FieldCount)
        For iC256596 = 0 to RsC256596.FieldCount - 1
            Select Case RsC256596.GetDataTypeName(iC256596).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C256596DataType(iC256596 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C256596DataType(iC256596 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C256596DataType(iC256596 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC256596
        RsC256596_EOF = Not RsC256596.Read()

        GoTo Comp_C256597

    Comp_C256597:

        ' FlagPed?
        sCurrComponent = "FlagPed?"
        C256597 = (( fn_ConvertValueCompiled(RsC256596(0), C256596DataType(1), AKB_DecimalPoint, False) = 6 ) or ( fn_ConvertValueCompiled(RsC256596(0), C256596DataType(1), AKB_DecimalPoint, False) = 453 ) or ( fn_ConvertValueCompiled(RsC256596(0), C256596DataType(1), AKB_DecimalPoint, False) = 7 ) or ( fn_ConvertValueCompiled(RsC256596(0), C256596DataType(1), AKB_DecimalPoint, False) = 8 )   or ( fn_ConvertValueCompiled(RsC256596(0), C256596DataType(1), AKB_DecimalPoint, False) = 11 ) or ( fn_ConvertValueCompiled(RsC256596(0), C256596DataType(1), AKB_DecimalPoint, False) = 12 ) or ( fn_ConvertValueCompiled(RsC256596(0), C256596DataType(1), AKB_DecimalPoint, False) = 9 ))
        C256597DataType = AKBTypeConst.cAKBTypeLogical
        If C256597 Then
            GoTo Comp_C256598
        Else
            GoTo Comp_C411793
        End If

    Comp_C256598:

        ' MsgFlag
        sCurrComponent = "MsgFlag"
        C256598 = 1
        C256598DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C29463

    Comp_C296995:

        ' EOF Itens
        sCurrComponent = "EOF Itens"
        C296995DataType = 4
        C296995 = RsC29457_EOF
        GoTo Comp_C296996

    Comp_C296996:

        ' Tem Itens?
        sCurrComponent = "Tem Itens?"
        C296996 = (fn_ConvertValueCompiled(C296995, C296995DataType, AKB_DecimalPoint, False) = 0)
        C296996DataType = AKBTypeConst.cAKBTypeLogical
        If C296996 Then
            GoTo Comp_C29458
        Else
            GoTo Comp_C29460
        End If

    Comp_C411793:

        ' Tipo_Faturamento
        sCurrComponent = "Tipo_Faturamento"
        C411793DataType = 0
        C411793DataType = C256596Datatype(2)
        If C411793DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC256596(1)) Then
          C411793 = Strings.RTrim(RsC256596(1))
        Else 
          C411793 = RsC256596(1)
        End If 

        GoTo Comp_C411998

    Comp_C411794:

        ' vGroupBy
        sCurrComponent = "vGroupBy"
        C411794 = "GROUP BY WF_PEDIDO_ITENS.Sigla_Prod ,   WF_PEDIDO_ITENS.Acesso ,   WF_PEDIDO_ITENS.Cod_Embalagem ,   WF_PEDIDO_ITENS.Tipo_Ped,  WF_PEDIDO_ITENS.Prazo_Entr_Cliente  HAVING COUNT(1) > 1"
        C411794DataType = 5
        GoTo Comp_C256596

    Comp_C411795:

        ' Tipo_Faturamento <> 'Interno'?
        sCurrComponent = "Tipo_Faturamento <> 'Interno'?"
        C411795 = (fn_ConvertValueCompiled(C411998, C411998DataType, AKB_DecimalPoint, False) = 1 AND fn_ConvertValueCompiled(C411793, C411793DataType, AKB_DecimalPoint, False) = "Interno")
        C411795DataType = AKBTypeConst.cAKBTypeLogical
        If C411795 Then
            GoTo Comp_C411796
        Else
            GoTo Comp_C29457
        End If

    Comp_C411796:

        ' AtrGroupBy
        sCurrComponent = "AtrGroupBy"
        C411796DataType = 4
        C411794 = fn_ConvertValueCompiled("GROUP BY WF_PEDIDO_ITENS.Sigla_Prod ,   WF_PEDIDO_ITENS.Acesso ,   WF_PEDIDO_ITENS.Cod_Embalagem ,   WF_PEDIDO_ITENS.Tipo_Ped  HAVING COUNT(1) > 1", 5, AKB_DecimalPoint)
        C411796 = True
        GoTo Comp_C29457

    Comp_C411998:

        ' Verifica_Prazo_Entrega_Cliente
        sCurrComponent = "Verifica_Prazo_Entrega_Cliente"
        C411998DataType = 0
        C411998DataType = C256596Datatype(3)
        If C411998DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC256596(2)) Then
          C411998 = Strings.RTrim(RsC256596(2))
        Else 
          C411998 = RsC256596(2)
        End If 

        GoTo Comp_C411795

    Exit_R3125:

        Exit Function

    Err_R3125:
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
                    fn_ConvertValueCompiled = IIf(CBool(vValue) = 1 Or CBool(vValue) = True, 1, 0)
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
    Public Shared Function GetDate(ByRef con As OracleConnection, ByRef txn As OracleTransaction) As Object
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select sysdate from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Query.Transaction = txn
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
