Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R4625

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

    Public Shared Function R4625(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P12869 As Double, ByVal P12870 As Double, ByVal P12872 As Double, ByVal P12873 As Double, ByVal P20587 As Double) As DataTable
    ' 
    ' 4625 - Verifica Restrições - Versão: 0
    ' 
        On Error GoTo Err_R4625
        Dim sCurrComponent as String

        Dim C55912 As Object
        Dim C55912DataType As Short
        Dim C55913DataType() As Short
        Dim QueryC55916 As New Object
        Dim RsC55916 As New Object
        Dim C55916DataType() As Short
        Dim RsC55916_EOF As Boolean
        Dim C55918 As Object
        Dim C55918DataType As Short
        Dim C55919 As Boolean
        Dim C55919DataType As Short
        Dim C55920 As Object
        Dim C55920DataType As Short
        Dim C55922 As Object
        Dim C55922DataType As Short
        Dim C56015 As Object
        Dim C56015DataType As Short
        Dim C56017 As Object
        Dim C56017DataType As Short
        Dim C56018 As Object
        Dim C56018DataType As Short
        Dim C56021 As Boolean
        Dim C56021DataType As Short
        Dim QueryC56022 As New Object
        Dim C56022 As Integer
        Dim C56022DataType As Short
        Dim C56051 As Object
        Dim C56051DataType As Short
        Dim C56053 As Boolean
        Dim C56053DataType As Short
        Dim C56058 As Object
        Dim C56058DataType As Short
        Dim QueryC56061 As New Object
        Dim C56061 As Integer
        Dim C56061DataType As Short
        Dim QueryC56062 As New Object
        Dim RsC56062 As New Object
        Dim C56062DataType() As Short
        Dim RsC56062_EOF As Boolean
        Dim C56064 As Object
        Dim C56064DataType As Short
        Dim C56065 As Object
        Dim C56065DataType As Short
        Dim C56066 As Boolean
        Dim C56066DataType As Short
        Dim C56067 As Object
        Dim C56067DataType As Short
        Dim C56068 As Object
        Dim C56068DataType As Short
        Dim C56069 As Object
        Dim C56069DataType As Short
        Dim C56071 As Object
        Dim C56071DataType As Short
        Dim QueryC56073 As New Object
        Dim C56073 As Integer
        Dim C56073DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C55912

    Comp_C55912:

        ' VRetorno
        sCurrComponent = "VRetorno"
        C55912 = 1
        C55912DataType = 4
        GoTo Comp_C55916

    Comp_C55913:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C55913 As DataTable = New DataTable()
        tb_C55913.Columns.Add("VRetorno" & "")
        Dim row_C55913 As DataRow = tb_C55913.NewRow()
        row_C55913(0) = C55912
        tb_C55913.Rows.Add(row_C55913)
        R4625 = tb_C55913
        ReDim C55913DataType(1)
        C55913DataType(1) = C55912DataType
        ReturnDataType = C55913DataType

        GoTo Exit_R4625

    Comp_C55916:

        ' SelRestrições
        sCurrComponent = "SelRestrições"
        QueryC55916 = con.CreateCommand()
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "SELECT "
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "WF_RESTRICAO.Id_Restricao , WF_RESTRICAO.Descricao ,  "
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "WF_RESTRICAO.Desconto_Maximo , WF_RESTRICAO.Custo_Financeiro "
        QueryC55916.CommandText = QueryC55916.CommandText & " " & ""
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "FROM AKBUSER01.WF_RESTRICAO , AKBUSER01.WF_RESTRICAO_TP_PRODUTO, AKBUSER01.WF_RESTRICAO_REPRES "
        QueryC55916.CommandText = QueryC55916.CommandText & " " & ""
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "WHERE WF_RESTRICAO.Id_Restricao = WF_RESTRICAO_TP_PRODUTO.Id_Restricao "
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "AND WF_RESTRICAO_REPRES.Id_Restricao = WF_RESTRICAO_TP_PRODUTO.Id_Restricao "
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "AND WF_RESTRICAO_REPRES.Cod_Repres = " & _ 
ConvertValue(P20587, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "AND  WF_RESTRICAO_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P12870, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC55916.CommandText = QueryC55916.CommandText & " " & "AND  WF_RESTRICAO.Ativa = 1"
        QueryC55916.CommandText = QueryC55916.CommandText & " " & ""
        QueryC55916.CommandText = QueryC55916.CommandText & " " & ""
        QueryC55916.Transaction = txn
        RsC55916 = QueryC55916.ExecuteReader()
        Dim iC55916 As Short
        ReDim C55916DataType(RsC55916.FieldCount)
        For iC55916 = 0 to RsC55916.FieldCount - 1
            Select Case RsC55916.GetDataTypeName(iC55916).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C55916DataType(iC55916 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C55916DataType(iC55916 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C55916DataType(iC55916 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC55916
        RsC55916_EOF = Not RsC55916.Read()

        GoTo Comp_C55918

    Comp_C55918:

        ' Eof
        sCurrComponent = "Eof"
        C55918DataType = 4
        C55918 = RsC55916_EOF
        GoTo Comp_C55919

    Comp_C55919:

        ' Eof=T
        sCurrComponent = "Eof=T"
        C55919 = (fn_ConvertValueCompiled(C55918, C55918DataType, AKB_DecimalPoint, False) = 1)
        C55919DataType = AKBTypeConst.cAKBTypeLogical
        If C55919 Then
            GoTo Comp_C55913
        Else
            GoTo Comp_C55922
        End If

    Comp_C55920:

        ' Next
        sCurrComponent = "Next"
        C55920DataType = 4
        RsC55916.Read()
        C55920 = True

        GoTo Comp_C55918

    Comp_C55922:

        ' IdRestrição
        sCurrComponent = "IdRestrição"
        C55922DataType = 0
        C55922 = RsC55916(0)
        C55922DataType = C55916Datatype(1)

        GoTo Comp_C56015

    Comp_C56015:

        ' Descrição
        sCurrComponent = "Descrição"
        C56015DataType = 0
        C56015 = RsC55916(1)
        C56015DataType = C55916Datatype(2)

        GoTo Comp_C56017

    Comp_C56017:

        ' DescontoMáximo
        sCurrComponent = "DescontoMáximo"
        C56017DataType = 0
        C56017 = RsC55916(2)
        C56017DataType = C55916Datatype(3)

        GoTo Comp_C56018

    Comp_C56018:

        ' CustoMáximo
        sCurrComponent = "CustoMáximo"
        C56018DataType = 0
        C56018 = RsC55916(3)
        C56018DataType = C55916Datatype(4)

        GoTo Comp_C56021

    Comp_C56021:

        ' ErrDesconto
        sCurrComponent = "ErrDesconto"
        C56021 = (fn_ConvertValueCompiled(P12872, 1, AKB_DecimalPoint, False) > fn_ConvertValueCompiled(C56017, C56017DataType, AKB_DecimalPoint, False))
        C56021DataType = AKBTypeConst.cAKBTypeLogical
        If C56021 Then
            GoTo Comp_C56051
        Else
            GoTo Comp_C56053
        End If

    Comp_C56022:

        ' LogDesconto
        sCurrComponent = "LogDesconto"
        QueryC56022 = con.CreateCommand()
        QueryC56022.CommandText = QueryC56022.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO "
        QueryC56022.CommandText = QueryC56022.CommandText & " " & ""
        QueryC56022.CommandText = QueryC56022.CommandText & " " & "SET  WF_PRE_PEDIDO.Error_Log = WF_PRE_PEDIDO.Error_Log || 'Restrição Desconto: ' || " & _ 
ConvertValue(P12872, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56022.CommandText = QueryC56022.CommandText & " " & ""
        QueryC56022.CommandText = QueryC56022.CommandText & " " & "WHERE  WF_PRE_PEDIDO.Id_PrePedido  = " & _ 
ConvertValue(P12869, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56022.CommandText = QueryC56022.CommandText & " " & ""
        QueryC56022.CommandText = QueryC56022.CommandText & " " & ""
        C56022 = QueryC56022.ExecuteNonQuery()
        C56022DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56053

    Comp_C56051:

        ' AttFalse1
        sCurrComponent = "AttFalse1"
        C56051DataType = 4
        C55912 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C56051 = True
        GoTo Comp_C56022

    Comp_C56053:

        ' ErrCusto
        sCurrComponent = "ErrCusto"
        C56053 = (fn_ConvertValueCompiled(P12873, 1, AKB_DecimalPoint, False) > fn_ConvertValueCompiled(C56018, C56018DataType, AKB_DecimalPoint, False))
        C56053DataType = AKBTypeConst.cAKBTypeLogical
        If C56053 Then
            GoTo Comp_C56058
        Else
            GoTo Comp_C56062
        End If

    Comp_C56058:

        ' AttFalse2
        sCurrComponent = "AttFalse2"
        C56058DataType = 4
        C55912 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C56058 = True
        GoTo Comp_C56061

    Comp_C56061:

        ' LogCusto
        sCurrComponent = "LogCusto"
        QueryC56061 = con.CreateCommand()
        QueryC56061.CommandText = QueryC56061.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO "
        QueryC56061.CommandText = QueryC56061.CommandText & " " & ""
        QueryC56061.CommandText = QueryC56061.CommandText & " " & "SET  WF_PRE_PEDIDO.Error_Log = WF_PRE_PEDIDO.Error_Log || 'Restrição Custo Financeiro: ' || " & _ 
ConvertValue(P12873, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56061.CommandText = QueryC56061.CommandText & " " & ""
        QueryC56061.CommandText = QueryC56061.CommandText & " " & "WHERE  WF_PRE_PEDIDO.Id_PrePedido  = " & _ 
ConvertValue(P12869, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56061.CommandText = QueryC56061.CommandText & " " & ""
        QueryC56061.CommandText = QueryC56061.CommandText & " " & ""
        C56061 = QueryC56061.ExecuteNonQuery()
        C56061DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56062

    Comp_C56062:

        ' SelDescontos
        sCurrComponent = "SelDescontos"
        QueryC56062 = con.CreateCommand()
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "SELECT"
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "WF_PRE_PED_ITENS_DESCONTO.Seq_Item , "
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "WF_PRE_PED_ITENS_DESCONTO.Tipo_Desconto , "
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "WF_PRE_PED_ITENS_DESCONTO.PORC_DESCONTO "
        QueryC56062.CommandText = QueryC56062.CommandText & " " & ""
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "FROM  AKBUSER01.WF_PRE_PED_ITENS_DESCONTO , AKBUSER01.WF_RESTRICAO_DESCONTOS "
        QueryC56062.CommandText = QueryC56062.CommandText & " " & ""
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "WHERE WF_RESTRICAO_DESCONTOS.Id_Restricao = " & _ 
ConvertValue(C55922, C55922DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "AND WF_PRE_PED_ITENS_DESCONTO.Id_PrePedido = " & _ 
ConvertValue(P12869, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56062.CommandText = QueryC56062.CommandText & " " & ""
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "AND WF_RESTRICAO_DESCONTOS.Tipo_Desconto = WF_PRE_PED_ITENS_DESCONTO.Tipo_Desconto "
        QueryC56062.CommandText = QueryC56062.CommandText & " " & ""
        QueryC56062.CommandText = QueryC56062.CommandText & " " & "AND WF_PRE_PED_ITENS_DESCONTO.PORC_DESCONTO > WF_RESTRICAO_DESCONTOS.Porc_Maximo "
        QueryC56062.CommandText = QueryC56062.CommandText & " " & ""
        QueryC56062.Transaction = txn
        RsC56062 = QueryC56062.ExecuteReader()
        Dim iC56062 As Short
        ReDim C56062DataType(RsC56062.FieldCount)
        For iC56062 = 0 to RsC56062.FieldCount - 1
            Select Case RsC56062.GetDataTypeName(iC56062).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56062DataType(iC56062 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56062DataType(iC56062 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56062DataType(iC56062 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56062
        RsC56062_EOF = Not RsC56062.Read()

        GoTo Comp_C56064

    Comp_C56064:

        ' Eof2
        sCurrComponent = "Eof2"
        C56064DataType = 4
        C56064 = RsC56062_EOF
        GoTo Comp_C56066

    Comp_C56065:

        ' Next2
        sCurrComponent = "Next2"
        C56065DataType = 4
        RsC56062.Read()
        C56065 = True

        GoTo Comp_C56064

    Comp_C56066:

        ' Eof2=1
        sCurrComponent = "Eof2=1"
        C56066 = (fn_ConvertValueCompiled(C56064, C56064DataType, AKB_DecimalPoint, False) = 1)
        C56066DataType = AKBTypeConst.cAKBTypeLogical
        If C56066 Then
            GoTo Comp_C55920
        Else
            GoTo Comp_C56071
        End If

    Comp_C56067:

        ' Item
        sCurrComponent = "Item"
        C56067DataType = 0
        C56067 = RsC56062(0)
        C56067DataType = C56062Datatype(1)

        GoTo Comp_C56069

    Comp_C56068:

        ' %Desconto
        sCurrComponent = "%Desconto"
        C56068DataType = 0
        C56068 = RsC56062(2)
        C56068DataType = C56062Datatype(3)

        GoTo Comp_C56073

    Comp_C56069:

        ' TpDesconto
        sCurrComponent = "TpDesconto"
        C56069DataType = 0
        C56069 = RsC56062(1)
        C56069DataType = C56062Datatype(2)

        GoTo Comp_C56068

    Comp_C56071:

        ' AttFalse3
        sCurrComponent = "AttFalse3"
        C56071DataType = 4
        C55912 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C56071 = True
        GoTo Comp_C56067

    Comp_C56073:

        ' LogDescontoItem
        sCurrComponent = "LogDescontoItem"
        QueryC56073 = con.CreateCommand()
        QueryC56073.CommandText = QueryC56073.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO "
        QueryC56073.CommandText = QueryC56073.CommandText & " " & ""
        QueryC56073.CommandText = QueryC56073.CommandText & " " & "SET  WF_PRE_PEDIDO.Error_Log = WF_PRE_PEDIDO.Error_Log || 'Restrição Desconto Item: ' || " & _ 
ConvertValue(C56067, C56067DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " || ' Tipo: ' || " & _ 
ConvertValue(C56069, C56069DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ||"
        QueryC56073.CommandText = QueryC56073.CommandText & " " & "   ' Desc: ' || " & _ 
ConvertValue(C56068, C56068DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56073.CommandText = QueryC56073.CommandText & " " & ""
        QueryC56073.CommandText = QueryC56073.CommandText & " " & "WHERE  WF_PRE_PEDIDO.Id_PrePedido  = " & _ 
ConvertValue(P12869, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56073.CommandText = QueryC56073.CommandText & " " & ""
        QueryC56073.CommandText = QueryC56073.CommandText & " " & ""
        C56073 = QueryC56073.ExecuteNonQuery()
        C56073DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56065

    Exit_R4625:

        Exit Function

    Err_R4625:
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
