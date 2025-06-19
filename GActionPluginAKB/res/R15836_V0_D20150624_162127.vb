Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R15836

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

    Public Shared Function R15836(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P56206 As Object, ByVal P58093 As Object) As DataTable
    ' 
    ' 15836 - Verificação/Bloqueio do Pré Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R15836
        Dim sCurrComponent as String

        Dim QueryC338494 As New Object
        Dim RsC338494 As New Object
        Dim C338494DataType() As Short
        Dim RsC338494_EOF As Boolean
        Dim C338498 As Boolean
        Dim C338498DataType As Short
        Dim C338501DataType() As Short
        Dim QueryC338505 As New Object
        Dim RsC338505 As New Object
        Dim C338505DataType() As Short
        Dim RsC338505_EOF As Boolean
        Dim C338511 As Boolean
        Dim C338511DataType As Short
        Dim QueryC338516 As New Object
        Dim RsC338516 As New Object
        Dim C338516DataType() As Short
        Dim RsC338516_EOF As Boolean
        Dim C338854 As Boolean
        Dim C338854DataType As Short
        Dim QueryC338903 As New Object
        Dim RsC338903 As New Object
        Dim C338903DataType() As Short
        Dim RsC338903_EOF As Boolean
        Dim C338906 As Boolean
        Dim C338906DataType As Short
        Dim QueryC338907 As New Object
        Dim RsC338907 As New Object
        Dim C338907DataType() As Short
        Dim RsC338907_EOF As Boolean
        Dim QueryC339066 As New Object
        Dim RsC339066 As New Object
        Dim C339066DataType() As Short
        Dim RsC339066_EOF As Boolean
        Dim C339069 As Boolean
        Dim C339069DataType As Short
        Dim C339070 As Object
        Dim C339070DataType As Short
        Dim QueryC339200 As New Object
        Dim RsC339200 As New Object
        Dim C339200DataType() As Short
        Dim RsC339200_EOF As Boolean
        Dim C339201 As Boolean
        Dim C339201DataType As Short
        Dim QueryC339203 As New Object
        Dim C339203 As Integer
        Dim C339203DataType As Short
        Dim C339204 As Object
        Dim C339204DataType As Short
        Dim C339206 As Object
        Dim C339206DataType As Short
        Dim C339207 As Object
        Dim C339207DataType As Short
        Dim C339208 As Object
        Dim C339208DataType As Short
        Dim C339209 As Object
        Dim C339209DataType As Short
        Dim C339210 As Object
        Dim C339210DataType As Short
        Dim C339211 As Object
        Dim C339211DataType As Short
        Dim C339212 As Object
        Dim C339212DataType As Short
        Dim C339213 As Object
        Dim C339213DataType As Short
        Dim C353432 As Object
        Dim C353432DataType As Short
        Dim C353433 As Boolean
        Dim C353433DataType As Short
        Dim C353435 As Object
        Dim C353435DataType As Short
        Dim C353436 As Object
        Dim C353436DataType As Short
        Dim C353451 As Object
        Dim C353451DataType As Short
        Dim C353458( , ) As Object
        Dim C353458DataType() As Short
        Dim C353458CurrentRow As Integer
        Dim C353460 As Object
        Dim C353460DataType As Short
        Dim C353461 As Object
        Dim C353461DataType As Short
        Dim C353462 As Object
        Dim C353462DataType As Short
        Dim C353463 As Object
        Dim C353463DataType As Short
        Dim C353467 As Object
        Dim C353467DataType As Short
        Dim C353474 As Object
        Dim C353474DataType As Short
        Dim C353475 As Object
        Dim C353475DataType As Short
        Dim C353476 As Object
        Dim C353476DataType As Short
        Dim C353477 As Boolean
        Dim C353477DataType As Short
        Dim C353478 As Object
        Dim C353478DataType As Short
        Dim C353482 As Object
        Dim C353482DataType As Short
        Dim C353580 As Object
        Dim C353580DataType As Short
        Dim C357575 As Boolean
        Dim C357575DataType As Short
        Dim QueryC357859 As New Object
        Dim RsC357859 As New Object
        Dim C357859DataType() As Short
        Dim RsC357859_EOF As Boolean
        Dim C357864 As Boolean
        Dim C357864DataType As Short
        Dim C357865 As Object
        Dim C357865DataType As Short
        Dim C357866 As Object
        Dim C357866DataType As Short
        Dim QueryC358845 As New Object
        Dim RsC358845 As New Object
        Dim C358845DataType() As Short
        Dim RsC358845_EOF As Boolean
        Dim QueryC358847 As New Object
        Dim RsC358847 As New Object
        Dim C358847DataType() As Short
        Dim RsC358847_EOF As Boolean
        Dim C358848 As Object
        Dim C358848DataType As Short
        Dim C358849 As Boolean
        Dim C358849DataType As Short
        Dim C358850 As Boolean
        Dim C358850DataType As Short
        Dim C358851 As Object
        Dim C358851DataType As Short
        Dim QueryC437817 As New Object
        Dim RsC437817 As New Object
        Dim C437817DataType() As Short
        Dim RsC437817_EOF As Boolean
        Dim C437818 As Boolean
        Dim C437818DataType As Short
        Dim QueryC443935 As New Object
        Dim RsC443935 As New Object
        Dim C443935DataType() As Short
        Dim RsC443935_EOF As Boolean
        Dim C443936 As Boolean
        Dim C443936DataType As Short
        Dim C443937 As Object
        Dim C443937DataType As Short
        Dim C443938 As Object
        Dim C443938DataType As Short
        Dim QueryC443939 As New Object
        Dim RsC443939 As New Object
        Dim C443939DataType() As Short
        Dim RsC443939_EOF As Boolean
        Dim C443940 As Boolean
        Dim C443940DataType As Short
        Dim C443941 As Object
        Dim C443941DataType As Short
        Dim C443942 As Object
        Dim C443942DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C353458

    Comp_C338494:

        ' SelPreCadastro
        sCurrComponent = "SelPreCadastro"
        QueryC338494 = con.CreateCommand()
        QueryC338494.CommandText = QueryC338494.CommandText & " " & "SELECT COUNT(WF_CLIENTES.Pre_Cadastro)"
        QueryC338494.CommandText = QueryC338494.CommandText & " " & "FROM  AKBUSER01.WF_CLIENTES "
        QueryC338494.CommandText = QueryC338494.CommandText & " " & "WHERE WF_CLIENTES.Pre_Cadastro = 1"
        QueryC338494.CommandText = QueryC338494.CommandText & " " & "AND WF_CLIENTES.Pre_Cadastro_Liberado = 0 "
        QueryC338494.CommandText = QueryC338494.CommandText & " " & "AND WF_CLIENTES.Cod_Cliente = (SELECT WF_PRE_PEDIDO.Cod_Cliente "
        QueryC338494.CommandText = QueryC338494.CommandText & " " & "                                                                                           FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC338494.CommandText = QueryC338494.CommandText & " " & "                                                                                           WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC338494.Transaction = txn
        RsC338494 = QueryC338494.ExecuteReader()
        Dim iC338494 As Short
        ReDim C338494DataType(RsC338494.FieldCount)
        For iC338494 = 0 to RsC338494.FieldCount - 1
            Select Case RsC338494.GetDataTypeName(iC338494).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C338494DataType(iC338494 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C338494DataType(iC338494 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C338494DataType(iC338494 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC338494
        RsC338494_EOF = Not RsC338494.Read()

        GoTo Comp_C338498

    Comp_C338498:

        ' Pré_Cadastro?
        sCurrComponent = "Pré_Cadastro?"
        C338498 = (fn_ConvertValueCompiled(RsC338494(0), C338494DataType(1), AKB_DecimalPoint, False) =1)
        C338498DataType = AKBTypeConst.cAKBTypeLogical
        If C338498 Then
            GoTo Comp_C353460
        Else
            GoTo Comp_C338505
        End If

    Comp_C338501:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C338501 As DataTable = New DataTable()
        tb_C338501.Columns.Add("vRet" & "")
        Dim row_C338501 As DataRow = tb_C338501.NewRow()
        row_C338501(0) = C339070
        tb_C338501.Rows.Add(row_C338501)
        R15836 = tb_C338501
        ReDim C338501DataType(1)
        C338501DataType(1) = C339070DataType
        ReturnDataType = C338501DataType

        GoTo Exit_R15836

    Comp_C338505:

        ' SelTransportadora
        sCurrComponent = "SelTransportadora"
        QueryC338505 = con.CreateCommand()
        QueryC338505.CommandText = QueryC338505.CommandText & " " & "SELECT NVL(WF_PRE_PEDIDO.Cod_Transp,0) "
        QueryC338505.CommandText = QueryC338505.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC338505.CommandText = QueryC338505.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC338505.Transaction = txn
        RsC338505 = QueryC338505.ExecuteReader()
        Dim iC338505 As Short
        ReDim C338505DataType(RsC338505.FieldCount)
        For iC338505 = 0 to RsC338505.FieldCount - 1
            Select Case RsC338505.GetDataTypeName(iC338505).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C338505DataType(iC338505 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C338505DataType(iC338505 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C338505DataType(iC338505 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC338505
        RsC338505_EOF = Not RsC338505.Read()

        GoTo Comp_C338511

    Comp_C338511:

        ' Transportadora?
        sCurrComponent = "Transportadora?"
        C338511 = (fn_ConvertValueCompiled(RsC338505(0), C338505DataType(1), AKB_DecimalPoint, False) <> 0)
        C338511DataType = AKBTypeConst.cAKBTypeLogical
        If C338511 Then
            GoTo Comp_C338516
        Else
            GoTo Comp_C353461
        End If

    Comp_C338516:

        ' SelForaDeLinha
        sCurrComponent = "SelForaDeLinha"
        QueryC338516 = con.CreateCommand()
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "SELECT COUNT(*)"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_ITENS, AKBUSER01.WF_TIPO_PED"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Tipo_Ped = WF_TIPO_PED.Tipo_Ped"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "AND WF_TIPO_PED.Pronta_Entrega = 0"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "AND EXISTS"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "(SELECT *"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS, AKBUSER01.WF_COLECAO"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Id_Colecao = WF_COLECAO.Id_Colecao"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC338516.CommandText = QueryC338516.CommandText & " " & "AND (WF_COLECAO_PRODUTOS.Data_Validade_Final < SYSDATE OR WF_COLECAO.Data_Validade_Final < SYSDATE)) "
        QueryC338516.Transaction = txn
        RsC338516 = QueryC338516.ExecuteReader()
        Dim iC338516 As Short
        ReDim C338516DataType(RsC338516.FieldCount)
        For iC338516 = 0 to RsC338516.FieldCount - 1
            Select Case RsC338516.GetDataTypeName(iC338516).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C338516DataType(iC338516 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C338516DataType(iC338516 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C338516DataType(iC338516 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC338516
        RsC338516_EOF = Not RsC338516.Read()

        GoTo Comp_C338854

    Comp_C338854:

        ' Fora de Linha?
        sCurrComponent = "Fora de Linha?"
        C338854 = (fn_ConvertValueCompiled(RsC338516(0), C338516DataType(1), AKB_DecimalPoint, False) > 0)
        C338854DataType = AKBTypeConst.cAKBTypeLogical
        If C338854 Then
            GoTo Comp_C353462
        Else
            GoTo Comp_C357575
        End If

    Comp_C338903:

        ' Sel Observação
        sCurrComponent = "Sel Observação"
        QueryC338903 = con.CreateCommand()
        QueryC338903.CommandText = QueryC338903.CommandText & " " & "SELECT COUNT(WF_PRE_PEDIDO_OBS.Obs)"
        QueryC338903.CommandText = QueryC338903.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_OBS, AKBUSER01.WF_PRE_PEDIDO "
        QueryC338903.CommandText = QueryC338903.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = WF_PRE_PEDIDO_OBS.Id_PrePedido"
        QueryC338903.CommandText = QueryC338903.CommandText & " " & "AND WF_PRE_PEDIDO_OBS.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC338903.Transaction = txn
        RsC338903 = QueryC338903.ExecuteReader()
        Dim iC338903 As Short
        ReDim C338903DataType(RsC338903.FieldCount)
        For iC338903 = 0 to RsC338903.FieldCount - 1
            Select Case RsC338903.GetDataTypeName(iC338903).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C338903DataType(iC338903 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C338903DataType(iC338903 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C338903DataType(iC338903 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC338903
        RsC338903_EOF = Not RsC338903.Read()

        GoTo Comp_C338906

    Comp_C338906:

        ' Observação?
        sCurrComponent = "Observação?"
        C338906 = (fn_ConvertValueCompiled(RsC338903(0), C338903DataType(1), AKB_DecimalPoint, False)  = 0)
        C338906DataType = AKBTypeConst.cAKBTypeLogical
        If C338906 Then
            GoTo Comp_C437817
        Else
            GoTo Comp_C353463
        End If

    Comp_C338907:

        ' SelCondPagtoCliente
        sCurrComponent = "SelCondPagtoCliente"
        QueryC338907 = con.CreateCommand()
        QueryC338907.CommandText = QueryC338907.CommandText & " " & "SELECT NVL(WF_CLIENTES_COND_PAGTO.Cod_Pagto ,0)"
        QueryC338907.CommandText = QueryC338907.CommandText & " " & "FROM AKBUSER01.WF_CLIENTES_COND_PAGTO "
        QueryC338907.CommandText = QueryC338907.CommandText & " " & "WHERE WF_CLIENTES_COND_PAGTO.Cod_Cliente = (SELECT WF_PRE_PEDIDO.Cod_Cliente  FROM AKBUSER01.WF_PRE_PEDIDO   WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC338907.CommandText = QueryC338907.CommandText & " " & "AND WF_CLIENTES_COND_PAGTO.Condicao_Default = 1"
        QueryC338907.CommandText = QueryC338907.CommandText & " " & ""
        QueryC338907.CommandText = QueryC338907.CommandText & " " & ""
        QueryC338907.Transaction = txn
        RsC338907 = QueryC338907.ExecuteReader()
        Dim iC338907 As Short
        ReDim C338907DataType(RsC338907.FieldCount)
        For iC338907 = 0 to RsC338907.FieldCount - 1
            Select Case RsC338907.GetDataTypeName(iC338907).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C338907DataType(iC338907 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C338907DataType(iC338907 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C338907DataType(iC338907 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC338907
        RsC338907_EOF = Not RsC338907.Read()

        GoTo Comp_C353432

    Comp_C339066:

        ' SelDescontos
        sCurrComponent = "SelDescontos"
        QueryC339066 = con.CreateCommand()
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "SELECT COUNT(*)  "
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "FROM DUAL "
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "WHERE 280 IN "
        QueryC339066.CommandText = QueryC339066.CommandText & " " & ""
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "(SELECT NVL(WF_TP_DESCONTO_CLIENTE.Tipo_Desconto ,0)"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_CLIENTE"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "WHERE WF_TP_DESCONTO_CLIENTE.Cod_Cliente = (SELECT WF_PRE_PEDIDO.Cod_Cliente FROM AKBUSER01.WF_PRE_PEDIDO  WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "AND WF_TP_DESCONTO_CLIENTE.usar_pedido = 1"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & ""
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "UNION ALL"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & ""
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "SELECT NVL(WF_TP_DESCONTO_REPRES.Tipo_Desconto ,0)"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_REPRES"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "WHERE WF_TP_DESCONTO_REPRES.Cod_Repres = (SELECT WF_PRE_PEDIDO.Cod_Repres FROM AKBUSER01.WF_PRE_PEDIDO  WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "AND WF_TP_DESCONTO_REPRES.usar_pedido = 1"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & ""
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "UNION ALL"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & ""
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "SELECT NVL( WF_TP_DESCONTO_DEPART.Tipo_Desconto,0) "
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_DEPART"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "WHERE WF_TP_DESCONTO_DEPART.Departamento = (SELECT WF_PRE_PEDIDO.Departamento "
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "AND WF_TP_DESCONTO_DEPART.usar_pedido = 1"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & ""
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "UNION ALL"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & ""
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "SELECT NVL(WF_TP_DESCONTO_ESTAB.Tipo_Desconto ,0)"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_ESTAB"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "WHERE WF_TP_DESCONTO_ESTAB.Cod_Pessoa_Estab_Juridico = (SELECT WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico "
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC339066.CommandText = QueryC339066.CommandText & " " & "AND WF_TP_DESCONTO_ESTAB.usar_pedido = 1 )"
        QueryC339066.Transaction = txn
        RsC339066 = QueryC339066.ExecuteReader()
        Dim iC339066 As Short
        ReDim C339066DataType(RsC339066.FieldCount)
        For iC339066 = 0 to RsC339066.FieldCount - 1
            Select Case RsC339066.GetDataTypeName(iC339066).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C339066DataType(iC339066 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C339066DataType(iC339066 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C339066DataType(iC339066 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC339066
        RsC339066_EOF = Not RsC339066.Read()

        GoTo Comp_C339069

    Comp_C339069:

        ' Desc280?
        sCurrComponent = "Desc280?"
        C339069 = (1 = 0)
        C339069DataType = AKBTypeConst.cAKBTypeLogical
        If C339069 Then
            GoTo Comp_C353474
        Else
            GoTo Comp_C339200
        End If

    Comp_C339070:

        ' vRet
        sCurrComponent = "vRet"
        C339070 = 0
        C339070DataType = 1
        GoTo Comp_C353482

    Comp_C339200:

        ' SelPrazo
        sCurrComponent = "SelPrazo"
        QueryC339200 = con.CreateCommand()
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "SELECT NVL(WF_TP_PRODUTO_SIGLA.Prazo_Dias,0), 'Mínimo'"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO, AKBUSER01.WF_PRE_PEDIDO_ITENS, AKBUSER01.WF_TP_PRODUTO_SIGLA"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND WF_PRE_PEDIDO.Tp_Produto = WF_TP_PRODUTO_SIGLA.Tp_Produto"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Sigla_Prod = WF_TP_PRODUTO_SIGLA.Sigla_Prod"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND NVL(WF_TP_PRODUTO_SIGLA.Prazo_Dias,0) > 0"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND (WF_PRE_PEDIDO_ITENS.Prazo_Entrega - WF_PRE_PEDIDO.Dt_Pedido) > NVL(WF_TP_PRODUTO_SIGLA.Prazo_Dias,0)"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & ""
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "UNION ALL"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & ""
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "SELECT NVL(WF_TIPO_PED.Qtde_Dias_Min_Entrega,0), 'Mínimo'"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_ITENS, AKBUSER01.WF_TIPO_PED, AKBUSER01.WF_PRE_PEDIDO"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Tipo_Ped = WF_TIPO_PED.Tipo_Ped"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND NVL(WF_TIPO_PED.Qtde_Dias_Min_Entrega,0) > 0"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND (WF_PRE_PEDIDO_ITENS.Prazo_Entrega - WF_PRE_PEDIDO.Dt_Pedido) > NVL(WF_TIPO_PED.Qtde_Dias_Min_Entrega,0)"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND NOT EXISTS ("
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "SELECT *"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "FROM AKBUSER01.WF_TP_PRODUTO_SIGLA"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "WHERE WF_PRE_PEDIDO.Tp_Produto = WF_TP_PRODUTO_SIGLA.Tp_Produto"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Sigla_Prod = WF_TP_PRODUTO_SIGLA.Sigla_Prod"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND NVL(WF_TP_PRODUTO_SIGLA.Prazo_Dias,0) > 0"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & ") "
        QueryC339200.CommandText = QueryC339200.CommandText & " " & ""
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "UNION ALL"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & ""
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "SELECT NVL(WF_TIPO_PED.Qtde_Dias_Entrega,0), 'Máximo'"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_ITENS, AKBUSER01.WF_TIPO_PED, AKBUSER01.WF_PRE_PEDIDO"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Tipo_Ped = WF_TIPO_PED.Tipo_Ped"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND NVL(WF_TIPO_PED.Qtde_Dias_Min_Entrega,0) > 0"
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC339200.CommandText = QueryC339200.CommandText & " " & "AND NVL(WF_TIPO_PED.Qtde_Dias_Entrega,0) < (WF_PRE_PEDIDO_ITENS.Prazo_Entrega - WF_PRE_PEDIDO.Dt_Pedido) "
        QueryC339200.Transaction = txn
        RsC339200 = QueryC339200.ExecuteReader()
        Dim iC339200 As Short
        ReDim C339200DataType(RsC339200.FieldCount)
        For iC339200 = 0 to RsC339200.FieldCount - 1
            Select Case RsC339200.GetDataTypeName(iC339200).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C339200DataType(iC339200 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C339200DataType(iC339200 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C339200DataType(iC339200 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC339200
        RsC339200_EOF = Not RsC339200.Read()

        GoTo Comp_C353435

    Comp_C339201:

        ' Tem Prazo?
        sCurrComponent = "Tem Prazo?"
        C339201 = (1 = 0)
        C339201DataType = AKBTypeConst.cAKBTypeLogical
        If C339201 Then
            GoTo Comp_C353475
        Else
            GoTo Comp_C353476
        End If

    Comp_C339203:

        ' InsLogPrePedido
        sCurrComponent = "InsLogPrePedido"
        QueryC339203 = con.CreateCommand()
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "INSERT"
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "INTO AKBUSER01.WF_PRE_PEDIDO_LOG_BLOQ"
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "  ("
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "    WF_PRE_PEDIDO_LOG_BLOQ.ID_LOG,"
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "    WF_PRE_PEDIDO_LOG_BLOQ.ID_PRE_PEDIDO,"
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "    WF_PRE_PEDIDO_LOG_BLOQ.OBS"
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "  )"
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "  VALUES"
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "  ("
        QueryC339203.CommandText = QueryC339203.CommandText & " " & "   decode( " & _ 
ConvertValue(P58093, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", null," & _ 
ConvertValue(C339204, C339204DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(P58093, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") ," & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ," & _ 
ConvertValue(C353580, C353580DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") "
        QueryC339203.Transaction = txn
        C339203 = QueryC339203.ExecuteNonQuery()
        C339203DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C338501

    Comp_C339204:

        ' IdentIdLog
        sCurrComponent = "IdentIdLog"
        C339204DataType = 1
        C339204 = ObjTable_NewIdentity (con, txn, "WF_PRE_PEDIDO_LOG_BLOQ")
        GoTo Comp_C339203

    Comp_C339206:

        ' Atr_Vlr_Obs_Cliente
        sCurrComponent = "Atr_Vlr_Obs_Cliente"
        C339206DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Possui Cliente do Pré-Cadastro (Novo Cliente)", 3, AKB_DecimalPoint)
        C339206 = True
        GoTo Comp_C338505

    Comp_C339207:

        ' Atr_Vlr_Obs_Transp
        sCurrComponent = "Atr_Vlr_Obs_Transp"
        C339207DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Não Possui Transportadora", 3, AKB_DecimalPoint)
        C339207 = True
        GoTo Comp_C338516

    Comp_C339208:

        ' Atr_Vlr_Obs_ForaLinha
        sCurrComponent = "Atr_Vlr_Obs_ForaLinha"
        C339208DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Possui Artigos Fora de Linha", 3, AKB_DecimalPoint)
        C339208 = True
        GoTo Comp_C357575

    Comp_C339209:

        ' Atr_Vlr_Obs_Obs
        sCurrComponent = "Atr_Vlr_Obs_Obs"
        C339209DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Possui  Observação", 3, AKB_DecimalPoint)
        C339209 = True
        GoTo Comp_C437817

    Comp_C339210:

        ' Atr_Vlr_Obs_CondPagto
        sCurrComponent = "Atr_Vlr_Obs_CondPagto"
        C339210DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Não Possui Condição de Pagto", 3, AKB_DecimalPoint)
        C339210 = True
        GoTo Comp_C357859

    Comp_C339211:

        ' Atr_Vlr_Obs_Desconto280
        sCurrComponent = "Atr_Vlr_Obs_Desconto280"
        C339211DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Possui Desconto de 2ª / 3ª Qualidade", 3, AKB_DecimalPoint)
        C339211 = True
        GoTo Comp_C339200

    Comp_C339212:

        ' Atr_Vlr_Obs_Pilotagem
        sCurrComponent = "Atr_Vlr_Obs_Pilotagem"
        C339212DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Fora do prazo " & C353451 & " de " & C353436 & " dias para entrega", 3, AKB_DecimalPoint)
        C339212 = True
        GoTo Comp_C353476

    Comp_C339213:

        ' Atr_Vlr_Ret
        sCurrComponent = "Atr_Vlr_Ret"
        C339213DataType = 4
        C339070 = fn_ConvertValueCompiled(1, 1, AKB_DecimalPoint)
        C339213 = True
        GoTo Comp_C338501

    Comp_C353432:

        ' EOF Cond PagtoCliente
        sCurrComponent = "EOF Cond PagtoCliente"
        C353432DataType = 4
        C353432 = RsC338907_EOF
        GoTo Comp_C353433

    Comp_C353433:

        ' Tem CondiçãoCliente?
        sCurrComponent = "Tem CondiçãoCliente?"
        C353433 = (fn_ConvertValueCompiled(C353432, C353432DataType, AKB_DecimalPoint, False) = 0)
        C353433DataType = AKBTypeConst.cAKBTypeLogical
        If C353433 Then
            GoTo Comp_C357859
        Else
            GoTo Comp_C358845
        End If

    Comp_C353435:

        ' EOF prazo
        sCurrComponent = "EOF prazo"
        C353435DataType = 4
        C353435 = RsC339200_EOF
        GoTo Comp_C339201

    Comp_C353436:

        ' vDias
        sCurrComponent = "vDias"
        C353436DataType = 0
        C353436 = RsC339200(0)
        C353436DataType = C339200Datatype(1)
        If C353436DataType = AKBTypeConst.cAKBTypeString Then
          C353436 = IIF(IsDBNull(C353436), C353436, Strings.RTrim(C353436))
        End If 

        GoTo Comp_C353451

    Comp_C353451:

        ' vTipoPrazo
        sCurrComponent = "vTipoPrazo"
        C353451DataType = 0
        C353451DataType = C339200Datatype(2)
        If C353451DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC339200(1)) Then
          C353451 = Strings.RTrim(RsC339200(1))
        Else 
          C353451 = RsC339200(1)
        End If 

        GoTo Comp_C339212

    Comp_C353458:

        ' matrizObs
        sCurrComponent = "matrizObs"
        ReDim C353458(1, 1)
        ReDim C353458DataType(1)
        C353458DataType(1) = 3
        C353458(1, 1) = "0"
        C353458CurrentRow = 1

        GoTo Comp_C339070

    Comp_C353460:

        ' addLine1
        sCurrComponent = "addLine1"
        C353460DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C353460 = True

        GoTo Comp_C339206

    Comp_C353461:

        ' addLine2
        sCurrComponent = "addLine2"
        C353461DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C353461 = True

        GoTo Comp_C339207

    Comp_C353462:

        ' addLine3
        sCurrComponent = "addLine3"
        C353462DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C353462 = True

        GoTo Comp_C339208

    Comp_C353463:

        ' addLine4
        sCurrComponent = "addLine4"
        C353463DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C353463 = True

        GoTo Comp_C339209

    Comp_C353467:

        ' addLine5
        sCurrComponent = "addLine5"
        C353467DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C353467 = True

        GoTo Comp_C339210

    Comp_C353474:

        ' addline7
        sCurrComponent = "addline7"
        C353474DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C353474 = True

        GoTo Comp_C339211

    Comp_C353475:

        ' addLine8
        sCurrComponent = "addLine8"
        C353475DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C353475 = True

        GoTo Comp_C353436

    Comp_C353476:

        ' MontaListaMatriz
        sCurrComponent = "MontaListaMatriz"
        C353476DataType = 5
        C353458CurrentRow = 1
        C353476 = ""
        Do Until C353458CurrentRow > UBound(C353458, 2)
            If RTrim(C353476) <> "" Then
                C353476 = C353476 & ", "
            End If
            C353476 = C353476 & ConvertValue(C353458(1, C353458CurrentRow), 0, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss")
            If Not (C353458CurrentRow > UBound(C353458, 2)) Then
                C353458CurrentRow = C353458CurrentRow + 1
            End If
        Loop

        GoTo Comp_C353478

    Comp_C353477:

        ' Teve Bloqueio?
        sCurrComponent = "Teve Bloqueio?"
        C353477 = (fn_ConvertValueCompiled(C353478, C353478DataType, AKB_DecimalPoint, False)  >0)
        C353477DataType = AKBTypeConst.cAKBTypeLogical
        If C353477 Then
            GoTo Comp_C353580
        Else
            GoTo Comp_C339213
        End If

    Comp_C353478:

        ' ComprMontaLista
        sCurrComponent = "ComprMontaLista"
        C353478DataType = 1
        C353478 = Len(C353476 & "")
        GoTo Comp_C353477

    Comp_C353482:

        ' LimpaMatriz
        sCurrComponent = "LimpaMatriz"
        C353482DataType = 4
        ReDim C353458(UBound(C353458DataType), 0)
        C353458CurrentRow = 1
        C353482 = True

        GoTo Comp_C338494

    Comp_C353580:

        ' Trocar'
        sCurrComponent = "Trocar'"
        C353580DataType = 3
        C353580 = Replace(C353476 , "'", "")
        GoTo Comp_C339204

    Comp_C357575:

        ' Desvio
        sCurrComponent = "Desvio"
        C357575 = (1=1)
        C357575DataType = AKBTypeConst.cAKBTypeLogical
        If C357575 Then
            GoTo Comp_C437817
        Else
            GoTo Comp_C338903
        End If

    Comp_C357859:

        ' SelPercentualDesc
        sCurrComponent = "SelPercentualDesc"
        QueryC357859 = con.CreateCommand()
        QueryC357859.CommandText = QueryC357859.CommandText & " " & "SELECT COUNT(WF_PRE_PEDIDO_ITENS.PORC_DESCONTO)"
        QueryC357859.CommandText = QueryC357859.CommandText & " " & ""
        QueryC357859.CommandText = QueryC357859.CommandText & " " & "FROM WF_PRE_PEDIDO, WF_PRE_PEDIDO_ITENS, WF_DEPARTAMENTO"
        QueryC357859.CommandText = QueryC357859.CommandText & " " & ""
        QueryC357859.CommandText = QueryC357859.CommandText & " " & "WHERE WF_PRE_PEDIDO.ID_PREPEDIDO = WF_PRE_PEDIDO_ITENS.ID_PREPEDIDO"
        QueryC357859.CommandText = QueryC357859.CommandText & " " & "AND WF_PRE_PEDIDO.DEPARTAMENTO = WF_DEPARTAMENTO.DEPARTAMENTO"
        QueryC357859.CommandText = QueryC357859.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.PORC_DESCONTO > NVL(WF_DEPARTAMENTO.PERC_MAX_DESCONT,0)"
        QueryC357859.CommandText = QueryC357859.CommandText & " " & "AND WF_PRE_PEDIDO.ID_PREPEDIDO = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC357859.CommandText = QueryC357859.CommandText & " " & ""
        QueryC357859.CommandText = QueryC357859.CommandText & " " & ""
        QueryC357859.CommandText = QueryC357859.CommandText & " " & ""
        QueryC357859.Transaction = txn
        RsC357859 = QueryC357859.ExecuteReader()
        Dim iC357859 As Short
        ReDim C357859DataType(RsC357859.FieldCount)
        For iC357859 = 0 to RsC357859.FieldCount - 1
            Select Case RsC357859.GetDataTypeName(iC357859).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C357859DataType(iC357859 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C357859DataType(iC357859 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C357859DataType(iC357859 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC357859
        RsC357859_EOF = Not RsC357859.Read()

        GoTo Comp_C357864

    Comp_C357864:

        ' %DescMaior?
        sCurrComponent = "%DescMaior?"
        C357864 = (fn_ConvertValueCompiled(RsC357859(0), C357859DataType(1), AKB_DecimalPoint, False)  >0)
        C357864DataType = AKBTypeConst.cAKBTypeLogical
        If C357864 Then
            GoTo Comp_C357865
        Else
            GoTo Comp_C443939
        End If

    Comp_C357865:

        ' addLine6
        sCurrComponent = "addLine6"
        C357865DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C357865 = True

        GoTo Comp_C357866

    Comp_C357866:

        ' Atr_Vlr_Obs_DescMax
        sCurrComponent = "Atr_Vlr_Obs_DescMax"
        C357866DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Possui Desconto Acima do Máximo Permitido por Depto", 3, AKB_DecimalPoint)
        C357866 = True
        GoTo Comp_C443939

    Comp_C358845:

        ' SelCondPagtoRepres
        sCurrComponent = "SelCondPagtoRepres"
        QueryC358845 = con.CreateCommand()
        QueryC358845.CommandText = QueryC358845.CommandText & " " & "SELECT NVL( WF_REPRES_COND_PAGTO.Cod_Pagto,0) FROM AKBUSER01.WF_REPRES_COND_PAGTO "
        QueryC358845.CommandText = QueryC358845.CommandText & " " & "WHERE WF_REPRES_COND_PAGTO.Cod_Repres = (SELECT WF_PRE_PEDIDO.Cod_Repres FROM AKBUSER01.WF_PRE_PEDIDO WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC358845.CommandText = QueryC358845.CommandText & " " & ""
        QueryC358845.CommandText = QueryC358845.CommandText & " " & ""
        QueryC358845.Transaction = txn
        RsC358845 = QueryC358845.ExecuteReader()
        Dim iC358845 As Short
        ReDim C358845DataType(RsC358845.FieldCount)
        For iC358845 = 0 to RsC358845.FieldCount - 1
            Select Case RsC358845.GetDataTypeName(iC358845).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C358845DataType(iC358845 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C358845DataType(iC358845 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C358845DataType(iC358845 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC358845
        RsC358845_EOF = Not RsC358845.Read()

        GoTo Comp_C358848

    Comp_C358847:

        ' SelCondPagtoDepto
        sCurrComponent = "SelCondPagtoDepto"
        QueryC358847 = con.CreateCommand()
        QueryC358847.CommandText = QueryC358847.CommandText & " " & "SELECT NVL( WF_DEPTO_COND_PAGTO.Cod_Pagto,0) FROM AKBUSER01.WF_DEPTO_COND_PAGTO "
        QueryC358847.CommandText = QueryC358847.CommandText & " " & "WHERE WF_DEPTO_COND_PAGTO.Departamento = (SELECT WF_PRE_PEDIDO.Departamento                         FROM AKBUSER01.WF_PRE_PEDIDO WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC358847.CommandText = QueryC358847.CommandText & " " & "AND WF_DEPTO_COND_PAGTO.Flag_Default = 1 "
        QueryC358847.Transaction = txn
        RsC358847 = QueryC358847.ExecuteReader()
        Dim iC358847 As Short
        ReDim C358847DataType(RsC358847.FieldCount)
        For iC358847 = 0 to RsC358847.FieldCount - 1
            Select Case RsC358847.GetDataTypeName(iC358847).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C358847DataType(iC358847 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C358847DataType(iC358847 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C358847DataType(iC358847 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC358847
        RsC358847_EOF = Not RsC358847.Read()

        GoTo Comp_C358851

    Comp_C358848:

        ' EOF Cond Pagto Repres
        sCurrComponent = "EOF Cond Pagto Repres"
        C358848DataType = 4
        C358848 = RsC358845_EOF
        GoTo Comp_C358849

    Comp_C358849:

        ' Tem CondPagtoRepres?
        sCurrComponent = "Tem CondPagtoRepres?"
        C358849 = (fn_ConvertValueCompiled(C358848, C358848DataType, AKB_DecimalPoint, False)  = 0)
        C358849DataType = AKBTypeConst.cAKBTypeLogical
        If C358849 Then
            GoTo Comp_C357859
        Else
            GoTo Comp_C358847
        End If

    Comp_C358850:

        ' Tem CondPagtoDepto?
        sCurrComponent = "Tem CondPagtoDepto?"
        C358850 = (fn_ConvertValueCompiled(C358851, C358851DataType, AKB_DecimalPoint, False)   = 0)
        C358850DataType = AKBTypeConst.cAKBTypeLogical
        If C358850 Then
            GoTo Comp_C357859
        Else
            GoTo Comp_C353467
        End If

    Comp_C358851:

        ' EOF Cond PagtoDepto
        sCurrComponent = "EOF Cond PagtoDepto"
        C358851DataType = 4
        C358851 = RsC358847_EOF
        GoTo Comp_C358850

    Comp_C437817:

        ' CondPgtPP
        sCurrComponent = "CondPgtPP"
        QueryC437817 = con.CreateCommand()
        QueryC437817.CommandText = QueryC437817.CommandText & " " & "select NVL(Cod_Pagto,0)"
        QueryC437817.CommandText = QueryC437817.CommandText & " " & "FROM WF_PRE_PEDIDO "
        QueryC437817.CommandText = QueryC437817.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC437817.Transaction = txn
        RsC437817 = QueryC437817.ExecuteReader()
        Dim iC437817 As Short
        ReDim C437817DataType(RsC437817.FieldCount)
        For iC437817 = 0 to RsC437817.FieldCount - 1
            Select Case RsC437817.GetDataTypeName(iC437817).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C437817DataType(iC437817 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C437817DataType(iC437817 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C437817DataType(iC437817 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC437817
        RsC437817_EOF = Not RsC437817.Read()

        GoTo Comp_C437818

    Comp_C437818:

        ' DESVIO1
        sCurrComponent = "DESVIO1"
        C437818 = (fn_ConvertValueCompiled(RsC437817(0), C437817DataType(1), AKB_DecimalPoint, False) = 0)
        C437818DataType = AKBTypeConst.cAKBTypeLogical
        If C437818 Then
            GoTo Comp_C338907
        Else
            GoTo Comp_C357859
        End If

    Comp_C443935:

        ' CliSemCodIdentDest
        sCurrComponent = "CliSemCodIdentDest"
        QueryC443935 = con.CreateCommand()
        QueryC443935.CommandText = QueryC443935.CommandText & " " & "SELECT COUNT(*) "
        QueryC443935.CommandText = QueryC443935.CommandText & " " & "FROM AKBUSER01.WF_PES_JURIDICA "
        QueryC443935.CommandText = QueryC443935.CommandText & " " & "WHERE (nvl( WF_PES_JURIDICA.Codigo_Ident_Destinatario,0) =  0)"
        QueryC443935.CommandText = QueryC443935.CommandText & " " & "AND WF_PES_JURIDICA.Cod_Pessoa = (SELECT WF_PRE_PEDIDO.Cod_Cliente "
        QueryC443935.CommandText = QueryC443935.CommandText & " " & "                                        FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC443935.CommandText = QueryC443935.CommandText & " " & "                                        WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC443935.Transaction = txn
        RsC443935 = QueryC443935.ExecuteReader()
        Dim iC443935 As Short
        ReDim C443935DataType(RsC443935.FieldCount)
        For iC443935 = 0 to RsC443935.FieldCount - 1
            Select Case RsC443935.GetDataTypeName(iC443935).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C443935DataType(iC443935 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C443935DataType(iC443935 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C443935DataType(iC443935 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC443935
        RsC443935_EOF = Not RsC443935.Read()

        GoTo Comp_C443936

    Comp_C443936:

        ' CliSemCodIdentDest?
        sCurrComponent = "CliSemCodIdentDest?"
        C443936 = (fn_ConvertValueCompiled(RsC443935(0), C443935DataType(1), AKB_DecimalPoint, False) > 0)
        C443936DataType = AKBTypeConst.cAKBTypeLogical
        If C443936 Then
            GoTo Comp_C443937
        Else
            GoTo Comp_C339066
        End If

    Comp_C443937:

        ' AddLine9
        sCurrComponent = "AddLine9"
        C443937DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C443937 = True

        GoTo Comp_C443938

    Comp_C443938:

        ' Atr_Vlr_Cli_Sem_Cod_Ident
        sCurrComponent = "Atr_Vlr_Cli_Sem_Cod_Ident"
        C443938DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Cliente Não Possui Cod Identificação IE - Pes Juridica", 3, AKB_DecimalPoint)
        C443938 = True
        GoTo Comp_C339066

    Comp_C443939:

        ' EmbInativaPed
        sCurrComponent = "EmbInativaPed"
        QueryC443939 = con.CreateCommand()
        QueryC443939.CommandText = QueryC443939.CommandText & " " & "SELECT COUNT(*)"
        QueryC443939.CommandText = QueryC443939.CommandText & " " & "FROM WF_PRE_PEDIDO PP, WF_PRE_PEDIDO_ITENS PPI, WF_EMB_COMP_VDA_PROD ECV"
        QueryC443939.CommandText = QueryC443939.CommandText & " " & "WHERE PP.Id_PrePedido = " & _ 
ConvertValue(P56206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC443939.CommandText = QueryC443939.CommandText & " " & "AND PP.Id_PrePedido = PPI.Id_PrePedido"
        QueryC443939.CommandText = QueryC443939.CommandText & " " & "AND PPI.Sigla_Prod = ECV.Sigla_Prod"
        QueryC443939.CommandText = QueryC443939.CommandText & " " & "AND PPI.Acesso = ECV.Acesso"
        QueryC443939.CommandText = QueryC443939.CommandText & " " & "AND PPI.Cod_Embalagem = ECV.Cod_Embalagem"
        QueryC443939.CommandText = QueryC443939.CommandText & " " & "AND ECV.Inativo = 1"
        QueryC443939.Transaction = txn
        RsC443939 = QueryC443939.ExecuteReader()
        Dim iC443939 As Short
        ReDim C443939DataType(RsC443939.FieldCount)
        For iC443939 = 0 to RsC443939.FieldCount - 1
            Select Case RsC443939.GetDataTypeName(iC443939).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C443939DataType(iC443939 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C443939DataType(iC443939 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C443939DataType(iC443939 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC443939
        RsC443939_EOF = Not RsC443939.Read()

        GoTo Comp_C443940

    Comp_C443940:

        ' EmbInativaPed?
        sCurrComponent = "EmbInativaPed?"
        C443940 = (fn_ConvertValueCompiled(RsC443939(0), C443939DataType(1), AKB_DecimalPoint, False) > 0)
        C443940DataType = AKBTypeConst.cAKBTypeLogical
        If C443940 Then
            GoTo Comp_C443941
        Else
            GoTo Comp_C443935
        End If

    Comp_C443941:

        ' AddLine10
        sCurrComponent = "AddLine10"
        C443941DataType = 4
        ReDim Preserve C353458(UBound(C353458DataType), UBound(C353458, 2) + 1)
        C353458CurrentRow = UBound(C353458, 2)
        C443941 = True

        GoTo Comp_C443942

    Comp_C443942:

        ' Atr_Emb_Inativa
        sCurrComponent = "Atr_Emb_Inativa"
        C443942DataType = 4
        C353458(1, C353458CurrentRow) = fn_ConvertValueCompiled("Pedido Possui Embalagem Inativa", 3, AKB_DecimalPoint)
        C443942 = True
        GoTo Comp_C443935

    Exit_R15836:

        Exit Function

    Err_R15836:
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
    Public Shared Function ObjTable_NewIdentity(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal TableName as String) As Long
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select SEQ_" & TableName & ".nextval from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Query.Transaction = txn
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            ObjTable_NewIdentity = Rs(0)
        End If
        Rs.Close()
    End Function

End Class
