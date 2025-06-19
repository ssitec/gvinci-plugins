Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R6668

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

    Public Shared Function R6668(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P20174 As Double, ByVal P20175 As Double, ByVal P47604 As Object, ByVal P47605 As Object, ByVal P47606 As Object, ByVal P47607 As Object) As DataTable
    ' 
    ' 6668 - Verifica o tipo de desconto no pedido - Versão: 0
    ' 
        'On Error GoTo Err_R6668
        Dim sCurrComponent as String

        Dim C95093 As Object
        Dim C95093DataType As Short
        Dim QueryC95094 As New Object
        Dim RsC95094 As New Object
        Dim C95094DataType() As Short
        Dim RsC95094_EOF As Boolean
        Dim C95095 As Boolean
        Dim C95095DataType As Short
        Dim C95096 As Short
        Dim C95096DataType As Short
        Dim C95096ReturnDataType() As Short

        Dim C95097 As Object
        Dim C95097DataType As Short
        Dim C95098DataType() As Short
        Dim QueryC95099 As New Object
        Dim RsC95099 As New Object
        Dim C95099DataType() As Short
        Dim RsC95099_EOF As Boolean
        Dim C95100 As Boolean
        Dim C95100DataType As Short
        Dim C95101 As Short
        Dim C95101DataType As Short
        Dim C95101ReturnDataType() As Short

        Dim C95102 As Object
        Dim C95102DataType As Short
        Dim C95112 As Object
        Dim C95112DataType As Short
        Dim C95113 As Object
        Dim C95113DataType As Short
        Dim QueryC266967 As New Object
        Dim RsC266967 As New Object
        Dim C266967DataType() As Short
        Dim RsC266967_EOF As Boolean
        Dim QueryC266968 As New Object
        Dim RsC266968 As New Object
        Dim C266968DataType() As Short
        Dim RsC266968_EOF As Boolean
        Dim C266969 As Boolean
        Dim C266969DataType As Short
        Dim C266970 As Short
        Dim C266970DataType As Short
        Dim C266970ReturnDataType() As Short

        Dim C266971 As Boolean
        Dim C266971DataType As Short
        Dim C266972 As Short
        Dim C266972DataType As Short
        Dim C266972ReturnDataType() As Short

        Dim C266973 As Object
        Dim C266973DataType As Short
        Dim C266974 As Object
        Dim C266974DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C95093

    Comp_C95093:

        ' Retorno
        sCurrComponent = "Retorno"
        C95093 = 1
        C95093DataType = 4
        GoTo Comp_C266968

    Comp_C95094:

        ' Itens
        sCurrComponent = "Itens"
        QueryC95094 = con.CreateCommand()
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "select "
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "WF_TIPO_DESC.Tipo_Desconto || ' - ' || WF_TIPO_DESC.Descr_Tipo_Desconto "
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "from AKBUSER01.WF_PED_ITENS_DESC , AKBUSER01.WF_TIPO_DESC "
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "where"
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "WF_PED_ITENS_DESC.Tp_Produto = " & _ 
ConvertValue(P20174, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "and WF_PED_ITENS_DESC.Pedido = " & _ 
ConvertValue(P20175, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "and WF_TIPO_DESC.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto "
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "and WF_TIPO_DESC.PERMISSAO_UTILIZACAO = 'Pedido'"
        QueryC95094.CommandText = QueryC95094.CommandText & " " & "and ROWNUM = 1"
        QueryC95094.Transaction = txn
        RsC95094 = QueryC95094.ExecuteReader()
        Dim iC95094 As Short
        ReDim C95094DataType(RsC95094.FieldCount)
        For iC95094 = 0 to RsC95094.FieldCount - 1
            Select Case RsC95094.GetDataTypeName(iC95094).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C95094DataType(iC95094 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C95094DataType(iC95094 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C95094DataType(iC95094 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC95094
        RsC95094_EOF = Not RsC95094.Read()

        GoTo Comp_C95112

    Comp_C95095:

        ' Itens_EOF = 1
        sCurrComponent = "Itens_EOF = 1"
        C95095 = (fn_ConvertValueCompiled(C95112, C95112DataType, AKB_DecimalPoint, False) = 1)
        C95095DataType = AKBTypeConst.cAKBTypeLogical
        If C95095 Then
            GoTo Comp_C95099
        Else
            GoTo Comp_C95096
        End If

    Comp_C95096:

        ' Msg_01
        sCurrComponent = "Msg_01"
        C95096 = 1
        C95096DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C95097

    Comp_C95097:

        ' att_retorno_01
        sCurrComponent = "att_retorno_01"
        C95097DataType = 4
        C95093 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C95097 = True
        GoTo Comp_C95099

    Comp_C95098:

        ' Retorno_01
        sCurrComponent = "Retorno_01"
        Dim tb_C95098 As DataTable = New DataTable()
        tb_C95098.Columns.Add("Retorno" & "")
        Dim row_C95098 As DataRow = tb_C95098.NewRow()
        row_C95098(0) = C95093
        tb_C95098.Rows.Add(row_C95098)
        R6668 = tb_C95098
        ReDim C95098DataType(1)
        C95098DataType(1) = C95093DataType
        ReturnDataType = C95098DataType

        GoTo Exit_R6668

    Comp_C95099:

        ' Pedido
        sCurrComponent = "Pedido"
        QueryC95099 = con.CreateCommand()
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "select "
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "WF_TIPO_DESC.Tipo_Desconto || ' - ' || WF_TIPO_DESC.Descr_Tipo_Desconto "
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "from AKBUSER01.WF_PED_SEQ_DESCONTO , AKBUSER01.WF_TIPO_DESC "
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "where"
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P20174, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "and WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P20175, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "and WF_TIPO_DESC.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto "
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "and WF_TIPO_DESC.PERMISSAO_UTILIZACAO = 'Itens'"
        QueryC95099.CommandText = QueryC95099.CommandText & " " & "and ROWNUM = 1"
        QueryC95099.Transaction = txn
        RsC95099 = QueryC95099.ExecuteReader()
        Dim iC95099 As Short
        ReDim C95099DataType(RsC95099.FieldCount)
        For iC95099 = 0 to RsC95099.FieldCount - 1
            Select Case RsC95099.GetDataTypeName(iC95099).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C95099DataType(iC95099 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C95099DataType(iC95099 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C95099DataType(iC95099 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC95099
        RsC95099_EOF = Not RsC95099.Read()

        GoTo Comp_C95113

    Comp_C95100:

        ' Pedido_EOF = 1
        sCurrComponent = "Pedido_EOF = 1"
        C95100 = (fn_ConvertValueCompiled(C95113, C95113DataType, AKB_DecimalPoint, False) = 1)
        C95100DataType = AKBTypeConst.cAKBTypeLogical
        If C95100 Then
            GoTo Comp_C95098
        Else
            GoTo Comp_C95101
        End If

    Comp_C95101:

        ' Msg_02
        sCurrComponent = "Msg_02"
        C95101 = 1
        C95101DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C95102

    Comp_C95102:

        ' att_retorno_02
        sCurrComponent = "att_retorno_02"
        C95102DataType = 4
        C95093 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C95102 = True
        GoTo Comp_C95098

    Comp_C95112:

        ' Itens_EOF
        sCurrComponent = "Itens_EOF"
        C95112DataType = 4
        C95112 = RsC95094_EOF
        GoTo Comp_C95095

    Comp_C95113:

        ' Pedido_EOF
        sCurrComponent = "Pedido_EOF"
        C95113DataType = 4
        C95113 = RsC95099_EOF
        GoTo Comp_C95100

    Comp_C266967:

        ' desconto_itens
        sCurrComponent = "desconto_itens"
        QueryC266967 = con.CreateCommand()
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "SELECT NVL( STRAGG( TRIM( WF_TIPO_DESC.Descr_Tipo_Desconto)||' '), 0)"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM "
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AKBUSER01.WF_PED_ITENS_DESC, "
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AKBUSER01.WF_TIPO_DESC"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_PED_ITENS_DESC.Tp_Produto = " & _ 
ConvertValue(P20174, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_PED_ITENS_DESC.Pedido = " & _ 
ConvertValue(P20175, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_PED_ITENS_DESC.Tipo_Desconto = WF_TIPO_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND ( EXISTS ( "
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "( SELECT WF_TP_DESCONTO_CLIENTE.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_CLIENTE"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_TP_DESCONTO_CLIENTE.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_CLIENTE.Cod_Cliente = " & _ 
ConvertValue(P47606, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_CLIENTE.nao_usar = 1)"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "UNION"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "( SELECT WF_TP_DESCONTO_REPRES.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_REPRES"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_TP_DESCONTO_REPRES.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_REPRES.Cod_Repres = " & _ 
ConvertValue(P47605, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_REPRES.nao_usar = 1)"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "UNION"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "( SELECT WF_TP_DESCONTO_DEPART.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_DEPART"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_TP_DESCONTO_DEPART.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_DEPART.Departamento = " & _ 
ConvertValue(P47604, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_DEPART.nao_usar = 1)"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "UNION"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "( SELECT WF_TP_DESCONTO_ESTAB.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_ESTAB"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_TP_DESCONTO_ESTAB.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_ESTAB.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P47607, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_ESTAB.nao_usar = 1)"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ")"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "OR NOT EXISTS ( "
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "( SELECT WF_TP_DESCONTO_CLIENTE.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_CLIENTE"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_TP_DESCONTO_CLIENTE.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_CLIENTE.Cod_Cliente = " & _ 
ConvertValue(P47606, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ")"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "UNION"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "( SELECT WF_TP_DESCONTO_REPRES.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_REPRES"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_TP_DESCONTO_REPRES.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_REPRES.Cod_Repres = " & _ 
ConvertValue(P47605, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "UNION"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "( SELECT WF_TP_DESCONTO_DEPART.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_DEPART"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_TP_DESCONTO_DEPART.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_DEPART.Departamento = " & _ 
ConvertValue(P47604, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "UNION"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ""
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "( SELECT WF_TP_DESCONTO_ESTAB.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_ESTAB"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "WHERE WF_TP_DESCONTO_ESTAB.Tipo_Desconto = WF_PED_ITENS_DESC.Tipo_Desconto"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & "AND WF_TP_DESCONTO_ESTAB.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P47607, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC266967.CommandText = QueryC266967.CommandText & " " & ") )"
        QueryC266967.Transaction = txn
        RsC266967 = QueryC266967.ExecuteReader()
        Dim iC266967 As Short
        ReDim C266967DataType(RsC266967.FieldCount)
        For iC266967 = 0 to RsC266967.FieldCount - 1
            Select Case RsC266967.GetDataTypeName(iC266967).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C266967DataType(iC266967 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C266967DataType(iC266967 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C266967DataType(iC266967 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC266967
        RsC266967_EOF = Not RsC266967.Read()

        GoTo Comp_C266971

    Comp_C266968:

        ' desconto_pedido
        sCurrComponent = "desconto_pedido"
        QueryC266968 = con.CreateCommand()
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "SELECT NVL( STRAGG( TRIM( WF_TIPO_DESC.Descr_Tipo_Desconto)||' '), 0)"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM "
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AKBUSER01.WF_PED_SEQ_DESCONTO, "
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AKBUSER01.WF_TIPO_DESC"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P20174, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P20175, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Tipo_Desconto = WF_TIPO_DESC.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND ( EXISTS( "
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "( SELECT WF_TP_DESCONTO_CLIENTE.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_CLIENTE"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_TP_DESCONTO_CLIENTE.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_CLIENTE.Cod_Cliente = " & _ 
ConvertValue(P47606, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_CLIENTE.nao_usar = 1)"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "UNION"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "( SELECT WF_TP_DESCONTO_REPRES.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_REPRES"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_TP_DESCONTO_REPRES.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_REPRES.Cod_Repres = " & _ 
ConvertValue(P47605, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_REPRES.nao_usar = 1)"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "UNION"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "( SELECT WF_TP_DESCONTO_DEPART.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_DEPART"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_TP_DESCONTO_DEPART.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_DEPART.Departamento = " & _ 
ConvertValue(P47604, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_DEPART.nao_usar = 1)"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "UNION"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "( SELECT WF_TP_DESCONTO_ESTAB.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_ESTAB"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_TP_DESCONTO_ESTAB.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_ESTAB.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P47607, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_ESTAB.nao_usar = 1)"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ")"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "OR NOT EXISTS( "
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "( SELECT WF_TP_DESCONTO_CLIENTE.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_CLIENTE"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_TP_DESCONTO_CLIENTE.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_CLIENTE.Cod_Cliente = " & _ 
ConvertValue(P47606, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "UNION"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "( SELECT WF_TP_DESCONTO_REPRES.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_REPRES"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_TP_DESCONTO_REPRES.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_REPRES.Cod_Repres = " & _ 
ConvertValue(P47605, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "UNION"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "( SELECT WF_TP_DESCONTO_DEPART.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_DEPART"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_TP_DESCONTO_DEPART.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_DEPART.Departamento = " & _ 
ConvertValue(P47604, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "UNION"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ""
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "( SELECT WF_TP_DESCONTO_ESTAB.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "FROM AKBUSER01.WF_TP_DESCONTO_ESTAB"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "WHERE WF_TP_DESCONTO_ESTAB.Tipo_Desconto = WF_PED_SEQ_DESCONTO.Tipo_Desconto"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & "AND WF_TP_DESCONTO_ESTAB.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P47607, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC266968.CommandText = QueryC266968.CommandText & " " & ") )"
        QueryC266968.Transaction = txn
        RsC266968 = QueryC266968.ExecuteReader()
        Dim iC266968 As Short
        ReDim C266968DataType(RsC266968.FieldCount)
        For iC266968 = 0 to RsC266968.FieldCount - 1
            Select Case RsC266968.GetDataTypeName(iC266968).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C266968DataType(iC266968 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C266968DataType(iC266968 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C266968DataType(iC266968 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC266968
        RsC266968_EOF = Not RsC266968.Read()

        GoTo Comp_C266969

    Comp_C266969:

        ' desconto_pedido = 0
        sCurrComponent = "desconto_pedido = 0"
        C266969 = (fn_ConvertValueCompiled(RsC266968(0), C266968DataType(1), AKB_DecimalPoint, False) = 0)
        C266969DataType = AKBTypeConst.cAKBTypeLogical
        If C266969 Then
            GoTo Comp_C266967
        Else
            GoTo Comp_C266970
        End If

    Comp_C266970:

        ' MSG1
        sCurrComponent = "MSG1"
        C266970 = 1
        C266970DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C266973

    Comp_C266971:

        ' desconto_itens = 0
        sCurrComponent = "desconto_itens = 0"
        C266971 = (fn_ConvertValueCompiled(RsC266967(0), C266967DataType(1), AKB_DecimalPoint, False) = 0)
        C266971DataType = AKBTypeConst.cAKBTypeLogical
        If C266971 Then
            GoTo Comp_C95094
        Else
            GoTo Comp_C266972
        End If

    Comp_C266972:

        ' MSG2
        sCurrComponent = "MSG2"
        C266972 = 1
        C266972DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C266974

    Comp_C266973:

        ' ATT_RETORNO_011
        sCurrComponent = "ATT_RETORNO_011"
        C266973DataType = 4
        C95093 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C266973 = True
        GoTo Comp_C266967

    Comp_C266974:

        ' ATT_RETORNO_0111
        sCurrComponent = "ATT_RETORNO_0111"
        C266974DataType = 4
        C95093 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C266974 = True
        GoTo Comp_C95094

    Exit_R6668:

        Exit Function

    Err_R6668:
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
