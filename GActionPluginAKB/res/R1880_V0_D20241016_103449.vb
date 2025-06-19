Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R1880

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

    Public Shared Function R1880(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P3333 As Double, ByVal P3334 As Object, ByVal P3337 As Double, ByVal P3358 As Object, ByVal P3360 As Object) As DataTable
    ' 
    ' 1880 - Verifica Preço Digitado nos Itens - Versão: 0
    ' 
        'On Error GoTo Err_R1880
        Dim sCurrComponent as String

        Dim QueryC12711 As New Object
        Dim RsC12711 As New Object
        Dim C12711DataType() As Short
        Dim RsC12711_EOF As Boolean
        Dim QueryC12739 As New Object
        Dim RsC12739 As New Object
        Dim C12739DataType() As Short
        Dim RsC12739_EOF As Boolean
        Dim C12740 As Object
        Dim C12740DataType As Short
        Dim C12741 As Object
        Dim C12741DataType As Short
        Dim C12742 As Boolean
        Dim C12742DataType As Short
        Dim C12744 As Object
        Dim C12744DataType As Short
        Dim C12745 As Object
        Dim C12745DataType As Short
        Dim C12746 As Boolean
        Dim C12746DataType As Short
        Dim C12748DataType() As Short
        Dim C12749 As Short
        Dim C12749DataType As Short
        Dim C12749ReturnDataType() As Short

        Dim C12758DataType() As Short
        Dim QueryC14373 As New Object
        Dim RsC14373 As New Object
        Dim C14373DataType() As Short
        Dim RsC14373_EOF As Boolean
        Dim C14453 As Object
        Dim C14453DataType As Short
        Dim C14454 As Boolean
        Dim C14454DataType As Short
        Dim QueryC14455 As New Object
        Dim RsC14455 As New Object
        Dim C14455DataType() As Short
        Dim RsC14455_EOF As Boolean
        Dim C14456 As Object
        Dim C14456DataType As Short
        Dim C14457 As Boolean
        Dim C14457DataType As Short
        Dim C14458DataType() As Short
        Dim C14459 As Short
        Dim C14459DataType As Short
        Dim C14459ReturnDataType() As Short

        Dim QueryC28754 As New Object
        Dim RsC28754 As New Object
        Dim C28754DataType() As Short
        Dim RsC28754_EOF As Boolean
        Dim C28755 As Boolean
        Dim C28755DataType As Short
        Dim C28756 As Boolean
        Dim C28756DataType As Short
        Dim QueryC35188 As New Object
        Dim RsC35188 As New Object
        Dim C35188DataType() As Short
        Dim RsC35188_EOF As Boolean
        Dim C35190 As Boolean
        Dim C35190DataType As Short
        Dim C35191 As Short
        Dim C35191DataType As Short
        Dim C35191ReturnDataType() As Short

        Dim C35192DataType() As Short
        Dim QueryC88206 As New Object
        Dim RsC88206 As New Object
        Dim C88206DataType() As Short
        Dim RsC88206_EOF As Boolean
        Dim C88407 As Boolean
        Dim C88407DataType As Short
        Dim C88408 As Object
        Dim C88408DataType As Short
        Dim C88409 As Object
        Dim C88409DataType As Short
        Dim C88410 As Object
        Dim C88410DataType As Short
        Dim C88411 As Object
        Dim C88411DataType As Short
        Dim QueryC89647 As New Object
        Dim RsC89647 As New Object
        Dim C89647DataType() As Short
        Dim RsC89647_EOF As Boolean
        Dim QueryC177610 As New Object
        Dim RsC177610 As New Object
        Dim C177610DataType() As Short
        Dim RsC177610_EOF As Boolean
        Dim C296825 As Object
        Dim C296825DataType As Short
        Dim C320108DataType() As Short
        Dim C320109 As Object
        Dim C320109DataType As Short
        P3334 = IIf(IsDBNull(P3334), 0, P3334)

        P3358 = IIf(IsDBNull(P3358), 1, P3358)

        P3360 = IIf(IsDBNull(P3360), 0, P3360)

        ReDim ReturnDatatype(0)

        GoTo Comp_C12740

    Comp_C12711:

        ' SelVda
        sCurrComponent = "SelVda"
        QueryC12711 = con.CreateCommand()
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "SELECT WF_ACESSOS.Descr_Acesso "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS , AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_TABPRECO_VDA_PRODUTO, AKBUSER01.WF_ACESSOS "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(P3334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Preco_Digitado = 1"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_PEDIDO_ITENS.Sigla_Prod = WF_ACESSOS.Sigla_Prod "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_PEDIDO_ITENS.Acesso = WF_ACESSOS.Acesso "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_TABPRECO_VDA_PRODUTO.Sigla_Prod (+) = WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_TABPRECO_VDA_PRODUTO.Acesso (+) = WF_PEDIDO_ITENS.Acesso "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_TABPRECO_VDA_PRODUTO.Cod_Embalagem (+) = WF_PEDIDO_ITENS.Cod_Embalagem "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_TABPRECO_VDA_PRODUTO.Tabela_Preco_Venda (+) =  " & _ 
ConvertValue(P3334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND EXISTS (SELECT * FROM  AKBUSER01.WF_PED_ITENS_DESC "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = WF_PED_ITENS_DESC.Tp_Produto "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = WF_PED_ITENS_DESC.Pedido "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Seq_Item  = WF_PED_ITENS_DESC.Seq_Item  )"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "And   NVL(WF_PEDIDO_ITENS.PRECO_SEM_DESCONTO,0) <= 0"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "And 0 =  ("
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "                  Select NVL(NAO_VALIDA_AGRUPAMENTO, 0)"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "                  from WF_TABELA_PRECO_VENDA"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "                  where TABELA_PRECO_VENDA = " & _ 
ConvertValue(P3334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "UNION ALL"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "SELECT WF_ACESSOS.Descr_Acesso "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS ,  AKBUSER01.WF_TABPRECO_VDA_PRODUTO, AKBUSER01.WF_TABELA_PRECO_VENDA, AKBUSER01.WF_ACESSOS "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(P3334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Preco_Digitado = 1"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_PEDIDO_ITENS.Sigla_Prod = WF_ACESSOS.Sigla_Prod "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_PEDIDO_ITENS.Acesso = WF_ACESSOS.Acesso "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_TABPRECO_VDA_PRODUTO.Sigla_Prod (+) = WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_TABPRECO_VDA_PRODUTO.Acesso (+) = WF_PEDIDO_ITENS.Acesso "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_TABPRECO_VDA_PRODUTO.Cod_Embalagem (+) = WF_PEDIDO_ITENS.Cod_Embalagem "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND  WF_TABPRECO_VDA_PRODUTO.Tabela_Preco_Venda (+) =  " & _ 
ConvertValue(P3334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND NOT EXISTS (SELECT * FROM  AKBUSER01.WF_PED_ITENS_DESC "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = WF_PED_ITENS_DESC.Tp_Produto "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = WF_PED_ITENS_DESC.Pedido "
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "AND WF_PEDIDO_ITENS.Seq_Item  = WF_PED_ITENS_DESC.Seq_Item  )"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "And  NVL(WF_PEDIDO_ITENS.PRECO_SEM_DESCONTO,0) <= 0"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "And 0 =  ("
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "                  Select NVL(NAO_VALIDA_AGRUPAMENTO, 0)"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "                  from WF_TABELA_PRECO_VENDA"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & "                  where TABELA_PRECO_VENDA = " & _ 
ConvertValue(P3334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.CommandText = QueryC12711.CommandText & " " & ""
        QueryC12711.Transaction = txn
        RsC12711 = QueryC12711.ExecuteReader()
        Dim iC12711 As Short
        ReDim C12711DataType(RsC12711.FieldCount)
        For iC12711 = 0 to RsC12711.FieldCount - 1
            Select Case RsC12711.GetDataTypeName(iC12711).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C12711DataType(iC12711 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C12711DataType(iC12711 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C12711DataType(iC12711 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC12711
        RsC12711_EOF = Not RsC12711.Read()

        GoTo Comp_C12745

    Comp_C12739:

        ' SelLote
        sCurrComponent = "SelLote"
        QueryC12739 = con.CreateCommand()
        QueryC12739.CommandText = QueryC12739.CommandText & " " & "SELECT WF_PED_SEQ_DESCONTO.Porc_Desconto FROM AKBUSER01.WF_PED_SEQ_DESCONTO WHERE WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PED_SEQ_DESCONTO.Seq = " & _ 
ConvertValue(P3358, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PED_SEQ_DESCONTO.Tipo_Desconto = 280 "
        QueryC12739.Transaction = txn
        RsC12739 = QueryC12739.ExecuteReader()
        Dim iC12739 As Short
        ReDim C12739DataType(RsC12739.FieldCount)
        For iC12739 = 0 to RsC12739.FieldCount - 1
            Select Case RsC12739.GetDataTypeName(iC12739).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C12739DataType(iC12739 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C12739DataType(iC12739 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C12739DataType(iC12739 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC12739
        RsC12739_EOF = Not RsC12739.Read()

        GoTo Comp_C12741

    Comp_C12740:

        ' VLote
        sCurrComponent = "VLote"
        C12740 = 0
        C12740DataType = 1
        GoTo Comp_C12739

    Comp_C12741:

        ' Eof
        sCurrComponent = "Eof"
        C12741DataType = 4
        C12741 = RsC12739_EOF
        GoTo Comp_C12742

    Comp_C12742:

        ' Eof=T
        sCurrComponent = "Eof=T"
        C12742 = (fn_ConvertValueCompiled(C12741, C12741DataType, AKB_DecimalPoint, False) =1)
        C12742DataType = AKBTypeConst.cAKBTypeLogical
        If C12742 Then
            GoTo Comp_C28754
        Else
            GoTo Comp_C12744
        End If

    Comp_C12744:

        ' UsaLote
        sCurrComponent = "UsaLote"
        C12744DataType = 4
        C12740 = fn_ConvertValueCompiled(RsC12739(0) , 1, AKB_DecimalPoint)
        C12744 = True
        GoTo Comp_C28754

    Comp_C12745:

        ' Eof1
        sCurrComponent = "Eof1"
        C12745DataType = 4
        C12745 = RsC12711_EOF
        GoTo Comp_C12746

    Comp_C12746:

        ' Eof1=T
        sCurrComponent = "Eof1=T"
        C12746 = (fn_ConvertValueCompiled(C12745, C12745DataType, AKB_DecimalPoint, False) =1)
        C12746DataType = AKBTypeConst.cAKBTypeLogical
        If C12746 Then
            GoTo Comp_C88408
        Else
            GoTo Comp_C88411
        End If

    Comp_C12748:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C12748 As DataTable = New DataTable()
        tb_C12748.Columns.Add("VLote" & "")
        Dim row_C12748 As DataRow = tb_C12748.NewRow()
        row_C12748(0) = C12740
        tb_C12748.Rows.Add(row_C12748)
        R1880 = tb_C12748
        ReDim C12748DataType(1)
        C12748DataType(1) = C12740DataType
        ReturnDataType = C12748DataType

        GoTo Exit_R1880

    Comp_C12749:

        ' MSG1
        sCurrComponent = "MSG1"
        C12749 = 1
        C12749DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C12748

    Comp_C12758:

        ' RET2
        sCurrComponent = "RET2"
        Dim tb_C12758 As DataTable = New DataTable()
        tb_C12758.Columns.Add("Eof2=T" & "")
        Dim row_C12758 As DataRow = tb_C12758.NewRow()
        row_C12758(0) = C14454
        tb_C12758.Rows.Add(row_C12758)
        R1880 = tb_C12758
        ReDim C12758DataType(1)
        C12758DataType(1) = C14454DataType
        ReturnDataType = C12758DataType

        GoTo Exit_R1880

    Comp_C14373:

        ' SelSer1
        sCurrComponent = "SelSer1"
        QueryC14373 = con.CreateCommand()
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "SELECT WF_ACESSOS.Descr_Acesso "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & ""
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "FROM AKBUSER01.WF_ACESSOS , AKBUSER01.WF_PEDIDO_ITENS , AKBUSER01.WF_TAB_PRECO_CTO_QTDE , AKBUSER01.WF_PEDIDO "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & ""
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Preco_Digitado = 0 "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & ""
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & ""
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_ACESSOS.Sigla_Prod = WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_ACESSOS.Acesso = WF_PEDIDO_ITENS.Acesso "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & ""
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Tabela_Preco_Custo  (+) = " & _ 
ConvertValue(P3360, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "  "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Sigla_Prod (+) = WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Acesso  (+) = WF_PEDIDO_ITENS.Acesso "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Cod_Embalagem  (+) = WF_PEDIDO_ITENS.Cod_Embalagem "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Qtd_Inicial  (+) <= WF_PEDIDO_ITENS.Qtde_Pedido "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Qtd_Final  (+) >= WF_PEDIDO_ITENS.Qtde_Pedido "
        QueryC14373.CommandText = QueryC14373.CommandText & " " & ""
        QueryC14373.CommandText = QueryC14373.CommandText & " " & "AND WF_TAB_PRECO_CTO_QTDE.Acesso IS NULL"
        QueryC14373.Transaction = txn
        RsC14373 = QueryC14373.ExecuteReader()
        Dim iC14373 As Short
        ReDim C14373DataType(RsC14373.FieldCount)
        For iC14373 = 0 to RsC14373.FieldCount - 1
            Select Case RsC14373.GetDataTypeName(iC14373).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C14373DataType(iC14373 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C14373DataType(iC14373 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C14373DataType(iC14373 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC14373
        RsC14373_EOF = Not RsC14373.Read()

        GoTo Comp_C14453

    Comp_C14453:

        ' Eof2
        sCurrComponent = "Eof2"
        C14453DataType = 4
        C14453 = RsC14373_EOF
        GoTo Comp_C14454

    Comp_C14454:

        ' Eof2=T
        sCurrComponent = "Eof2=T"
        C14454 = (fn_ConvertValueCompiled(C14453, C14453DataType, AKB_DecimalPoint, False) =1)
        C14454DataType = AKBTypeConst.cAKBTypeLogical
        If C14454 Then
            GoTo Comp_C12758
        Else
            GoTo Comp_C14455
        End If

    Comp_C14455:

        ' SelSer2
        sCurrComponent = "SelSer2"
        QueryC14455 = con.CreateCommand()
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "SELECT WF_ACESSOS.Descr_Acesso "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "FROM AKBUSER01.WF_ACESSOS , AKBUSER01.WF_PEDIDO , AKBUSER01.WF_PEDIDO_ITENS ,AKBUSER01.WF_TAB_PRC_CUSTO_ACESSO "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "WHERE WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND   WF_PEDIDO.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO_ITENS.Preco_Digitado = 0 "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO_ITENS.Sigla_Prod = WF_ACESSOS.Sigla_Prod "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO_ITENS.Acesso = WF_ACESSOS.Acesso "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO_ITENS.Sigla_Prod = WF_TAB_PRC_CUSTO_ACESSO.Sigla_Prod  (+) "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO_ITENS.Acesso = WF_TAB_PRC_CUSTO_ACESSO.Acesso  (+) "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO_ITENS.Cod_Embalagem = WF_TAB_PRC_CUSTO_ACESSO.Cod_Embalagem  (+) "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_TAB_PRC_CUSTO_ACESSO.Tabela_Preco_Custo (+) = " & _ 
ConvertValue(P3360, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND WF_PEDIDO_ITENS.Seq_Item NOT IN "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "(SELECT WF_PEDIDO_ITENS.Seq_Item  "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "FROM AKBUSER01.WF_ACESSOS , AKBUSER01.WF_PEDIDO_ITENS , AKBUSER01.WF_TAB_PRECO_CTO_QTDE , AKBUSER01.WF_PEDIDO "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Preco_Digitado = 0 "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_ACESSOS.Sigla_Prod = WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_ACESSOS.Acesso = WF_PEDIDO_ITENS.Acesso "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Tabela_Preco_Custo  = " & _ 
ConvertValue(P3360, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Sigla_Prod  = WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Acesso  = WF_PEDIDO_ITENS.Acesso "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Cod_Embalagem  = WF_PEDIDO_ITENS.Cod_Embalagem "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Qtd_Inicial  <= WF_PEDIDO_ITENS.Qtde_Pedido "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_TAB_PRECO_CTO_QTDE.Qtd_Final   >= WF_PEDIDO_ITENS.Qtde_Pedido "
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ")"
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & "AND  WF_TAB_PRC_CUSTO_ACESSO.Acesso  IS NULL"
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.CommandText = QueryC14455.CommandText & " " & ""
        QueryC14455.Transaction = txn
        RsC14455 = QueryC14455.ExecuteReader()
        Dim iC14455 As Short
        ReDim C14455DataType(RsC14455.FieldCount)
        For iC14455 = 0 to RsC14455.FieldCount - 1
            Select Case RsC14455.GetDataTypeName(iC14455).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C14455DataType(iC14455 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C14455DataType(iC14455 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C14455DataType(iC14455 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC14455
        RsC14455_EOF = Not RsC14455.Read()

        GoTo Comp_C14456

    Comp_C14456:

        ' Eof3
        sCurrComponent = "Eof3"
        C14456DataType = 4
        C14456 = RsC14455_EOF
        GoTo Comp_C14457

    Comp_C14457:

        ' Eof3=T
        sCurrComponent = "Eof3=T"
        C14457 = (fn_ConvertValueCompiled(C14456, C14456DataType, AKB_DecimalPoint, False) =1)
        C14457DataType = AKBTypeConst.cAKBTypeLogical
        If C14457 Then
            GoTo Comp_C14458
        Else
            GoTo Comp_C14459
        End If

    Comp_C14458:

        ' RET3
        sCurrComponent = "RET3"
        Dim tb_C14458 As DataTable = New DataTable()
        tb_C14458.Columns.Add("Eof3=T" & "")
        Dim row_C14458 As DataRow = tb_C14458.NewRow()
        row_C14458(0) = C14457
        tb_C14458.Rows.Add(row_C14458)
        R1880 = tb_C14458
        ReDim C14458DataType(1)
        C14458DataType(1) = C14457DataType
        ReturnDataType = C14458DataType

        GoTo Exit_R1880

    Comp_C14459:

        ' MSG2
        sCurrComponent = "MSG2"
        C14459 = 1
        C14459DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C14458

    Comp_C28754:

        ' TpTab
        sCurrComponent = "TpTab"
        QueryC28754 = con.CreateCommand()
        QueryC28754.CommandText = QueryC28754.CommandText & " " & "SELECT NVL(WF_TP_PRODUTO.Tabela_Preco , 'Nenhum')"
        QueryC28754.CommandText = QueryC28754.CommandText & " " & "FROM AKBUSER01.WF_TP_PRODUTO "
        QueryC28754.CommandText = QueryC28754.CommandText & " " & "WHERE WF_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC28754.Transaction = txn
        RsC28754 = QueryC28754.ExecuteReader()
        Dim iC28754 As Short
        ReDim C28754DataType(RsC28754.FieldCount)
        For iC28754 = 0 to RsC28754.FieldCount - 1
            Select Case RsC28754.GetDataTypeName(iC28754).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C28754DataType(iC28754 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C28754DataType(iC28754 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C28754DataType(iC28754 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC28754
        RsC28754_EOF = Not RsC28754.Read()

        GoTo Comp_C35188

    Comp_C28755:

        ' TpTab=Venda
        sCurrComponent = "TpTab=Venda"
        C28755 = (fn_ConvertValueCompiled(RsC28754(0), C28754DataType(1), AKB_DecimalPoint, False) = "Venda")
        C28755DataType = AKBTypeConst.cAKBTypeLogical
        If C28755 Then
            GoTo Comp_C88409
        Else
            GoTo Comp_C28756
        End If

    Comp_C28756:

        ' TpTab=Serviço
        sCurrComponent = "TpTab=Serviço"
        C28756 = (fn_ConvertValueCompiled(RsC28754(0), C28754DataType(1), AKB_DecimalPoint, False) = "Serviço")
        C28756DataType = AKBTypeConst.cAKBTypeLogical
        If C28756 Then
            GoTo Comp_C14373
        Else
            GoTo Comp_C320109
        End If

    Comp_C35188:

        ' SelQualid
        sCurrComponent = "SelQualid"
        QueryC35188 = con.CreateCommand()
        QueryC35188.CommandText = QueryC35188.CommandText & " " & "SELECT COUNT (WF_PEDIDO_ITENS.Sigla_Prod)"
        QueryC35188.CommandText = QueryC35188.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS"
        QueryC35188.CommandText = QueryC35188.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC35188.CommandText = QueryC35188.CommandText & " " & ""
        QueryC35188.CommandText = QueryC35188.CommandText & " " & "AND NOT EXISTS"
        QueryC35188.CommandText = QueryC35188.CommandText & " " & " (SELECT * FROM AKBUSER01.WF_QUALID_ITEM_SIGLA_PROD "
        QueryC35188.CommandText = QueryC35188.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Sigla_Prod = WF_QUALID_ITEM_SIGLA_PROD.Sigla_Prod "
        QueryC35188.CommandText = QueryC35188.CommandText & " " & "AND WF_PEDIDO_ITENS.Qualid_Item_Estoque = WF_QUALID_ITEM_SIGLA_PROD.Qualid_Item_Estoque)"
        QueryC35188.Transaction = txn
        RsC35188 = QueryC35188.ExecuteReader()
        Dim iC35188 As Short
        ReDim C35188DataType(RsC35188.FieldCount)
        For iC35188 = 0 to RsC35188.FieldCount - 1
            Select Case RsC35188.GetDataTypeName(iC35188).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C35188DataType(iC35188 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C35188DataType(iC35188 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C35188DataType(iC35188 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC35188
        RsC35188_EOF = Not RsC35188.Read()

        GoTo Comp_C35190

    Comp_C35190:

        ' DESVIO1
        sCurrComponent = "DESVIO1"
        C35190 = (fn_ConvertValueCompiled(RsC35188(0), C35188DataType(1), AKB_DecimalPoint, False) = 0)
        C35190DataType = AKBTypeConst.cAKBTypeLogical
        If C35190 Then
            GoTo Comp_C177610
        Else
            GoTo Comp_C35191
        End If

    Comp_C35191:

        ' MSG3
        sCurrComponent = "MSG3"
        C35191 = 1
        C35191DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C35192

    Comp_C35192:

        ' RET4
        sCurrComponent = "RET4"
        Dim tb_C35192 As DataTable = New DataTable()
        tb_C35192.Columns.Add("DESVIO1" & "")
        Dim row_C35192 As DataRow = tb_C35192.NewRow()
        row_C35192(0) = C35190
        tb_C35192.Rows.Add(row_C35192)
        R1880 = tb_C35192
        ReDim C35192DataType(1)
        C35192DataType(1) = C35190DataType
        ReturnDataType = C35192DataType

        GoTo Exit_R1880

    Comp_C88206:

        ' Sel_Agup
        sCurrComponent = "Sel_Agup"
        QueryC88206 = con.CreateCommand()
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "Select WF_Pedido_Itens.Acesso, WF_Pedido_Itens.Sigla_Prod,"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "              WF_Pedido_Itens.COD_EMBALAGEM"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "from WF_Pedido_Itens"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "where not exists ( Select WF_TABPRECO_VDA_PRODUTO.Acesso,"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                                                 WF_TABPRECO_VDA_PRODUTO.Sigla_Prod, "
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                                                 WF_TABPRECO_VDA_PRODUTO.COD_EMBALAGEM"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                                  from WF_TABPRECO_VDA_PRODUTO "
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                                 where WF_TABPRECO_VDA_PRODUTO.Sigla_Prod = WF_Pedido_Itens.Sigla_Prod"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                                 and WF_TABPRECO_VDA_PRODUTO.Acesso = WF_Pedido_Itens.Acesso"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                                 and WF_TABPRECO_VDA_PRODUTO.COD_EMBALAGEM = WF_Pedido_Itens.COD_EMBALAGEM "
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                                and WF_TABPRECO_VDA_PRODUTO.TABELA_PRECO_VENDA = " & _ 
ConvertValue(P3334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "and 0 = ("
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                 Select NVL(NAO_VALIDA_AGRUPAMENTO, 0)"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                 from WF_TABELA_PRECO_VENDA"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "                 where TABELA_PRECO_VENDA = " & _ 
ConvertValue(P3334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & ""
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "and WF_Pedido_Itens.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "and WF_Pedido_Itens.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "and WF_Pedido_Itens.FLAG_POS_PED not in (7, 8, 10, 11, 12, 13)"
        QueryC88206.CommandText = QueryC88206.CommandText & " " & "and WF_Pedido_Itens.Pedido > 0"
        QueryC88206.Transaction = txn
        RsC88206 = QueryC88206.ExecuteReader()
        Dim iC88206 As Short
        ReDim C88206DataType(RsC88206.FieldCount)
        For iC88206 = 0 to RsC88206.FieldCount - 1
            Select Case RsC88206.GetDataTypeName(iC88206).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C88206DataType(iC88206 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C88206DataType(iC88206 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C88206DataType(iC88206 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC88206
        RsC88206_EOF = Not RsC88206.Read()

        GoTo Comp_C296825

    Comp_C88407:

        ' EOF SelAgrup = 1
        sCurrComponent = "EOF SelAgrup = 1"
        C88407 = (fn_ConvertValueCompiled(C296825, C296825DataType, AKB_DecimalPoint, False) = 1)
        C88407DataType = AKBTypeConst.cAKBTypeLogical
        If C88407 Then
            GoTo Comp_C12711
        Else
            GoTo Comp_C89647
        End If

    Comp_C88408:

        ' AtribTrue
        sCurrComponent = "AtribTrue"
        C88408DataType = 4
        C12740 = fn_ConvertValueCompiled(1, 1, AKB_DecimalPoint)
        C88408 = True
        GoTo Comp_C12748

    Comp_C88409:

        ' vMsg
        sCurrComponent = "vMsg"
        C88409 = System.DBNull.Value
        C88409DataType = 3
        GoTo Comp_C88206

    Comp_C88410:

        ' AtribAgrup
        sCurrComponent = "AtribAgrup"
        C88410DataType = 4
        C88409 = fn_ConvertValueCompiled(RsC89647(0) , 3, AKB_DecimalPoint)
        C88410 = True
        GoTo Comp_C12749

    Comp_C88411:

        ' AtribVDA
        sCurrComponent = "AtribVDA"
        C88411DataType = 4
        C88409(1) = fn_ConvertValueCompiled(RsC12711(0) , 3, AKB_DecimalPoint)
        C88411 = True
        GoTo Comp_C12749

    Comp_C89647:

        ' Sel_Descr_Acesso
        sCurrComponent = "Sel_Descr_Acesso"
        QueryC89647 = con.CreateCommand()
        QueryC89647.CommandText = QueryC89647.CommandText & " " & "Select WF_ACESSOS.Descr_Acesso "
        QueryC89647.CommandText = QueryC89647.CommandText & " " & "from AKBUSER01.WF_ACESSOS "
        QueryC89647.CommandText = QueryC89647.CommandText & " " & "where WF_ACESSOS.Acesso = " & _ 
ConvertValue(RsC88206(0), C88206DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC89647.Transaction = txn
        RsC89647 = QueryC89647.ExecuteReader()
        Dim iC89647 As Short
        ReDim C89647DataType(RsC89647.FieldCount)
        For iC89647 = 0 to RsC89647.FieldCount - 1
            Select Case RsC89647.GetDataTypeName(iC89647).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C89647DataType(iC89647 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C89647DataType(iC89647 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C89647DataType(iC89647 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC89647
        RsC89647_EOF = Not RsC89647.Read()

        GoTo Comp_C88410

    Comp_C177610:

        ' Sel_CustoFinan
        sCurrComponent = "Sel_CustoFinan"
        QueryC177610 = con.CreateCommand()
        QueryC177610.CommandText = QueryC177610.CommandText & " " & "SELECT NVL(WF_PEDIDO.Porc_Custo_Financ_Incluso,0)"
        QueryC177610.CommandText = QueryC177610.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO "
        QueryC177610.CommandText = QueryC177610.CommandText & " " & "WHERE WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P3337, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC177610.CommandText = QueryC177610.CommandText & " " & "AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P3333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC177610.Transaction = txn
        RsC177610 = QueryC177610.ExecuteReader()
        Dim iC177610 As Short
        ReDim C177610DataType(RsC177610.FieldCount)
        For iC177610 = 0 to RsC177610.FieldCount - 1
            Select Case RsC177610.GetDataTypeName(iC177610).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C177610DataType(iC177610 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C177610DataType(iC177610 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C177610DataType(iC177610 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC177610
        RsC177610_EOF = Not RsC177610.Read()

        GoTo Comp_C28755

    Comp_C296825:

        ' EOF SelAgrup
        sCurrComponent = "EOF SelAgrup"
        C296825DataType = 4
        C296825 = RsC88206_EOF
        GoTo Comp_C88407

    Comp_C320108:

        ' RET5
        sCurrComponent = "RET5"
        Dim tb_C320108 As DataTable = New DataTable()
        tb_C320108.Columns.Add("VAR1" & "")
        Dim row_C320108 As DataRow = tb_C320108.NewRow()
        row_C320108(0) = C320109
        tb_C320108.Rows.Add(row_C320108)
        R1880 = tb_C320108
        ReDim C320108DataType(1)
        C320108DataType(1) = C320109DataType
        ReturnDataType = C320108DataType

        GoTo Exit_R1880

    Comp_C320109:

        ' VAR1
        sCurrComponent = "VAR1"
        C320109 = 1
        C320109DataType = 4
        GoTo Comp_C320108

    Exit_R1880:

        Exit Function

    Err_R1880:
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
