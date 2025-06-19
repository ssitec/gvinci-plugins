Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R16268

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

    Public Shared Function R16268(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P58159 As String, ByRef P70895() As Object, ByVal P70896 As Object, ByRef P70897() As Object) As DataTable
    ' 
    ' 16268 - Gera Pedido e Carrega Grid Automático - Versão: 0
    ' 
        'On Error GoTo Err_R16268
        Dim sCurrComponent as String

        Dim C353064 As DataTable
        Dim C353064CurrentRow As Integer
        Dim C353064DataType() As Short
        Dim QueryC353081 As New Object
        Dim RsC353081 As New Object
        Dim C353081DataType() As Short
        Dim RsC353081_EOF As Boolean
        Dim C353126 As Object
        Dim C353126DataType As Short
        Dim C353127 As Object
        Dim C353127DataType As Short
        Dim C353128 As Object
        Dim C353128DataType As Short
        Dim C353129 As Object
        Dim C353129DataType As Short
        Dim C353130 As Boolean
        Dim C353130DataType As Short
        Dim C353175DataType() As Short
        Dim C353219 As Boolean
        Dim C353219DataType As Short
        Dim C353220 As DataTable
        Dim C353220CurrentRow As Integer
        Dim C353220DataType() As Short
        Dim C353228 As Object
        Dim C353228DataType As Short
        Dim QueryC353229 As New Object
        Dim RsC353229 As New Object
        Dim C353229DataType() As Short
        Dim RsC353229_EOF As Boolean
        Dim C353230 As Boolean
        Dim C353230DataType As Short
        Dim C353231 As Boolean
        Dim C353231DataType As Short
        Dim C353232 As Boolean
        Dim C353232DataType As Short
        Dim C353241 As Object
        Dim C353241DataType As Short
        Dim C353348( , ) As Object
        Dim C353348DataType() As Short
        Dim C353348CurrentRow As Integer
        Dim QueryC353506 As New Object
        Dim RsC353506 As New Object
        Dim C353506DataType() As Short
        Dim RsC353506_EOF As Boolean
        Dim C353507 As Object
        Dim C353507DataType As Short
        Dim C353508 As Boolean
        Dim C353508DataType As Short
        Dim C437834 As Object
        Dim C437834DataType As Short
        Dim C437835 As Boolean
        Dim C437835DataType As Short
        Dim C437836 As Object
        Dim C437836DataType As Short
        Dim C437839 As DataTable
        Dim C437839CurrentRow As Integer
        Dim C437839DataType() As Short
        Dim C447118 As Boolean
        Dim C447118DataType As Short
        Dim C447119 As Boolean
        Dim C447119DataType As Short
        Dim QueryC541200 As New Object
        Dim RsC541200 As New Object
        Dim C541200DataType() As Short
        Dim RsC541200_EOF As Boolean
        Dim C541201 As Boolean
        Dim C541201DataType As Short
        Dim C541202 As Short
        Dim C541202DataType As Short
        Dim C541202ReturnDataType() As Short

        Dim C579654 As Object
        Dim C579654DataType As Short
        Dim C579655 As Object
        Dim C579655DataType As Short
        Dim P70895Current As Integer

        Dim P70897Current As Integer

        ReDim ReturnDatatype(0)

        GoTo Comp_C541200

    Comp_C353064:

        ' Verificação/Bloqueio do Pré Pedido
        sCurrComponent = "Verificação/Bloqueio do Pré Pedido"
        C353064 = clsRuleDynamicallyCompiled_R15836.R15836(con, txn, msg, IIf(Not IsDBNull(C353129), C353129, System.DBNull.Value), IIf(Not IsDBNull(C353241), C353241, System.DBNull.Value))
        C353064CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C353064) Then
          iColumns = C353064.Columns.Count
        End If
        ReDim C353064DataType(iColumns)
        For iCol = 1 To iColumns
            C353064DataType(iCol) = clsRuleDynamicallyCompiled_R15836.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C353219

    Comp_C353081:

        ' SelPedidos
        sCurrComponent = "SelPedidos"
        QueryC353081 = con.CreateCommand()
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "SELECT  DISTINCT Id_PrePedido FROM ("
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "SELECT "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & " DISTINCT WF_PRE_PEDIDO.Id_PrePedido ,"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "dbms_random.random ID"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & ""
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "FROM "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "WF_PRE_PEDIDO ,"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "WF_CONDICAO_PGTO ,"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "WF_ZONAS, WF_CLIENTES ,"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "WF_PRE_PEDIDO_ITENS, "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "WF_TABELA_PRECO_VENDA,"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "WF_TAB_PRECO_CUSTO,"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "WF_REPRESENTANTE"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & ""
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "WHERE "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "        WF_PRE_PEDIDO.Cod_Pagto = WF_CONDICAO_PGTO.Cod_Pagto  (+) "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_PRE_PEDIDO.Cod_Cliente  = WF_CLIENTES.Cod_CLIENTE(+)"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_CLIENTES.Cod_Zona = WF_ZONAS.Cod_Zona (+) "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.ID_PREPEDIDO = WF_PRE_PEDIDO.ID_PREPEDIDO"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_REPRESENTANTE.Cod_Repres = WF_PRE_PEDIDO.Cod_Repres"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico IN (0" & _ 
ConvertValue(P58159, 5, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_TABELA_PRECO_VENDA.TABELA_PRECO_VENDA (+) = WF_PRE_PEDIDO.TABELA_PRECO_VENDA"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & ""
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_TAB_PRECO_CUSTO.TABELA_PRECO_CUSTO (+) = WF_PRE_PEDIDO.TABELA_PRECO_CUSTO"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & ""
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_PRE_PEDIDO.LIBERADO_REPRESENTANTE = 1 "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND NVL(WF_PRE_PEDIDO.Gerou_Pedido, 0) = 0"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & ""
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_PRE_PEDIDO.PROFORMA_CANCELADA = 0"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & " "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "AND WF_PRE_PEDIDO.INATIVO_PDA =0"
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "" & _ 
ConvertValue(C437839.rows(C437839CurrentRow - 1)(0), C437839DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "" & _ 
ConvertValue(C437834, C437834DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC353081.CommandText = QueryC353081.CommandText & " " & "ORDER BY 2)"
        QueryC353081.Transaction = txn
        RsC353081 = QueryC353081.ExecuteReader()
        Dim iC353081 As Short
        ReDim C353081DataType(RsC353081.FieldCount)
        For iC353081 = 0 to RsC353081.FieldCount - 1
            Select Case RsC353081.GetDataTypeName(iC353081).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C353081DataType(iC353081 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C353081DataType(iC353081 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C353081DataType(iC353081 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC353081
        RsC353081_EOF = Not RsC353081.Read()

        GoTo Comp_C353126

    Comp_C353126:

        ' PrimeiroValor
        sCurrComponent = "PrimeiroValor"
        C353126DataType = 4

        GoTo Comp_C353127

    Comp_C353127:

        ' FimValores
        sCurrComponent = "FimValores"
        C353127DataType = 4
        C353127 = RsC353081_EOF
        GoTo Comp_C353130

    Comp_C353128:

        ' PróximoValor
        sCurrComponent = "PróximoValor"
        C353128DataType = 4
        RsC353081_EOF = Not RsC353081.Read()
        C353128 = True

        GoTo Comp_C353127

    Comp_C353129:

        ' Id_PrePedido
        sCurrComponent = "Id_PrePedido"
        C353129DataType = 0
        C353129 = RsC353081(0)
        C353129DataType = C353081Datatype(1)
        If C353129DataType = AKBTypeConst.cAKBTypeString Then
          C353129 = IIF(IsDBNull(C353129), C353129, Strings.RTrim(C353129))
        End If 

        GoTo Comp_C353230

    Comp_C353130:

        ' Fim?
        sCurrComponent = "Fim?"
        C353130 = (fn_ConvertValueCompiled(C353127, C353127DataType, AKB_DecimalPoint, False) =1)
        C353130DataType = AKBTypeConst.cAKBTypeLogical
        If C353130 Then
            GoTo Comp_C353175
        Else
            GoTo Comp_C353129
        End If

    Comp_C353175:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C353175 As DataTable = New DataTable()
        tb_C353175.Columns.Add("RetornoTexto" & "")
        tb_C353175.Columns.Add("RetornoIdLog" & "")
        ReDim ReturnDatatype(C353348.GetLength(1))
        For j As Integer = 1 To C353348.GetLength(0) - 1
            Dim row_C353175 As DataRow = tb_C353175.NewRow()
            Dim iC353175 As Short
            For iC353175 = 0 To 1
                row_C353175(iC353175) = C353348(j, iC353175)
                ReturnDatatype(iC353175 + 1) = C353348DataType(iC353175)
            Next iC353175
            tb_C353175.Rows.Add(row_C353175)
        Next j
        R16268 = tb_C353175

        GoTo Exit_R16268

    Comp_C353219:

        ' NaoBloquear?
        sCurrComponent = "NaoBloquear?"
        C353219 = (fn_ConvertValueCompiled(C353064.rows(C353064CurrentRow - 1)(0), C353064DataType(1), AKB_DecimalPoint, False) = 0)
        C353219DataType = AKBTypeConst.cAKBTypeLogical
        If C353219 Then
            GoTo Comp_C353231
        Else
            GoTo Comp_C353220
        End If

    Comp_C353220:

        ' Gerar Pedido do Pré Pedido - 1 Pedido
        sCurrComponent = "Gerar Pedido do Pré Pedido - 1 Pedido"
        C353220 = clsRuleDynamicallyCompiled_R4557.R4557(con, txn, msg, IIf(Not IsDBNull(C353129), C353129, System.DBNull.Value), System.DBNull.Value, System.DBNull.Value, fn_ConvertValueCompiled( "0", 4, AKB_DecimalPoint), System.DBNull.Value, fn_ConvertValueCompiled( "0", 1, AKB_DecimalPoint), System.DBNull.Value, IIf(Not IsDBNull(P58159), P58159, System.DBNull.Value), System.DBNull.Value, System.DBNull.Value, System.DBNull.Value)
        C353220CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C353220) Then
          iColumns = C353220.Columns.Count
        End If
        ReDim C353220DataType(iColumns)
        For iCol = 1 To iColumns
            C353220DataType(iCol) = clsRuleDynamicallyCompiled_R4557.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C447118

    Comp_C353228:

        ' varRetorno
        sCurrComponent = "varRetorno"
        C353228 = "Pronto."
        C353228DataType = 3
        GoTo Comp_C353506

    Comp_C353229:

        ' SelPrePedido
        sCurrComponent = "SelPrePedido"
        QueryC353229 = con.CreateCommand()
        QueryC353229.CommandText = QueryC353229.CommandText & " " & "SELECT "
        QueryC353229.CommandText = QueryC353229.CommandText & " " & "WF_PRE_PEDIDO.GEROU_PEDIDO"
        QueryC353229.CommandText = QueryC353229.CommandText & " " & ""
        QueryC353229.CommandText = QueryC353229.CommandText & " " & "FROM "
        QueryC353229.CommandText = QueryC353229.CommandText & " " & "WF_PRE_PEDIDO "
        QueryC353229.CommandText = QueryC353229.CommandText & " " & ""
        QueryC353229.CommandText = QueryC353229.CommandText & " " & "WHERE WF_PRE_PEDIDO.ID_PREPEDIDO = " & _ 
ConvertValue(C353129, C353129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC353229.CommandText = QueryC353229.CommandText & " " & "FOR UPDATE "
        QueryC353229.Transaction = txn
        RsC353229 = QueryC353229.ExecuteReader()
        Dim iC353229 As Short
        ReDim C353229DataType(RsC353229.FieldCount)
        For iC353229 = 0 to RsC353229.FieldCount - 1
            Select Case RsC353229.GetDataTypeName(iC353229).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C353229DataType(iC353229 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C353229DataType(iC353229 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C353229DataType(iC353229 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC353229
        RsC353229_EOF = Not RsC353229.Read()

        GoTo Comp_C353232

    Comp_C353230:

        ' Início
        sCurrComponent = "Início"
        txn = con.BeginTransaction
        C353230 = True
        C353230DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C353229

    Comp_C353231:

        ' Fim
        sCurrComponent = "Fim"
        txn.Commit()
        C353231 = True
        C353231DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C353128

    Comp_C353232:

        ' Pedido Gerado?
        sCurrComponent = "Pedido Gerado?"
        C353232 = (fn_ConvertValueCompiled(RsC353229(0), C353229DataType(1), AKB_DecimalPoint, False)  = 1)
        C353232DataType = AKBTypeConst.cAKBTypeLogical
        If C353232 Then
            GoTo Comp_C353231
        Else
            GoTo Comp_C353064
        End If

    Comp_C353241:

        ' IdentLog
        sCurrComponent = "IdentLog"
        C353241DataType = 1
        C353241 = ObjTable_NewIdentity (con, txn, "WF_PRE_PEDIDO_LOG_BLOQ")
        GoTo Comp_C579654

    Comp_C353348:

        ' matrizRetorno
        sCurrComponent = "matrizRetorno"
        ReDim C353348(1, 3)
        ReDim C353348DataType(2)
        C353348DataType(1) = 3
        C353348(1, 1) = "0"
        C353348DataType(2) = 1
        C353348(2, 1) = 0
        C353348CurrentRow = 1

        GoTo Comp_C541201

    Comp_C353506:

        ' SelGeraPedidoAutomatico
        sCurrComponent = "SelGeraPedidoAutomatico"
        QueryC353506 = con.CreateCommand()
        QueryC353506.CommandText = QueryC353506.CommandText & " " & "SELECT MAX(WF_ESTAB_JURIDICO_PARAM.GERAR_PED_PRE_PED_AUT2)"
        QueryC353506.CommandText = QueryC353506.CommandText & " " & "FROM WF_ESTAB_JURIDICO_PARAM"
        QueryC353506.CommandText = QueryC353506.CommandText & " " & "WHERE WF_ESTAB_JURIDICO_PARAM.COD_PESSOA_ESTAB_JURIDICO IN (0" & _ 
ConvertValue(P58159, 5, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") "
        QueryC353506.CommandText = QueryC353506.CommandText & " " & ""
        QueryC353506.Transaction = txn
        RsC353506 = QueryC353506.ExecuteReader()
        Dim iC353506 As Short
        ReDim C353506DataType(RsC353506.FieldCount)
        For iC353506 = 0 to RsC353506.FieldCount - 1
            Select Case RsC353506.GetDataTypeName(iC353506).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C353506DataType(iC353506 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C353506DataType(iC353506 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C353506DataType(iC353506 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC353506
        RsC353506_EOF = Not RsC353506.Read()

        GoTo Comp_C353507

    Comp_C353507:

        ' vFlagGera
        sCurrComponent = "vFlagGera"
        C353507DataType = 0
        C353507 = RsC353506(0)
        C353507DataType = C353506Datatype(1)
        If C353507DataType = AKBTypeConst.cAKBTypeString Then
          C353507 = IIF(IsDBNull(C353507), C353507, Strings.RTrim(C353507))
        End If 

        GoTo Comp_C353241

    Comp_C353508:

        ' Gera Pedido Automático?
        sCurrComponent = "Gera Pedido Automático?"
        C353508 = (fn_ConvertValueCompiled(C353507, C353507DataType, AKB_DecimalPoint, False) =1)
        C353508DataType = AKBTypeConst.cAKBTypeLogical
        If C353508 Then
            GoTo Comp_C353081
        Else
            GoTo Comp_C353175
        End If

    Comp_C437834:

        ' QueryCelula
        sCurrComponent = "QueryCelula"
        C437834 = "AND 1 = 1"
        C437834DataType = 5
        GoTo Comp_C437835

    Comp_C437835:

        ' Celula>0?
        sCurrComponent = "Celula>0?"
        C437835 = (fn_ConvertValueCompiled(P70896, 1, AKB_DecimalPoint, False) > 0)
        C437835DataType = AKBTypeConst.cAKBTypeLogical
        If C437835 Then
            GoTo Comp_C437836
        Else
            GoTo Comp_C437839
        End If

    Comp_C437836:

        ' ATRIBUIVA1
        sCurrComponent = "ATRIBUIVA1"
        C437836DataType = 4
        C437834 = fn_ConvertValueCompiled("AND WF_REPRESENTANTE.Celula = " & P70896 , 5, AKB_DecimalPoint)
        C437836 = True
        GoTo Comp_C437839

    Comp_C437839:

        ' ListaDpto
        sCurrComponent = "ListaDpto"
        C437839 = clsRuleDynamicallyCompiled_R14036.R14036(con, txn, msg, P70895, P70897)
        C437839CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C437839) Then
          iColumns = C437839.Columns.Count
        End If
        ReDim C437839DataType(iColumns)
        For iCol = 1 To iColumns
            C437839DataType(iCol) = clsRuleDynamicallyCompiled_R14036.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C353228

    Comp_C447118:

        ' DESVIO1
        sCurrComponent = "DESVIO1"
        C447118 = (fn_ConvertValueCompiled(C353220.rows(C353220CurrentRow - 1)(0), C353220DataType(1), AKB_DecimalPoint, False) = 1)
        C447118DataType = AKBTypeConst.cAKBTypeLogical
        If C447118 Then
            GoTo Comp_C353231
        Else
            GoTo Comp_C447119
        End If

    Comp_C447119:

        ' CancTrans
        sCurrComponent = "CancTrans"
        txn.Rollback()
        C447119 = True
        C447119DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C353128

    Comp_C541200:

        ' SelUtilizaWS
        sCurrComponent = "SelUtilizaWS"
        QueryC541200 = con.CreateCommand()
        QueryC541200.CommandText = QueryC541200.CommandText & " " & "SELECT NVL(utiliza_sintegraws,0)"
        QueryC541200.CommandText = QueryC541200.CommandText & " " & "FROM wf_estab_juridico"
        QueryC541200.CommandText = QueryC541200.CommandText & " " & "WHERE  cod_pessoa_estab_juridico = " & _ 
ConvertValue(P58159, 5, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC541200.Transaction = txn
        RsC541200 = QueryC541200.ExecuteReader()
        Dim iC541200 As Short
        ReDim C541200DataType(RsC541200.FieldCount)
        For iC541200 = 0 to RsC541200.FieldCount - 1
            Select Case RsC541200.GetDataTypeName(iC541200).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C541200DataType(iC541200 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C541200DataType(iC541200 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C541200DataType(iC541200 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC541200
        RsC541200_EOF = Not RsC541200.Read()

        GoTo Comp_C353348

    Comp_C541201:

        ' UtilizaWS?
        sCurrComponent = "UtilizaWS?"
        C541201 = (fn_ConvertValueCompiled(RsC541200(0), C541200DataType(1), AKB_DecimalPoint, False) = 1)
        C541201DataType = AKBTypeConst.cAKBTypeLogical
        If C541201 Then
            GoTo Comp_C541202
        Else
            GoTo Comp_C437834
        End If

    Comp_C541202:

        ' MsgSintegraWS
        sCurrComponent = "MsgSintegraWS"
        Dim row_C541202 As DataRow = msg.NewRow()
        row_C541202(0) = "Não pode usar a geração automática, pois o estabelecimento utiliza a consulta automática do Sintegra." & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "Favor gerar os pedidos usando os botões: Gerar Pedido - Apenas 1 ou selecione os pedidos a gerar no grid e use o botão Gerar Pedido - Todos." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C541202)
        C541202DataType = AKBTypeConst.cAKBTypeNumeric
        C541202 = 1

        GoTo Comp_C353175

    Comp_C579654:

        ' ATRIBUIVA2
        sCurrComponent = "ATRIBUIVA2"
        C579654DataType = 4
        C353348(1, C353348CurrentRow) = fn_ConvertValueCompiled(C353228 , 3, AKB_DecimalPoint)
        C579654 = True
        GoTo Comp_C579655

    Comp_C579655:

        ' ATRIBUIVA3
        sCurrComponent = "ATRIBUIVA3"
        C579655DataType = 4
        C353348(2, C353348CurrentRow) = fn_ConvertValueCompiled(C353241 , 1, AKB_DecimalPoint)
        C579655 = True
        GoTo Comp_C353508

    Exit_R16268:

        Exit Function

    Err_R16268:
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
