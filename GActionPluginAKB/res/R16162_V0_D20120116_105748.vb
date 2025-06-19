Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R16162

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

    Public Shared Function R16162(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P57496 As Double) As DataTable
    ' 
    ' 16162 - Regras do Pré Pedido - Versão: 0
    ' 
        On Error GoTo Err_R16162
        Dim sCurrComponent as String

        Dim QueryC346951 As New Object
        Dim RsC346951 As New Object
        Dim C346951DataType() As Short
        Dim RsC346951_EOF As Boolean
        Dim C346952 As Object
        Dim C346952DataType As Short
        Dim QueryC346961 As New Object
        Dim RsC346961 As New Object
        Dim C346961DataType() As Short
        Dim RsC346961_EOF As Boolean
        Dim C346963 As Object
        Dim C346963DataType As Short
        Dim C346964 As Object
        Dim C346964DataType As Short
        Dim C346965 As Boolean
        Dim C346965DataType As Short
        Dim C346966 As Object
        Dim C346966DataType As Short
        Dim QueryC346967 As New Object
        Dim C346967 As Integer
        Dim C346967DataType As Short
        Dim C346971 As Boolean
        Dim C346971DataType As Short
        Dim C346972DataType() As Short
        Dim QueryC346973 As New Object
        Dim C346973 As Integer
        Dim C346973DataType As Short
        Dim C347075 As DataTable
        Dim C347075CurrentRow As Integer
        Dim C347075DataType() As Short
        Dim C347076 As Object
        Dim C347076DataType As Short
        Dim C347077 As Boolean
        Dim C347077DataType As Short
        Dim C347078 As DataTable
        Dim C347078CurrentRow As Integer
        Dim C347078DataType() As Short
        Dim C347079 As Object
        Dim C347079DataType As Short
        Dim C347080 As Boolean
        Dim C347080DataType As Short
        Dim C347081 As DataTable
        Dim C347081CurrentRow As Integer
        Dim C347081DataType() As Short
        Dim C347082 As Object
        Dim C347082DataType As Short
        Dim C347083 As Object
        Dim C347083DataType As Short
        Dim C347084 As Object
        Dim C347084DataType As Short
        Dim C347085 As Object
        Dim C347085DataType As Short
        Dim C347086 As Boolean
        Dim C347086DataType As Short
        Dim C347087 As DataTable
        Dim C347087CurrentRow As Integer
        Dim C347087DataType() As Short
        Dim C347088 As Object
        Dim C347088DataType As Short
        Dim C347089 As Object
        Dim C347089DataType As Short
        Dim C347090 As Boolean
        Dim C347090DataType As Short
        Dim C347091 As DataTable
        Dim C347091CurrentRow As Integer
        Dim C347091DataType() As Short
        Dim C347092 As Boolean
        Dim C347092DataType As Short
        Dim C347093 As Object
        Dim C347093DataType As Short
        Dim C347094 As DataTable
        Dim C347094CurrentRow As Integer
        Dim C347094DataType() As Short
        Dim C347095 As Object
        Dim C347095DataType As Short
        Dim C347097 As Boolean
        Dim C347097DataType As Short
        Dim C347098 As Object
        Dim C347098DataType As Short
        Dim C347152 As Boolean
        Dim C347152DataType As Short
        Dim C347153 As Object
        Dim C347153DataType As Short
        Dim QueryC350045 As New Object
        Dim C350045 As Integer
        Dim C350045DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C346952

    Comp_C346951:

        ' selDadosPrePedido
        sCurrComponent = "selDadosPrePedido"
        QueryC346951 = con.CreateCommand()
        QueryC346951.CommandText = QueryC346951.CommandText & " " & "SELECT WF_PRE_PEDIDO.Cod_Cliente, WF_PRE_PEDIDO.Tp_Produto, WF_PRE_PEDIDO.Tabela_Preco_Venda , WF_PRE_PEDIDO.Tabela_Preco_Custo , WF_PRE_PEDIDO.Prazo_Entr_Cliente , WF_PRE_PEDIDO.Porc_Custo_Financ_Incluso , WF_PRE_PEDIDO.Cod_Pagto , WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico , WF_PRE_PEDIDO.Dt_Pedido , WF_PRE_PEDIDO.Porc_Custo_Financ , WF_PRE_PEDIDO.Cod_Repres "
        QueryC346951.CommandText = QueryC346951.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO"
        QueryC346951.CommandText = QueryC346951.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P57496, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC346951.Transaction = txn
        RsC346951 = QueryC346951.ExecuteReader()
        Dim iC346951 As Short
        ReDim C346951DataType(RsC346951.FieldCount)
        For iC346951 = 0 to RsC346951.FieldCount - 1
            Select Case RsC346951.GetDataTypeName(iC346951).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C346951DataType(iC346951 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C346951DataType(iC346951 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C346951DataType(iC346951 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC346951
        RsC346951_EOF = Not RsC346951.Read()

        GoTo Comp_C346963

    Comp_C346952:

        ' vRet
        sCurrComponent = "vRet"
        C346952 = 0
        C346952DataType = 4
        GoTo Comp_C346951

    Comp_C346961:

        ' selPorcAdicionalCliente
        sCurrComponent = "selPorcAdicionalCliente"
        QueryC346961 = con.CreateCommand()
        QueryC346961.CommandText = QueryC346961.CommandText & " " & "SELECT NVL(WF_CLIENTES.PORC_ADICIONAL_QTDE_PEDIDO,0) , NVL(WF_CLIENTES.Nao_Aceita_Qtde_Fat_Maior_Ped,0)"
        QueryC346961.CommandText = QueryC346961.CommandText & " " & "FROM AKBUSER01.WF_CLIENTES "
        QueryC346961.CommandText = QueryC346961.CommandText & " " & "WHERE WF_CLIENTES.Cod_Cliente = " & _ 
ConvertValue(C346963, C346963DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC346961.Transaction = txn
        RsC346961 = QueryC346961.ExecuteReader()
        Dim iC346961 As Short
        ReDim C346961DataType(RsC346961.FieldCount)
        For iC346961 = 0 to RsC346961.FieldCount - 1
            Select Case RsC346961.GetDataTypeName(iC346961).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C346961DataType(iC346961 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C346961DataType(iC346961 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C346961DataType(iC346961 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC346961
        RsC346961_EOF = Not RsC346961.Read()

        GoTo Comp_C346966

    Comp_C346963:

        ' vCodCliente
        sCurrComponent = "vCodCliente"
        C346963DataType = 0
        C346963 = RsC346951(0)
        C346963DataType = C346951Datatype(1)

        GoTo Comp_C347076

    Comp_C346964:

        ' vNãoAceitaPorcAdicionalCliente
        sCurrComponent = "vNãoAceitaPorcAdicionalCliente"
        C346964DataType = 0
        C346964 = RsC346961(1)
        C346964DataType = C346961Datatype(2)

        GoTo Comp_C346965

    Comp_C346965:

        ' Aceita PorcAdicionalCliente?
        sCurrComponent = "Aceita PorcAdicionalCliente?"
        C346965 = (fn_ConvertValueCompiled(C346964, C346964DataType, AKB_DecimalPoint, False) = 0)
        C346965DataType = AKBTypeConst.cAKBTypeLogical
        If C346965 Then
            GoTo Comp_C346971
        Else
            GoTo Comp_C350045
        End If

    Comp_C346966:

        ' vPorcAdicionalCliente
        sCurrComponent = "vPorcAdicionalCliente"
        C346966DataType = 0
        C346966 = RsC346961(0)
        C346966DataType = C346961Datatype(1)

        GoTo Comp_C346964

    Comp_C346967:

        ' upPorcAdicionalCliente
        sCurrComponent = "upPorcAdicionalCliente"
        QueryC346967 = con.CreateCommand()
        QueryC346967.CommandText = QueryC346967.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC346967.CommandText = QueryC346967.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.PORC_ADICIONAL_QTDE_PEDIDO = " & _ 
ConvertValue(C346966, C346966DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC346967.CommandText = QueryC346967.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P57496, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        C346967 = QueryC346967.ExecuteNonQuery()
        C346967DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C347075

    Comp_C346971:

        ' Tem PorcAdicionalCliente?
        sCurrComponent = "Tem PorcAdicionalCliente?"
        C346971 = (fn_ConvertValueCompiled(C346966, C346966DataType, AKB_DecimalPoint, False) > 0)
        C346971DataType = AKBTypeConst.cAKBTypeLogical
        If C346971 Then
            GoTo Comp_C346967
        Else
            GoTo Comp_C346973
        End If

    Comp_C346972:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C346972 As DataTable = New DataTable()
        tb_C346972.Columns.Add("vRet" & "")
        Dim row_C346972 As DataRow = tb_C346972.NewRow()
        row_C346972(0) = C346952
        tb_C346972.Rows.Add(row_C346972)
        R16162 = tb_C346972
        ReDim C346972DataType(1)
        C346972DataType(1) = C346952DataType
        ReturnDataType = C346972DataType

        GoTo Exit_R16162

    Comp_C346973:

        ' upPorcAdicionalSigla
        sCurrComponent = "upPorcAdicionalSigla"
        QueryC346973 = con.CreateCommand()
        QueryC346973.CommandText = QueryC346973.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC346973.CommandText = QueryC346973.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.PORC_ADICIONAL_QTDE_PEDIDO = "
        QueryC346973.CommandText = QueryC346973.CommandText & " " & "(SELECT WF_SIGLA_PROD_AGRUP.Porc_Adicional_Qtde_Pedidos"
        QueryC346973.CommandText = QueryC346973.CommandText & " " & "FROM AKBUSER01.WF_SIGLA_PROD_AGRUP "
        QueryC346973.CommandText = QueryC346973.CommandText & " " & "WHERE WF_SIGLA_PROD_AGRUP.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC346973.CommandText = QueryC346973.CommandText & " " & "AND NVL(WF_SIGLA_PROD_AGRUP.Nao_Aceita_Qtde_Fat_Maior_Ped,0) = 0)"
        QueryC346973.CommandText = QueryC346973.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P57496, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        C346973 = QueryC346973.ExecuteNonQuery()
        C346973DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C347075

    Comp_C347075:

        ' Atual_desconto_Pedido
        sCurrComponent = "Atual_desconto_Pedido"
        'C347075 = clsRuleDynamicallyCompiled_R5779.R5779(con, txn, IIf(Not IsDBNull(P57496), P57496, System.DBNull.Value), IIf(Not IsDBNull(C347076), C347076, System.DBNull.Value))
        C347075CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C347075) Then
          iColumns = C347075.Columns.Count
        End If
        ReDim C347075DataType(iColumns)
        For iCol = 1 To iColumns
            'C347075DataType(iCol) = clsRuleDynamicallyCompiled_R5779.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C347077

    Comp_C347076:

        ' vTipoProduto
        sCurrComponent = "vTipoProduto"
        C347076DataType = 0
        C347076 = RsC346951(1)
        C347076DataType = C346951Datatype(2)

        GoTo Comp_C347079

    Comp_C347077:

        ' Atual_desconto_Pedido = 1
        sCurrComponent = "Atual_desconto_Pedido = 1"
        C347077 = (fn_ConvertValueCompiled(C347075.rows(0)(C347075CurrentRow - 1), C347075DataType(1), AKB_DecimalPoint, False) = 1)
        C347077DataType = AKBTypeConst.cAKBTypeLogical
        If C347077 Then
            GoTo Comp_C347078
        Else
            GoTo Comp_C346972
        End If

    Comp_C347078:

        ' Verifica_Preço_Digitado
        sCurrComponent = "Verifica_Preço_Digitado"
        'C347078 = clsRuleDynamicallyCompiled_R9825.R9825(con, txn, IIf(Not IsDBNull(P57496), P57496, System.DBNull.Value), IIf(Not IsDBNull(C347079), C347079, System.DBNull.Value), IIf(Not IsDBNull(C347076), C347076, System.DBNull.Value), System.DBNull.Value, IIf(Not IsDBNull(C347082), C347082, System.DBNull.Value))
        C347078CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C347078) Then
          iColumns = C347078.Columns.Count
        End If
        ReDim C347078DataType(iColumns)
        For iCol = 1 To iColumns
            'C347078DataType(iCol) = clsRuleDynamicallyCompiled_R9825.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C347080

    Comp_C347079:

        ' vTabelaVenda
        sCurrComponent = "vTabelaVenda"
        C347079DataType = 0
        C347079 = RsC346951(2)
        C347079DataType = C346951Datatype(3)

        GoTo Comp_C347082

    Comp_C347080:

        ' Verifica_Preço_Digitado = 1
        sCurrComponent = "Verifica_Preço_Digitado = 1"
        C347080 = (fn_ConvertValueCompiled(C347078.rows(0)(C347078CurrentRow - 1), C347078DataType(1), AKB_DecimalPoint, False) = 1)
        C347080DataType = AKBTypeConst.cAKBTypeLogical
        If C347080 Then
            GoTo Comp_C347081
        Else
            GoTo Comp_C346972
        End If

    Comp_C347081:

        ' Atualiza_Total_Itens_EDI
        sCurrComponent = "Atualiza_Total_Itens_EDI"
        'C347081 = clsRuleDynamicallyCompiled_R5783.R5783(con, txn, IIf(Not IsDBNull(C347079), C347079, System.DBNull.Value), IIf(Not IsDBNull(C347082), C347082, System.DBNull.Value), IIf(Not IsDBNull(C347084), C347084, System.DBNull.Value), IIf(Not IsDBNull(C347083), C347083, System.DBNull.Value), IIf(Not IsDBNull(C347085), C347085, System.DBNull.Value), IIf(Not IsDBNull(P57496), P57496, System.DBNull.Value), IIf(Not IsDBNull(C347076), C347076, System.DBNull.Value))
        C347081CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C347081) Then
          iColumns = C347081.Columns.Count
        End If
        ReDim C347081DataType(iColumns)
        For iCol = 1 To iColumns
            'C347081DataType(iCol) = clsRuleDynamicallyCompiled_R5783.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C347086

    Comp_C347082:

        ' vTabelaCusto
        sCurrComponent = "vTabelaCusto"
        C347082DataType = 0
        C347082 = RsC346951(3)
        C347082DataType = C346951Datatype(4)

        GoTo Comp_C347083

    Comp_C347083:

        ' vPrazoCliente
        sCurrComponent = "vPrazoCliente"
        C347083DataType = 0
        C347083 = RsC346951(4)
        C347083DataType = C346951Datatype(5)

        GoTo Comp_C347084

    Comp_C347084:

        ' vCustoItens
        sCurrComponent = "vCustoItens"
        C347084DataType = 0
        C347084 = RsC346951(5)
        C347084DataType = C346951Datatype(6)

        GoTo Comp_C347085

    Comp_C347085:

        ' vCondPagto
        sCurrComponent = "vCondPagto"
        C347085DataType = 0
        C347085 = RsC346951(6)
        C347085DataType = C346951Datatype(7)

        GoTo Comp_C347088

    Comp_C347086:

        ' Atualiza_Total_Itens_EDI = 1
        sCurrComponent = "Atualiza_Total_Itens_EDI = 1"
        C347086 = (fn_ConvertValueCompiled(C347081.rows(0)(C347081CurrentRow - 1), C347081DataType(1), AKB_DecimalPoint, False) = 1)
        C347086DataType = AKBTypeConst.cAKBTypeLogical
        If C347086 Then
            GoTo Comp_C347087
        Else
            GoTo Comp_C346972
        End If

    Comp_C347087:

        ' Verifica_Vencimentos
        sCurrComponent = "Verifica_Vencimentos"
        'C347087 = clsRuleDynamicallyCompiled_R10095.R10095(con, txn, IIf(Not IsDBNull(P57496), P57496, System.DBNull.Value), IIf(Not IsDBNull(C347088), C347088, System.DBNull.Value), IIf(Not IsDBNull(C347089), C347089, System.DBNull.Value))
        C347087CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C347087) Then
          iColumns = C347087.Columns.Count
        End If
        ReDim C347087DataType(iColumns)
        For iCol = 1 To iColumns
            'C347087DataType(iCol) = clsRuleDynamicallyCompiled_R10095.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C347090

    Comp_C347088:

        ' vCodEstab
        sCurrComponent = "vCodEstab"
        C347088DataType = 0
        C347088 = RsC346951(7)
        C347088DataType = C346951Datatype(8)

        GoTo Comp_C347089

    Comp_C347089:

        ' vDataPedido
        sCurrComponent = "vDataPedido"
        C347089DataType = 0
        C347089 = RsC346951(8)
        C347089DataType = C346951Datatype(9)

        GoTo Comp_C347093

    Comp_C347090:

        ' Verifica_Vencimentos = 1
        sCurrComponent = "Verifica_Vencimentos = 1"
        C347090 = (fn_ConvertValueCompiled(C347087.rows(0)(C347087CurrentRow - 1), C347087DataType(1), AKB_DecimalPoint, False) = 1)
        C347090DataType = AKBTypeConst.cAKBTypeLogical
        If C347090 Then
            GoTo Comp_C347091
        Else
            GoTo Comp_C346972
        End If

    Comp_C347091:

        ' Zera_Custo_Financeiro
        sCurrComponent = "Zera_Custo_Financeiro"
        'C347091 = clsRuleDynamicallyCompiled_R6487.R6487(con, txn, IIf(Not IsDBNull(C347093), C347093, System.DBNull.Value), IIf(Not IsDBNull(P57496), P57496, System.DBNull.Value))
        C347091CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C347091) Then
          iColumns = C347091.Columns.Count
        End If
        ReDim C347091DataType(iColumns)
        For iCol = 1 To iColumns
            'C347091DataType(iCol) = clsRuleDynamicallyCompiled_R6487.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C347092

    Comp_C347092:

        ' Zera_Custo_Financeiro = 1
        sCurrComponent = "Zera_Custo_Financeiro = 1"
        C347092 = (fn_ConvertValueCompiled(C347091.rows(0)(C347091CurrentRow - 1), C347091DataType(1), AKB_DecimalPoint, False) = 1)
        C347092DataType = AKBTypeConst.cAKBTypeLogical
        If C347092 Then
            GoTo Comp_C347094
        Else
            GoTo Comp_C346972
        End If

    Comp_C347093:

        ' vPorcCustoFinanceiro
        sCurrComponent = "vPorcCustoFinanceiro"
        C347093DataType = 0
        C347093 = RsC346951(9)
        C347093DataType = C346951Datatype(10)

        GoTo Comp_C347095

    Comp_C347094:

        ' Verifica_Email_Repres_Cliente
        sCurrComponent = "Verifica_Email_Repres_Cliente"
        'C347094 = clsRuleDynamicallyCompiled_R14926.R14926(con, txn, IIf(Not IsDBNull(C347076), C347076, System.DBNull.Value), IIf(Not IsDBNull(C347095), C347095, System.DBNull.Value), IIf(Not IsDBNull(C346963), C346963, System.DBNull.Value), IIf(Not IsDBNull(C347088), C347088, System.DBNull.Value))
        C347094CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C347094) Then
          iColumns = C347094.Columns.Count
        End If
        ReDim C347094DataType(iColumns)
        For iCol = 1 To iColumns
            'C347094DataType(iCol) = clsRuleDynamicallyCompiled_R14926.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C347097

    Comp_C347095:

        ' vCodRepres
        sCurrComponent = "vCodRepres"
        C347095DataType = 0
        C347095 = RsC346951(10)
        C347095DataType = C346951Datatype(11)

        GoTo Comp_C347152

    Comp_C347097:

        ' Verifica_Email_Repres_Cliente = 1
        sCurrComponent = "Verifica_Email_Repres_Cliente = 1"
        C347097 = (fn_ConvertValueCompiled(C347094.rows(0)(C347094CurrentRow - 1), C347094DataType(1), AKB_DecimalPoint, False) = 1)
        C347097DataType = AKBTypeConst.cAKBTypeLogical
        If C347097 Then
            GoTo Comp_C347098
        Else
            GoTo Comp_C346972
        End If

    Comp_C347098:

        ' atTrue
        sCurrComponent = "atTrue"
        C347098DataType = 4
        C346952 = fn_ConvertValueCompiled(1, 4, AKB_DecimalPoint)
        C347098 = True
        GoTo Comp_C346972

    Comp_C347152:

        ' Hudtelfa?
        sCurrComponent = "Hudtelfa?"
        C347152 = (fn_ConvertValueCompiled(C347088, C347088DataType, AKB_DecimalPoint, False) = 6)
        C347152DataType = AKBTypeConst.cAKBTypeLogical
        If C347152 Then
            GoTo Comp_C346961
        Else
            GoTo Comp_C347153
        End If

    Comp_C347153:

        ' atSair
        sCurrComponent = "atSair"
        C347153DataType = 4
        C346952 = fn_ConvertValueCompiled(1, 4, AKB_DecimalPoint)
        C347153 = True
        GoTo Comp_C346972

    Comp_C350045:

        ' upPorcAdicionalCliente0
        sCurrComponent = "upPorcAdicionalCliente0"
        QueryC350045 = con.CreateCommand()
        QueryC350045.CommandText = QueryC350045.CommandText & " " & "Update AKBUSER01.WF_PRE_PEDIDO_ITENS "
        QueryC350045.CommandText = QueryC350045.CommandText & " " & "Set WF_PRE_PEDIDO_ITENS.PORC_ADICIONAL_QTDE_PEDIDO = 0"
        QueryC350045.CommandText = QueryC350045.CommandText & " " & "Where WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P57496, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        C350045 = QueryC350045.ExecuteNonQuery()
        C350045DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C347075

    Exit_R16162:

        Exit Function

    Err_R16162:
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
