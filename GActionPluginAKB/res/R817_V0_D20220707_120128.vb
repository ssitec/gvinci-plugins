Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R817

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

    Public Shared Function R817(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P1250 As Double, ByVal P1251 As Double, ByVal P1252 As Object) As DataTable
    ' 
    ' 817 - Atualiza desconto Pedido Seq - Versão: 0
    ' 
        'On Error GoTo Err_R817
        Dim sCurrComponent as String

        Dim QueryC4629 As New Object
        Dim RsC4629 As New Object
        Dim C4629DataType() As Short
        Dim RsC4629_EOF As Boolean
        Dim C4630 As Object
        Dim C4630DataType As Short
        Dim C4631 As Object
        Dim C4631DataType As Short
        Dim C4632 As Boolean
        Dim C4632DataType As Short
        Dim C4633 As Object
        Dim C4633DataType As Short
        Dim C4635 As Object
        Dim C4635DataType As Short
        Dim QueryC4637 As New Object
        Dim C4637 As Integer
        Dim C4637DataType As Short
        Dim C4639DataType() As Short
        Dim C11609 As Object
        Dim C11609DataType As Short
        Dim QueryC11612 As New Object
        Dim C11612 As Integer
        Dim C11612DataType As Short
        Dim C11618 As Object
        Dim C11618DataType As Short
        Dim QueryC11625 As New Object
        Dim RsC11625 As New Object
        Dim C11625DataType() As Short
        Dim RsC11625_EOF As Boolean
        Dim C11626 As Object
        Dim C11626DataType As Short
        Dim C11627 As Boolean
        Dim C11627DataType As Short
        Dim C11629 As Object
        Dim C11629DataType As Short
        Dim QueryC11630 As New Object
        Dim C11630 As Integer
        Dim C11630DataType As Short
        Dim C11632 As Object
        Dim C11632DataType As Short
        Dim C12541 As Object
        Dim C12541DataType As Short
        Dim C12542 As Object
        Dim C12542DataType As Short
        Dim C12543 As Boolean
        Dim C12543DataType As Short
        Dim C12544 As Object
        Dim C12544DataType As Short
        Dim C12545 As Object
        Dim C12545DataType As Short
        Dim C12547 As Object
        Dim C12547DataType As Short
        Dim C12548 As Object
        Dim C12548DataType As Short
        Dim C12549 As Boolean
        Dim C12549DataType As Short
        Dim C12550 As Object
        Dim C12550DataType As Short
        Dim C12551 As Boolean
        Dim C12551DataType As Short
        Dim C12577 As Object
        Dim C12577DataType As Short
        Dim C12579 As Boolean
        Dim C12579DataType As Short
        Dim C12801 As Boolean
        Dim C12801DataType As Short
        Dim C12802DataType() As Short
        Dim QueryC56938 As New Object
        Dim RsC56938 As New Object
        Dim C56938DataType() As Short
        Dim RsC56938_EOF As Boolean
        Dim C56939 As Object
        Dim C56939DataType As Short
        Dim C56940 As Object
        Dim C56940DataType As Short
        Dim QueryC56942 As New Object
        Dim RsC56942 As New Object
        Dim C56942DataType() As Short
        Dim RsC56942_EOF As Boolean
        Dim QueryC56944 As New Object
        Dim RsC56944 As New Object
        Dim C56944DataType() As Short
        Dim RsC56944_EOF As Boolean
        Dim QueryC56946 As New Object
        Dim RsC56946 As New Object
        Dim C56946DataType() As Short
        Dim RsC56946_EOF As Boolean
        Dim QueryC56947 As New Object
        Dim RsC56947 As New Object
        Dim C56947DataType() As Short
        Dim RsC56947_EOF As Boolean
        Dim QueryC56950 As New Object
        Dim RsC56950 As New Object
        Dim C56950DataType() As Short
        Dim RsC56950_EOF As Boolean
        Dim QueryC546334 As New Object
        Dim C546334 As Integer
        Dim C546334DataType As Short
        P1252 = IIf(IsDBNull(P1252), 1, P1252)

        ReDim ReturnDatatype(0)

        GoTo Comp_C56938

    Comp_C4629:

        ' SelDesc
        sCurrComponent = "SelDesc"
        QueryC4629 = con.CreateCommand()
        QueryC4629.CommandText = QueryC4629.CommandText & " " & "SELECT WF_PED_SEQ_DESCONTO.Porc_Desconto , WF_PED_SEQ_DESCONTO.Tipo_Desconto "
        QueryC4629.CommandText = QueryC4629.CommandText & " " & ""
        QueryC4629.CommandText = QueryC4629.CommandText & " " & ""
        QueryC4629.CommandText = QueryC4629.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_DESCONTO "
        QueryC4629.CommandText = QueryC4629.CommandText & " " & ""
        QueryC4629.CommandText = QueryC4629.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P1250, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC4629.CommandText = QueryC4629.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P1251, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC4629.CommandText = QueryC4629.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Seq = " & _ 
ConvertValue(P1252, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC4629.CommandText = QueryC4629.CommandText & " " & ""
        QueryC4629.CommandText = QueryC4629.CommandText & " " & "ORDER BY WF_PED_SEQ_DESCONTO.Porc_Desconto "
        QueryC4629.CommandText = QueryC4629.CommandText & " " & " "
        QueryC4629.Transaction = txn
        RsC4629 = QueryC4629.ExecuteReader()
        Dim iC4629 As Short
        ReDim C4629DataType(RsC4629.FieldCount)
        For iC4629 = 0 to RsC4629.FieldCount - 1
            Select Case RsC4629.GetDataTypeName(iC4629).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C4629DataType(iC4629 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C4629DataType(iC4629 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C4629DataType(iC4629 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC4629
        RsC4629_EOF = Not RsC4629.Read()

        GoTo Comp_C4630

    Comp_C4630:

        ' Eof
        sCurrComponent = "Eof"
        C4630DataType = 4
        C4630 = RsC4629_EOF
        GoTo Comp_C4632

    Comp_C4631:

        ' Valor
        sCurrComponent = "Valor"
        C4631 = 100
        C4631DataType = 1
        GoTo Comp_C12542

    Comp_C4632:

        ' Eof=F
        sCurrComponent = "Eof=F"
        C4632 = (fn_ConvertValueCompiled(C4630, C4630DataType, AKB_DecimalPoint, False) = 0)
        C4632DataType = AKBTypeConst.cAKBTypeLogical
        If C4632 Then
            GoTo Comp_C11609
        Else
            GoTo Comp_C56944
        End If

    Comp_C4633:

        ' Next
        sCurrComponent = "Next"
        C4633DataType = 4
        RsC4629_EOF = Not RsC4629.Read()
        C4633 = True

        GoTo Comp_C4630

    Comp_C4635:

        ' GuardaValor
        sCurrComponent = "GuardaValor"
        C4635DataType = 4
        C4631 = fn_ConvertValueCompiled(RsC56942(0) , 1, AKB_DecimalPoint)
        C4635 = True
        GoTo Comp_C4633

    Comp_C4637:

        ' UpDescPed
        sCurrComponent = "UpDescPed"
        QueryC4637 = con.CreateCommand()
        QueryC4637.CommandText = QueryC4637.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_SEQ "
        QueryC4637.CommandText = QueryC4637.CommandText & " " & ""
        QueryC4637.CommandText = QueryC4637.CommandText & " " & "SET WF_PEDIDO_SEQ.Porc_Desc_Ped = " & _ 
ConvertValue(RsC56944(0), C56944DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC4637.CommandText = QueryC4637.CommandText & " " & ""
        QueryC4637.CommandText = QueryC4637.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto = " & _ 
ConvertValue(P1250, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC4637.CommandText = QueryC4637.CommandText & " " & "AND WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P1251, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC4637.CommandText = QueryC4637.CommandText & " " & "AND WF_PEDIDO_SEQ.Seq = " & _ 
ConvertValue(P1252, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC4637.CommandText = QueryC4637.CommandText & " " & ""
        QueryC4637.Transaction = txn
        C4637 = QueryC4637.ExecuteNonQuery()
        C4637DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C546334

    Comp_C4639:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C4639 As DataTable = New DataTable()
        tb_C4639.Columns.Add("EOF2 DESCITEM=T" & "")
        Dim row_C4639 As DataRow = tb_C4639.NewRow()
        row_C4639(0) = C12579
        tb_C4639.Rows.Add(row_C4639)
        R817 = tb_C4639
        ReDim C4639DataType(1)
        C4639DataType(1) = C12579DataType
        ReturnDataType = C4639DataType

        GoTo Exit_R817

    Comp_C11609:

        ' Desc
        sCurrComponent = "Desc"
        C11609DataType = 0
        C11609 = RsC4629(0)
        C11609DataType = C4629Datatype(1)
        If C11609DataType = AKBTypeConst.cAKBTypeString Then
          C11609 = IIF(IsDBNull(C11609), C11609, Strings.RTrim(C11609))
        End If 

        GoTo Comp_C12541

    Comp_C11612:

        ' UpDescItem
        sCurrComponent = "UpDescItem"
        QueryC11612 = con.CreateCommand()
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC11612.CommandText = QueryC11612.CommandText & " " & ""
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "SET WF_PEDIDO_ITENS.Porc_Desconto = " & _ 
ConvertValue(RsC56944(0), C56944DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11612.CommandText = QueryC11612.CommandText & " " & ""
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P1250, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P1251, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "AND WF_PEDIDO_ITENS.Item_Bonificado = 0"
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "AND NOT EXISTS (SELECT * FROM wf_tabpreco_vda_produto, WF_PEDIDO"
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "WHERE WF_PEDIDO.PEDIDO = WF_PEDIDO_ITENS.PEDIDO"
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "AND WF_PEDIDO.TABELA_PRECO_VENDA = wf_tabpreco_vda_produto.TABELA_PRECO_VENDA"
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "AND WF_PEDIDO_ITENS.ACESSO = wf_tabpreco_vda_produto.ACESSO"
        QueryC11612.CommandText = QueryC11612.CommandText & " " & "AND nvl(wf_tabpreco_vda_produto.ACESSO_SERVICO,0) <> 0)"
        QueryC11612.Transaction = txn
        C11612 = QueryC11612.ExecuteNonQuery()
        C11612DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C11625

    Comp_C11618:

        ' ZeraValor
        sCurrComponent = "ZeraValor"
        C11618DataType = 4
        C4631 = fn_ConvertValueCompiled(100, 1, AKB_DecimalPoint)
        C11618 = True
        GoTo Comp_C56950

    Comp_C11625:

        ' SelDescItem
        sCurrComponent = "SelDescItem"
        QueryC11625 = con.CreateCommand()
        QueryC11625.CommandText = QueryC11625.CommandText & " " & "SELECT WF_PED_ITENS_DESC.Porc_Desconto , WF_PED_ITENS_DESC.Seq_Item "
        QueryC11625.CommandText = QueryC11625.CommandText & " " & "FROM AKBUSER01.WF_PED_ITENS_DESC "
        QueryC11625.CommandText = QueryC11625.CommandText & " " & "WHERE WF_PED_ITENS_DESC.Tp_Produto = " & _ 
ConvertValue(P1250, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11625.CommandText = QueryC11625.CommandText & " " & "AND  WF_PED_ITENS_DESC.Pedido = " & _ 
ConvertValue(P1251, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11625.CommandText = QueryC11625.CommandText & " " & "AND  WF_PED_ITENS_DESC.Tipo_Desconto <> 280 "
        QueryC11625.CommandText = QueryC11625.CommandText & " " & "ORDER BY WF_PED_ITENS_DESC.Seq_Item "
        QueryC11625.Transaction = txn
        RsC11625 = QueryC11625.ExecuteReader()
        Dim iC11625 As Short
        ReDim C11625DataType(RsC11625.FieldCount)
        For iC11625 = 0 to RsC11625.FieldCount - 1
            Select Case RsC11625.GetDataTypeName(iC11625).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C11625DataType(iC11625 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C11625DataType(iC11625 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C11625DataType(iC11625 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC11625
        RsC11625_EOF = Not RsC11625.Read()

        GoTo Comp_C11626

    Comp_C11626:

        ' Eof DescItem
        sCurrComponent = "Eof DescItem"
        C11626DataType = 4
        C11626 = RsC11625_EOF
        GoTo Comp_C11627

    Comp_C11627:

        ' Eof DescItem=T
        sCurrComponent = "Eof DescItem=T"
        C11627 = (fn_ConvertValueCompiled(C11626, C11626DataType, AKB_DecimalPoint, False) =1)
        C11627DataType = AKBTypeConst.cAKBTypeLogical
        If C11627 Then
            GoTo Comp_C12551
        Else
            GoTo Comp_C12545
        End If

    Comp_C11629:

        ' Next DescItem
        sCurrComponent = "Next DescItem"
        C11629DataType = 4
        RsC11625_EOF = Not RsC11625.Read()
        C11629 = True

        GoTo Comp_C11626

    Comp_C11630:

        ' UpDescItem2
        sCurrComponent = "UpDescItem2"
        QueryC11630 = con.CreateCommand()
        QueryC11630.CommandText = QueryC11630.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC11630.CommandText = QueryC11630.CommandText & " " & ""
        QueryC11630.CommandText = QueryC11630.CommandText & " " & "SET WF_PEDIDO_ITENS.Porc_Desconto = " & _ 
ConvertValue(RsC56947(0), C56947DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11630.CommandText = QueryC11630.CommandText & " " & ""
        QueryC11630.CommandText = QueryC11630.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P1250, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11630.CommandText = QueryC11630.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P1251, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11630.CommandText = QueryC11630.CommandText & " " & "AND  WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C12548, C12548DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC11630.CommandText = QueryC11630.CommandText & " " & "AND  WF_PEDIDO_ITENS.Item_Bonificado = 0"
        QueryC11630.Transaction = txn
        C11630 = QueryC11630.ExecuteNonQuery()
        C11630DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C12579

    Comp_C11632:

        ' GuardaValor2
        sCurrComponent = "GuardaValor2"
        C11632DataType = 4
        C4631 = fn_ConvertValueCompiled(RsC56950(0) , 1, AKB_DecimalPoint)
        C11632 = True
        GoTo Comp_C11629

    Comp_C12541:

        ' TpDesc
        sCurrComponent = "TpDesc"
        C12541DataType = 0
        C12541DataType = C4629Datatype(2)
        If C12541DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC4629(1)) Then
          C12541 = Strings.RTrim(RsC4629(1))
        Else 
          C12541 = RsC4629(1)
        End If 

        GoTo Comp_C12543

    Comp_C12542:

        ' VDesc2
        sCurrComponent = "VDesc2"
        C12542 = 0
        C12542DataType = 1
        GoTo Comp_C4629

    Comp_C12543:

        ' Desc=280
        sCurrComponent = "Desc=280"
        C12543 = (fn_ConvertValueCompiled(C12541, C12541DataType, AKB_DecimalPoint, False) =280 OR fn_ConvertValueCompiled(C12541, C12541DataType, AKB_DecimalPoint, False) =281)
        C12543DataType = AKBTypeConst.cAKBTypeLogical
        If C12543 Then
            GoTo Comp_C12544
        Else
            GoTo Comp_C56942
        End If

    Comp_C12544:

        ' AtDesc2
        sCurrComponent = "AtDesc2"
        C12544DataType = 4
        C12542 = fn_ConvertValueCompiled(C11609 , 1, AKB_DecimalPoint)
        C12544 = True
        GoTo Comp_C56942

    Comp_C12545:

        ' DescItem
        sCurrComponent = "DescItem"
        C12545DataType = 0
        C12545 = RsC11625(0)
        C12545DataType = C11625Datatype(1)
        If C12545DataType = AKBTypeConst.cAKBTypeString Then
          C12545 = IIF(IsDBNull(C12545), C12545, Strings.RTrim(C12545))
        End If 

        GoTo Comp_C12547

    Comp_C12547:

        ' SeqItem
        sCurrComponent = "SeqItem"
        C12547DataType = 0
        C12547DataType = C11625Datatype(2)
        If C12547DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC11625(1)) Then
          C12547 = Strings.RTrim(RsC11625(1))
        Else 
          C12547 = RsC11625(1)
        End If 

        GoTo Comp_C12549

    Comp_C12548:

        ' VSeqItem
        sCurrComponent = "VSeqItem"
        C12548 = 0
        C12548DataType = 1
        GoTo Comp_C4637

    Comp_C12549:

        ' Item<>VItem
        sCurrComponent = "Item<>VItem"
        C12549 = (fn_ConvertValueCompiled(C12547, C12547DataType, AKB_DecimalPoint, False) <>fn_ConvertValueCompiled(C12548, C12548DataType, AKB_DecimalPoint, False))
        C12549DataType = AKBTypeConst.cAKBTypeLogical
        If C12549 Then
            GoTo Comp_C12551
        Else
            GoTo Comp_C56950
        End If

    Comp_C12550:

        ' NovoItem
        sCurrComponent = "NovoItem"
        C12550DataType = 4
        C12548 = fn_ConvertValueCompiled(C12547 , 1, AKB_DecimalPoint)
        C12550 = True
        GoTo Comp_C11618

    Comp_C12551:

        ' VItem>0
        sCurrComponent = "VItem>0"
        C12551 = (fn_ConvertValueCompiled(C12548, C12548DataType, AKB_DecimalPoint, False)  >0)
        C12551DataType = AKBTypeConst.cAKBTypeLogical
        If C12551 Then
            GoTo Comp_C56946
        Else
            GoTo Comp_C12801
        End If

    Comp_C12577:

        ' GValor2
        sCurrComponent = "GValor2"
        C12577DataType = 4
        C4631 = fn_ConvertValueCompiled(RsC56946(0) , 1, AKB_DecimalPoint)
        C12577 = True
        GoTo Comp_C56947

    Comp_C12579:

        ' EOF2=T
        sCurrComponent = "EOF2=T"
        C12579 = (fn_ConvertValueCompiled(C11626, C11626DataType, AKB_DecimalPoint, False) =1)
        C12579DataType = AKBTypeConst.cAKBTypeLogical
        If C12579 Then
            GoTo Comp_C4639
        Else
            GoTo Comp_C12550
        End If

    Comp_C12801:

        ' EOF2
        sCurrComponent = "EOF2"
        C12801 = (fn_ConvertValueCompiled(C11626, C11626DataType, AKB_DecimalPoint, False) =1)
        C12801DataType = AKBTypeConst.cAKBTypeLogical
        If C12801 Then
            GoTo Comp_C12802
        Else
            GoTo Comp_C12550
        End If

    Comp_C12802:

        ' RET2
        sCurrComponent = "RET2"
        Dim tb_C12802 As DataTable = New DataTable()
        tb_C12802.Columns.Add("EOF DESCITEM=T1" & "")
        Dim row_C12802 As DataRow = tb_C12802.NewRow()
        row_C12802(0) = C12801
        tb_C12802.Rows.Add(row_C12802)
        R817 = tb_C12802
        ReDim C12802DataType(1)
        C12802DataType(1) = C12801DataType
        ReturnDataType = C12802DataType

        GoTo Exit_R817

    Comp_C56938:

        ' SelParâmetros
        sCurrComponent = "SelParâmetros"
        QueryC56938 = con.CreateCommand()
        QueryC56938.CommandText = QueryC56938.CommandText & " " & "SELECT WF_TP_PRODUTO.Desconto_Decimal , WF_TP_PRODUTO.Trunca_Desconto FROM AKBUSER01.WF_TP_PRODUTO WHERE WF_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P1250, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56938.Transaction = txn
        RsC56938 = QueryC56938.ExecuteReader()
        Dim iC56938 As Short
        ReDim C56938DataType(RsC56938.FieldCount)
        For iC56938 = 0 to RsC56938.FieldCount - 1
            Select Case RsC56938.GetDataTypeName(iC56938).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56938DataType(iC56938 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56938DataType(iC56938 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56938DataType(iC56938 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56938
        RsC56938_EOF = Not RsC56938.Read()

        GoTo Comp_C56939

    Comp_C56939:

        ' TruncaDesconto
        sCurrComponent = "TruncaDesconto"
        C56939DataType = 0
        C56939DataType = C56938Datatype(2)
        If C56939DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56938(1)) Then
          C56939 = Strings.RTrim(RsC56938(1))
        Else 
          C56939 = RsC56938(1)
        End If 

        GoTo Comp_C56940

    Comp_C56940:

        ' DecimaisDesconto
        sCurrComponent = "DecimaisDesconto"
        C56940DataType = 0
        C56940 = RsC56938(0)
        C56940DataType = C56938Datatype(1)
        If C56940DataType = AKBTypeConst.cAKBTypeString Then
          C56940 = IIF(IsDBNull(C56940), C56940, Strings.RTrim(C56940))
        End If 

        GoTo Comp_C4631

    Comp_C56942:

        ' CalcDesc
        sCurrComponent = "CalcDesc"
        QueryC56942 = con.CreateCommand()
        QueryC56942.CommandText = QueryC56942.CommandText & " " & "SELECT  "
        QueryC56942.CommandText = QueryC56942.CommandText & " " & ""
        QueryC56942.CommandText = QueryC56942.CommandText & " " & "DECODE( " & _ 
ConvertValue(C56939, C56939DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC56942.CommandText = QueryC56942.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C11609, C11609DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC56942.CommandText = QueryC56942.CommandText & " " & "ROUND( " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C11609, C11609DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC56942.CommandText = QueryC56942.CommandText & " " & ""
        QueryC56942.CommandText = QueryC56942.CommandText & " " & "FROM DUAL"
        QueryC56942.CommandText = QueryC56942.CommandText & " " & ""
        QueryC56942.Transaction = txn
        RsC56942 = QueryC56942.ExecuteReader()
        Dim iC56942 As Short
        ReDim C56942DataType(RsC56942.FieldCount)
        For iC56942 = 0 to RsC56942.FieldCount - 1
            Select Case RsC56942.GetDataTypeName(iC56942).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56942DataType(iC56942 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56942DataType(iC56942 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56942DataType(iC56942 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56942
        RsC56942_EOF = Not RsC56942.Read()

        GoTo Comp_C4635

    Comp_C56944:

        ' DescTotal
        sCurrComponent = "DescTotal"
        QueryC56944 = con.CreateCommand()
        QueryC56944.CommandText = QueryC56944.CommandText & " " & "SELECT  "
        QueryC56944.CommandText = QueryC56944.CommandText & " " & ""
        QueryC56944.CommandText = QueryC56944.CommandText & " " & "DECODE( " & _ 
ConvertValue(C56939, C56939DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC56944.CommandText = QueryC56944.CommandText & " " & "TRUNC( 100-" & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC56944.CommandText = QueryC56944.CommandText & " " & "ROUND( 100-" & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC56944.CommandText = QueryC56944.CommandText & " " & ""
        QueryC56944.CommandText = QueryC56944.CommandText & " " & "FROM DUAL"
        QueryC56944.CommandText = QueryC56944.CommandText & " " & ""
        QueryC56944.Transaction = txn
        RsC56944 = QueryC56944.ExecuteReader()
        Dim iC56944 As Short
        ReDim C56944DataType(RsC56944.FieldCount)
        For iC56944 = 0 to RsC56944.FieldCount - 1
            Select Case RsC56944.GetDataTypeName(iC56944).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56944DataType(iC56944 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56944DataType(iC56944 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56944DataType(iC56944 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56944
        RsC56944_EOF = Not RsC56944.Read()

        GoTo Comp_C12548

    Comp_C56946:

        ' CALC DESCITEM2
        sCurrComponent = "CALC DESCITEM2"
        QueryC56946 = con.CreateCommand()
        QueryC56946.CommandText = QueryC56946.CommandText & " " & "SELECT  "
        QueryC56946.CommandText = QueryC56946.CommandText & " " & ""
        QueryC56946.CommandText = QueryC56946.CommandText & " " & "DECODE( " & _ 
ConvertValue(C56939, C56939DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC56946.CommandText = QueryC56946.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C12542, C12542DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC56946.CommandText = QueryC56946.CommandText & " " & "ROUND( " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C12542, C12542DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC56946.CommandText = QueryC56946.CommandText & " " & ""
        QueryC56946.CommandText = QueryC56946.CommandText & " " & "FROM DUAL"
        QueryC56946.CommandText = QueryC56946.CommandText & " " & ""
        QueryC56946.Transaction = txn
        RsC56946 = QueryC56946.ExecuteReader()
        Dim iC56946 As Short
        ReDim C56946DataType(RsC56946.FieldCount)
        For iC56946 = 0 to RsC56946.FieldCount - 1
            Select Case RsC56946.GetDataTypeName(iC56946).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56946DataType(iC56946 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56946DataType(iC56946 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56946DataType(iC56946 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56946
        RsC56946_EOF = Not RsC56946.Read()

        GoTo Comp_C12577

    Comp_C56947:

        ' Desc Total Item
        sCurrComponent = "Desc Total Item"
        QueryC56947 = con.CreateCommand()
        QueryC56947.CommandText = QueryC56947.CommandText & " " & "SELECT  "
        QueryC56947.CommandText = QueryC56947.CommandText & " " & ""
        QueryC56947.CommandText = QueryC56947.CommandText & " " & "DECODE( " & _ 
ConvertValue(C56939, C56939DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC56947.CommandText = QueryC56947.CommandText & " " & "TRUNC( 100 - " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC56947.CommandText = QueryC56947.CommandText & " " & "ROUND( 100 - " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC56947.CommandText = QueryC56947.CommandText & " " & ""
        QueryC56947.CommandText = QueryC56947.CommandText & " " & "FROM DUAL"
        QueryC56947.CommandText = QueryC56947.CommandText & " " & ""
        QueryC56947.Transaction = txn
        RsC56947 = QueryC56947.ExecuteReader()
        Dim iC56947 As Short
        ReDim C56947DataType(RsC56947.FieldCount)
        For iC56947 = 0 to RsC56947.FieldCount - 1
            Select Case RsC56947.GetDataTypeName(iC56947).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56947DataType(iC56947 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56947DataType(iC56947 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56947DataType(iC56947 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56947
        RsC56947_EOF = Not RsC56947.Read()

        GoTo Comp_C11630

    Comp_C56950:

        ' CalcDescItem
        sCurrComponent = "CalcDescItem"
        QueryC56950 = con.CreateCommand()
        QueryC56950.CommandText = QueryC56950.CommandText & " " & "SELECT  "
        QueryC56950.CommandText = QueryC56950.CommandText & " " & ""
        QueryC56950.CommandText = QueryC56950.CommandText & " " & "DECODE( " & _ 
ConvertValue(C56939, C56939DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC56950.CommandText = QueryC56950.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C12545, C12545DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC56950.CommandText = QueryC56950.CommandText & " " & "ROUND( " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C12545, C12545DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C4631, C4631DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C56940, C56940DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC56950.CommandText = QueryC56950.CommandText & " " & ""
        QueryC56950.CommandText = QueryC56950.CommandText & " " & "FROM DUAL"
        QueryC56950.CommandText = QueryC56950.CommandText & " " & ""
        QueryC56950.Transaction = txn
        RsC56950 = QueryC56950.ExecuteReader()
        Dim iC56950 As Short
        ReDim C56950DataType(RsC56950.FieldCount)
        For iC56950 = 0 to RsC56950.FieldCount - 1
            Select Case RsC56950.GetDataTypeName(iC56950).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56950DataType(iC56950 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56950DataType(iC56950 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56950DataType(iC56950 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56950
        RsC56950_EOF = Not RsC56950.Read()

        GoTo Comp_C11632

    Comp_C546334:

        ' ZeraDescItens
        sCurrComponent = "ZeraDescItens"
        QueryC546334 = con.CreateCommand()
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC546334.CommandText = QueryC546334.CommandText & " " & ""
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "SET WF_PEDIDO_ITENS.Porc_Desconto = 0"
        QueryC546334.CommandText = QueryC546334.CommandText & " " & ""
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P1250, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P1251, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "AND NOT EXISTS (SELECT * FROM wf_tabpreco_vda_produto, WF_PEDIDO"
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "WHERE WF_PEDIDO.PEDIDO = WF_PEDIDO_ITENS.PEDIDO"
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "AND WF_PEDIDO.TABELA_PRECO_VENDA = wf_tabpreco_vda_produto.TABELA_PRECO_VENDA"
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "AND WF_PEDIDO_ITENS.ACESSO = wf_tabpreco_vda_produto.ACESSO"
        QueryC546334.CommandText = QueryC546334.CommandText & " " & "AND nvl(wf_tabpreco_vda_produto.ACESSO_SERVICO,0) <> 0)"
        QueryC546334.Transaction = txn
        C546334 = QueryC546334.ExecuteNonQuery()
        C546334DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C11612

    Exit_R817:

        Exit Function

    Err_R817:
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
