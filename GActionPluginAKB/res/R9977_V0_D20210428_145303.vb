Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R9977

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

    Public Shared Function R9977(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P31857 As Double, ByVal P31858 As Double, ByVal P60141 As Object, ByVal P65133 As Object) As DataTable
    ' 
    ' 9977 - Verifica Coleção dos Itens do Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R9977
        Dim sCurrComponent as String

        Dim QueryC174732 As New Object
        Dim RsC174732 As New Object
        Dim C174732DataType() As Short
        Dim RsC174732_EOF As Boolean
        Dim C174736DataType() As Short
        Dim C174737 As Object
        Dim C174737DataType As Short
        Dim C174740 As Object
        Dim C174740DataType As Short
        Dim C174744 As DataTable
        Dim C174744CurrentRow As Integer
        Dim C174744DataType() As Short
        Dim C174745 As Boolean
        Dim C174745DataType As Short
        Dim C174859 As Object
        Dim C174859DataType As Short
        Dim C278978 As Boolean
        Dim C278978DataType As Short
        Dim C296812 As Object
        Dim C296812DataType As Short
        Dim QueryC332488 As New Object
        Dim RsC332488 As New Object
        Dim C332488DataType() As Short
        Dim RsC332488_EOF As Boolean
        Dim C332489 As Boolean
        Dim C332489DataType As Short
        Dim C332492 As Object
        Dim C332492DataType As Short
        Dim C340807 As Object
        Dim C340807DataType As Short
        Dim C340891 As DataTable
        Dim C340891CurrentRow As Integer
        Dim C340891DataType() As Short
        Dim C340930 As Object
        Dim C340930DataType As Short
        Dim C340933 As Object
        Dim C340933DataType As Short
        Dim C340934 As Boolean
        Dim C340934DataType As Short
        Dim C341055 As Object
        Dim C341055DataType As Short
        Dim C341112 As Object
        Dim C341112DataType As Short
        Dim QueryC355118 As New Object
        Dim RsC355118 As New Object
        Dim C355118DataType() As Short
        Dim RsC355118_EOF As Boolean
        Dim C355119 As Boolean
        Dim C355119DataType As Short
        Dim C367293 As Boolean
        Dim C367293DataType As Short
        Dim QueryC367294 As New Object
        Dim RsC367294 As New Object
        Dim C367294DataType() As Short
        Dim RsC367294_EOF As Boolean
        Dim C367300 As Boolean
        Dim C367300DataType As Short
        Dim C367301 As Object
        Dim C367301DataType As Short
        Dim C368875 As Object
        Dim C368875DataType As Short
        Dim QueryC372222 As New Object
        Dim C372222 As Integer
        Dim C372222DataType As Short
        Dim C377024 As Object
        Dim C377024DataType As Short
        Dim C393054 As Object
        Dim C393054DataType As Short
        Dim C406123 As Boolean
        Dim C406123DataType As Short
        Dim QueryC419525 As New Object
        Dim RsC419525 As New Object
        Dim C419525DataType() As Short
        Dim RsC419525_EOF As Boolean
        Dim C419526 As Boolean
        Dim C419526DataType As Short
        Dim QueryC419528 As New Object
        Dim C419528 As Integer
        Dim C419528DataType As Short
        P65133 = IIf(IsDBNull(P65133), 1, P65133)

        ReDim ReturnDatatype(0)

        GoTo Comp_C174737

    Comp_C174732:

        ' SelItensPedido
        sCurrComponent = "SelItensPedido"
        QueryC174732 = con.CreateCommand()
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "SELECT   WF_PEDIDO_ITENS.Sigla_Prod , "
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "WF_PEDIDO_ITENS.Acesso , "
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "WF_PEDIDO_ITENS.Seq_Item , "
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "TRUNC(WF_PEDIDO.Dt_Pedido) , "
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "NVL(WF_PEDIDO_ITENS.Id_Colecao,0) , "
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "WF_PEDIDO.Tabela_Preco_Venda ,"
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "WF_PEDIDO_ITENS.Cod_Embalagem "
        QueryC174732.CommandText = QueryC174732.CommandText & " " & ""
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS , AKBUSER01.WF_PEDIDO"
        QueryC174732.CommandText = QueryC174732.CommandText & " " & ""
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = WF_PEDIDO.Tp_Produto"
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = WF_PEDIDO.Pedido"
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P31858, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P31857, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174732.CommandText = QueryC174732.CommandText & " " & "AND WF_PEDIDO_ITENS.Flag_Pos_Ped <> 8"
        QueryC174732.Transaction = txn
        RsC174732 = QueryC174732.ExecuteReader()
        Dim iC174732 As Short
        ReDim C174732DataType(RsC174732.FieldCount)
        For iC174732 = 0 to RsC174732.FieldCount - 1
            Select Case RsC174732.GetDataTypeName(iC174732).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C174732DataType(iC174732 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C174732DataType(iC174732 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C174732DataType(iC174732 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC174732
        RsC174732_EOF = Not RsC174732.Read()

        GoTo Comp_C296812

    Comp_C174736:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C174736 As DataTable = New DataTable()
        tb_C174736.Columns.Add("vRet" & "")
        Dim row_C174736 As DataRow = tb_C174736.NewRow()
        row_C174736(0) = C174737
        tb_C174736.Rows.Add(row_C174736)
        R9977 = tb_C174736
        ReDim C174736DataType(1)
        C174736DataType(1) = C174737DataType
        ReturnDataType = C174736DataType

        GoTo Exit_R9977

    Comp_C174737:

        ' vRet
        sCurrComponent = "vRet"
        C174737 = 1
        C174737DataType = 4
        GoTo Comp_C341112

    Comp_C174740:

        ' NextSelItens
        sCurrComponent = "NextSelItens"
        C174740DataType = 4
        RsC174732_EOF = Not RsC174732.Read()
        C174740 = True

        GoTo Comp_C296812

    Comp_C174744:

        ' Valida Coleção
        sCurrComponent = "Valida Coleção"
        C174744 = clsRuleDynamicallyCompiled_R9971.R9971(con, txn, IIf(Not IsDBNull(C340807), C340807, System.DBNull.Value))
        C174744CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C174744) Then
          iColumns = C174744.Columns.Count
        End If
        ReDim C174744DataType(iColumns)
        For iCol = 1 To iColumns
            C174744DataType(iCol) = clsRuleDynamicallyCompiled_R9971.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C174745

    Comp_C174745:

        ' Flag Deve Pertencer a uma Coleção = 1?
        sCurrComponent = "Flag Deve Pertencer a uma Coleção = 1?"
        C174745 = (fn_ConvertValueCompiled(C174744.rows(C174744CurrentRow - 1)(0), C174744DataType(1), AKB_DecimalPoint, False) = 0)
        C174745DataType = AKBTypeConst.cAKBTypeLogical
        If C174745 Then
            GoTo Comp_C367301
        Else
            GoTo Comp_C332488
        End If

    Comp_C174859:

        ' ValAcesso
        sCurrComponent = "ValAcesso"
        C174859DataType = 0
        C174859DataType = C174732Datatype(2)
        If C174859DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC174732(1)) Then
          C174859 = Strings.RTrim(RsC174732(1))
        Else 
          C174859 = RsC174732(1)
        End If 

        GoTo Comp_C340930

    Comp_C278978:

        ' TpProd=Manual?
        sCurrComponent = "TpProd=Manual?"
        C278978 = (( fn_ConvertValueCompiled(P31857, 1, AKB_DecimalPoint, False) = 200 ) or ( fn_ConvertValueCompiled(P31857, 1, AKB_DecimalPoint, False) = 201 ))
        C278978DataType = AKBTypeConst.cAKBTypeLogical
        If C278978 Then
            GoTo Comp_C174736
        Else
            GoTo Comp_C355118
        End If

    Comp_C296812:

        ' EOFSelItens
        sCurrComponent = "EOFSelItens"
        C296812DataType = 4
        C296812 = RsC174732_EOF
        GoTo Comp_C340934

    Comp_C332488:

        ' SelProntaEntrega
        sCurrComponent = "SelProntaEntrega"
        QueryC332488 = con.CreateCommand()
        QueryC332488.CommandText = QueryC332488.CommandText & " " & "SELECT WF_TIPO_PED.Pronta_Entrega"
        QueryC332488.CommandText = QueryC332488.CommandText & " " & ""
        QueryC332488.CommandText = QueryC332488.CommandText & " " & "FROM AKBUSER01.WF_TIPO_PED, AKBUSER01.WF_PEDIDO_ITENS"
        QueryC332488.CommandText = QueryC332488.CommandText & " " & ""
        QueryC332488.CommandText = QueryC332488.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tipo_Ped = WF_TIPO_PED.Tipo_Ped"
        QueryC332488.CommandText = QueryC332488.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P31858, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC332488.CommandText = QueryC332488.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P31857, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC332488.CommandText = QueryC332488.CommandText & " " & "AND WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C340930, C340930DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC332488.Transaction = txn
        RsC332488 = QueryC332488.ExecuteReader()
        Dim iC332488 As Short
        ReDim C332488DataType(RsC332488.FieldCount)
        For iC332488 = 0 to RsC332488.FieldCount - 1
            Select Case RsC332488.GetDataTypeName(iC332488).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C332488DataType(iC332488 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C332488DataType(iC332488 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C332488DataType(iC332488 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC332488
        RsC332488_EOF = Not RsC332488.Read()

        GoTo Comp_C367294

    Comp_C332489:

        ' É Pronta Entrega?
        sCurrComponent = "É Pronta Entrega?"
        C332489 = (fn_ConvertValueCompiled(RsC332488(0), C332488DataType(1), AKB_DecimalPoint, False) = 1)
        C332489DataType = AKBTypeConst.cAKBTypeLogical
        If C332489 Then
            GoTo Comp_C367293
        Else
            GoTo Comp_C340891
        End If

    Comp_C332492:

        ' AtbFalse_vRet
        sCurrComponent = "AtbFalse_vRet"
        C332492DataType = 4
        C174737 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C332492 = True
        GoTo Comp_C174736

    Comp_C340807:

        ' ValSigla
        sCurrComponent = "ValSigla"
        C340807DataType = 0
        C340807 = RsC174732(0)
        C340807DataType = C174732Datatype(1)
        If C340807DataType = AKBTypeConst.cAKBTypeString Then
          C340807 = IIF(IsDBNull(C340807), C340807, Strings.RTrim(C340807))
        End If 

        GoTo Comp_C174859

    Comp_C340891:

        ' Atualiza Coleções Pedidos
        sCurrComponent = "Atualiza Coleções Pedidos"
        C340891 = clsRuleDynamicallyCompiled_R15919.R15919(con, txn, IIf(Not IsDBNull(C340807), C340807, System.DBNull.Value), IIf(Not IsDBNull(C174859), C174859, System.DBNull.Value), IIf(Not IsDBNull(P31858), P31858, System.DBNull.Value), IIf(Not IsDBNull(C340933), C340933, System.DBNull.Value), IIf(Not IsDBNull(P31857), P31857, System.DBNull.Value), IIf(Not IsDBNull(C340930), C340930, System.DBNull.Value), IIf(Not IsDBNull(C341055), C341055, System.DBNull.Value), IIf(Not IsDBNull(C341112), C341112, System.DBNull.Value), IIf(Not IsDBNull(C368875), C368875, System.DBNull.Value), IIf(Not IsDBNull(C377024), C377024, System.DBNull.Value), IIf(Not IsDBNull(C393054), C393054, System.DBNull.Value), IIf(Not IsDBNull(P60141), P60141, System.DBNull.Value), IIf(Not IsDBNull(P65133), P65133, System.DBNull.Value))
        C340891CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C340891) Then
          iColumns = C340891.Columns.Count
        End If
        ReDim C340891DataType(iColumns)
        For iCol = 1 To iColumns
            C340891DataType(iCol) = clsRuleDynamicallyCompiled_R15919.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C406123

    Comp_C340930:

        ' ValSeqIten
        sCurrComponent = "ValSeqIten"
        C340930DataType = 0
        C340930DataType = C174732Datatype(3)
        If C340930DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC174732(2)) Then
          C340930 = Strings.RTrim(RsC174732(2))
        Else 
          C340930 = RsC174732(2)
        End If 

        GoTo Comp_C340933

    Comp_C340933:

        ' ValDataGeração
        sCurrComponent = "ValDataGeração"
        C340933DataType = 0
        C340933DataType = C174732Datatype(4)
        If C340933DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC174732(3)) Then
          C340933 = Strings.RTrim(RsC174732(3))
        Else 
          C340933 = RsC174732(3)
        End If 

        GoTo Comp_C341055

    Comp_C340934:

        ' FimItensPedido?
        sCurrComponent = "FimItensPedido?"
        C340934 = (fn_ConvertValueCompiled(C296812, C296812DataType, AKB_DecimalPoint, False) = 1)
        C340934DataType = AKBTypeConst.cAKBTypeLogical
        If C340934 Then
            GoTo Comp_C174736
        Else
            GoTo Comp_C340807
        End If

    Comp_C341055:

        ' ValIdColeção
        sCurrComponent = "ValIdColeção"
        C341055DataType = 0
        C341055DataType = C174732Datatype(5)
        If C341055DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC174732(4)) Then
          C341055 = Strings.RTrim(RsC174732(4))
        Else 
          C341055 = RsC174732(4)
        End If 

        GoTo Comp_C368875

    Comp_C341112:

        ' vExibirMsg
        sCurrComponent = "vExibirMsg"
        C341112 = 0
        C341112DataType = 1
        GoTo Comp_C377024

    Comp_C355118:

        ' SelTpFaturamento
        sCurrComponent = "SelTpFaturamento"
        QueryC355118 = con.CreateCommand()
        QueryC355118.CommandText = QueryC355118.CommandText & " " & "SELECT DISTINCT WF_TP_PRODUTO.Tp_Tab_Preco"
        QueryC355118.CommandText = QueryC355118.CommandText & " " & ""
        QueryC355118.CommandText = QueryC355118.CommandText & " " & "FROM AKBUSER01.WF_TP_PRODUTO"
        QueryC355118.CommandText = QueryC355118.CommandText & " " & ""
        QueryC355118.CommandText = QueryC355118.CommandText & " " & "WHERE WF_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P31857, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC355118.Transaction = txn
        RsC355118 = QueryC355118.ExecuteReader()
        Dim iC355118 As Short
        ReDim C355118DataType(RsC355118.FieldCount)
        For iC355118 = 0 to RsC355118.FieldCount - 1
            Select Case RsC355118.GetDataTypeName(iC355118).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C355118DataType(iC355118 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C355118DataType(iC355118 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C355118DataType(iC355118 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC355118
        RsC355118_EOF = Not RsC355118.Read()

        GoTo Comp_C355119

    Comp_C355119:

        ' Tp Faturamento = Interno?
        sCurrComponent = "Tp Faturamento = Interno?"
        C355119 = (fn_ConvertValueCompiled(RsC355118(0), C355118DataType(1), AKB_DecimalPoint, False) = "Interno")
        C355119DataType = AKBTypeConst.cAKBTypeLogical
        If C355119 Then
            GoTo Comp_C174736
        Else
            GoTo Comp_C174732
        End If

    Comp_C367293:

        ' Flag Produto Pronta Entrega com Coleção?
        sCurrComponent = "Flag Produto Pronta Entrega com Coleção?"
        C367293 = (fn_ConvertValueCompiled(RsC367294(0), C367294DataType(1), AKB_DecimalPoint, False) = 1)
        C367293DataType = AKBTypeConst.cAKBTypeLogical
        If C367293 Then
            GoTo Comp_C340891
        Else
            GoTo Comp_C174740
        End If

    Comp_C367294:

        ' SelFlagProdutoProntaEntrComColeção
        sCurrComponent = "SelFlagProdutoProntaEntrComColeção"
        QueryC367294 = con.CreateCommand()
        QueryC367294.CommandText = QueryC367294.CommandText & " " & "SELECT WF_ESTAB_JURIDICO_PARAM.prod_pronta_entr_com_colecao"
        QueryC367294.CommandText = QueryC367294.CommandText & " " & ""
        QueryC367294.CommandText = QueryC367294.CommandText & " " & "FROM AKBUSER01.WF_ESTAB_JURIDICO_PARAM"
        QueryC367294.CommandText = QueryC367294.CommandText & " " & ""
        QueryC367294.CommandText = QueryC367294.CommandText & " " & "WHERE WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P60141, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC367294.Transaction = txn
        RsC367294 = QueryC367294.ExecuteReader()
        Dim iC367294 As Short
        ReDim C367294DataType(RsC367294.FieldCount)
        For iC367294 = 0 to RsC367294.FieldCount - 1
            Select Case RsC367294.GetDataTypeName(iC367294).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C367294DataType(iC367294 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C367294DataType(iC367294 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C367294DataType(iC367294 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC367294
        RsC367294_EOF = Not RsC367294.Read()

        GoTo Comp_C332489

    Comp_C367300:

        ' IdColeção = 0?
        sCurrComponent = "IdColeção = 0?"
        C367300 = (fn_ConvertValueCompiled(C341055, C341055DataType, AKB_DecimalPoint, False) = 0)
        C367300DataType = AKBTypeConst.cAKBTypeLogical
        If C367300 Then
            GoTo Comp_C372222
        Else
            GoTo Comp_C174744
        End If

    Comp_C367301:

        ' AtbTrue_vDevePertencerColeção
        sCurrComponent = "AtbTrue_vDevePertencerColeção"
        C367301DataType = 4
        C377024 = fn_ConvertValueCompiled(1, 1, AKB_DecimalPoint)
        C367301 = True
        GoTo Comp_C332488

    Comp_C368875:

        ' TabPreçoVenda
        sCurrComponent = "TabPreçoVenda"
        C368875DataType = 0
        C368875DataType = C174732Datatype(6)
        If C368875DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC174732(5)) Then
          C368875 = Strings.RTrim(RsC174732(5))
        Else 
          C368875 = RsC174732(5)
        End If 

        GoTo Comp_C393054

    Comp_C372222:

        ' UpdateColeçao
        sCurrComponent = "UpdateColeçao"
        QueryC372222 = con.CreateCommand()
        QueryC372222.CommandText = QueryC372222.CommandText & " " & "UPDATE WF_PEDIDO_ITENS"
        QueryC372222.CommandText = QueryC372222.CommandText & " " & "SET WF_PEDIDO_ITENS.ID_COLECAO = NULL,"
        QueryC372222.CommandText = QueryC372222.CommandText & " " & "WF_PEDIDO_ITENS.SIGLA_PROD2 = NULL,"
        QueryC372222.CommandText = QueryC372222.CommandText & " " & "WF_PEDIDO_ITENS.ACESSO2 = NULL,"
        QueryC372222.CommandText = QueryC372222.CommandText & " " & "WF_PEDIDO_ITENS.COD_EMBALAGEM2 = NULL "
        QueryC372222.CommandText = QueryC372222.CommandText & " " & "WHERE WF_PEDIDO_ITENS.TP_PRODUTO = " & _ 
ConvertValue(P31857, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC372222.CommandText = QueryC372222.CommandText & " " & "AND WF_PEDIDO_ITENS.SEQ_ITEM = " & _ 
ConvertValue(C340930, C340930DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC372222.CommandText = QueryC372222.CommandText & " " & "AND WF_PEDIDO_ITENS.PEDIDO = " & _ 
ConvertValue(P31858, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC372222.Transaction = txn
        C372222 = QueryC372222.ExecuteNonQuery()
        C372222DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C174744

    Comp_C377024:

        ' vDevePertencerColeção
        sCurrComponent = "vDevePertencerColeção"
        C377024 = 0
        C377024DataType = 1
        GoTo Comp_C278978

    Comp_C393054:

        ' Cod_Embalagem
        sCurrComponent = "Cod_Embalagem"
        C393054DataType = 0
        C393054DataType = C174732Datatype(7)
        If C393054DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC174732(6)) Then
          C393054 = Strings.RTrim(RsC174732(6))
        Else 
          C393054 = RsC174732(6)
        End If 

        GoTo Comp_C419525

    Comp_C406123:

        ' Atualiza Coleções Pedidos = 1?
        sCurrComponent = "Atualiza Coleções Pedidos = 1?"
        C406123 = (fn_ConvertValueCompiled(C340891.rows(C340891CurrentRow - 1)(0), C340891DataType(1), AKB_DecimalPoint, False) = 1)
        C406123DataType = AKBTypeConst.cAKBTypeLogical
        If C406123 Then
            GoTo Comp_C174740
        Else
            GoTo Comp_C332492
        End If

    Comp_C419525:

        ' selTabPreçoNop
        sCurrComponent = "selTabPreçoNop"
        QueryC419525 = con.CreateCommand()
        QueryC419525.CommandText = QueryC419525.CommandText & " " & "SELECT COUNT(WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico) "
        QueryC419525.CommandText = QueryC419525.CommandText & " " & "FROM AKBUSER01.WF_ESTAB_JURIDICO"
        QueryC419525.CommandText = QueryC419525.CommandText & " " & "WHERE WF_ESTAB_JURIDICO.Tabela_Preco_Venda = " & _ 
ConvertValue(C368875, C368875DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC419525.Transaction = txn
        RsC419525 = QueryC419525.ExecuteReader()
        Dim iC419525 As Short
        ReDim C419525DataType(RsC419525.FieldCount)
        For iC419525 = 0 to RsC419525.FieldCount - 1
            Select Case RsC419525.GetDataTypeName(iC419525).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C419525DataType(iC419525 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C419525DataType(iC419525 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C419525DataType(iC419525 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC419525
        RsC419525_EOF = Not RsC419525.Read()

        GoTo Comp_C419526

    Comp_C419526:

        ' Tabela Preço NOP?
        sCurrComponent = "Tabela Preço NOP?"
        C419526 = (fn_ConvertValueCompiled(RsC419525(0), C419525DataType(1), AKB_DecimalPoint, False) > 0)
        C419526DataType = AKBTypeConst.cAKBTypeLogical
        If C419526 Then
            GoTo Comp_C419528
        Else
            GoTo Comp_C367300
        End If

    Comp_C419528:

        ' UpdateColeçãoNOP
        sCurrComponent = "UpdateColeçãoNOP"
        QueryC419528 = con.CreateCommand()
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "UPDATE WF_PEDIDO_ITENS"
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "SET WF_PEDIDO_ITENS.ID_COLECAO = (SELECT Max(WF_COLECAO_PRODUTOS.Id_Colecao )"
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PEDIDO_ITENS.Sigla_Prod  "
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PEDIDO_ITENS.Acesso "
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PEDIDO_ITENS.Cod_Embalagem ),"
        QueryC419528.CommandText = QueryC419528.CommandText & " " & ""
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "WF_PEDIDO_ITENS.SIGLA_PROD2 = WF_PEDIDO_ITENS.Sigla_Prod ,"
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "WF_PEDIDO_ITENS.ACESSO2 =  WF_PEDIDO_ITENS.Acesso ,"
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "WF_PEDIDO_ITENS.COD_EMBALAGEM2 = WF_PEDIDO_ITENS.Cod_Embalagem  "
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "WHERE WF_PEDIDO_ITENS.TP_PRODUTO = " & _ 
ConvertValue(P31857, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "AND WF_PEDIDO_ITENS.SEQ_ITEM = " & _ 
ConvertValue(C340930, C340930DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC419528.CommandText = QueryC419528.CommandText & " " & "AND WF_PEDIDO_ITENS.PEDIDO = " & _ 
ConvertValue(P31858, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC419528.Transaction = txn
        C419528 = QueryC419528.ExecuteNonQuery()
        C419528DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C367301

    Exit_R9977:

        Exit Function

    Err_R9977:
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
