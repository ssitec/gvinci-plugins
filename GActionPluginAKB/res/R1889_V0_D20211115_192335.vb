Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R1889

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

    Public Shared Function R1889(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P3374 As Double, ByVal P13730 As Double) As DataTable
    ' 
    ' 1889 - Gera Espelho - Versão: 0
    ' 
        'On Error GoTo Err_R1889
        Dim sCurrComponent as String

        Dim QueryC12814 As New Object
        Dim C12814 As Integer
        Dim C12814DataType As Short
        Dim QueryC12821 As New Object
        Dim C12821 As Integer
        Dim C12821DataType As Short
        Dim QueryC12822 As New Object
        Dim C12822 As Integer
        Dim C12822DataType As Short
        Dim QueryC12824 As New Object
        Dim C12824 As Integer
        Dim C12824DataType As Short
        Dim QueryC12825 As New Object
        Dim C12825 As Integer
        Dim C12825DataType As Short
        Dim C12826 As Object
        Dim C12826DataType As Short
        Dim C12827DataType() As Short
        Dim QueryC12828 As New Object
        Dim C12828 As Integer
        Dim C12828DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C12826

    Comp_C12814:

        ' DiasEnt
        sCurrComponent = "DiasEnt"
        QueryC12814 = con.CreateCommand()
        QueryC12814.CommandText = QueryC12814.CommandText & " " & "INSERT INTO AKBUSER01.WF_DIAS_NAO_ENTREGA"
        QueryC12814.CommandText = QueryC12814.CommandText & " " & " "
        QueryC12814.CommandText = QueryC12814.CommandText & " " & "(WF_DIAS_NAO_ENTREGA.Tp_Produto , WF_DIAS_NAO_ENTREGA.Pedido , "
        QueryC12814.CommandText = QueryC12814.CommandText & " " & "WF_DIAS_NAO_ENTREGA.Cod_Dia )"
        QueryC12814.CommandText = QueryC12814.CommandText & " " & ""
        QueryC12814.CommandText = QueryC12814.CommandText & " " & "SELECT WF_DIAS_NAO_ENTREGA.Tp_Produto , WF_DIAS_NAO_ENTREGA.Pedido * (-1) , "
        QueryC12814.CommandText = QueryC12814.CommandText & " " & "WF_DIAS_NAO_ENTREGA.Cod_Dia "
        QueryC12814.CommandText = QueryC12814.CommandText & " " & ""
        QueryC12814.CommandText = QueryC12814.CommandText & " " & "FROM AKBUSER01.WF_DIAS_NAO_ENTREGA "
        QueryC12814.CommandText = QueryC12814.CommandText & " " & ""
        QueryC12814.CommandText = QueryC12814.CommandText & " " & "WHERE  WF_DIAS_NAO_ENTREGA.Pedido = " & _ 
ConvertValue(P3374, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12814.CommandText = QueryC12814.CommandText & " " & "AND  WF_DIAS_NAO_ENTREGA.Tp_Produto = " & _ 
ConvertValue(P13730, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12814.CommandText = QueryC12814.CommandText & " " & ""
        QueryC12814.Transaction = txn
        C12814 = QueryC12814.ExecuteNonQuery()
        C12814DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C12821

    Comp_C12821:

        ' PedItens
        sCurrComponent = "PedItens"
        QueryC12821 = con.CreateCommand()
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_ITENS "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & ""
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "(WF_PEDIDO_ITENS.Tp_Produto , WF_PEDIDO_ITENS.Pedido , WF_PEDIDO_ITENS.Seq_Item , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Sigla_Prod , WF_PEDIDO_ITENS.Acesso , WF_PEDIDO_ITENS.Cod_Embalagem , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Prazo_Entrega , WF_PEDIDO_ITENS.Prazo_Entr_Cliente , WF_PEDIDO_ITENS.Porc_Desconto ,"
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Pedido , WF_PEDIDO_ITENS.Qtde_Atendida , WF_PEDIDO_ITENS.Tipo_Ped , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Flag_Pos_Ped , WF_PEDIDO_ITENS.Seq_Agrup_Compose , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Preco_Unit , WF_PEDIDO_ITENS.Preco_Digitado , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Qualid_Item_Estoque ,"
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Produto_Exclusivo ) "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & ""
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "SELECT WF_PEDIDO_ITENS.Tp_Produto , WF_PEDIDO_ITENS.Pedido * (-1) , WF_PEDIDO_ITENS.Seq_Item , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Sigla_Prod , WF_PEDIDO_ITENS.Acesso , WF_PEDIDO_ITENS.Cod_Embalagem , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Prazo_Entrega , WF_PEDIDO_ITENS.Prazo_Entr_Cliente , WF_PEDIDO_ITENS.Porc_Desconto , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Pedido , WF_PEDIDO_ITENS.Qtde_Atendida , WF_PEDIDO_ITENS.Tipo_Ped , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "299 , WF_PEDIDO_ITENS.Seq_Agrup_Compose , "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Preco_Unit , WF_PEDIDO_ITENS.Preco_Digitado , WF_PEDIDO_ITENS.Qualid_Item_Estoque ,"
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WF_PEDIDO_ITENS.Produto_Exclusivo "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & ""
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & ""
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P3374, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & "AND  WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P13730, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12821.CommandText = QueryC12821.CommandText & " " & ""
        QueryC12821.Transaction = txn
        C12821 = QueryC12821.ExecuteNonQuery()
        C12821DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C12822

    Comp_C12822:

        ' PedItensDesc
        sCurrComponent = "PedItensDesc"
        QueryC12822 = con.CreateCommand()
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_ITENS_DESC "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & ""
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "(WF_PED_ITENS_DESC.Tp_Produto , WF_PED_ITENS_DESC.Pedido , "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "WF_PED_ITENS_DESC.Seq_Item , WF_PED_ITENS_DESC.Tipo_Desconto , "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "WF_PED_ITENS_DESC.Porc_Desconto )"
        QueryC12822.CommandText = QueryC12822.CommandText & " " & ""
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "SELECT WF_PED_ITENS_DESC.Tp_Produto , WF_PED_ITENS_DESC.Pedido * (-1), "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "WF_PED_ITENS_DESC.Seq_Item , WF_PED_ITENS_DESC.Tipo_Desconto , "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "WF_PED_ITENS_DESC.Porc_Desconto "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "FROM AKBUSER01.WF_PED_ITENS_DESC "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & ""
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "WHERE WF_PED_ITENS_DESC.Pedido  =  " & _ 
ConvertValue(P3374, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & "AND WF_PED_ITENS_DESC.Tp_Produto = " & _ 
ConvertValue(P13730, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12822.CommandText = QueryC12822.CommandText & " " & ""
        QueryC12822.Transaction = txn
        C12822 = QueryC12822.ExecuteNonQuery()
        C12822DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C12827

    Comp_C12824:

        ' PedSeq
        sCurrComponent = "PedSeq"
        QueryC12824 = con.CreateCommand()
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_SEQ "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & ""
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "(WF_PEDIDO_SEQ.Tp_Produto , WF_PEDIDO_SEQ.Pedido , WF_PEDIDO_SEQ.Seq , "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WF_PEDIDO_SEQ.Prazo_Entrega_Fabrica , WF_PEDIDO_SEQ.Porc_Desc_Ped , "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WF_PEDIDO_SEQ.Cod_Embalagem , WF_PEDIDO_SEQ.Qtde_Embalagem , "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WF_PEDIDO_SEQ.Qtde_Volumes , WF_PEDIDO_SEQ.Tara , WF_PEDIDO_SEQ.Peso_Bruto , "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WF_PEDIDO_SEQ.Porc_Comissao ) "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & ""
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "SELECT WF_PEDIDO_SEQ.Tp_Produto , WF_PEDIDO_SEQ.Pedido * (-1) , WF_PEDIDO_SEQ.Seq , "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WF_PEDIDO_SEQ.Prazo_Entrega_Fabrica , WF_PEDIDO_SEQ.Porc_Desc_Ped , "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WF_PEDIDO_SEQ.Cod_Embalagem , WF_PEDIDO_SEQ.Qtde_Embalagem , "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WF_PEDIDO_SEQ.Qtde_Volumes , WF_PEDIDO_SEQ.Tara , WF_PEDIDO_SEQ.Peso_Bruto , "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WF_PEDIDO_SEQ.Porc_Comissao "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & ""
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_SEQ "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & ""
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P3374, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & "AND  WF_PEDIDO_SEQ.Tp_Produto = " & _ 
ConvertValue(P13730, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12824.CommandText = QueryC12824.CommandText & " " & ""
        QueryC12824.Transaction = txn
        C12824 = QueryC12824.ExecuteNonQuery()
        C12824DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C12825

    Comp_C12825:

        ' PedSeqDesc
        sCurrComponent = "PedSeqDesc"
        QueryC12825 = con.CreateCommand()
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_SEQ_DESCONTO "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & ""
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "(WF_PED_SEQ_DESCONTO.Tp_Produto , WF_PED_SEQ_DESCONTO.Pedido , "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "WF_PED_SEQ_DESCONTO.Seq , WF_PED_SEQ_DESCONTO.Tipo_Desconto , "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "WF_PED_SEQ_DESCONTO.Porc_Desconto)"
        QueryC12825.CommandText = QueryC12825.CommandText & " " & ""
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "SELECT WF_PED_SEQ_DESCONTO.Tp_Produto , WF_PED_SEQ_DESCONTO.Pedido * (-1), "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "WF_PED_SEQ_DESCONTO.Seq , WF_PED_SEQ_DESCONTO.Tipo_Desconto , "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "WF_PED_SEQ_DESCONTO.Porc_Desconto "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & ""
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_DESCONTO "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & ""
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P3374, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & "AND  WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P13730, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12825.CommandText = QueryC12825.CommandText & " " & ""
        QueryC12825.Transaction = txn
        C12825 = QueryC12825.ExecuteNonQuery()
        C12825DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C12814

    Comp_C12826:

        ' VTrue
        sCurrComponent = "VTrue"
        C12826 = 1
        C12826DataType = 4
        GoTo Comp_C12828

    Comp_C12827:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C12827 As DataTable = New DataTable()
        tb_C12827.Columns.Add("VTrue" & "")
        Dim row_C12827 As DataRow = tb_C12827.NewRow()
        row_C12827(0) = C12826
        tb_C12827.Rows.Add(row_C12827)
        R1889 = tb_C12827
        ReDim C12827DataType(1)
        C12827DataType(1) = C12826DataType
        ReturnDataType = C12827DataType

        GoTo Exit_R1889

    Comp_C12828:

        ' Pedido
        sCurrComponent = "Pedido"
        QueryC12828 = con.CreateCommand()
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & ""
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "( WF_PEDIDO.Tp_Produto , WF_PEDIDO.Pedido , WF_PEDIDO.Dt_Pedido , WF_PEDIDO.Cod_Cliente ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & " WF_PEDIDO.Cod_Repres, WF_PEDIDO.Num_Ped_Empresa , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Num_Ped_Cliente , WF_PEDIDO.Prazo_Entr_Cliente , WF_PEDIDO.Entrega_Parcial , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Obs_Comercial , WF_PEDIDO.Obs_Producao , WF_PEDIDO.Custo_Financ_no_Fat , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Flag_Pos_Ped , WF_PEDIDO.Porc_Custo_Financ , WF_PEDIDO.Cod_Pessoa_Estab_Juridico , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Cod_Redespacho , WF_PEDIDO.Cod_Frete_Redesp , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Cod_Transp , WF_PEDIDO.Cod_Frete , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Cod_Pagto , WF_PEDIDO.Tabela_Preco_Venda , WF_PEDIDO.Liberado_Credito ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & " WF_PEDIDO.Data_Manut_Credito , WF_PEDIDO.Usuario_Manut_Credito , WF_PEDIDO.Liberado_Diretoria , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Data_Manut_Cred_Diretoria , WF_PEDIDO.Usuario_Manut_Cred_Diretoria , WF_PEDIDO.Tabela_Preco_Custo , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Total_Itens , WF_PEDIDO.Verificado , WF_PEDIDO.Usuario_Verificacao , WF_PEDIDO.Data_Verificacao , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Porc_Custo_Financ_Incluso, WF_PEDIDO.Data_Base ,WF_PEDIDO.Data_Base_Faturamento  ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Departamento , WF_PEDIDO.Valor_Seguro , WF_PEDIDO.Valor_Frete ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Despesas_Acessorias , WF_PEDIDO.Local_Embarque ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Local_Destino , WF_PEDIDO.Gastos_Embarque ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.CODIGO_VIA,  WF_PEDIDO.Porc_Pg_Antec_Prod ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Valor_Antec_Prod , WF_PEDIDO.Pagto_Antecipado ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Bloquear_Producao , WF_PEDIDO.Bloquear_Faturamento ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Porc_Pg_Antec_Fat , WF_PEDIDO.Valor_Antec_Fat,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.CODIGO_DETALHE_PAGTO,  WF_PEDIDO.COD_FRETE_EXP,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.BONIFICACAO, WF_PEDIDO.Dias_Entrega, WF_PEDIDO.UF_Embarque , WF_PEDIDO.Desp_Acess_Entra_ICMS  )"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & ""
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "SELECT WF_PEDIDO.Tp_Produto , WF_PEDIDO.Pedido*(-1) , WF_PEDIDO.Dt_Pedido , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Cod_Cliente , WF_PEDIDO.Cod_Repres , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "'-' || SUBSTR(WF_PEDIDO.Num_Ped_Empresa , 1, 14),"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Num_Ped_Cliente , WF_PEDIDO.Prazo_Entr_Cliente , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Entrega_Parcial , WF_PEDIDO.Obs_Comercial , WF_PEDIDO.Obs_Producao , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Custo_Financ_no_Fat , 299 , WF_PEDIDO.Porc_Custo_Financ , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Cod_Pessoa_Estab_Juridico , WF_PEDIDO.Cod_Redespacho , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Cod_Frete_Redesp , WF_PEDIDO.Cod_Transp , WF_PEDIDO.Cod_Frete , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Cod_Pagto , WF_PEDIDO.Tabela_Preco_Venda , WF_PEDIDO.Liberado_Credito , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Data_Manut_Credito , WF_PEDIDO.Usuario_Manut_Credito , WF_PEDIDO.Liberado_Diretoria , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Data_Manut_Cred_Diretoria , WF_PEDIDO.Usuario_Manut_Cred_Diretoria , WF_PEDIDO.Tabela_Preco_Custo , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Total_Itens , WF_PEDIDO.Verificado , WF_PEDIDO.Usuario_Verificacao , WF_PEDIDO.Data_Verificacao , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Porc_Custo_Financ_Incluso ,WF_PEDIDO.Data_Base ,WF_PEDIDO.Data_Base_Faturamento ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Departamento, WF_PEDIDO.Valor_Seguro , WF_PEDIDO.Valor_Frete , WF_PEDIDO.Despesas_Acessorias ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Local_Embarque , WF_PEDIDO.Local_Destino ,WF_PEDIDO.Gastos_Embarque ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.CODIGO_VIA , WF_PEDIDO.Porc_Pg_Antec_Prod ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Valor_Antec_Prod , WF_PEDIDO.Pagto_Antecipado ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Bloquear_Producao , WF_PEDIDO.Bloquear_Faturamento , "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.Porc_Pg_Antec_Fat , WF_PEDIDO.Valor_Antec_Fat ,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.CODIGO_DETALHE_PAGTO, WF_PEDIDO.COD_FRETE_EXP,"
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WF_PEDIDO.BONIFICACAO, WF_PEDIDO.Dias_Entrega, WF_PEDIDO.UF_Embarque, WF_PEDIDO.Desp_Acess_Entra_ICMS  "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & ""
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & ""
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "WHERE WF_PEDIDO.Pedido = " & _ 
ConvertValue(P3374, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P13730, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12828.CommandText = QueryC12828.CommandText & " " & ""
        QueryC12828.Transaction = txn
        C12828 = QueryC12828.ExecuteNonQuery()
        C12828DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C12824

    Exit_R1889:

        Exit Function

    Err_R1889:
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
