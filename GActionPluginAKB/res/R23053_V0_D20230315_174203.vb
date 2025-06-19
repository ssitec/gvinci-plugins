Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R23053

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

    Public Shared Function R23053(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P89318 As Object, ByVal P89319 As Double, ByVal P89320 As Object, ByVal P89321 As Object, ByVal P89323 As Object) As DataTable
    ' 
    ' 23053 - Valida Parametros de Descontos - com Função - Versão: 0
    ' 
        On Error GoTo Err_R23053
        Dim sCurrComponent as String

        Dim QueryC556138 As New Object
        Dim RsC556138 As New Object
        Dim C556138DataType() As Short
        Dim RsC556138_EOF As Boolean
        Dim C556139 As Object
        Dim C556139DataType As Short
        Dim C556140 As Object
        Dim C556140DataType As Short
        Dim C556145 As Object
        Dim C556145DataType As Short
        Dim C556147 As Object
        Dim C556147DataType As Short
        Dim C556148 As Object
        Dim C556148DataType As Short
        Dim QueryC556149 As New Object
        Dim RsC556149 As New Object
        Dim C556149DataType() As Short
        Dim RsC556149_EOF As Boolean
        Dim C556150 As Object
        Dim C556150DataType As Short
        Dim C556151 As Object
        Dim C556151DataType As Short
        Dim C556152 As Object
        Dim C556152DataType As Short
        Dim C556153 As Object
        Dim C556153DataType As Short
        Dim C556154 As Object
        Dim C556154DataType As Short
        Dim C556155 As Boolean
        Dim C556155DataType As Short
        Dim C556156DataType() As Short
        Dim C556157 As Object
        Dim C556157DataType As Short
        Dim C556161 As Object
        Dim C556161DataType As Short
        Dim C556178 As Object
        Dim C556178DataType As Short
        Dim C556179 As Object
        Dim C556179DataType As Short
        Dim C556180 As Object
        Dim C556180DataType As Short
        Dim C556181 As Object
        Dim C556181DataType As Short
        Dim C556182 As Object
        Dim C556182DataType As Short
        Dim C556183 As Object
        Dim C556183DataType As Short
        Dim C556257 As Object
        Dim C556257DataType As Short
        Dim C556258 As Object
        Dim C556258DataType As Short
        Dim QueryC556262 As New Object
        Dim C556262 As Integer
        Dim C556262DataType As Short
        Dim C556307 As DataTable
        Dim C556307CurrentRow As Integer
        Dim C556307DataType() As Short
        Dim QueryC556308 As New Object
        Dim RsC556308 As New Object
        Dim C556308DataType() As Short
        Dim RsC556308_EOF As Boolean
        Dim C556309 As Boolean
        Dim C556309DataType As Short
        Dim QueryC556310 As New Object
        Dim C556310 As Integer
        Dim C556310DataType As Short
        Dim C556311 As Object
        Dim C556311DataType As Short
        Dim QueryC556312 As New Object
        Dim C556312 As Integer
        Dim C556312DataType As Short
        Dim C556313 As Object
        Dim C556313DataType As Short
        Dim C556314 As Object
        Dim C556314DataType As Short
        Dim C556315 As Boolean
        Dim C556315DataType As Short
        Dim QueryC556316 As New Object
        Dim C556316 As Integer
        Dim C556316DataType As Short
        Dim C556317 As Object
        Dim C556317DataType As Short
        Dim C556318 As Object
        Dim C556318DataType As Short
        Dim C556319 As Object
        Dim C556319DataType As Short
        Dim C556320 As Object
        Dim C556320DataType As Short
        Dim C556321 As Object
        Dim C556321DataType As Short
        Dim C556322 As Object
        Dim C556322DataType As Short
        Dim C556323 As Object
        Dim C556323DataType As Short
        Dim C556324 As Object
        Dim C556324DataType As Short
        Dim C556325 As Object
        Dim C556325DataType As Short
        Dim C556326 As Object
        Dim C556326DataType As Short
        Dim C556327 As Object
        Dim C556327DataType As Short
        Dim C556328 As Object
        Dim C556328DataType As Short
        Dim C556329 As Object
        Dim C556329DataType As Short
        Dim C556330 As Object
        Dim C556330DataType As Short
        Dim C556331 As Object
        Dim C556331DataType As Short
        Dim C556332 As Object
        Dim C556332DataType As Short
        Dim C556333 As Object
        Dim C556333DataType As Short
        Dim C556334 As Object
        Dim C556334DataType As Short
        Dim C556335 As Object
        Dim C556335DataType As Short
        Dim C556336 As Object
        Dim C556336DataType As Short
        Dim QueryC556337 As New Object
        Dim RsC556337 As New Object
        Dim C556337DataType() As Short
        Dim RsC556337_EOF As Boolean
        Dim C556338 As Object
        Dim C556338DataType As Short
        Dim C556339 As Boolean
        Dim C556339DataType As Short
        Dim QueryC556340 As New Object
        Dim RsC556340 As New Object
        Dim C556340DataType() As Short
        Dim RsC556340_EOF As Boolean
        Dim C556341 As Object
        Dim C556341DataType As Short
        Dim QueryC556342 As New Object
        Dim C556342 As Integer
        Dim C556342DataType As Short
        Dim C556343 As Boolean
        Dim C556343DataType As Short
        Dim QueryC556345 As New Object
        Dim RsC556345 As New Object
        Dim C556345DataType() As Short
        Dim RsC556345_EOF As Boolean
        Dim C556346 As Object
        Dim C556346DataType As Short
        Dim C556347 As Boolean
        Dim C556347DataType As Short
        Dim QueryC560127 As New Object
        Dim RsC560127 As New Object
        Dim C560127DataType() As Short
        Dim RsC560127_EOF As Boolean
        Dim QueryC560128 As New Object
        Dim RsC560128 As New Object
        Dim C560128DataType() As Short
        Dim RsC560128_EOF As Boolean
        Dim C560129 As Object
        Dim C560129DataType As Short
        Dim C560130 As Boolean
        Dim C560130DataType As Short
        Dim C560131 As Object
        Dim C560131DataType As Short
        Dim C560132 As Object
        Dim C560132DataType As Short
        Dim C560334 As Object
        Dim C560334DataType As Short
        Dim C560335 As Object
        Dim C560335DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C556338

    Comp_C556138:

        ' SelDadosPed
        sCurrComponent = "SelDadosPed"
        QueryC556138 = con.CreateCommand()
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "SELECT WF_PRE_PEDIDO.tabela_preco_venda, wf_cidades.sigla_estado, wf_pre_pedido.cod_cliente,"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "(select ROUND((SUM(quant_dias)/wf_condicao_pgto.qtde_parcelas),0) FROM wf_ocorrencias_cond_pagto"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "WHERE wf_condicao_pgto.cod_pagto = wf_ocorrencias_cond_pagto.cod_pagto) DIASMEDIO,"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "wf_tabela_preco_venda.prazo_medio_max_sem_custo, nvl(wf_condicao_pgto.A_Vista_Param_Tabela,0)"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "FROM WF_PRE_PEDIDO, WF_PESSOAS, WF_CIDADES, wf_condicao_pgto, wf_tabela_preco_venda"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "WHERE id_prepedido = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND WF_PESSOAS.COD_PESSOA = WF_PRE_PEDIDO.COD_CLIENTE"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND wf_pessoas.codigo_cidade = wf_cidades.codigo_cidade"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND wf_pre_pedido.cod_pagto = wf_condicao_pgto.cod_pagto"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "and wf_tabela_preco_venda.tabela_preco_venda = wf_pre_pedido.tabela_preco_venda"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND " & _ 
ConvertValue(P89319, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 1"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & ""
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "UNION"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & ""
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "SELECT DISTINCT " & _ 
ConvertValue(P89320, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", CID.Sigla_Estado, CLI.Cod_Cliente, 0, TV.prazo_medio_max_sem_custo, 0"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "FROM WF_PARAM_TAB_PRECO PA, WF_CLI_PARAM_TABELA CLI, WF_PESSOAS PE, WF_CIDADES CID,"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "wf_tabela_preco_venda TV"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "WHERE PA.Status = 'APROVADO'"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND TV.Tabela_Preco_Venda = " & _ 
ConvertValue(P89320, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND PA.Tabela_Preco_Venda = CLI.Tabela_Preco_Venda"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND PA.Seq = CLI.Seq"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND PE.Cod_Pessoa = CLI.Cod_Cliente"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND CLI.Cod_Cliente = " & _ 
ConvertValue(P89321, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND PE.Codigo_Cidade = CID.Codigo_Cidade"
        QueryC556138.CommandText = QueryC556138.CommandText & " " & "AND " & _ 
ConvertValue(P89319, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 2"
        RsC556138 = QueryC556138.ExecuteReader()
        Dim iC556138 As Short
        ReDim C556138DataType(RsC556138.FieldCount)
        For iC556138 = 0 to RsC556138.FieldCount - 1
            Select Case RsC556138.GetDataTypeName(iC556138).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556138DataType(iC556138 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556138DataType(iC556138 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556138DataType(iC556138 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556138
        RsC556138_EOF = Not RsC556138.Read()

        GoTo Comp_C556139

    Comp_C556139:

        ' vTabPreço
        sCurrComponent = "vTabPreço"
        C556139DataType = 0
        C556139 = RsC556138(0)
        C556139DataType = C556138Datatype(1)

        GoTo Comp_C556140

    Comp_C556140:

        ' vUF
        sCurrComponent = "vUF"
        C556140DataType = 0
        C556140 = RsC556138(1)
        C556140DataType = C556138Datatype(2)

        GoTo Comp_C556147

    Comp_C556145:

        ' vSeq
        sCurrComponent = "vSeq"
        C556145 = System.DBNull.Value
        C556145DataType = 1
        GoTo Comp_C556138

    Comp_C556147:

        ' vCodCliente
        sCurrComponent = "vCodCliente"
        C556147DataType = 0
        C556147 = RsC556138(2)
        C556147DataType = C556138Datatype(3)

        GoTo Comp_C556148

    Comp_C556148:

        ' vDiasMedio
        sCurrComponent = "vDiasMedio"
        C556148DataType = 0
        C556148 = RsC556138(3)
        C556148DataType = C556138Datatype(4)

        GoTo Comp_C556314

    Comp_C556149:

        ' SelItens
        sCurrComponent = "SelItens"
        QueryC556149 = con.CreateCommand()
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "SELECT ITENS.seq_item, ITENS.preco_sem_desconto, ITENS.preco_unit, ITENS.porc_desconto, ITENS.id_colecao,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT  EIO.Espec_Item_Operac"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_PADRAO_PRODUTO_ITENS PPI, WF_TABPRECO_VDA_PRODUTO TP, WF_ESPEC_ITENS_OPER EIO, WF_PRE_PEDIDO PED"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Padrao_Produto = PPI.Padrao_Produto"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Tabela_Preco_Venda = PED.Tabela_Preco_Venda"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PPI.Item_Oper = 137"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PPI.Espec_Item_Operac = EIO.Espec_Item_Operac"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PED.ID_PREPEDIDO = ITENS.ID_PREPEDIDO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Acesso = TP.Acesso) codcomercial,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT  EIO.Espec_Item_Operac"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_PADRAO_PRODUTO_ITENS PPI, WF_TABPRECO_VDA_PRODUTO TP, WF_ESPEC_ITENS_OPER EIO, WF_PRE_PEDIDO PED"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Padrao_Produto = PPI.Padrao_Produto"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Tabela_Preco_Venda = PED.Tabela_Preco_Venda"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PPI.Item_Oper = 139"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PPI.Espec_Item_Operac = EIO.Espec_Item_Operac"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PED.ID_PREPEDIDO = ITENS.ID_PREPEDIDO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Acesso = TP.Acesso) largura,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT TP.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_TABPRECO_VDA_PRODUTO TP, WF_TABELA_EMBALAGEM T,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WF_EMB_COMP_VDA_PROD ECV, WF_PRE_PEDIDO PED"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Tabela_Preco_Venda = PED.Tabela_Preco_Venda"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.COD_EMBALAGEM = T.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Sigla_Prod = ECV.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Acesso = ECV.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.COD_EMBALAGEM = ECV.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PED.ID_PREPEDIDO = ITENS.ID_PREPEDIDO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Acesso = TP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.COD_EMBALAGEM = TP.COD_EMBALAGEM) CodEmb,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT T1.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_PADRAO_PRODUTO_ITENS PPI, WF_TABPRECO_VDA_PRODUTO TP, WF_TABELA_EMBALAGEM T, WF_TABELA_EMBALAGEM T1,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WF_EMB_COMP_VDA_PROD ECV, WF_PRE_PEDIDO PED"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Padrao_Produto = PPI.Padrao_Produto"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Tabela_Preco_Venda = PED.Tabela_Preco_Venda"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.COD_EMBALAGEM = T.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Sigla_Prod = ECV.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Acesso = ECV.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.COD_EMBALAGEM = ECV.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ECV.Cod_Embalagem_Int_Emb = T1.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PED.ID_PREPEDIDO = ITENS.ID_PREPEDIDO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Acesso = TP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.COD_EMBALAGEM = TP.COD_EMBALAGEM) CodEmbInt,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT ECV.Fator_Conv_Cpr_Inter"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_PADRAO_PRODUTO_ITENS PPI, WF_TABPRECO_VDA_PRODUTO TP, WF_TABELA_EMBALAGEM T, WF_TABELA_EMBALAGEM T1,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WF_EMB_COMP_VDA_PROD ECV, WF_PRE_PEDIDO PED"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Padrao_Produto = PPI.Padrao_Produto"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Tabela_Preco_Venda = PED.Tabela_Preco_Venda"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.COD_EMBALAGEM = T.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Sigla_Prod = ECV.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.Acesso = ECV.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP.COD_EMBALAGEM = ECV.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ECV.Cod_Embalagem_Int_Emb = T1.COD_EMBALAGEM (+)"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PED.ID_PREPEDIDO = ITENS.ID_PREPEDIDO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Sigla_Prod = TP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.Acesso = TP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ITENS.COD_EMBALAGEM = TP.COD_EMBALAGEM) FatorConv,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "ITENS.SIGLA_PROD,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "ITENS.ACESSO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_PRE_PEDIDO_ITENS ITENS"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE ITENS.id_prepedido = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND " & _ 
ConvertValue(P89319, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 1"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & ""
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "UNION"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & ""
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "SELECT 0, TP1.Preco, 0, 0, "
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT MAX(CP.Id_Colecao)"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_COLECAO_PRODUTOS CP"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE TP1.Sigla_Prod = CP.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.Acesso = CP.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.Cod_Embalagem = CP.Cod_Embalagem"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND CP.Data_Validade_Inicial <= SYSDATE"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND (CP.Data_Validade_Final >= SYSDATE OR CP.Data_Validade_Final IS NULL))"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "id_colecao,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT  EIO.Espec_Item_Operac"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_PADRAO_PRODUTO_ITENS PPI, WF_ESPEC_ITENS_OPER EIO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP1.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP1.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Padrao_Produto = PPI.Padrao_Produto"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PPI.Item_Oper = 137"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PPI.Espec_Item_Operac = EIO.Espec_Item_Operac) codcomercial,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT  EIO.Espec_Item_Operac"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_PADRAO_PRODUTO_ITENS PPI,WF_ESPEC_ITENS_OPER EIO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP1.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP1.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Padrao_Produto = PPI.Padrao_Produto"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PPI.Item_Oper = 139"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND PPI.Espec_Item_Operac = EIO.Espec_Item_Operac) largura,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT TP1.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_TABELA_EMBALAGEM T, WF_EMB_COMP_VDA_PROD ECV"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP1.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP1.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.COD_EMBALAGEM = T.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.Sigla_Prod = ECV.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.Acesso = ECV.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.COD_EMBALAGEM = ECV.COD_EMBALAGEM) CodEmb,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT T1.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_TABELA_EMBALAGEM T, WF_TABELA_EMBALAGEM T1,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WF_EMB_COMP_VDA_PROD ECV"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP1.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP1.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.COD_EMBALAGEM = T.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.Sigla_Prod = ECV.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.Acesso = ECV.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.COD_EMBALAGEM = ECV.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ECV.Cod_Embalagem_Int_Emb = T1.COD_EMBALAGEM) CodEmbInt,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "(SELECT DISTINCT ECV.Fator_Conv_Cpr_Inter"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_ACESSOS A, WF_PADRAO_PRODUTO_ITENS PPI, WF_TABELA_EMBALAGEM T, WF_TABELA_EMBALAGEM T1,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WF_EMB_COMP_VDA_PROD ECV"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE A.Sigla_Prod = TP1.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Acesso = TP1.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND A.Padrao_Produto = PPI.Padrao_Produto"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.COD_EMBALAGEM = T.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.Sigla_Prod = ECV.Sigla_Prod"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.Acesso = ECV.Acesso"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND TP1.COD_EMBALAGEM = ECV.COD_EMBALAGEM"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND ECV.Cod_Embalagem_Int_Emb = T1.COD_EMBALAGEM (+)) FatorConv,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "TP1.SIGLA_PROD,"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "TP1.ACESSO"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "FROM WF_TABPRECO_VDA_PRODUTO TP1"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "WHERE TP1.Tabela_Preco_Venda = " & _ 
ConvertValue(P89323, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND NVL(TP1.Inativo,0) = 0"
        QueryC556149.CommandText = QueryC556149.CommandText & " " & "AND " & _ 
ConvertValue(P89319, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 2"
        RsC556149 = QueryC556149.ExecuteReader()
        Dim iC556149 As Short
        ReDim C556149DataType(RsC556149.FieldCount)
        For iC556149 = 0 to RsC556149.FieldCount - 1
            Select Case RsC556149.GetDataTypeName(iC556149).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556149DataType(iC556149 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556149DataType(iC556149 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556149DataType(iC556149 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556149
        RsC556149_EOF = Not RsC556149.Read()

        GoTo Comp_C556153

    Comp_C556150:

        ' PreçoUnit
        sCurrComponent = "PreçoUnit"
        C556150DataType = 0
        C556150 = RsC556149(2)
        C556150DataType = C556149Datatype(3)

        GoTo Comp_C556183

    Comp_C556151:

        ' PreçoSemDesc
        sCurrComponent = "PreçoSemDesc"
        C556151DataType = 0
        C556151 = RsC556149(1)
        C556151DataType = C556149Datatype(2)

        GoTo Comp_C556150

    Comp_C556152:

        ' SeqItem
        sCurrComponent = "SeqItem"
        C556152DataType = 0
        C556152 = RsC556149(0)
        C556152DataType = C556149Datatype(1)

        GoTo Comp_C556151

    Comp_C556153:

        ' PrimItens
        sCurrComponent = "PrimItens"
        C556153DataType = 4

        GoTo Comp_C556154

    Comp_C556154:

        ' FimItens
        sCurrComponent = "FimItens"
        C556154DataType = 4
        C556154 = RsC556149_EOF
        GoTo Comp_C556155

    Comp_C556155:

        ' FImItens?
        sCurrComponent = "FImItens?"
        C556155 = (fn_ConvertValueCompiled(C556154, C556154DataType, AKB_DecimalPoint, False) = 1)
        C556155DataType = AKBTypeConst.cAKBTypeLogical
        If C556155 Then
            GoTo Comp_C556343
        Else
            GoTo Comp_C556152
        End If

    Comp_C556156:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C556156 As DataTable = New DataTable()
        tb_C556156.Columns.Add("vRet" & "")
        Dim row_C556156 As DataRow = tb_C556156.NewRow()
        row_C556156(0) = C556311
        tb_C556156.Rows.Add(row_C556156)
        R23053 = tb_C556156
        ReDim C556156DataType(1)
        C556156DataType(1) = C556311DataType
        ReturnDataType = C556156DataType

        GoTo Exit_R23053

    Comp_C556157:

        ' PróxItens
        sCurrComponent = "PróxItens"
        C556157DataType = 4
        RsC556149.Read()
        C556157 = True

        GoTo Comp_C556154

    Comp_C556161:

        ' IdColeção
        sCurrComponent = "IdColeção"
        C556161DataType = 0
        C556161 = RsC556149(4)
        C556161DataType = C556149Datatype(5)

        GoTo Comp_C556182

    Comp_C556178:

        ' FatorConv
        sCurrComponent = "FatorConv"
        C556178DataType = 0
        C556178 = RsC556149(9)
        C556178DataType = C556149Datatype(10)

        GoTo Comp_C556258

    Comp_C556179:

        ' CodEmbInt
        sCurrComponent = "CodEmbInt"
        C556179DataType = 0
        C556179 = RsC556149(8)
        C556179DataType = C556149Datatype(9)

        GoTo Comp_C556178

    Comp_C556180:

        ' CodEmb
        sCurrComponent = "CodEmb"
        C556180DataType = 0
        C556180 = RsC556149(7)
        C556180DataType = C556149Datatype(8)

        GoTo Comp_C556179

    Comp_C556181:

        ' Largura
        sCurrComponent = "Largura"
        C556181DataType = 0
        C556181 = RsC556149(6)
        C556181DataType = C556149Datatype(7)

        GoTo Comp_C556180

    Comp_C556182:

        ' CodComerc
        sCurrComponent = "CodComerc"
        C556182DataType = 0
        C556182 = RsC556149(5)
        C556182DataType = C556149Datatype(6)

        GoTo Comp_C556181

    Comp_C556183:

        ' PorcDesc
        sCurrComponent = "PorcDesc"
        C556183DataType = 0
        C556183 = RsC556149(3)
        C556183DataType = C556149Datatype(4)

        GoTo Comp_C556161

    Comp_C556257:

        ' Acesso
        sCurrComponent = "Acesso"
        C556257DataType = 0
        C556257 = RsC556149(11)
        C556257DataType = C556149Datatype(12)

        GoTo Comp_C556345

    Comp_C556258:

        ' SiglaProd
        sCurrComponent = "SiglaProd"
        C556258DataType = 0
        C556258 = RsC556149(10)
        C556258DataType = C556149Datatype(11)

        GoTo Comp_C556257

    Comp_C556262:

        ' UpItens
        sCurrComponent = "UpItens"
        QueryC556262 = con.CreateCommand()
        QueryC556262.CommandText = QueryC556262.CommandText & " " & "UPDATE wf_pre_pedido_itens SET"
        QueryC556262.CommandText = QueryC556262.CommandText & " " & "OBS_BLOQ_PARAM_TAB = 'Sequencia não encontrada'"
        QueryC556262.CommandText = QueryC556262.CommandText & " " & "WHERE ID_PREPEDIDO = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556262.CommandText = QueryC556262.CommandText & " " & "AND SEQ_ITEM = " & _ 
ConvertValue(C556152, C556152DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556262.CommandText = QueryC556262.CommandText & " " & "AND " & _ 
ConvertValue(P89319, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 1"
        C556262 = QueryC556262.ExecuteNonQuery()
        C556262DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556337

    Comp_C556307:

        ' ValidaPreço
        sCurrComponent = "ValidaPreço"
        C556307 = clsRuleDynamicallyCompiled_R23005.R23005(con, txn, IIf(Not IsDBNull(C556145), C556145, System.DBNull.Value), IIf(Not IsDBNull(P89318), P89318, System.DBNull.Value), IIf(Not IsDBNull(C556139), C556139, System.DBNull.Value), IIf(Not IsDBNull(C556148), C556148, System.DBNull.Value), IIf(Not IsDBNull(C556152), C556152, System.DBNull.Value), IIf(Not IsDBNull(C556151), C556151, System.DBNull.Value), IIf(Not IsDBNull(C556150), C556150, System.DBNull.Value), IIf(Not IsDBNull(C556313), C556313, System.DBNull.Value), IIf(Not IsDBNull(P89319), P89319, System.DBNull.Value), IIf(Not IsDBNull(C556140), C556140, System.DBNull.Value), IIf(Not IsDBNull(C556258), C556258, System.DBNull.Value), IIf(Not IsDBNull(C556257), C556257, System.DBNull.Value), IIf(Not IsDBNull(C556180), C556180, System.DBNull.Value), IIf(Not IsDBNull(C556147), C556147, System.DBNull.Value))
        C556307CurrentRow = 1
        If Not IsNothing(C556307) Then
          iColumns = C556307.Columns.Count
        End If
        ReDim C556307DataType(iColumns)
        For iCol = 1 To iColumns
            C556307DataType(iCol) = clsRuleDynamicallyCompiled_R23005.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C556157

    Comp_C556308:

        ' SelItemBloq
        sCurrComponent = "SelItemBloq"
        QueryC556308 = con.CreateCommand()
        QueryC556308.CommandText = QueryC556308.CommandText & " " & "SELECT COUNT(*)"
        QueryC556308.CommandText = QueryC556308.CommandText & " " & "FROM WF_PRE_PEDIDO_ITENS"
        QueryC556308.CommandText = QueryC556308.CommandText & " " & "WHERE id_prepedido = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556308.CommandText = QueryC556308.CommandText & " " & "AND OBS_BLOQ_PARAM_TAB <> 'OK'"
        RsC556308 = QueryC556308.ExecuteReader()
        Dim iC556308 As Short
        ReDim C556308DataType(RsC556308.FieldCount)
        For iC556308 = 0 to RsC556308.FieldCount - 1
            Select Case RsC556308.GetDataTypeName(iC556308).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556308DataType(iC556308 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556308DataType(iC556308 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556308DataType(iC556308 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556308
        RsC556308_EOF = Not RsC556308.Read()

        GoTo Comp_C556309

    Comp_C556309:

        ' TemBloqueio?
        sCurrComponent = "TemBloqueio?"
        C556309 = (fn_ConvertValueCompiled(RsC556308(0), C556308DataType(1), AKB_DecimalPoint, False) > 0)
        C556309DataType = AKBTypeConst.cAKBTypeLogical
        If C556309 Then
            GoTo Comp_C556312
        Else
            GoTo Comp_C556156
        End If

    Comp_C556310:

        ' UpBloqPrePedido
        sCurrComponent = "UpBloqPrePedido"
        QueryC556310 = con.CreateCommand()
        QueryC556310.CommandText = QueryC556310.CommandText & " " & "UPDATE WF_PRE_PEDIDO SET"
        QueryC556310.CommandText = QueryC556310.CommandText & " " & "BLOQUEADO_PRECO_FORA = 1"
        QueryC556310.CommandText = QueryC556310.CommandText & " " & "WHERE ID_PREPEDIDO = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        C556310 = QueryC556310.ExecuteNonQuery()
        C556310DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556156

    Comp_C556311:

        ' vRet
        sCurrComponent = "vRet"
        C556311 = 1
        C556311DataType = 1
        GoTo Comp_C556145

    Comp_C556312:

        ' UpObsBloq
        sCurrComponent = "UpObsBloq"
        QueryC556312 = con.CreateCommand()
        QueryC556312.CommandText = QueryC556312.CommandText & " " & "UPDATE wf_pre_pedido_itens SET"
        QueryC556312.CommandText = QueryC556312.CommandText & " " & "OBS_BLOQ_PARAM_TAB = 'S/ SEQ'"
        QueryC556312.CommandText = QueryC556312.CommandText & " " & "WHERE ID_PREPEDIDO = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556312.CommandText = QueryC556312.CommandText & " " & "AND OBS_BLOQ_PARAM_TAB IS NULL"
        C556312 = QueryC556312.ExecuteNonQuery()
        C556312DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556310

    Comp_C556313:

        ' vAvista
        sCurrComponent = "vAvista"
        C556313DataType = 0
        C556313 = RsC556138(5)
        C556313DataType = C556138Datatype(6)

        GoTo Comp_C560129

    Comp_C556314:

        ' vPrazoMedioTabPreco
        sCurrComponent = "vPrazoMedioTabPreco"
        C556314DataType = 0
        C556314 = RsC556138(4)
        C556314DataType = C556138Datatype(5)

        GoTo Comp_C556313

    Comp_C556315:

        ' DiasMedioMaior?
        sCurrComponent = "DiasMedioMaior?"
        C556315 = (fn_ConvertValueCompiled(C556148, C556148DataType, AKB_DecimalPoint, False) > fn_ConvertValueCompiled(C560129, C560129DataType, AKB_DecimalPoint, False))
        C556315DataType = AKBTypeConst.cAKBTypeLogical
        If C556315 Then
            GoTo Comp_C556316
        Else
            GoTo Comp_C556308
        End If

    Comp_C556316:

        ' UpBloqDiasMedio
        sCurrComponent = "UpBloqDiasMedio"
        QueryC556316 = con.CreateCommand()
        QueryC556316.CommandText = QueryC556316.CommandText & " " & "UPDATE WF_PRE_PEDIDO SET"
        QueryC556316.CommandText = QueryC556316.CommandText & " " & "BLOQUEADO_PRECO_FORA = 1,"
        QueryC556316.CommandText = QueryC556316.CommandText & " " & "OBS_PRECO_FORA = 'Dias medios maior que permitido - DM calculado: " & _ 
ConvertValue(C556148, C556148DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ". DM tab ou Estab: " & _ 
ConvertValue(C560129, C560129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " '"
        QueryC556316.CommandText = QueryC556316.CommandText & " " & "WHERE ID_PREPEDIDO = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        C556316 = QueryC556316.ExecuteNonQuery()
        C556316DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556308

    Comp_C556317:

        ' ClienteUFNE
        sCurrComponent = "ClienteUFNE"
        C556317 = "and not exists (select * from WF_CLI_PARAM_TABELA where WF_CLI_PARAM_TABELA.Tabela_Preco_Venda = WF_UF_PARAM_TABELA.Tabela_Preco_Venda and WF_CLI_PARAM_TABELA.SEQ = WF_UF_PARAM_TABELA.SEQ)        "
        C556317DataType = 5
        GoTo Comp_C556333

    Comp_C556318:

        ' ColecUFNE
        sCurrComponent = "ColecUFNE"
        C556318 = "and not exists (select * from WF_COLEC_PARAM_TABELA where WF_COLEC_PARAM_TABELA.Tabela_Preco_Venda = WF_UF_PARAM_TABELA.Tabela_Preco_Venda and WF_COLEC_PARAM_TABELA.SEQ = WF_UF_PARAM_TABELA.SEQ)"
        C556318DataType = 5
        GoTo Comp_C556319

    Comp_C556319:

        ' ColecCliNE
        sCurrComponent = "ColecCliNE"
        C556319 = "and not exists (select * from WF_COLEC_PARAM_TABELA where WF_COLEC_PARAM_TABELA.Tabela_Preco_Venda = WF_CLI_PARAM_TABELA.Tabela_Preco_Venda and WF_COLEC_PARAM_TABELA.SEQ = WF_CLI_PARAM_TABELA.SEQ)"
        C556319DataType = 5
        GoTo Comp_C556320

    Comp_C556320:

        ' ColecCodComerNE
        sCurrComponent = "ColecCodComerNE"
        C556320 = "and not exists (select * from WF_COLEC_PARAM_TABELA where WF_COLEC_PARAM_TABELA.Tabela_Preco_Venda = WF_COF_COM_PARAM_TABELA.Tabela_Preco_Venda and WF_COLEC_PARAM_TABELA.SEQ = WF_COF_COM_PARAM_TABELA.SEQ)"
        C556320DataType = 5
        GoTo Comp_C556323

    Comp_C556321:

        ' CodComerCliNE
        sCurrComponent = "CodComerCliNE"
        C556321 = "and not exists (select * from WF_COF_COM_PARAM_TABELA where WF_COF_COM_PARAM_TABELA.Tabela_Preco_Venda = WF_CLI_PARAM_TABELA.Tabela_Preco_Venda and WF_COF_COM_PARAM_TABELA.SEQ = WF_CLI_PARAM_TABELA.SEQ)  "
        C556321DataType = 5
        GoTo Comp_C556322

    Comp_C556322:

        ' CodComerUFNE
        sCurrComponent = "CodComerUFNE"
        C556322 = "and not exists (select * from WF_COF_COM_PARAM_TABELA where WF_COF_COM_PARAM_TABELA.Tabela_Preco_Venda = WF_UF_PARAM_TABELA.Tabela_Preco_Venda and WF_COF_COM_PARAM_TABELA.SEQ = WF_UF_PARAM_TABELA.SEQ)"
        C556322DataType = 5
        GoTo Comp_C556335

    Comp_C556323:

        ' ColecLargNE
        sCurrComponent = "ColecLargNE"
        C556323 = "and not exists (select * from WF_COLEC_PARAM_TABELA where WF_COLEC_PARAM_TABELA.Tabela_Preco_Venda = WF_LARGURA_PARAM_TABELA.Tabela_Preco_Venda and WF_COLEC_PARAM_TABELA.SEQ = WF_LARGURA_PARAM_TABELA.SEQ)"
        C556323DataType = 5
        GoTo Comp_C556336

    Comp_C556324:

        ' ClienteLargNE
        sCurrComponent = "ClienteLargNE"
        C556324 = "and not exists (select * from WF_CLI_PARAM_TABELA where WF_CLI_PARAM_TABELA.Tabela_Preco_Venda = WF_LARGURA_PARAM_TABELA.Tabela_Preco_Venda and WF_CLI_PARAM_TABELA.SEQ = WF_LARGURA_PARAM_TABELA.SEQ)"
        C556324DataType = 5
        GoTo Comp_C556325

    Comp_C556325:

        ' ClienteCodComerNE
        sCurrComponent = "ClienteCodComerNE"
        C556325 = "and not exists (select * from WF_CLI_PARAM_TABELA where WF_CLI_PARAM_TABELA.Tabela_Preco_Venda = WF_COF_COM_PARAM_TABELA.Tabela_Preco_Venda and WF_CLI_PARAM_TABELA.SEQ = WF_COF_COM_PARAM_TABELA.SEQ)"
        C556325DataType = 5
        GoTo Comp_C556326

    Comp_C556326:

        ' ClienteColecNE
        sCurrComponent = "ClienteColecNE"
        C556326 = "and not exists (select * from WF_CLI_PARAM_TABELA where WF_CLI_PARAM_TABELA.Tabela_Preco_Venda = WF_COLEC_PARAM_TABELA.Tabela_Preco_Venda and WF_CLI_PARAM_TABELA.SEQ = WF_COLEC_PARAM_TABELA.SEQ)"
        C556326DataType = 5
        GoTo Comp_C556317

    Comp_C556327:

        ' CodComerLargNE
        sCurrComponent = "CodComerLargNE"
        C556327 = "and not exists (select * from WF_COF_COM_PARAM_TABELA where WF_COF_COM_PARAM_TABELA.Tabela_Preco_Venda = WF_LARGURA_PARAM_TABELA.Tabela_Preco_Venda and WF_COF_COM_PARAM_TABELA.SEQ = WF_LARGURA_PARAM_TABELA.SEQ)"
        C556327DataType = 5
        GoTo Comp_C556328

    Comp_C556328:

        ' CodComerColecNE
        sCurrComponent = "CodComerColecNE"
        C556328 = "and not exists (select * from WF_COF_COM_PARAM_TABELA where WF_COF_COM_PARAM_TABELA.Tabela_Preco_Venda = WF_COLEC_PARAM_TABELA.Tabela_Preco_Venda and WF_COF_COM_PARAM_TABELA.SEQ = WF_COLEC_PARAM_TABELA.SEQ)"
        C556328DataType = 5
        GoTo Comp_C556321

    Comp_C556329:

        ' LargUFNE
        sCurrComponent = "LargUFNE"
        C556329 = "and not exists (select * from WF_LARGURA_PARAM_TABELA where WF_LARGURA_PARAM_TABELA.Tabela_Preco_Venda = WF_UF_PARAM_TABELA.Tabela_Preco_Venda and WF_LARGURA_PARAM_TABELA.SEQ = WF_UF_PARAM_TABELA.SEQ)"
        C556329DataType = 5
        GoTo Comp_C556330

    Comp_C556330:

        ' LargCliNE
        sCurrComponent = "LargCliNE"
        C556330 = "and not exists (select * from WF_LARGURA_PARAM_TABELA where WF_LARGURA_PARAM_TABELA.Tabela_Preco_Venda = WF_CLI_PARAM_TABELA.Tabela_Preco_Venda and WF_LARGURA_PARAM_TABELA.SEQ = WF_CLI_PARAM_TABELA.SEQ)"
        C556330DataType = 5
        GoTo Comp_C556331

    Comp_C556331:

        ' LargColecNE
        sCurrComponent = "LargColecNE"
        C556331 = "and not exists (select * from WF_LARGURA_PARAM_TABELA where WF_LARGURA_PARAM_TABELA.Tabela_Preco_Venda = WF_COLEC_PARAM_TABELA.Tabela_Preco_Venda and WF_LARGURA_PARAM_TABELA.SEQ = WF_COLEC_PARAM_TABELA.SEQ)"
        C556331DataType = 5
        GoTo Comp_C556332

    Comp_C556332:

        ' LargCodComerNE
        sCurrComponent = "LargCodComerNE"
        C556332 = "and not exists (select * from WF_LARGURA_PARAM_TABELA where WF_LARGURA_PARAM_TABELA.Tabela_Preco_Venda = WF_COF_COM_PARAM_TABELA.Tabela_Preco_Venda and WF_LARGURA_PARAM_TABELA.SEQ = WF_COF_COM_PARAM_TABELA.SEQ)"
        C556332DataType = 5
        GoTo Comp_C556334

    Comp_C556333:

        ' CliEmbNE
        sCurrComponent = "CliEmbNE"
        C556333 = "and not exists (select * from WF_CLI_PARAM_TABELA where WF_CLI_PARAM_TABELA.Tabela_Preco_Venda = WF_EMBAL_PARAM_TABELA.Tabela_Preco_Venda and WF_CLI_PARAM_TABELA.SEQ = WF_EMBAL_PARAM_TABELA.SEQ) "
        C556333DataType = 5
        GoTo Comp_C556311

    Comp_C556334:

        ' LargEmbNE
        sCurrComponent = "LargEmbNE"
        C556334 = "and not exists (select * from WF_LARGURA_PARAM_TABELA where WF_LARGURA_PARAM_TABELA.Tabela_Preco_Venda = WF_EMBAL_PARAM_TABELA.Tabela_Preco_Venda and WF_LARGURA_PARAM_TABELA.SEQ = WF_EMBAL_PARAM_TABELA.SEQ)    "
        C556334DataType = 5
        GoTo Comp_C556327

    Comp_C556335:

        ' CodComerEmbNE
        sCurrComponent = "CodComerEmbNE"
        C556335 = "and not exists (select * from WF_COF_COM_PARAM_TABELA where WF_COF_COM_PARAM_TABELA.Tabela_Preco_Venda = WF_EMBAL_PARAM_TABELA.Tabela_Preco_Venda and WF_COF_COM_PARAM_TABELA.SEQ = WF_EMBAL_PARAM_TABELA.SEQ)"
        C556335DataType = 5
        GoTo Comp_C556318

    Comp_C556336:

        ' ColecEmbNE
        sCurrComponent = "ColecEmbNE"
        C556336 = "and not exists (select * from WF_COLEC_PARAM_TABELA where WF_COLEC_PARAM_TABELA.Tabela_Preco_Venda = WF_EMBAL_PARAM_TABELA.Tabela_Preco_Venda and WF_COLEC_PARAM_TABELA.SEQ = WF_EMBAL_PARAM_TABELA.SEQ)"
        C556336DataType = 5
        GoTo Comp_C556324

    Comp_C556337:

        ' SelUF
        sCurrComponent = "SelUF"
        QueryC556337 = con.CreateCommand()
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "SELECT COUNT(*)"
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "FROM WF_UF_PARAM_TABELA"
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "WHERE tabela_preco_venda = " & _ 
ConvertValue(C556139, C556139DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND sigla_estado = " & _ 
ConvertValue(C556140, C556140DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND EXISTS  (SELECT * FROM WF_PARAM_TAB_PRECO"
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "WHERE WF_UF_PARAM_TABELA.SEQ = WF_PARAM_TAB_PRECO.SEQ"
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND WF_UF_PARAM_TABELA.Tabela_Preco_Venda = WF_PARAM_TAB_PRECO.Tabela_Preco_Venda  "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND WF_PARAM_TAB_PRECO.Validade_Inicial <= SYSDATE "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND ( WF_PARAM_TAB_PRECO.Validade_Final >= SYSDATE + 1 OR WF_PARAM_TAB_PRECO.Validade_Final IS NULL)"
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND WF_PARAM_TAB_PRECO.Status = 'APROVADO')"
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND NOT EXISTS (SELECT * FROM WF_PROD_EXC_PARAM_TPV"
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "WHERE Tabela_preco_venda = " & _ 
ConvertValue(C556139, C556139DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND sigla_prod = " & _ 
ConvertValue(C556258, C556258DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND ACESSO = " & _ 
ConvertValue(C556257, C556257DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND COD_EMBALAGEM = " & _ 
ConvertValue(C556180, C556180DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "AND WF_PROD_EXC_PARAM_TPV.SEQ = WF_UF_PARAM_TABELA.SEQ)"
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "" & _ 
ConvertValue(C556317, C556317DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "" & _ 
ConvertValue(C556318, C556318DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "" & _ 
ConvertValue(C556322, C556322DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "" & _ 
ConvertValue(C556329, C556329DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556337.CommandText = QueryC556337.CommandText & " " & "" & _ 
ConvertValue(C556338, C556338DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        RsC556337 = QueryC556337.ExecuteReader()
        Dim iC556337 As Short
        ReDim C556337DataType(RsC556337.FieldCount)
        For iC556337 = 0 to RsC556337.FieldCount - 1
            Select Case RsC556337.GetDataTypeName(iC556337).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556337DataType(iC556337 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556337DataType(iC556337 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556337DataType(iC556337 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556337
        RsC556337_EOF = Not RsC556337.Read()

        GoTo Comp_C556339

    Comp_C556338:

        ' EmbUFNE
        sCurrComponent = "EmbUFNE"
        C556338 = "and not exists (select * from  WF_EMBAL_PARAM_TABELA where  WF_EMBAL_PARAM_TABELA.Tabela_Preco_Venda = WF_UF_PARAM_TABELA.Tabela_Preco_Venda and  WF_EMBAL_PARAM_TABELA.SEQ = WF_UF_PARAM_TABELA.SEQ) "
        C556338DataType = 5
        GoTo Comp_C556329

    Comp_C556339:

        ' SelUF=1
        sCurrComponent = "SelUF=1"
        C556339 = (fn_ConvertValueCompiled(RsC556337(0), C556337DataType(1), AKB_DecimalPoint, False) = 1)
        C556339DataType = AKBTypeConst.cAKBTypeLogical
        If C556339 Then
            GoTo Comp_C556340
        Else
            GoTo Comp_C556157
        End If

    Comp_C556340:

        ' SelSeqUF
        sCurrComponent = "SelSeqUF"
        QueryC556340 = con.CreateCommand()
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "SELECT SEQ"
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "FROM WF_UF_PARAM_TABELA"
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "WHERE tabela_preco_venda = " & _ 
ConvertValue(C556139, C556139DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND sigla_estado = " & _ 
ConvertValue(C556140, C556140DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND EXISTS  (SELECT * FROM WF_PARAM_TAB_PRECO"
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "WHERE WF_UF_PARAM_TABELA.SEQ = WF_PARAM_TAB_PRECO.SEQ"
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND WF_UF_PARAM_TABELA.Tabela_Preco_Venda = WF_PARAM_TAB_PRECO.Tabela_Preco_Venda  "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND WF_PARAM_TAB_PRECO.Validade_Inicial <= SYSDATE "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND ( WF_PARAM_TAB_PRECO.Validade_Final >= SYSDATE + 1 OR WF_PARAM_TAB_PRECO.Validade_Final IS NULL)"
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND WF_PARAM_TAB_PRECO.Status = 'APROVADO')"
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND NOT EXISTS (SELECT * FROM WF_PROD_EXC_PARAM_TPV"
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "WHERE Tabela_preco_venda = " & _ 
ConvertValue(C556139, C556139DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND sigla_prod = " & _ 
ConvertValue(C556258, C556258DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND ACESSO = " & _ 
ConvertValue(C556257, C556257DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND COD_EMBALAGEM = " & _ 
ConvertValue(C556180, C556180DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "AND WF_PROD_EXC_PARAM_TPV.SEQ = WF_UF_PARAM_TABELA.SEQ)"
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "" & _ 
ConvertValue(C556317, C556317DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "" & _ 
ConvertValue(C556318, C556318DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "" & _ 
ConvertValue(C556322, C556322DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "" & _ 
ConvertValue(C556329, C556329DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556340.CommandText = QueryC556340.CommandText & " " & "" & _ 
ConvertValue(C556338, C556338DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        RsC556340 = QueryC556340.ExecuteReader()
        Dim iC556340 As Short
        ReDim C556340DataType(RsC556340.FieldCount)
        For iC556340 = 0 to RsC556340.FieldCount - 1
            Select Case RsC556340.GetDataTypeName(iC556340).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556340DataType(iC556340 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556340DataType(iC556340 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556340DataType(iC556340 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556340
        RsC556340_EOF = Not RsC556340.Read()

        GoTo Comp_C556341

    Comp_C556341:

        ' attSeqUF
        sCurrComponent = "attSeqUF"
        C556341DataType = 4
        C556145 = fn_ConvertValueCompiled(RsC556340(0) , 1, AKB_DecimalPoint)
        C556341 = True
        GoTo Comp_C556342

    Comp_C556342:

        ' UpNullBloq
        sCurrComponent = "UpNullBloq"
        QueryC556342 = con.CreateCommand()
        QueryC556342.CommandText = QueryC556342.CommandText & " " & "UPDATE wf_pre_pedido_itens SET"
        QueryC556342.CommandText = QueryC556342.CommandText & " " & "OBS_BLOQ_PARAM_TAB = null"
        QueryC556342.CommandText = QueryC556342.CommandText & " " & "WHERE ID_PREPEDIDO = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556342.CommandText = QueryC556342.CommandText & " " & "AND SEQ_ITEM = " & _ 
ConvertValue(C556152, C556152DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556342.CommandText = QueryC556342.CommandText & " " & "AND " & _ 
ConvertValue(P89319, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 1"
        C556342 = QueryC556342.ExecuteNonQuery()
        C556342DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556307

    Comp_C556343:

        ' TPREGRA=2?
        sCurrComponent = "TPREGRA=2?"
        C556343 = (fn_ConvertValueCompiled(P89319, 1, AKB_DecimalPoint, False) = 2)
        C556343DataType = AKBTypeConst.cAKBTypeLogical
        If C556343 Then
            GoTo Comp_C556156
        Else
            GoTo Comp_C556315
        End If

    Comp_C556345:

        ' SeqFunc
        sCurrComponent = "SeqFunc"
        QueryC556345 = con.CreateCommand()
        QueryC556345.CommandText = QueryC556345.CommandText & " " & "SELECT AKBWEB.VALIDA_TAB_PRECO_VIRTUAL(NVL(" & _ 
ConvertValue(P89320, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "," & _ 
ConvertValue(C556139, C556139DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") , " & _ 
ConvertValue(C556140, C556140DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556147, C556147DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC556345.CommandText = QueryC556345.CommandText & " " & "NULL, " & _ 
ConvertValue(C556258, C556258DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556257, C556257DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556180, C556180DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) SEQ FROM DUAL"
        QueryC556345.CommandText = QueryC556345.CommandText & " " & ""
        RsC556345 = QueryC556345.ExecuteReader()
        Dim iC556345 As Short
        ReDim C556345DataType(RsC556345.FieldCount)
        For iC556345 = 0 to RsC556345.FieldCount - 1
            Select Case RsC556345.GetDataTypeName(iC556345).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556345DataType(iC556345 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556345DataType(iC556345 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556345DataType(iC556345 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556345
        RsC556345_EOF = Not RsC556345.Read()

        GoTo Comp_C556346

    Comp_C556346:

        ' ATRIBUIVA1
        sCurrComponent = "ATRIBUIVA1"
        C556346DataType = 4
        C556145 = fn_ConvertValueCompiled(RsC556345(0) , 1, AKB_DecimalPoint)
        C556346 = True
        GoTo Comp_C556347

    Comp_C556347:

        ' DESVIO2
        sCurrComponent = "DESVIO2"
        C556347 = (fn_ConvertValueCompiled(C556145, C556145DataType, AKB_DecimalPoint, False) Is System.DBNull.Value OR fn_ConvertValueCompiled(C556145, C556145DataType, AKB_DecimalPoint, False)=-1)
        C556347DataType = AKBTypeConst.cAKBTypeLogical
        If C556347 Then
            GoTo Comp_C556262
        Else
            GoTo Comp_C556307
        End If

    Comp_C560127:

        ' PMVCliente
        sCurrComponent = "PMVCliente"
        QueryC560127 = con.CreateCommand()
        QueryC560127.CommandText = QueryC560127.CommandText & " " & "SELECT NVL(WF_CLIENTES.Prazo_Medio_de_Venda,0) FROM AKBUSER01.WF_CLIENTES WHERE WF_CLIENTES.Cod_Cliente = " & _ 
ConvertValue(C556147, C556147DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        RsC560127 = QueryC560127.ExecuteReader()
        Dim iC560127 As Short
        ReDim C560127DataType(RsC560127.FieldCount)
        For iC560127 = 0 to RsC560127.FieldCount - 1
            Select Case RsC560127.GetDataTypeName(iC560127).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C560127DataType(iC560127 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C560127DataType(iC560127 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C560127DataType(iC560127 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC560127
        RsC560127_EOF = Not RsC560127.Read()

        GoTo Comp_C560334

    Comp_C560128:

        ' FlagPMVEstab
        sCurrComponent = "FlagPMVEstab"
        QueryC560128 = con.CreateCommand()
        QueryC560128.CommandText = QueryC560128.CommandText & " " & "SELECT NVL(E.Valida_PMV_em_Pedidos,0)"
        QueryC560128.CommandText = QueryC560128.CommandText & " " & "FROM WF_PRE_PEDIDO PP, WF_ESTAB_JURIDICO_PARAM E"
        QueryC560128.CommandText = QueryC560128.CommandText & " " & "WHERE PP.Id_PrePedido = " & _ 
ConvertValue(P89318, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC560128.CommandText = QueryC560128.CommandText & " " & "AND PP.Cod_Pessoa_Estab_Juridico = E.Cod_Pessoa_Estab_Juridico"
        QueryC560128.CommandText = QueryC560128.CommandText & " " & ""
        RsC560128 = QueryC560128.ExecuteReader()
        Dim iC560128 As Short
        ReDim C560128DataType(RsC560128.FieldCount)
        For iC560128 = 0 to RsC560128.FieldCount - 1
            Select Case RsC560128.GetDataTypeName(iC560128).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C560128DataType(iC560128 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C560128DataType(iC560128 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C560128DataType(iC560128 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC560128
        RsC560128_EOF = Not RsC560128.Read()

        GoTo Comp_C560335

    Comp_C560129:

        ' vPMV
        sCurrComponent = "vPMV"
        C560129 = 0
        C560129DataType = 1
        GoTo Comp_C560127

    Comp_C560130:

        ' DESVIO3
        sCurrComponent = "DESVIO3"
        C560130 = (fn_ConvertValueCompiled(C560335, C560335DataType, AKB_DecimalPoint, False) = 0)
        C560130DataType = AKBTypeConst.cAKBTypeLogical
        If C560130 Then
            GoTo Comp_C560131
        Else
            GoTo Comp_C560132
        End If

    Comp_C560131:

        ' ATRIBUIVA2
        sCurrComponent = "ATRIBUIVA2"
        C560131DataType = 4
        C560129 = fn_ConvertValueCompiled(C556314 , 1, AKB_DecimalPoint)
        C560131 = True
        GoTo Comp_C556149

    Comp_C560132:

        ' ATRIBUIVA3
        sCurrComponent = "ATRIBUIVA3"
        C560132DataType = 4
        C560129 = fn_ConvertValueCompiled(C560334 , 1, AKB_DecimalPoint)
        C560132 = True
        GoTo Comp_C556149

    Comp_C560334:

        ' vPMVCliente
        sCurrComponent = "vPMVCliente"
        C560334DataType = 0
        C560334 = RsC560127(0)
        C560334DataType = C560127Datatype(1)

        GoTo Comp_C560128

    Comp_C560335:

        ' vFlagPMVEstab
        sCurrComponent = "vFlagPMVEstab"
        C560335DataType = 0
        C560335 = RsC560128(0)
        C560335DataType = C560128Datatype(1)

        GoTo Comp_C560130

    Exit_R23053:

        Exit Function

    Err_R23053:
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
