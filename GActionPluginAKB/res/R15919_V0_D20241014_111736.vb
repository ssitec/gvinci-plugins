Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R15919

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

    Public Shared Function R15919(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P56542 As String, ByVal P56543 As Double, ByVal P56550 As Double, ByVal P56551 As Object, ByVal P56552 As Double, ByVal P56553 As Double, ByVal P56570 As Object, ByVal P56593 As Object, ByVal P60421 As Object, ByVal P61917 As Object, ByVal P64355 As Double, ByVal P65013 As Object, ByVal P65134 As Object) As DataTable
    ' 
    ' 15919 - Atualiza Coleções de Pedidos Sub - Versão: 0
    ' 
        'On Error GoTo Err_R15919
        Dim sCurrComponent as String

        Dim C340894DataType() As Short
        Dim C340907 As Boolean
        Dim C340907DataType As Short
        Dim QueryC340908 As New Object
        Dim C340908 As Integer
        Dim C340908DataType As Short
        Dim C341119 As Object
        Dim C341119DataType As Short
        Dim C343377 As Short
        Dim C343377DataType As Short
        Dim C343377ReturnDataType() As Short

        Dim QueryC368901 As New Object
        Dim RsC368901 As New Object
        Dim C368901DataType() As Short
        Dim RsC368901_EOF As Boolean
        Dim QueryC368902 As New Object
        Dim RsC368902 As New Object
        Dim C368902DataType() As Short
        Dim RsC368902_EOF As Boolean
        Dim C368903 As Boolean
        Dim C368903DataType As Short
        Dim C368905 As Object
        Dim C368905DataType As Short
        Dim C368916 As Object
        Dim C368916DataType As Short
        Dim C368918 As Object
        Dim C368918DataType As Short
        Dim QueryC375990 As New Object
        Dim RsC375990 As New Object
        Dim C375990DataType() As Short
        Dim RsC375990_EOF As Boolean
        Dim C375991 As Object
        Dim C375991DataType As Short
        Dim C375992 As Object
        Dim C375992DataType As Short
        Dim C375993 As Boolean
        Dim C375993DataType As Short
        Dim C376550 As Object
        Dim C376550DataType As Short
        Dim C376554 As Object
        Dim C376554DataType As Short
        Dim QueryC376700 As New Object
        Dim RsC376700 As New Object
        Dim C376700DataType() As Short
        Dim RsC376700_EOF As Boolean
        Dim C376922 As Object
        Dim C376922DataType As Short
        Dim C376923 As Object
        Dim C376923DataType As Short
        Dim C376924 As Object
        Dim C376924DataType As Short
        Dim C376925 As Boolean
        Dim C376925DataType As Short
        Dim C393445 As Object
        Dim C393445DataType As Short
        Dim C393446 As Object
        Dim C393446DataType As Short
        Dim C393447 As Object
        Dim C393447DataType As Short
        Dim C393448 As Object
        Dim C393448DataType As Short
        Dim C393449 As Object
        Dim C393449DataType As Short
        Dim C393450 As Object
        Dim C393450DataType As Short
        Dim C393452 As Object
        Dim C393452DataType As Short
        Dim C393454 As Object
        Dim C393454DataType As Short
        Dim C393455 As Object
        Dim C393455DataType As Short
        Dim C393456 As Object
        Dim C393456DataType As Short
        Dim C393457 As Object
        Dim C393457DataType As Short
        Dim C393458 As Object
        Dim C393458DataType As Short
        Dim C393459 As Object
        Dim C393459DataType As Short
        Dim C393460 As Object
        Dim C393460DataType As Short
        Dim C393461 As Object
        Dim C393461DataType As Short
        Dim QueryC397751 As New Object
        Dim RsC397751 As New Object
        Dim C397751DataType() As Short
        Dim RsC397751_EOF As Boolean
        Dim C397755 As Boolean
        Dim C397755DataType As Short
        Dim C397759 As Short
        Dim C397759DataType As Short
        Dim C397759ReturnDataType() As Short

        Dim C398040 As Object
        Dim C398040DataType As Short
        Dim QueryC398041 As New Object
        Dim RsC398041 As New Object
        Dim C398041DataType() As Short
        Dim RsC398041_EOF As Boolean
        Dim C398042 As Object
        Dim C398042DataType As Short
        Dim C398043 As Boolean
        Dim C398043DataType As Short
        Dim QueryC398044 As New Object
        Dim RsC398044 As New Object
        Dim C398044DataType() As Short
        Dim RsC398044_EOF As Boolean
        Dim C398045 As Boolean
        Dim C398045DataType As Short
        Dim C398046 As Object
        Dim C398046DataType As Short
        Dim C398638 As Boolean
        Dim C398638DataType As Short
        Dim C398643 As Boolean
        Dim C398643DataType As Short
        Dim C406119 As Boolean
        Dim C406119DataType As Short
        Dim C406120 As Object
        Dim C406120DataType As Short
        Dim C406121 As Object
        Dim C406121DataType As Short
        Dim QueryC540561 As New Object
        Dim RsC540561 As New Object
        Dim C540561DataType() As Short
        Dim RsC540561_EOF As Boolean
        Dim QueryC541693 As New Object
        Dim RsC541693 As New Object
        Dim C541693DataType() As Short
        Dim RsC541693_EOF As Boolean
        Dim C541694 As Boolean
        Dim C541694DataType As Short
        Dim C541696 As Object
        Dim C541696DataType As Short
        Dim QueryC541697 As New Object
        Dim RsC541697 As New Object
        Dim C541697DataType() As Short
        Dim RsC541697_EOF As Boolean
        Dim QueryC541698 As New Object
        Dim C541698 As Integer
        Dim C541698DataType As Short
        P65134 = IIf(IsDBNull(P65134), 1, P65134)

        ReDim ReturnDatatype(0)

        GoTo Comp_C406121

    Comp_C340894:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C340894 As DataTable = New DataTable()
        tb_C340894.Columns.Add("vRetorno" & "")
        Dim row_C340894 As DataRow = tb_C340894.NewRow()
        row_C340894(0) = C406121
        tb_C340894.Rows.Add(row_C340894)
        R15919 = tb_C340894
        ReDim C340894DataType(1)
        C340894DataType(1) = C406121DataType
        ReturnDataType = C340894DataType

        GoTo Exit_R15919

    Comp_C340907:

        ' SemColeção?
        sCurrComponent = "SemColeção?"
        C340907 = (fn_ConvertValueCompiled(RsC368902(0), C368902DataType(1), AKB_DecimalPoint, False) = 0)
        C340907DataType = AKBTypeConst.cAKBTypeLogical
        If C340907 Then
            GoTo Comp_C397751
        Else
            GoTo Comp_C368903
        End If

    Comp_C340908:

        ' UpdateColecaoAntiga
        sCurrComponent = "UpdateColecaoAntiga"
        QueryC340908 = con.CreateCommand()
        QueryC340908.CommandText = QueryC340908.CommandText & " " & "UPDATE WF_PEDIDO_ITENS"
        QueryC340908.CommandText = QueryC340908.CommandText & " " & "SET WF_PEDIDO_ITENS.ID_COLECAO = " & _ 
ConvertValue(C368916, C368916DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ","
        QueryC340908.CommandText = QueryC340908.CommandText & " " & "WF_PEDIDO_ITENS.SIGLA_PROD2 = " & _ 
ConvertValue(C393445, C393445DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ","
        QueryC340908.CommandText = QueryC340908.CommandText & " " & "WF_PEDIDO_ITENS.ACESSO2 = " & _ 
ConvertValue(C393446, C393446DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ","
        QueryC340908.CommandText = QueryC340908.CommandText & " " & "WF_PEDIDO_ITENS.COD_EMBALAGEM2 = " & _ 
ConvertValue(C393447, C393447DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC340908.CommandText = QueryC340908.CommandText & " " & "WHERE WF_PEDIDO_ITENS.TP_PRODUTO = " & _ 
ConvertValue(P56552, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC340908.CommandText = QueryC340908.CommandText & " " & "AND WF_PEDIDO_ITENS.SEQ_ITEM = " & _ 
ConvertValue(P56553, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC340908.CommandText = QueryC340908.CommandText & " " & "AND WF_PEDIDO_ITENS.PEDIDO = " & _ 
ConvertValue(P56550, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC340908.Transaction = txn
        C340908 = QueryC340908.ExecuteNonQuery()
        C340908DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C340894

    Comp_C341119:

        ' AtrFalsoSemErroAtualização
        sCurrComponent = "AtrFalsoSemErroAtualização"
        C341119DataType = 4
        C406121 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C341119 = True
        GoTo Comp_C343377

    Comp_C343377:

        ' MsgColeçãoEmBranco
        sCurrComponent = "MsgColeçãoEmBranco"
        C343377 = 1
        C343377DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C340894

    Comp_C368901:

        ' SelColeção
        sCurrComponent = "SelColeção"
        QueryC368901 = con.CreateCommand()
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "SELECT   WF_COLECAO.Id_Colecao,"
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "WF_COLECAO_PRODUTOS.Sigla_Prod ,"
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "WF_COLECAO_PRODUTOS.Acesso ,"
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "WF_COLECAO_PRODUTOS.Cod_Embalagem"
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS, AKBUSER01.WF_COLECAO"
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Id_Colecao  = WF_COLECAO.Id_Colecao"
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "AND (WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL "
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "OR  WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Sigla_Prod = " & _ 
ConvertValue(P56542, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = " & _ 
ConvertValue(P56543, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368901.CommandText = QueryC368901.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = " & _ 
ConvertValue(C398040, C398040DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368901.Transaction = txn
        RsC368901 = QueryC368901.ExecuteReader()
        Dim iC368901 As Short
        ReDim C368901DataType(RsC368901.FieldCount)
        For iC368901 = 0 to RsC368901.FieldCount - 1
            Select Case RsC368901.GetDataTypeName(iC368901).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C368901DataType(iC368901 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C368901DataType(iC368901 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C368901DataType(iC368901 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC368901
        RsC368901_EOF = Not RsC368901.Read()

        GoTo Comp_C368905

    Comp_C368902:

        ' SelN°ColeçõesProduto
        sCurrComponent = "SelN°ColeçõesProduto"
        QueryC368902 = con.CreateCommand()
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "SELECT   COUNT (WF_COLECAO.Id_Colecao)"
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS, AKBUSER01.WF_COLECAO"
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Id_Colecao  = WF_COLECAO.Id_Colecao"
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "AND (WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL  "
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "OR  WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Sigla_Prod = " & _ 
ConvertValue(P56542, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = " & _ 
ConvertValue(P56543, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368902.CommandText = QueryC368902.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = " & _ 
ConvertValue(C398040, C398040DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368902.Transaction = txn
        RsC368902 = QueryC368902.ExecuteReader()
        Dim iC368902 As Short
        ReDim C368902DataType(RsC368902.FieldCount)
        For iC368902 = 0 to RsC368902.FieldCount - 1
            Select Case RsC368902.GetDataTypeName(iC368902).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C368902DataType(iC368902 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C368902DataType(iC368902 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C368902DataType(iC368902 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC368902
        RsC368902_EOF = Not RsC368902.Read()

        GoTo Comp_C340907

    Comp_C368903:

        ' Coleção > 1?
        sCurrComponent = "Coleção > 1?"
        C368903 = (fn_ConvertValueCompiled(RsC368902(0), C368902DataType(1), AKB_DecimalPoint, False) > 1 AND fn_ConvertValueCompiled(RsC540561(0), C540561DataType(1), AKB_DecimalPoint, False)  = 0)
        C368903DataType = AKBTypeConst.cAKBTypeLogical
        If C368903 Then
            GoTo Comp_C376700
        Else
            GoTo Comp_C368901
        End If

    Comp_C368905:

        ' ValColeção
        sCurrComponent = "ValColeção"
        C368905DataType = 0
        C368905 = RsC368901(0)
        C368905DataType = C368901Datatype(1)
        If C368905DataType = AKBTypeConst.cAKBTypeString Then
          C368905 = IIF(IsDBNull(C368905), C368905, Strings.RTrim(C368905))
        End If 

        GoTo Comp_C393456

    Comp_C368916:

        ' vIDColeção
        sCurrComponent = "vIDColeção"
        C368916 = System.DBNull.Value
        C368916DataType = 1
        GoTo Comp_C393445

    Comp_C368918:

        ' AtrId_Coleção2
        sCurrComponent = "AtrId_Coleção2"
        C368918DataType = 4
        C368916 = fn_ConvertValueCompiled(RsC368901(0) , 1, AKB_DecimalPoint)
        C368918 = True
        GoTo Comp_C393459

    Comp_C375990:

        ' SelIdColeçãoxTabVenda
        sCurrComponent = "SelIdColeçãoxTabVenda"
        QueryC375990 = con.CreateCommand()
        QueryC375990.CommandText = QueryC375990.CommandText & " " & "SELECT WF_TABELA_PRECO_VENDA.Id_Colecao"
        QueryC375990.CommandText = QueryC375990.CommandText & " " & "FROM AKBUSER01.WF_TABELA_PRECO_VENDA"
        QueryC375990.CommandText = QueryC375990.CommandText & " " & "WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(P60421, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC375990.Transaction = txn
        RsC375990 = QueryC375990.ExecuteReader()
        Dim iC375990 As Short
        ReDim C375990DataType(RsC375990.FieldCount)
        For iC375990 = 0 to RsC375990.FieldCount - 1
            Select Case RsC375990.GetDataTypeName(iC375990).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C375990DataType(iC375990 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C375990DataType(iC375990 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C375990DataType(iC375990 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC375990
        RsC375990_EOF = Not RsC375990.Read()

        GoTo Comp_C376550

    Comp_C375991:

        ' Id_Coleção_Tabela_Preço_Venda
        sCurrComponent = "Id_Coleção_Tabela_Preço_Venda"
        C375991DataType = 0
        C375991 = RsC375990(0)
        C375991DataType = C375990Datatype(1)
        If C375991DataType = AKBTypeConst.cAKBTypeString Then
          C375991 = IIF(IsDBNull(C375991), C375991, Strings.RTrim(C375991))
        End If 

        GoTo Comp_C375992

    Comp_C375992:

        ' Atr_Id_Coleção_Tabela_Preço_Venda
        sCurrComponent = "Atr_Id_Coleção_Tabela_Preço_Venda"
        C375992DataType = 4
        C368916 = fn_ConvertValueCompiled(C375991 , 1, AKB_DecimalPoint)
        C375992 = True
        GoTo Comp_C376554

    Comp_C375993:

        ' Tb Venda sem Coleção?
        sCurrComponent = "Tb Venda sem Coleção?"
        C375993 = (fn_ConvertValueCompiled(C376554, C376554DataType, AKB_DecimalPoint, False) = 1)
        C375993DataType = AKBTypeConst.cAKBTypeLogical
        If C375993 Then
            GoTo Comp_C398643
        Else
            GoTo Comp_C340908
        End If

    Comp_C376550:

        ' Fim SelIdColeção
        sCurrComponent = "Fim SelIdColeção"
        C376550DataType = 4
        C376550 = RsC375990_EOF
        GoTo Comp_C406119

    Comp_C376554:

        ' ÉNuloIDColeçãoTabPreço
        sCurrComponent = "ÉNuloIDColeçãoTabPreço"
        C376554DataType = 4
        C376554 = IsDBNull (C375991)
        GoTo Comp_C375993

    Comp_C376700:

        ' SelColeção Venda x Produro
        sCurrComponent = "SelColeção Venda x Produro"
        QueryC376700 = con.CreateCommand()
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "SELECT WF_COLECAO.Id_Colecao ,"
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "WF_COLECAO_PRODUTOS.Sigla_Prod ,"
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "WF_COLECAO_PRODUTOS.Acesso ,"
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "WF_COLECAO_PRODUTOS.Cod_Embalagem"
        QueryC376700.CommandText = QueryC376700.CommandText & " " & ""
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "FROM AKBUSER01.WF_COLECAO , "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AKBUSER01.WF_TABELA_PRECO_VENDA , "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AKBUSER01.WF_COLECAO_PRODUTOS "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & ""
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "WHERE WF_COLECAO.Id_Colecao = WF_TABELA_PRECO_VENDA.Id_Colecao "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Id_Colecao = WF_TABELA_PRECO_VENDA.Id_Colecao"
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AND WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(P60421, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Sigla_Prod = " & _ 
ConvertValue(P56542, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = " & _ 
ConvertValue(P56543, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = " & _ 
ConvertValue(C398040, C398040DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AND (WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL  "
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "OR  WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC376700.CommandText = QueryC376700.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC376700.Transaction = txn
        RsC376700 = QueryC376700.ExecuteReader()
        Dim iC376700 As Short
        ReDim C376700DataType(RsC376700.FieldCount)
        For iC376700 = 0 to RsC376700.FieldCount - 1
            Select Case RsC376700.GetDataTypeName(iC376700).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C376700DataType(iC376700 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C376700DataType(iC376700 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C376700DataType(iC376700 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC376700
        RsC376700_EOF = Not RsC376700.Read()

        GoTo Comp_C376922

    Comp_C376922:

        ' FimSelColeção Venda x Produro
        sCurrComponent = "FimSelColeção Venda x Produro"
        C376922DataType = 4
        C376922 = RsC376700_EOF
        GoTo Comp_C376925

    Comp_C376923:

        ' Id_Coleção1
        sCurrComponent = "Id_Coleção1"
        C376923DataType = 0
        C376923 = RsC376700(0)
        C376923DataType = C376700Datatype(1)
        If C376923DataType = AKBTypeConst.cAKBTypeString Then
          C376923 = IIF(IsDBNull(C376923), C376923, Strings.RTrim(C376923))
        End If 

        GoTo Comp_C393448

    Comp_C376924:

        ' AtrId_Coleção1
        sCurrComponent = "AtrId_Coleção1"
        C376924DataType = 4
        C368916 = fn_ConvertValueCompiled(C376923 , 1, AKB_DecimalPoint)
        C376924 = True
        GoTo Comp_C393452

    Comp_C376925:

        ' SelColeção Venda x Produto Vazio?
        sCurrComponent = "SelColeção Venda x Produto Vazio?"
        C376925 = (fn_ConvertValueCompiled(C376922, C376922DataType, AKB_DecimalPoint, False) = 1)
        C376925DataType = AKBTypeConst.cAKBTypeLogical
        If C376925 Then
            GoTo Comp_C541693
        Else
            GoTo Comp_C376923
        End If

    Comp_C393445:

        ' vSigla_Prod_Colec
        sCurrComponent = "vSigla_Prod_Colec"
        C393445 = System.DBNull.Value
        C393445DataType = 3
        GoTo Comp_C393446

    Comp_C393446:

        ' vAcesso_Colec
        sCurrComponent = "vAcesso_Colec"
        C393446 = System.DBNull.Value
        C393446DataType = 1
        GoTo Comp_C393447

    Comp_C393447:

        ' vCod_Embalagem_Colec
        sCurrComponent = "vCod_Embalagem_Colec"
        C393447 = System.DBNull.Value
        C393447DataType = 1
        GoTo Comp_C398040

    Comp_C393448:

        ' Sigla_Prod_Colec1
        sCurrComponent = "Sigla_Prod_Colec1"
        C393448DataType = 0
        C393448DataType = C376700Datatype(2)
        If C393448DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC376700(1)) Then
          C393448 = Strings.RTrim(RsC376700(1))
        Else 
          C393448 = RsC376700(1)
        End If 

        GoTo Comp_C393449

    Comp_C393449:

        ' Acesso_Colec1
        sCurrComponent = "Acesso_Colec1"
        C393449DataType = 0
        C393449DataType = C376700Datatype(3)
        If C393449DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC376700(2)) Then
          C393449 = Strings.RTrim(RsC376700(2))
        Else 
          C393449 = RsC376700(2)
        End If 

        GoTo Comp_C393450

    Comp_C393450:

        ' Cod_Emb_Colec1
        sCurrComponent = "Cod_Emb_Colec1"
        C393450DataType = 0
        C393450DataType = C376700Datatype(4)
        If C393450DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC376700(3)) Then
          C393450 = Strings.RTrim(RsC376700(3))
        Else 
          C393450 = RsC376700(3)
        End If 

        GoTo Comp_C376924

    Comp_C393452:

        ' AtrSigla_Prod_Colec1
        sCurrComponent = "AtrSigla_Prod_Colec1"
        C393452DataType = 4
        C393445 = fn_ConvertValueCompiled(C393448 , 3, AKB_DecimalPoint)
        C393452 = True
        GoTo Comp_C393454

    Comp_C393454:

        ' AtrAcesso_Colec1
        sCurrComponent = "AtrAcesso_Colec1"
        C393454DataType = 4
        C393446 = fn_ConvertValueCompiled(C393449 , 1, AKB_DecimalPoint)
        C393454 = True
        GoTo Comp_C393455

    Comp_C393455:

        ' AtrCod_Emb_Colec1
        sCurrComponent = "AtrCod_Emb_Colec1"
        C393455DataType = 4
        C393447 = fn_ConvertValueCompiled(C393450 , 1, AKB_DecimalPoint)
        C393455 = True
        GoTo Comp_C340908

    Comp_C393456:

        ' Sigla_Prod_Colec2
        sCurrComponent = "Sigla_Prod_Colec2"
        C393456DataType = 0
        C393456DataType = C368901Datatype(2)
        If C393456DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC368901(1)) Then
          C393456 = Strings.RTrim(RsC368901(1))
        Else 
          C393456 = RsC368901(1)
        End If 

        GoTo Comp_C393457

    Comp_C393457:

        ' Acesso_Colec2
        sCurrComponent = "Acesso_Colec2"
        C393457DataType = 0
        C393457DataType = C368901Datatype(3)
        If C393457DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC368901(2)) Then
          C393457 = Strings.RTrim(RsC368901(2))
        Else 
          C393457 = RsC368901(2)
        End If 

        GoTo Comp_C393458

    Comp_C393458:

        ' Cod_Emb_Colec2
        sCurrComponent = "Cod_Emb_Colec2"
        C393458DataType = 0
        C393458DataType = C368901Datatype(4)
        If C393458DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC368901(3)) Then
          C393458 = Strings.RTrim(RsC368901(3))
        Else 
          C393458 = RsC368901(3)
        End If 

        GoTo Comp_C368918

    Comp_C393459:

        ' AtrSigla_Prod_Colec2
        sCurrComponent = "AtrSigla_Prod_Colec2"
        C393459DataType = 4
        C393445 = fn_ConvertValueCompiled(C393456 , 3, AKB_DecimalPoint)
        C393459 = True
        GoTo Comp_C393460

    Comp_C393460:

        ' AtrAcesso_Colec2
        sCurrComponent = "AtrAcesso_Colec2"
        C393460DataType = 4
        C393446 = fn_ConvertValueCompiled(C393457 , 1, AKB_DecimalPoint)
        C393460 = True
        GoTo Comp_C393461

    Comp_C393461:

        ' AtrCod_Emb_Colec2
        sCurrComponent = "AtrCod_Emb_Colec2"
        C393461DataType = 4
        C393447 = fn_ConvertValueCompiled(C393458 , 1, AKB_DecimalPoint)
        C393461 = True
        GoTo Comp_C340908

    Comp_C397751:

        ' SelAssumeColTabPreço
        sCurrComponent = "SelAssumeColTabPreço"
        QueryC397751 = con.CreateCommand()
        QueryC397751.CommandText = QueryC397751.CommandText & " " & "SELECT NVL(WF_Estab_Juridico_Param.Assume_Col_Tab_P_Prod_S_Col,0)"
        QueryC397751.CommandText = QueryC397751.CommandText & " " & "FROM WF_Estab_Juridico_Param"
        QueryC397751.CommandText = QueryC397751.CommandText & " " & "WHERE WF_Estab_Juridico_Param.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P65013, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC397751.Transaction = txn
        RsC397751 = QueryC397751.ExecuteReader()
        Dim iC397751 As Short
        ReDim C397751DataType(RsC397751.FieldCount)
        For iC397751 = 0 to RsC397751.FieldCount - 1
            Select Case RsC397751.GetDataTypeName(iC397751).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C397751DataType(iC397751 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C397751DataType(iC397751 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C397751DataType(iC397751 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC397751
        RsC397751_EOF = Not RsC397751.Read()

        GoTo Comp_C397755

    Comp_C397755:

        ' AssumeColTabPreço = 1?
        sCurrComponent = "AssumeColTabPreço = 1?"
        C397755 = (fn_ConvertValueCompiled(RsC397751(0), C397751DataType(1), AKB_DecimalPoint, False) = 1)
        C397755DataType = AKBTypeConst.cAKBTypeLogical
        If C397755 Then
            GoTo Comp_C375990
        Else
            GoTo Comp_C398643
        End If

    Comp_C397759:

        ' Msg_Produto_Sem_Coleção
        sCurrComponent = "Msg_Produto_Sem_Coleção"
        C397759 = 1
        C397759DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C340894

    Comp_C398040:

        ' vCod_Emb
        sCurrComponent = "vCod_Emb"
        C398040 = P64355 & " "
        C398040DataType = 1
        GoTo Comp_C540561

    Comp_C398041:

        ' SelEmbInterna
        sCurrComponent = "SelEmbInterna"
        QueryC398041 = con.CreateCommand()
        QueryC398041.CommandText = QueryC398041.CommandText & " " & "SELECT WF_Emb_Comp_Vda_Prod.Cod_Embalagem_Int_Emb"
        QueryC398041.CommandText = QueryC398041.CommandText & " " & "FROM WF_Emb_Comp_Vda_Prod"
        QueryC398041.CommandText = QueryC398041.CommandText & " " & "WHERE WF_Emb_Comp_Vda_Prod.Sigla_Prod = " & _ 
ConvertValue(P56542, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398041.CommandText = QueryC398041.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Acesso =  " & _ 
ConvertValue(P56543, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398041.CommandText = QueryC398041.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Cod_Embalagem = " & _ 
ConvertValue(P64355, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398041.CommandText = QueryC398041.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Cod_Embalagem_Int_Emb IS NOT NULL"
        QueryC398041.Transaction = txn
        RsC398041 = QueryC398041.ExecuteReader()
        Dim iC398041 As Short
        ReDim C398041DataType(RsC398041.FieldCount)
        For iC398041 = 0 to RsC398041.FieldCount - 1
            Select Case RsC398041.GetDataTypeName(iC398041).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C398041DataType(iC398041 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C398041DataType(iC398041 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C398041DataType(iC398041 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC398041
        RsC398041_EOF = Not RsC398041.Read()

        GoTo Comp_C398042

    Comp_C398042:

        ' Fim_Emb_Interna
        sCurrComponent = "Fim_Emb_Interna"
        C398042DataType = 4
        C398042 = RsC398041_EOF
        GoTo Comp_C398043

    Comp_C398043:

        ' Fim_Emb_Interna = 1?
        sCurrComponent = "Fim_Emb_Interna = 1?"
        C398043 = (fn_ConvertValueCompiled(C398042, C398042DataType, AKB_DecimalPoint, False) = 1)
        C398043DataType = AKBTypeConst.cAKBTypeLogical
        If C398043 Then
            GoTo Comp_C368902
        Else
            GoTo Comp_C398044
        End If

    Comp_C398044:

        ' VerificaSeProdutoEmbTemColeção
        sCurrComponent = "VerificaSeProdutoEmbTemColeção"
        QueryC398044 = con.CreateCommand()
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "SELECT COUNT(*)"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "FROM WF_Emb_Comp_Vda_Prod,"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "  WF_Colecao_Produtos,"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "  WF_Colecao"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "WHERE WF_Emb_Comp_Vda_Prod.Sigla_Prod = " & _ 
ConvertValue(P56542, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Acesso = " & _ 
ConvertValue(P56543, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Cod_Embalagem = " & _ 
ConvertValue(P64355, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Sigla_Prod = WF_Colecao_Produtos.Sigla_Prod"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Acesso = WF_Colecao_Produtos.Acesso"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Cod_Embalagem = WF_Colecao_Produtos.Cod_Embalagem"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND WF_Colecao.Id_Colecao = WF_Colecao_Produtos.Id_Colecao"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND (WF_Colecao.Data_Validade_Final IS NULL"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "OR WF_Colecao.Data_Validade_Final >= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND (WF_Colecao_Produtos.Data_Validade_Final IS NULL"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "OR WF_Colecao_Produtos.Data_Validade_Final >= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND WF_Colecao.Data_Validade_Inicial <= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398044.CommandText = QueryC398044.CommandText & " " & "AND WF_Colecao_Produtos.Data_Validade_Inicial <= " & _ 
ConvertValue(P56551, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC398044.Transaction = txn
        RsC398044 = QueryC398044.ExecuteReader()
        Dim iC398044 As Short
        ReDim C398044DataType(RsC398044.FieldCount)
        For iC398044 = 0 to RsC398044.FieldCount - 1
            Select Case RsC398044.GetDataTypeName(iC398044).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C398044DataType(iC398044 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C398044DataType(iC398044 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C398044DataType(iC398044 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC398044
        RsC398044_EOF = Not RsC398044.Read()

        GoTo Comp_C398045

    Comp_C398045:

        ' ProdutoEmbNãoTemColeção?
        sCurrComponent = "ProdutoEmbNãoTemColeção?"
        C398045 = (fn_ConvertValueCompiled(RsC398044(0), C398044DataType(1), AKB_DecimalPoint, False) = 0)
        C398045DataType = AKBTypeConst.cAKBTypeLogical
        If C398045 Then
            GoTo Comp_C398046
        Else
            GoTo Comp_C368902
        End If

    Comp_C398046:

        ' AtrEmbInterna
        sCurrComponent = "AtrEmbInterna"
        C398046DataType = 4
        C398040 = fn_ConvertValueCompiled(RsC398041(0) , 1, AKB_DecimalPoint)
        C398046 = True
        GoTo Comp_C368902

    Comp_C398638:

        ' MostrarMensagens? = 1
        sCurrComponent = "MostrarMensagens? = 1"
        C398638 = (fn_ConvertValueCompiled(P65134, 4, AKB_DecimalPoint, False) = 1 AND fn_ConvertValueCompiled(P61917, 1, AKB_DecimalPoint, False) = 1 AND fn_ConvertValueCompiled(P65013, 1, AKB_DecimalPoint, False)  <> 6)
        C398638DataType = AKBTypeConst.cAKBTypeLogical
        If C398638 Then
            GoTo Comp_C341119
        Else
            GoTo Comp_C340894
        End If

    Comp_C398643:

        ' Mostrar_Mensagem_Nenhuma_Coleção?
        sCurrComponent = "Mostrar_Mensagem_Nenhuma_Coleção?"
        C398643 = (fn_ConvertValueCompiled(P65134, 4, AKB_DecimalPoint, False) = 1 AND fn_ConvertValueCompiled(P61917, 1, AKB_DecimalPoint, False) = 1)
        C398643DataType = AKBTypeConst.cAKBTypeLogical
        If C398643 Then
            GoTo Comp_C406120
        Else
            GoTo Comp_C340894
        End If

    Comp_C406119:

        ' FimSelIdColeção = 1?
        sCurrComponent = "FimSelIdColeção = 1?"
        C406119 = (fn_ConvertValueCompiled(C376550, C376550DataType, AKB_DecimalPoint, False) = 1)
        C406119DataType = AKBTypeConst.cAKBTypeLogical
        If C406119 Then
            GoTo Comp_C398643
        Else
            GoTo Comp_C375991
        End If

    Comp_C406120:

        ' AtrFalsoSemErroAtualização2
        sCurrComponent = "AtrFalsoSemErroAtualização2"
        C406120DataType = 4
        C406121 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C406120 = True
        GoTo Comp_C397759

    Comp_C406121:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C406121 = 1
        C406121DataType = 4
        GoTo Comp_C368916

    Comp_C540561:

        ' SelProforma
        sCurrComponent = "SelProforma"
        QueryC540561 = con.CreateCommand()
        QueryC540561.CommandText = QueryC540561.CommandText & " " & "SELECT NVL(PROFORMA,0)"
        QueryC540561.CommandText = QueryC540561.CommandText & " " & "FROM WF_PRE_PEDIDO"
        QueryC540561.CommandText = QueryC540561.CommandText & " " & "WHERE PEDIDO = " & _ 
ConvertValue(P56550, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC540561.Transaction = txn
        RsC540561 = QueryC540561.ExecuteReader()
        Dim iC540561 As Short
        ReDim C540561DataType(RsC540561.FieldCount)
        For iC540561 = 0 to RsC540561.FieldCount - 1
            Select Case RsC540561.GetDataTypeName(iC540561).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C540561DataType(iC540561 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C540561DataType(iC540561 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C540561DataType(iC540561 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC540561
        RsC540561_EOF = Not RsC540561.Read()

        GoTo Comp_C398041

    Comp_C541693:

        ' FlagEcommerce
        sCurrComponent = "FlagEcommerce"
        QueryC541693 = con.CreateCommand()
        QueryC541693.CommandText = QueryC541693.CommandText & " " & "SELECT NVL(Pedido_Ecommerce,0)"
        QueryC541693.CommandText = QueryC541693.CommandText & " " & "FROM WF_PEDIDO"
        QueryC541693.CommandText = QueryC541693.CommandText & " " & "WHERE Tp_Produto = " & _ 
ConvertValue(P56552, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC541693.CommandText = QueryC541693.CommandText & " " & "AND Pedido = " & _ 
ConvertValue(P56550, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC541693.Transaction = txn
        RsC541693 = QueryC541693.ExecuteReader()
        Dim iC541693 As Short
        ReDim C541693DataType(RsC541693.FieldCount)
        For iC541693 = 0 to RsC541693.FieldCount - 1
            Select Case RsC541693.GetDataTypeName(iC541693).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C541693DataType(iC541693 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C541693DataType(iC541693 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C541693DataType(iC541693 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC541693
        RsC541693_EOF = Not RsC541693.Read()

        GoTo Comp_C541694

    Comp_C541694:

        ' FlagEcommerce=0?
        sCurrComponent = "FlagEcommerce=0?"
        C541694 = (fn_ConvertValueCompiled(RsC541693(0), C541693DataType(1), AKB_DecimalPoint, False) = 0)
        C541694DataType = AKBTypeConst.cAKBTypeLogical
        If C541694 Then
            GoTo Comp_C398638
        Else
            GoTo Comp_C541696
        End If

    Comp_C541696:

        ' Sysdate
        sCurrComponent = "Sysdate"
        C541696DataType = 2
        C541696 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C541697

    Comp_C541697:

        ' IdCol
        sCurrComponent = "IdCol"
        QueryC541697 = con.CreateCommand()
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "SELECT Max(WF_COLECAO.Id_Colecao )"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "FROM WF_COLECAO_PRODUTOS, WF_COLECAO, WF_PEDIDO_ITENS"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Id_Colecao = WF_COLECAO.Id_Colecao"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PEDIDO_ITENS.SIGLA_PROD"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PEDIDO_ITENS.ACESSO"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PEDIDO_ITENS.COD_EMBALAGEM"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "AND WF_COLECAO.DATA_VALIDADE_INICIAL = (SELECT MAX(WF_COLECAO.DATA_VALIDADE_INICIAL)"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "                                        FROM WF_COLECAO_PRODUTOS, WF_COLECAO"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "                                        WHERE WF_COLECAO_PRODUTOS.Id_Colecao = WF_COLECAO.Id_Colecao"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "                                        AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PEDIDO_ITENS.SIGLA_PROD "
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "                                        AND WF_COLECAO_PRODUTOS.Acesso = WF_PEDIDO_ITENS.ACESSO"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "                                        AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PEDIDO_ITENS.COD_EMBALAGEM"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "                                        AND WF_COLECAO.DATA_VALIDADE_INICIAL <= " & _ 
ConvertValue(C541696, C541696DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "AND WF_PEDIDO_ITENS.TP_PRODUTO = " & _ 
ConvertValue(P56552, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "AND WF_PEDIDO_ITENS.SEQ_ITEM = " & _ 
ConvertValue(P56553, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC541697.CommandText = QueryC541697.CommandText & " " & "AND WF_PEDIDO_ITENS.PEDIDO = " & _ 
ConvertValue(P56550, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC541697.Transaction = txn
        RsC541697 = QueryC541697.ExecuteReader()
        Dim iC541697 As Short
        ReDim C541697DataType(RsC541697.FieldCount)
        For iC541697 = 0 to RsC541697.FieldCount - 1
            Select Case RsC541697.GetDataTypeName(iC541697).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C541697DataType(iC541697 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C541697DataType(iC541697 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C541697DataType(iC541697 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC541697
        RsC541697_EOF = Not RsC541697.Read()

        GoTo Comp_C541698

    Comp_C541698:

        ' UpColEcommerce
        sCurrComponent = "UpColEcommerce"
        QueryC541698 = con.CreateCommand()
        QueryC541698.CommandText = QueryC541698.CommandText & " " & "UPDATE WF_PEDIDO_ITENS"
        QueryC541698.CommandText = QueryC541698.CommandText & " " & "SET WF_PEDIDO_ITENS.ID_COLECAO = " & _ 
ConvertValue(RsC541697(0), C541697DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC541698.CommandText = QueryC541698.CommandText & " " & "WHERE WF_PEDIDO_ITENS.TP_PRODUTO = " & _ 
ConvertValue(P56552, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC541698.CommandText = QueryC541698.CommandText & " " & "AND WF_PEDIDO_ITENS.SEQ_ITEM = " & _ 
ConvertValue(P56553, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC541698.CommandText = QueryC541698.CommandText & " " & "AND WF_PEDIDO_ITENS.PEDIDO = " & _ 
ConvertValue(P56550, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC541698.Transaction = txn
        C541698 = QueryC541698.ExecuteNonQuery()
        C541698DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C340894

    Exit_R15919:

        Exit Function

    Err_R15919:
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
