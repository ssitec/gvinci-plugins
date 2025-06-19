Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R9614

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

    Public Shared Function R9614(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P30484 As String, ByVal P30485 As Double, ByVal P30487 As Double, ByVal P64892 As Object) As DataTable
    ' 
    ' 9614 - Retorna tipo de Pedido do Acesso/Embalagem - Versão: 0
    ' 
        'On Error GoTo Err_R9614
        Dim sCurrComponent as String

        Dim QueryC164646 As New Object
        Dim RsC164646 As New Object
        Dim C164646DataType() As Short
        Dim RsC164646_EOF As Boolean
        Dim C164651DataType() As Short
        Dim C164655 As Object
        Dim C164655DataType As Short
        Dim C164656 As DataTable
        Dim C164656CurrentRow As Integer
        Dim C164656DataType() As Short
        Dim C388296 As Object
        Dim C388296DataType As Short
        Dim C388297 As Boolean
        Dim C388297DataType As Short
        Dim C388298 As Object
        Dim C388298DataType As Short
        Dim C388299 As Object
        Dim C388299DataType As Short
        Dim C388300 As Boolean
        Dim C388300DataType As Short
        Dim C394786 As Object
        Dim C394786DataType As Short
        Dim QueryC398047 As New Object
        Dim RsC398047 As New Object
        Dim C398047DataType() As Short
        Dim RsC398047_EOF As Boolean
        Dim C398048 As Object
        Dim C398048DataType As Short
        Dim C398049 As Boolean
        Dim C398049DataType As Short
        Dim C398050 As Object
        Dim C398050DataType As Short
        Dim QueryC398051 As New Object
        Dim RsC398051 As New Object
        Dim C398051DataType() As Short
        Dim RsC398051_EOF As Boolean
        Dim C398052 As Boolean
        Dim C398052DataType As Short
        Dim C398053 As Object
        Dim C398053DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C164655

    Comp_C164646:

        ' Sel_Tp_Pedido
        sCurrComponent = "Sel_Tp_Pedido"
        QueryC164646 = con.CreateCommand()
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "SELECT WF_COLECAO.Tipo_Ped,"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "WF_TIPO_PED.Desc_Tipo_Ped, 1 SEQ"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS, AKBUSER01.WF_COLECAO , AKBUSER01.WF_TIPO_PED"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Id_Colecao  = WF_COLECAO.Id_Colecao"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND " & _ 
ConvertValue(C394786, C394786DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " >= WF_COLECAO_PRODUTOS.Data_Validade_Inicial"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND (WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL  "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "    OR " & _ 
ConvertValue(C394786, C394786DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " < WF_COLECAO_PRODUTOS.Data_Validade_Final + 1)"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Sigla_Prod = " & _ 
ConvertValue(P30484, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = " & _ 
ConvertValue(P30485, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem =  " & _ 
ConvertValue(C398050, C398050DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_COLECAO.Tipo_Ped = WF_TIPO_PED.Tipo_Ped"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_COLECAO.Tipo_Ped IS NOT NULL"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & ""
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "UNION ALL"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & ""
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "SELECT WF_COLECAO.Tipo_Ped,"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "  WF_TIPO_PED.Desc_Tipo_Ped, 2 SEQ"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "FROM AKBUSER01.WF_TABELA_PRECO_VENDA,"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "  AKBUSER01.WF_COLECAO,"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "  AKBUSER01.WF_TIPO_PED"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "WHERE WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO.Id_Colecao"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_COLECAO.Tipo_Ped = WF_TIPO_PED.Tipo_Ped"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(P64892, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & ""
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "UNION ALL"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & ""
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "SELECT WF_LOGIN_CONFIG_PEDIDO.Tipo_Ped , WF_TIPO_PED.Desc_Tipo_Ped , 3 SEQ"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "FROM AKBUSER01.WF_LOGIN_CONFIG_PEDIDO , AKBUSER01.WF_TIPO_PED "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "WHERE WF_LOGIN_CONFIG_PEDIDO.Login = " & _ 
ConvertValue(C164655, C164655DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_TIPO_PED.Tipo_Ped = WF_LOGIN_CONFIG_PEDIDO.Tipo_Ped "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND NOT EXISTS ( SELECT WF_EMB_COMP_VDA_PROD.Tipo_Ped , "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "                 WF_TIPO_PED.Desc_Tipo_Ped "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "                 FROM AKBUSER01.WF_EMB_COMP_VDA_PROD , AKBUSER01.WF_TIPO_PED "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "                 WHERE WF_EMB_COMP_VDA_PROD.Sigla_Prod = " & _ 
ConvertValue(P30484, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "                 AND WF_EMB_COMP_VDA_PROD.Acesso = " & _ 
ConvertValue(P30485, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "                 AND WF_EMB_COMP_VDA_PROD.Cod_Embalagem = " & _ 
ConvertValue(C398050, C398050DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "                 AND WF_TIPO_PED.Tipo_Ped = WF_EMB_COMP_VDA_PROD.Tipo_Ped "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "                 AND WF_EMB_COMP_VDA_PROD.Tipo_Ped IS NOT NULL "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "                 )"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "						  "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "UNION ALL"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & ""
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "SELECT WF_EMB_COMP_VDA_PROD.Tipo_Ped , "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "WF_TIPO_PED.Desc_Tipo_Ped , 4 SEQ"
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "FROM AKBUSER01.WF_EMB_COMP_VDA_PROD , AKBUSER01.WF_TIPO_PED "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "WHERE WF_EMB_COMP_VDA_PROD.Sigla_Prod = " & _ 
ConvertValue(P30484, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_EMB_COMP_VDA_PROD.Acesso = " & _ 
ConvertValue(P30485, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_EMB_COMP_VDA_PROD.Cod_Embalagem = " & _ 
ConvertValue(C398050, C398050DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "AND WF_TIPO_PED.Tipo_Ped = WF_EMB_COMP_VDA_PROD.Tipo_Ped "
        QueryC164646.CommandText = QueryC164646.CommandText & " " & ""
        QueryC164646.CommandText = QueryC164646.CommandText & " " & "ORDER BY SEQ"
        QueryC164646.Transaction = txn
        RsC164646 = QueryC164646.ExecuteReader()
        Dim iC164646 As Short
        ReDim C164646DataType(RsC164646.FieldCount)
        For iC164646 = 0 to RsC164646.FieldCount - 1
            Select Case RsC164646.GetDataTypeName(iC164646).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C164646DataType(iC164646 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C164646DataType(iC164646 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C164646DataType(iC164646 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC164646
        RsC164646_EOF = Not RsC164646.Read()

        GoTo Comp_C388296

    Comp_C164651:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C164651 As DataTable = New DataTable()
        tb_C164651.Columns.Add("Tipo Pedido" & "")
        tb_C164651.Columns.Add("Descrição Tipo Pedido" & "")
        ReDim ReturnDatatype(RsC164646.FieldCount)
        Do Until RsC164646_EOF
            Dim row_C164651 As DataRow = tb_C164651.NewRow()
            Dim iC164651 As Short
            For iC164651 = 0 To 1
                row_C164651(iC164651) = RsC164646(iC164651)
                ReturnDatatype(iC164651 + 1) = C164646DataType(iC164651)
            Next iC164651
            tb_C164651.Rows.Add(row_C164651)
            RsC164646_EOF = Not RsC164646.Read()
        Loop
        R9614 = tb_C164651

        GoTo Exit_R9614

    Comp_C164655:

        ' Usuário
        sCurrComponent = "Usuário"
        C164655DataType = 3
        C164655 = Mid(con.ConnectionString, InStr(con.ConnectionString, "User Id=")+8, Instr(InStr(con.ConnectionString, "User Id="), con.ConnectionString, ";")-InStr(con.ConnectionString, "User Id=")-8)
        GoTo Comp_C394786

    Comp_C164656:

        ' Tipo_Ped_Emb
        sCurrComponent = "Tipo_Ped_Emb"
        C164656 = clsRuleDynamicallyCompiled_R5577.R5577(con, txn, IIf(Not IsDBNull(C388298), C388298, System.DBNull.Value))
        C164656CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C164656) Then
          iColumns = C164656.Columns.Count
        End If
        ReDim C164656DataType(iColumns)
        For iCol = 1 To iColumns
            C164656DataType(iCol) = clsRuleDynamicallyCompiled_R5577.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C164651

    Comp_C388296:

        ' EOF_Tp_Pedido
        sCurrComponent = "EOF_Tp_Pedido"
        C388296DataType = 4
        C388296 = RsC164646_EOF
        GoTo Comp_C388297

    Comp_C388297:

        ' Not_EOF_Tp_Pedido?
        sCurrComponent = "Not_EOF_Tp_Pedido?"
        C388297 = (fn_ConvertValueCompiled(C388296, C388296DataType, AKB_DecimalPoint, False) = 0)
        C388297DataType = AKBTypeConst.cAKBTypeLogical
        If C388297 Then
            GoTo Comp_C388298
        Else
            GoTo Comp_C164651
        End If

    Comp_C388298:

        ' vTipo_Pedido
        sCurrComponent = "vTipo_Pedido"
        C388298DataType = 0
        C388298 = RsC164646(0)
        C388298DataType = C164646Datatype(1)
        If C388298DataType = AKBTypeConst.cAKBTypeString Then
          C388298 = IIF(IsDBNull(C388298), C388298, Strings.RTrim(C388298))
        End If 

        GoTo Comp_C388299

    Comp_C388299:

        ' Nulo_Tipo_Pedido
        sCurrComponent = "Nulo_Tipo_Pedido"
        C388299DataType = 4
        C388299 = IsDBNull (C388298)
        GoTo Comp_C388300

    Comp_C388300:

        ' Tem_Tipo_Pedido?
        sCurrComponent = "Tem_Tipo_Pedido?"
        C388300 = (fn_ConvertValueCompiled(C388299, C388299DataType, AKB_DecimalPoint, False) = 0)
        C388300DataType = AKBTypeConst.cAKBTypeLogical
        If C388300 Then
            GoTo Comp_C164656
        Else
            GoTo Comp_C164651
        End If

    Comp_C394786:

        ' Data_Sistema
        sCurrComponent = "Data_Sistema"
        C394786DataType = 2
        C394786 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C398050

    Comp_C398047:

        ' SelEmbInterna
        sCurrComponent = "SelEmbInterna"
        QueryC398047 = con.CreateCommand()
        QueryC398047.CommandText = QueryC398047.CommandText & " " & "SELECT WF_Emb_Comp_Vda_Prod.Cod_Embalagem_Int_Emb"
        QueryC398047.CommandText = QueryC398047.CommandText & " " & "FROM WF_Emb_Comp_Vda_Prod"
        QueryC398047.CommandText = QueryC398047.CommandText & " " & "WHERE WF_Emb_Comp_Vda_Prod.Sigla_Prod = " & _ 
ConvertValue(P30484, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398047.CommandText = QueryC398047.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Acesso =  " & _ 
ConvertValue(P30485, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398047.CommandText = QueryC398047.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Cod_Embalagem = " & _ 
ConvertValue(P30487, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398047.CommandText = QueryC398047.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Cod_Embalagem_Int_Emb IS NOT NULL"
        QueryC398047.Transaction = txn
        RsC398047 = QueryC398047.ExecuteReader()
        Dim iC398047 As Short
        ReDim C398047DataType(RsC398047.FieldCount)
        For iC398047 = 0 to RsC398047.FieldCount - 1
            Select Case RsC398047.GetDataTypeName(iC398047).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C398047DataType(iC398047 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C398047DataType(iC398047 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C398047DataType(iC398047 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC398047
        RsC398047_EOF = Not RsC398047.Read()

        GoTo Comp_C398048

    Comp_C398048:

        ' Fim_Emb_Interna
        sCurrComponent = "Fim_Emb_Interna"
        C398048DataType = 4
        C398048 = RsC398047_EOF
        GoTo Comp_C398049

    Comp_C398049:

        ' Fim_Emb_Interna = 1?
        sCurrComponent = "Fim_Emb_Interna = 1?"
        C398049 = (fn_ConvertValueCompiled(C398048, C398048DataType, AKB_DecimalPoint, False) = 1)
        C398049DataType = AKBTypeConst.cAKBTypeLogical
        If C398049 Then
            GoTo Comp_C164646
        Else
            GoTo Comp_C398051
        End If

    Comp_C398050:

        ' vCod_Emb
        sCurrComponent = "vCod_Emb"
        C398050 = P30487 & " "
        C398050DataType = 1
        GoTo Comp_C398047

    Comp_C398051:

        ' VerificaSeProdutoEmbTemColeção
        sCurrComponent = "VerificaSeProdutoEmbTemColeção"
        QueryC398051 = con.CreateCommand()
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "SELECT COUNT(*)"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "FROM WF_Emb_Comp_Vda_Prod,"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "  WF_Colecao_Produtos,"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "  WF_Colecao"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "WHERE WF_Emb_Comp_Vda_Prod.Sigla_Prod = " & _ 
ConvertValue(P30484, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Acesso = " & _ 
ConvertValue(P30485, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Cod_Embalagem = " & _ 
ConvertValue(P30487, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Sigla_Prod = WF_Colecao_Produtos.Sigla_Prod"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Acesso = WF_Colecao_Produtos.Acesso"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "AND WF_Emb_Comp_Vda_Prod.Cod_Embalagem = WF_Colecao_Produtos.Cod_Embalagem"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "AND WF_Colecao.Id_Colecao = WF_Colecao_Produtos.Id_Colecao"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "AND (WF_Colecao.Data_Validade_Final IS NULL"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "OR WF_Colecao.Data_Validade_Final >= Sysdate)"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "AND (WF_Colecao_Produtos.Data_Validade_Final IS NULL"
        QueryC398051.CommandText = QueryC398051.CommandText & " " & "OR WF_Colecao_Produtos.Data_Validade_Final >= Sysdate)"
        QueryC398051.Transaction = txn
        RsC398051 = QueryC398051.ExecuteReader()
        Dim iC398051 As Short
        ReDim C398051DataType(RsC398051.FieldCount)
        For iC398051 = 0 to RsC398051.FieldCount - 1
            Select Case RsC398051.GetDataTypeName(iC398051).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C398051DataType(iC398051 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C398051DataType(iC398051 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C398051DataType(iC398051 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC398051
        RsC398051_EOF = Not RsC398051.Read()

        GoTo Comp_C398052

    Comp_C398052:

        ' VerificaSeProdutoEmbTemColeção = 0?
        sCurrComponent = "VerificaSeProdutoEmbTemColeção = 0?"
        C398052 = (fn_ConvertValueCompiled(RsC398051(0), C398051DataType(1), AKB_DecimalPoint, False) = 0)
        C398052DataType = AKBTypeConst.cAKBTypeLogical
        If C398052 Then
            GoTo Comp_C398053
        Else
            GoTo Comp_C164646
        End If

    Comp_C398053:

        ' AtrEmbInterna
        sCurrComponent = "AtrEmbInterna"
        C398053DataType = 4
        C398050 = fn_ConvertValueCompiled(RsC398047(0) , 1, AKB_DecimalPoint)
        C398053 = True
        GoTo Comp_C164646

    Exit_R9614:

        Exit Function

    Err_R9614:
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
