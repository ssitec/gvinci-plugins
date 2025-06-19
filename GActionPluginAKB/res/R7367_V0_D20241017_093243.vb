Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R7367

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

    Public Shared Function R7367(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P22313 As Double, ByVal P22314 As Double) As DataTable
    ' 
    ' 7367 - Verifica Embalagem Inativa no Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R7367
        Dim sCurrComponent as String

        Dim QueryC108826 As New Object
        Dim RsC108826 As New Object
        Dim C108826DataType() As Short
        Dim RsC108826_EOF As Boolean
        Dim C108827 As Boolean
        Dim C108827DataType As Short
        Dim C108828 As Object
        Dim C108828DataType As Short
        Dim C108829DataType() As Short
        Dim C108830 As Short
        Dim C108830DataType As Short
        Dim C108830ReturnDataType() As Short

        Dim C108832 As Object
        Dim C108832DataType As Short
        Dim C108834 As Object
        Dim C108834DataType As Short
        Dim C108835 As Object
        Dim C108835DataType As Short
        Dim C108836 As Object
        Dim C108836DataType As Short
        Dim C108837 As Object
        Dim C108837DataType As Short
        Dim C108838 As Object
        Dim C108838DataType As Short
        Dim C108839 As Object
        Dim C108839DataType As Short
        Dim C108840 As Object
        Dim C108840DataType As Short
        Dim C108841 As Object
        Dim C108841DataType As Short
        Dim C108842 As Boolean
        Dim C108842DataType As Short
        Dim C108843 As Short
        Dim C108843DataType As Short
        Dim C108843ReturnDataType() As Short

        Dim C108844 As Object
        Dim C108844DataType As Short
        Dim QueryC388245 As New Object
        Dim RsC388245 As New Object
        Dim C388245DataType() As Short
        Dim RsC388245_EOF As Boolean
        Dim C388246 As Boolean
        Dim C388246DataType As Short
        Dim C388372 As Object
        Dim C388372DataType As Short
        Dim C388373 As Boolean
        Dim C388373DataType As Short
        Dim C426198 As Object
        Dim C426198DataType As Short
        Dim C510716 As Boolean
        Dim C510716DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C108828

    Comp_C108826:

        ' SelItens
        sCurrComponent = "SelItens"
        QueryC108826 = con.CreateCommand()
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "SELECT WF_PEDIDO_ITENS.Sigla_Prod , WF_PEDIDO_ITENS.Acesso , WF_PEDIDO_ITENS.Cod_Embalagem ,WF_TABELA_EMBALAGEM.Descr_Embalagem, WF_PEDIDO.Cod_Pessoa_Estab_Juridico"
        QueryC108826.CommandText = QueryC108826.CommandText & " " & ""
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS , AKBUSER01.WF_TABELA_EMBALAGEM , AKBUSER01.WF_EMB_COMP_VDA_PROD , AKBUSER01.WF_PEDIDO"
        QueryC108826.CommandText = QueryC108826.CommandText & " " & ""
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P22313, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P22314, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_PEDIDO_ITENS.Cod_Embalagem = WF_TABELA_EMBALAGEM.Cod_Embalagem "
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_PEDIDO_ITENS.Sigla_Prod = WF_EMB_COMP_VDA_PROD.Sigla_Prod "
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_PEDIDO_ITENS.Acesso = WF_EMB_COMP_VDA_PROD.Acesso "
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_PEDIDO_ITENS.Cod_Embalagem = WF_EMB_COMP_VDA_PROD.Cod_Embalagem "
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_EMB_COMP_VDA_PROD.Inativo = 1"
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_PEDIDO_ITENS.Flag_Pos_Ped not in ( 7, 8, 10 )"
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido"
        QueryC108826.CommandText = QueryC108826.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto"
        QueryC108826.Transaction = txn
        RsC108826 = QueryC108826.ExecuteReader()
        Dim iC108826 As Short
        ReDim C108826DataType(RsC108826.FieldCount)
        For iC108826 = 0 to RsC108826.FieldCount - 1
            Select Case RsC108826.GetDataTypeName(iC108826).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C108826DataType(iC108826 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C108826DataType(iC108826 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C108826DataType(iC108826 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC108826
        RsC108826_EOF = Not RsC108826.Read()

        GoTo Comp_C108840

    Comp_C108827:

        ' SelItens=0
        sCurrComponent = "SelItens=0"
        C108827 = (fn_ConvertValueCompiled(C108841, C108841DataType, AKB_DecimalPoint, False) = 1)
        C108827DataType = AKBTypeConst.cAKBTypeLogical
        If C108827 Then
            GoTo Comp_C108842
        Else
            GoTo Comp_C108834
        End If

    Comp_C108828:

        ' VTrue
        sCurrComponent = "VTrue"
        C108828 = 1
        C108828DataType = 4
        GoTo Comp_C510716

    Comp_C108829:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C108829 As DataTable = New DataTable()
        tb_C108829.Columns.Add("VTrue" & "")
        Dim row_C108829 As DataRow = tb_C108829.NewRow()
        row_C108829(0) = C108828
        tb_C108829.Rows.Add(row_C108829)
        R7367 = tb_C108829
        ReDim C108829DataType(1)
        C108829DataType(1) = C108828DataType
        ReturnDataType = C108829DataType

        GoTo Exit_R7367

    Comp_C108830:

        ' MsgConfirma
        sCurrComponent = "MsgConfirma"
        C108830 = 1
        C108830DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C108832

    Comp_C108832:

        ' RetFalse
        sCurrComponent = "RetFalse"
        C108832DataType = 4
        C108828 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C108832 = True
        GoTo Comp_C108839

    Comp_C108834:

        ' Sigla_Prod
        sCurrComponent = "Sigla_Prod"
        C108834DataType = 0
        C108834 = RsC108826(0)
        C108834DataType = C108826Datatype(1)
        If C108834DataType = AKBTypeConst.cAKBTypeString Then
          C108834 = IIF(IsDBNull(C108834), C108834, Strings.RTrim(C108834))
        End If 

        GoTo Comp_C108835

    Comp_C108835:

        ' Acesso
        sCurrComponent = "Acesso"
        C108835DataType = 0
        C108835DataType = C108826Datatype(2)
        If C108835DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC108826(1)) Then
          C108835 = Strings.RTrim(RsC108826(1))
        Else 
          C108835 = RsC108826(1)
        End If 

        GoTo Comp_C108836

    Comp_C108836:

        ' Cod_Embalagem
        sCurrComponent = "Cod_Embalagem"
        C108836DataType = 0
        C108836DataType = C108826Datatype(3)
        If C108836DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC108826(2)) Then
          C108836 = Strings.RTrim(RsC108826(2))
        Else 
          C108836 = RsC108826(2)
        End If 

        GoTo Comp_C108838

    Comp_C108837:

        ' Flag
        sCurrComponent = "Flag"
        C108837 = 0
        C108837DataType = 1
        GoTo Comp_C108826

    Comp_C108838:

        ' DescEmb
        sCurrComponent = "DescEmb"
        C108838DataType = 0
        C108838DataType = C108826Datatype(4)
        If C108838DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC108826(3)) Then
          C108838 = Strings.RTrim(RsC108826(3))
        Else 
          C108838 = RsC108826(3)
        End If 

        GoTo Comp_C426198

    Comp_C108839:

        ' attFlag
        sCurrComponent = "attFlag"
        C108839DataType = 4
        C108837 = fn_ConvertValueCompiled(1, 1, AKB_DecimalPoint)
        C108839 = True
        GoTo Comp_C108844

    Comp_C108840:

        ' PRIMEIROV1
        sCurrComponent = "PRIMEIROV1"
        C108840DataType = 4

        GoTo Comp_C108841

    Comp_C108841:

        ' FIMVALORE1
        sCurrComponent = "FIMVALORE1"
        C108841DataType = 4
        C108841 = RsC108826_EOF
        GoTo Comp_C108827

    Comp_C108842:

        ' EmbInativa?
        sCurrComponent = "EmbInativa?"
        C108842 = (fn_ConvertValueCompiled(C108837, C108837DataType, AKB_DecimalPoint, False) = 0)
        C108842DataType = AKBTypeConst.cAKBTypeLogical
        If C108842 Then
            GoTo Comp_C108829
        Else
            GoTo Comp_C108843
        End If

    Comp_C108843:

        ' MsgEmbInativa
        sCurrComponent = "MsgEmbInativa"
        C108843 = 1
        C108843DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C108829

    Comp_C108844:

        ' PRÓXIMOVA1
        sCurrComponent = "PRÓXIMOVA1"
        C108844DataType = 4
        RsC108826_EOF = Not RsC108826.Read()
        C108844 = True

        GoTo Comp_C108841

    Comp_C388245:

        ' SelEstoqueDisponivel
        sCurrComponent = "SelEstoqueDisponivel"
        QueryC388245 = con.CreateCommand()
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "SELECT QTDE"
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "FROM (SELECT WF_ACESSO_ESTOQUE.Qtd_Emb_Mov_Interna_1 - WF_ACESSO_ESTOQUE.Qtde_Utilizada - NVL(SUM(WF_ACESSO_ESTOQUE_OP.Qtd_Emb_Mov_Interna_1), 0) QTDE"
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "FROM AKBUSER01.WF_ACESSO_ESTOQUE , AKBUSER01.WF_ACESSO_ESTOQUE_OP "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "WHERE  WF_ACESSO_ESTOQUE.Sigla_Prod = " & _ 
ConvertValue(C108834, C108834DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "AND  WF_ACESSO_ESTOQUE.Acesso = " & _ 
ConvertValue(C108835, C108835DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "AND WF_ACESSO_ESTOQUE.Cod_Embalagem = " & _ 
ConvertValue(C108836, C108836DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "AND WF_ACESSO_ESTOQUE.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(C426198, C426198DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "AND  (WF_ACESSO_ESTOQUE.Fase_Acesso_Estoque = 1 OR WF_ACESSO_ESTOQUE.Fase_Acesso_Estoque = 14)"
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "AND  WF_ACESSO_ESTOQUE.Sigla_Prod_Estoque = WF_ACESSO_ESTOQUE_OP.Sigla_Prod_Estoque (+)"
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "AND  WF_ACESSO_ESTOQUE.Cod_Ac_Estoque = WF_ACESSO_ESTOQUE_OP.Cod_Ac_Estoque (+)"
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "AND  WF_ACESSO_ESTOQUE.Seq_Div_Estoque = WF_ACESSO_ESTOQUE_OP.Seq_Div_Estoque (+)"
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "AND  WF_ACESSO_ESTOQUE_OP.Confirmado (+) = 0 "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "GROUP BY WF_ACESSO_ESTOQUE.Cod_Ac_Estoque , "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "WF_ACESSO_ESTOQUE.Sigla_Prod_Estoque , "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "WF_ACESSO_ESTOQUE.Seq_Div_Estoque , "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "WF_ACESSO_ESTOQUE.Qtd_Emb_Mov_Interna_1 , "
        QueryC388245.CommandText = QueryC388245.CommandText & " " & "WF_ACESSO_ESTOQUE.Qtde_Utilizada) "
        QueryC388245.Transaction = txn
        RsC388245 = QueryC388245.ExecuteReader()
        Dim iC388245 As Short
        ReDim C388245DataType(RsC388245.FieldCount)
        For iC388245 = 0 to RsC388245.FieldCount - 1
            Select Case RsC388245.GetDataTypeName(iC388245).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C388245DataType(iC388245 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C388245DataType(iC388245 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C388245DataType(iC388245 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC388245
        RsC388245_EOF = Not RsC388245.Read()

        GoTo Comp_C388372

    Comp_C388246:

        ' EstoqueDisponivel?
        sCurrComponent = "EstoqueDisponivel?"
        C388246 = (fn_ConvertValueCompiled(RsC388245(0), C388245DataType(1), AKB_DecimalPoint, False) > 0)
        C388246DataType = AKBTypeConst.cAKBTypeLogical
        If C388246 Then
            GoTo Comp_C108844
        Else
            GoTo Comp_C108830
        End If

    Comp_C388372:

        ' SemEstoqueDisponivel
        sCurrComponent = "SemEstoqueDisponivel"
        C388372DataType = 4
        C388372 = RsC388245_EOF
        GoTo Comp_C388373

    Comp_C388373:

        ' SemEstoqueDisponivel?
        sCurrComponent = "SemEstoqueDisponivel?"
        C388373 = (fn_ConvertValueCompiled(C388372, C388372DataType, AKB_DecimalPoint, False) = 1)
        C388373DataType = AKBTypeConst.cAKBTypeLogical
        If C388373 Then
            GoTo Comp_C108830
        Else
            GoTo Comp_C388246
        End If

    Comp_C426198:

        ' EstabJuridico
        sCurrComponent = "EstabJuridico"
        C426198DataType = 0
        C426198DataType = C108826Datatype(5)
        If C426198DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC108826(4)) Then
          C426198 = Strings.RTrim(RsC108826(4))
        Else 
          C426198 = RsC108826(4)
        End If 

        GoTo Comp_C388245

    Comp_C510716:

        ' DESVIO1
        sCurrComponent = "DESVIO1"
        C510716 = (1 = 1)
        C510716DataType = AKBTypeConst.cAKBTypeLogical
        If C510716 Then
            GoTo Comp_C108829
        Else
            GoTo Comp_C108837
        End If

    Exit_R7367:

        Exit Function

    Err_R7367:
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
