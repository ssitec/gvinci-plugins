Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R17759

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

    Public Shared Function R17759(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P64541 As Double, ByVal P64542 As Double) As DataTable
    ' 
    ' 17759 - Grava Produto Exclusivo Pre-Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R17759
        Dim sCurrComponent as String

        Dim QueryC394372 As New Object
        Dim RsC394372 As New Object
        Dim C394372DataType() As Short
        Dim RsC394372_EOF As Boolean
        Dim C394373 As Object
        Dim C394373DataType As Short
        Dim C394374 As Object
        Dim C394374DataType As Short
        Dim C394375 As Object
        Dim C394375DataType As Short
        Dim C394376 As Boolean
        Dim C394376DataType As Short
        Dim C394377DataType() As Short
        Dim C394385 As Object
        Dim C394385DataType As Short
        Dim C394386 As Object
        Dim C394386DataType As Short
        Dim C394387 As Object
        Dim C394387DataType As Short
        Dim C394388 As Object
        Dim C394388DataType As Short
        Dim QueryC394389 As New Object
        Dim C394389 As Integer
        Dim C394389DataType As Short
        Dim C406921 As Object
        Dim C406921DataType As Short
        Dim QueryC406922 As New Object
        Dim RsC406922 As New Object
        Dim C406922DataType() As Short
        Dim RsC406922_EOF As Boolean
        Dim C406923 As Object
        Dim C406923DataType As Short
        Dim C406934 As Object
        Dim C406934DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C394373

    Comp_C394372:

        ' SelProd
        sCurrComponent = "SelProd"
        QueryC394372 = con.CreateCommand()
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "SELECT WF_PRE_PEDIDO_ITENS.Sigla_Prod, WF_PRE_PEDIDO_ITENS.Acesso, WF_PRE_PEDIDO.Tp_Produto, WF_PRE_PEDIDO.Tabela_Preco_Venda"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "FROM WF_PRE_PEDIDO_ITENS, WF_PRE_PEDIDO, WF_TP_PRODUTO"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P64541, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "AND NOT EXISTS ( SELECT * FROM WF_PROD_EXC"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "                WHERE WF_PRE_PEDIDO_ITENS.Sigla_Prod = WF_PROD_EXC.Sigla_Prod"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "                AND WF_PRE_PEDIDO_ITENS.Acesso = WF_PROD_EXC.Acesso"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "                AND WF_PROD_EXC.Cod_Cliente = WF_PRE_PEDIDO.Cod_Cliente"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "                AND WF_PROD_EXC.Data_Inicial <= " & _ 
ConvertValue(C394374, C394374DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "                AND (WF_PROD_EXC.Data_Final >= " & _ 
ConvertValue(C394374, C394374DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  OR WF_PROD_EXC.Data_Final IS NULL))"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "AND EXISTS ( SELECT * FROM WF_TAB_PRECO_PRAZO"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "             WHERE WF_TAB_PRECO_PRAZO.Tabela_Preco_Venda = WF_PRE_PEDIDO.Tabela_Preco_Venda"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "             AND WF_TAB_PRECO_PRAZO.Tp_Produto = WF_PRE_PEDIDO.Tp_Produto"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "             AND WF_TAB_PRECO_PRAZO.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "             AND WF_TAB_PRECO_PRAZO.Prod_Exc_Prim_Pedido = 1 )"
        QueryC394372.CommandText = QueryC394372.CommandText & " " & "AND WF_PRE_PEDIDO.Tp_Produto = WF_TP_PRODUTO.Tp_Produto"
        QueryC394372.Transaction = txn
        RsC394372 = QueryC394372.ExecuteReader()
        Dim iC394372 As Short
        ReDim C394372DataType(RsC394372.FieldCount)
        For iC394372 = 0 to RsC394372.FieldCount - 1
            Select Case RsC394372.GetDataTypeName(iC394372).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C394372DataType(iC394372 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C394372DataType(iC394372 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C394372DataType(iC394372 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC394372
        RsC394372_EOF = Not RsC394372.Read()

        GoTo Comp_C394375

    Comp_C394373:

        ' Ret
        sCurrComponent = "Ret"
        C394373 = 1
        C394373DataType = 4
        GoTo Comp_C394374

    Comp_C394374:

        ' DataSys
        sCurrComponent = "DataSys"
        C394374DataType = 2
        C394374 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C394372

    Comp_C394375:

        ' FimSelProd
        sCurrComponent = "FimSelProd"
        C394375DataType = 4
        C394375 = RsC394372_EOF
        GoTo Comp_C394376

    Comp_C394376:

        ' FimSelProd?
        sCurrComponent = "FimSelProd?"
        C394376 = (fn_ConvertValueCompiled(C394375, C394375DataType, AKB_DecimalPoint, False) = 1)
        C394376DataType = AKBTypeConst.cAKBTypeLogical
        If C394376 Then
            GoTo Comp_C394377
        Else
            GoTo Comp_C394386
        End If

    Comp_C394377:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C394377 As DataTable = New DataTable()
        tb_C394377.Columns.Add("Ret" & "")
        Dim row_C394377 As DataRow = tb_C394377.NewRow()
        row_C394377(0) = C394373
        tb_C394377.Rows.Add(row_C394377)
        R17759 = tb_C394377
        ReDim C394377DataType(1)
        C394377DataType(1) = C394373DataType
        ReturnDataType = C394377DataType

        GoTo Exit_R17759

    Comp_C394385:

        ' ProSelProd
        sCurrComponent = "ProSelProd"
        C394385DataType = 4
        RsC394372_EOF = Not RsC394372.Read()
        C394385 = True

        GoTo Comp_C394375

    Comp_C394386:

        ' Sigla
        sCurrComponent = "Sigla"
        C394386DataType = 0
        C394386 = RsC394372(0)
        C394386DataType = C394372Datatype(1)
        If C394386DataType = AKBTypeConst.cAKBTypeString Then
          C394386 = IIF(IsDBNull(C394386), C394386, Strings.RTrim(C394386))
        End If 

        GoTo Comp_C394387

    Comp_C394387:

        ' Acesso
        sCurrComponent = "Acesso"
        C394387DataType = 0
        C394387DataType = C394372Datatype(2)
        If C394387DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394372(1)) Then
          C394387 = Strings.RTrim(RsC394372(1))
        Else 
          C394387 = RsC394372(1)
        End If 

        GoTo Comp_C406921

    Comp_C394388:

        ' Ident
        sCurrComponent = "Ident"
        C394388DataType = 1
        C394388 = ObjTable_NewIdentity (con, "WF_PROD_EXC")
        GoTo Comp_C394389

    Comp_C394389:

        ' INS1
        sCurrComponent = "INS1"
        QueryC394389 = con.CreateCommand()
        QueryC394389.CommandText = QueryC394389.CommandText & " " & "INSERT INTO AKBUSER01.WF_PROD_EXC ( WF_PROD_EXC.Sigla_Prod , WF_PROD_EXC.Acesso , WF_PROD_EXC.Sequencia , WF_PROD_EXC.Cod_Cliente , WF_PROD_EXC.Data_Inicial , WF_PROD_EXC.Data_Final ) VALUES( " & _ 
ConvertValue(C394386, C394386DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C394387, C394387DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C394388, C394388DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P64542, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C394374, C394374DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C406934, C406934DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC394389.Transaction = txn
        C394389 = QueryC394389.ExecuteNonQuery()
        C394389DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C394385

    Comp_C406921:

        ' TpProd
        sCurrComponent = "TpProd"
        C406921DataType = 0
        C406921DataType = C394372Datatype(3)
        If C406921DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394372(2)) Then
          C406921 = Strings.RTrim(RsC394372(2))
        Else 
          C406921 = RsC394372(2)
        End If 

        GoTo Comp_C406923

    Comp_C406922:

        ' MesExcl
        sCurrComponent = "MesExcl"
        QueryC406922 = con.CreateCommand()
        QueryC406922.CommandText = QueryC406922.CommandText & " " & "SELECT NVL(Meses_Validade_Prod_Exclusivo,6)"
        QueryC406922.CommandText = QueryC406922.CommandText & " " & "from WF_TAB_PRECO_PRAZO "
        QueryC406922.CommandText = QueryC406922.CommandText & " " & "WHERE Tabela_Preco_Venda = " & _ 
ConvertValue(C406923, C406923DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC406922.CommandText = QueryC406922.CommandText & " " & " AND Tp_Produto = " & _ 
ConvertValue(C406921, C406921DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC406922.CommandText = QueryC406922.CommandText & " " & " AND Sigla_Prod = " & _ 
ConvertValue(C394386, C394386DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC406922.CommandText = QueryC406922.CommandText & " " & " AND Prod_Exc_Prim_Pedido = 1"
        QueryC406922.Transaction = txn
        RsC406922 = QueryC406922.ExecuteReader()
        Dim iC406922 As Short
        ReDim C406922DataType(RsC406922.FieldCount)
        For iC406922 = 0 to RsC406922.FieldCount - 1
            Select Case RsC406922.GetDataTypeName(iC406922).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C406922DataType(iC406922 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C406922DataType(iC406922 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C406922DataType(iC406922 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC406922
        RsC406922_EOF = Not RsC406922.Read()

        GoTo Comp_C406934

    Comp_C406923:

        ' TabPreco
        sCurrComponent = "TabPreco"
        C406923DataType = 0
        C406923DataType = C394372Datatype(4)
        If C406923DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394372(3)) Then
          C406923 = Strings.RTrim(RsC394372(3))
        Else 
          C406923 = RsC394372(3)
        End If 

        GoTo Comp_C406922

    Comp_C406934:

        ' DtValidFinal
        sCurrComponent = "DtValidFinal"
        C406934DataType = 2
        C406934 = DateAdd("m", RsC406922(0), C394374)
        GoTo Comp_C394388

    Exit_R17759:

        Exit Function

    Err_R17759:
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
