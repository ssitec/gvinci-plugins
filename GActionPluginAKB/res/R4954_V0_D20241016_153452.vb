Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R4954

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

    Public Shared Function R4954(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P14088 As Double, ByVal P14091 As Double) As DataTable
    ' 
    ' 4954 - Checa T. de Produto X T. de Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R4954
        Dim sCurrComponent as String

        Dim QueryC63012 As New Object
        Dim RsC63012 As New Object
        Dim C63012DataType() As Short
        Dim RsC63012_EOF As Boolean
        Dim C63016DataType() As Short
        Dim QueryC63017 As New Object
        Dim RsC63017 As New Object
        Dim C63017DataType() As Short
        Dim RsC63017_EOF As Boolean
        Dim C63020 As Boolean
        Dim C63020DataType As Short
        Dim C63027 As Object
        Dim C63027DataType As Short
        Dim C63028 As Short
        Dim C63028DataType As Short
        Dim C63028ReturnDataType() As Short

        Dim C63031 As Object
        Dim C63031DataType As Short
        Dim C63032 As Object
        Dim C63032DataType As Short
        Dim C63043 As Boolean
        Dim C63043DataType As Short
        Dim C63044 As Object
        Dim C63044DataType As Short
        Dim C63124 As Boolean
        Dim C63124DataType As Short
        Dim C63126 As Object
        Dim C63126DataType As Short
        Dim QueryC86538 As New Object
        Dim RsC86538 As New Object
        Dim C86538DataType() As Short
        Dim RsC86538_EOF As Boolean
        Dim C86543 As Boolean
        Dim C86543DataType As Short
        Dim C86545 As Short
        Dim C86545DataType As Short
        Dim C86545ReturnDataType() As Short

        Dim C320275 As Short
        Dim C320275DataType As Short
        Dim C320275ReturnDataType() As Short

        Dim QueryC320276 As New Object
        Dim RsC320276 As New Object
        Dim C320276DataType() As Short
        Dim RsC320276_EOF As Boolean
        Dim C320277 As Boolean
        Dim C320277DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C63031

    Comp_C63012:

        ' sel
        sCurrComponent = "sel"
        QueryC63012 = con.CreateCommand()
        QueryC63012.CommandText = QueryC63012.CommandText & " " & "SELECT COUNT(1)  "
        QueryC63012.CommandText = QueryC63012.CommandText & " " & ""
        QueryC63012.CommandText = QueryC63012.CommandText & " " & "FROM AKBUSER01.WF_PRODUTO_PEDIDO "
        QueryC63012.CommandText = QueryC63012.CommandText & " " & ""
        QueryC63012.CommandText = QueryC63012.CommandText & " " & "WHERE"
        QueryC63012.CommandText = QueryC63012.CommandText & " " & "WF_PRODUTO_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P14088, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC63012.CommandText = QueryC63012.CommandText & " " & ""
        QueryC63012.Transaction = txn
        RsC63012 = QueryC63012.ExecuteReader()
        Dim iC63012 As Short
        ReDim C63012DataType(RsC63012.FieldCount)
        For iC63012 = 0 to RsC63012.FieldCount - 1
            Select Case RsC63012.GetDataTypeName(iC63012).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C63012DataType(iC63012 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C63012DataType(iC63012 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C63012DataType(iC63012 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC63012
        RsC63012_EOF = Not RsC63012.Read()

        GoTo Comp_C63020

    Comp_C63016:

        ' fim
        sCurrComponent = "fim"
        Dim tb_C63016 As DataTable = New DataTable()
        tb_C63016.Columns.Add("var1" & "")
        Dim row_C63016 As DataRow = tb_C63016.NewRow()
        row_C63016(0) = C63031
        tb_C63016.Rows.Add(row_C63016)
        R4954 = tb_C63016
        ReDim C63016DataType(1)
        C63016DataType(1) = C63031DataType
        ReturnDataType = C63016DataType

        GoTo Exit_R4954

    Comp_C63017:

        ' sel2
        sCurrComponent = "sel2"
        QueryC63017 = con.CreateCommand()
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "SELECT WF_PEDIDO_ITENS.Sigla_Prod , WF_PEDIDO_ITENS.Acesso "
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS"
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P14091, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P14088, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "AND NOT EXISTS "
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "(SELECT *  "
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "FROM AKBUSER01.WF_PRODUTO_PEDIDO "
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "WHERE WF_PRODUTO_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P14088, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC63017.CommandText = QueryC63017.CommandText & " " & "AND WF_PEDIDO_ITENS.Tipo_Ped = WF_PRODUTO_PEDIDO.Tipo_Ped )"
        QueryC63017.CommandText = QueryC63017.CommandText & " " & ""
        QueryC63017.Transaction = txn
        RsC63017 = QueryC63017.ExecuteReader()
        Dim iC63017 As Short
        ReDim C63017DataType(RsC63017.FieldCount)
        For iC63017 = 0 to RsC63017.FieldCount - 1
            Select Case RsC63017.GetDataTypeName(iC63017).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C63017DataType(iC63017 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C63017DataType(iC63017 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C63017DataType(iC63017 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC63017
        RsC63017_EOF = Not RsC63017.Read()

        GoTo Comp_C63126

    Comp_C63020:

        ' sel=0
        sCurrComponent = "sel=0"
        C63020 = (fn_ConvertValueCompiled(RsC63012(0), C63012DataType(1), AKB_DecimalPoint, False) = 0)
        C63020DataType = AKBTypeConst.cAKBTypeLogical
        If C63020 Then
            GoTo Comp_C63016
        Else
            GoTo Comp_C63017
        End If

    Comp_C63027:

        ' sigla
        sCurrComponent = "sigla"
        C63027DataType = 0
        C63027 = RsC63017(0)
        C63027DataType = C63017Datatype(1)
        If C63027DataType = AKBTypeConst.cAKBTypeString Then
          C63027 = IIF(IsDBNull(C63027), C63027, Strings.RTrim(C63027))
        End If 

        GoTo Comp_C63032

    Comp_C63028:

        ' msgPergunta
        sCurrComponent = "msgPergunta"
        C63028 = 6
        C63028DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C63043

    Comp_C63031:

        ' var1
        sCurrComponent = "var1"
        C63031 = 1
        C63031DataType = 4
        GoTo Comp_C320276

    Comp_C63032:

        ' acesso
        sCurrComponent = "acesso"
        C63032DataType = 0
        C63032DataType = C63017Datatype(2)
        If C63032DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC63017(1)) Then
          C63032 = Strings.RTrim(RsC63017(1))
        Else 
          C63032 = RsC63017(1)
        End If 

        GoTo Comp_C86538

    Comp_C63043:

        ' msg1 = S
        sCurrComponent = "msg1 = S"
        C63043 = (fn_ConvertValueCompiled(C63028, C63028DataType, AKB_DecimalPoint, False) = 6)
        C63043DataType = AKBTypeConst.cAKBTypeLogical
        If C63043 Then
            GoTo Comp_C63016
        Else
            GoTo Comp_C63044
        End If

    Comp_C63044:

        ' attvar1
        sCurrComponent = "attvar1"
        C63044DataType = 4
        C63031 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C63044 = True
        GoTo Comp_C63016

    Comp_C63124:

        ' eof2 = 1
        sCurrComponent = "eof2 = 1"
        C63124 = (fn_ConvertValueCompiled(C63126, C63126DataType, AKB_DecimalPoint, False) = 1)
        C63124DataType = AKBTypeConst.cAKBTypeLogical
        If C63124 Then
            GoTo Comp_C63016
        Else
            GoTo Comp_C63027
        End If

    Comp_C63126:

        ' eof2
        sCurrComponent = "eof2"
        C63126DataType = 4
        C63126 = RsC63017_EOF
        GoTo Comp_C63124

    Comp_C86538:

        ' Sel_Barra
        sCurrComponent = "Sel_Barra"
        QueryC86538 = con.CreateCommand()
        QueryC86538.CommandText = QueryC86538.CommandText & " " & "Select NVL(BARRA_NAO_PED_PROD, 0)"
        QueryC86538.CommandText = QueryC86538.CommandText & " " & "from WF_TP_Produto"
        QueryC86538.CommandText = QueryC86538.CommandText & " " & "where TP_Produto = " & _ 
ConvertValue(P14088, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC86538.CommandText = QueryC86538.CommandText & " " & ""
        QueryC86538.Transaction = txn
        RsC86538 = QueryC86538.ExecuteReader()
        Dim iC86538 As Short
        ReDim C86538DataType(RsC86538.FieldCount)
        For iC86538 = 0 to RsC86538.FieldCount - 1
            Select Case RsC86538.GetDataTypeName(iC86538).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C86538DataType(iC86538 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C86538DataType(iC86538 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C86538DataType(iC86538 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC86538
        RsC86538_EOF = Not RsC86538.Read()

        GoTo Comp_C86543

    Comp_C86543:

        ' Sel_Barra=0
        sCurrComponent = "Sel_Barra=0"
        C86543 = (fn_ConvertValueCompiled(RsC86538(0), C86538DataType(1), AKB_DecimalPoint, False) = 0)
        C86543DataType = AKBTypeConst.cAKBTypeLogical
        If C86543 Then
            GoTo Comp_C63028
        Else
            GoTo Comp_C86545
        End If

    Comp_C86545:

        ' MsgBarra
        sCurrComponent = "MsgBarra"
        C86545 = 1
        C86545DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C63044

    Comp_C320275:

        ' msgTipoPedido
        sCurrComponent = "msgTipoPedido"
        C320275 = 1
        C320275DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C63044

    Comp_C320276:

        ' selTipoPedido
        sCurrComponent = "selTipoPedido"
        QueryC320276 = con.CreateCommand()
        QueryC320276.CommandText = QueryC320276.CommandText & " " & "Select count(*)"
        QueryC320276.CommandText = QueryC320276.CommandText & " " & "from AKBUSER01.WF_PEDIDO_ITENS "
        QueryC320276.CommandText = QueryC320276.CommandText & " " & "where WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P14091, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC320276.CommandText = QueryC320276.CommandText & " " & "and WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P14088, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC320276.CommandText = QueryC320276.CommandText & " " & "and  ( WF_PEDIDO_ITENS.Tipo_Ped is null or WF_PEDIDO_ITENS.Tipo_Ped = 0 )"
        QueryC320276.Transaction = txn
        RsC320276 = QueryC320276.ExecuteReader()
        Dim iC320276 As Short
        ReDim C320276DataType(RsC320276.FieldCount)
        For iC320276 = 0 to RsC320276.FieldCount - 1
            Select Case RsC320276.GetDataTypeName(iC320276).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C320276DataType(iC320276 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C320276DataType(iC320276 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C320276DataType(iC320276 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC320276
        RsC320276_EOF = Not RsC320276.Read()

        GoTo Comp_C320277

    Comp_C320277:

        ' Sem Tipo de Pedido
        sCurrComponent = "Sem Tipo de Pedido"
        C320277 = (fn_ConvertValueCompiled(RsC320276(0), C320276DataType(1), AKB_DecimalPoint, False) > 0)
        C320277DataType = AKBTypeConst.cAKBTypeLogical
        If C320277 Then
            GoTo Comp_C320275
        Else
            GoTo Comp_C63012
        End If

    Exit_R4954:

        Exit Function

    Err_R4954:
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
