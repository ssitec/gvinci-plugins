Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R17504

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

    Public Shared Function R17504(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P63327 As Double, ByVal P63425 As Double) As DataTable
    ' 
    ' 17504 - Atualiza tipo de pedido parametrizado - Versão: 0
    ' 
        'On Error GoTo Err_R17504
        Dim sCurrComponent as String

        Dim QueryC386029 As New Object
        Dim RsC386029 As New Object
        Dim C386029DataType() As Short
        Dim RsC386029_EOF As Boolean
        Dim C386031 As Object
        Dim C386031DataType As Short
        Dim C386032 As Boolean
        Dim C386032DataType As Short
        Dim C386033DataType() As Short
        Dim QueryC386038 As New Object
        Dim C386038 As Integer
        Dim C386038DataType As Short
        Dim C386039 As Object
        Dim C386039DataType As Short
        Dim C386396 As Object
        Dim C386396DataType As Short
        Dim C386397 As Object
        Dim C386397DataType As Short
        Dim C394991 As Object
        Dim C394991DataType As Short
        Dim C394992 As Object
        Dim C394992DataType As Short
        Dim C394993 As Object
        Dim C394993DataType As Short
        Dim C395618 As DataTable
        Dim C395618CurrentRow As Integer
        Dim C395618DataType() As Short
        Dim C395619 As Object
        Dim C395619DataType As Short
        Dim C395620 As Object
        Dim C395620DataType As Short
        Dim C395621 As Boolean
        Dim C395621DataType As Short
        Dim C395848 As Object
        Dim C395848DataType As Short
        Dim QueryC397000 As New Object
        Dim RsC397000 As New Object
        Dim C397000DataType() As Short
        Dim RsC397000_EOF As Boolean
        Dim QueryC569312 As New Object
        Dim RsC569312 As New Object
        Dim C569312DataType() As Short
        Dim RsC569312_EOF As Boolean
        Dim C569313 As Boolean
        Dim C569313DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C386397

    Comp_C386029:

        ' Sel_Itens_Pedido
        sCurrComponent = "Sel_Itens_Pedido"
        QueryC386029 = con.CreateCommand()
        QueryC386029.CommandText = QueryC386029.CommandText & " " & "SELECT Wf_Pedido_Itens.Sigla_Prod,"
        QueryC386029.CommandText = QueryC386029.CommandText & " " & "  Wf_Pedido_Itens.Acesso,"
        QueryC386029.CommandText = QueryC386029.CommandText & " " & "  Wf_Pedido_Itens.Cod_Embalagem,"
        QueryC386029.CommandText = QueryC386029.CommandText & " " & "  Wf_Pedido_Itens.Seq_Item"
        QueryC386029.CommandText = QueryC386029.CommandText & " " & "FROM Wf_Pedido_Itens"
        QueryC386029.CommandText = QueryC386029.CommandText & " " & "WHERE Wf_Pedido_Itens.Pedido = " & _ 
ConvertValue(P63327, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC386029.CommandText = QueryC386029.CommandText & " " & "AND Wf_Pedido_Itens.Tp_Produto = " & _ 
ConvertValue(P63425, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC386029.Transaction = txn
        RsC386029 = QueryC386029.ExecuteReader()
        Dim iC386029 As Short
        ReDim C386029DataType(RsC386029.FieldCount)
        For iC386029 = 0 to RsC386029.FieldCount - 1
            Select Case RsC386029.GetDataTypeName(iC386029).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C386029DataType(iC386029 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C386029DataType(iC386029 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C386029DataType(iC386029 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC386029
        RsC386029_EOF = Not RsC386029.Read()

        GoTo Comp_C386031

    Comp_C386031:

        ' Fim_Itens_Pedido
        sCurrComponent = "Fim_Itens_Pedido"
        C386031DataType = 4
        C386031 = RsC386029_EOF
        GoTo Comp_C386032

    Comp_C386032:

        ' É_Fim_Itens_Pedido?
        sCurrComponent = "É_Fim_Itens_Pedido?"
        C386032 = (fn_ConvertValueCompiled(C386031, C386031DataType, AKB_DecimalPoint, False) = 1)
        C386032DataType = AKBTypeConst.cAKBTypeLogical
        If C386032 Then
            GoTo Comp_C386033
        Else
            GoTo Comp_C394991
        End If

    Comp_C386033:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C386033 As DataTable = New DataTable()
        tb_C386033.Columns.Add("vRetorno" & "")
        Dim row_C386033 As DataRow = tb_C386033.NewRow()
        row_C386033(0) = C386397
        tb_C386033.Rows.Add(row_C386033)
        R17504 = tb_C386033
        ReDim C386033DataType(1)
        C386033DataType(1) = C386397DataType
        ReturnDataType = C386033DataType

        GoTo Exit_R17504

    Comp_C386038:

        ' Up_TpPedido
        sCurrComponent = "Up_TpPedido"
        QueryC386038 = con.CreateCommand()
        QueryC386038.CommandText = QueryC386038.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS"
        QueryC386038.CommandText = QueryC386038.CommandText & " " & "SET WF_PEDIDO_ITENS.Tipo_Ped = " & _ 
ConvertValue(C395619, C395619DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC386038.CommandText = QueryC386038.CommandText & " " & "PRAZO_ENTREGA = decode(" & _ 
ConvertValue(C395619, C395619DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",2,null,PRAZO_ENTREGA)"
        QueryC386038.CommandText = QueryC386038.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P63327, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC386038.CommandText = QueryC386038.CommandText & " " & "AND WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C386396, C386396DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC386038.CommandText = QueryC386038.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P63425, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC386038.Transaction = txn
        C386038 = QueryC386038.ExecuteNonQuery()
        C386038DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C386039

    Comp_C386039:

        ' Próximo_Item
        sCurrComponent = "Próximo_Item"
        C386039DataType = 4
        RsC386029_EOF = Not RsC386029.Read()
        C386039 = True

        GoTo Comp_C386031

    Comp_C386396:

        ' Seq_Item
        sCurrComponent = "Seq_Item"
        C386396DataType = 0
        C386396DataType = C386029Datatype(4)
        If C386396DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC386029(3)) Then
          C386396 = Strings.RTrim(RsC386029(3))
        Else 
          C386396 = RsC386029(3)
        End If 

        GoTo Comp_C395618

    Comp_C386397:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C386397 = 1
        C386397DataType = 4
        GoTo Comp_C569312

    Comp_C394991:

        ' Sigla_Produto
        sCurrComponent = "Sigla_Produto"
        C394991DataType = 0
        C394991 = RsC386029(0)
        C394991DataType = C386029Datatype(1)
        If C394991DataType = AKBTypeConst.cAKBTypeString Then
          C394991 = IIF(IsDBNull(C394991), C394991, Strings.RTrim(C394991))
        End If 

        GoTo Comp_C394992

    Comp_C394992:

        ' Acesso
        sCurrComponent = "Acesso"
        C394992DataType = 0
        C394992DataType = C386029Datatype(2)
        If C394992DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC386029(1)) Then
          C394992 = Strings.RTrim(RsC386029(1))
        Else 
          C394992 = RsC386029(1)
        End If 

        GoTo Comp_C394993

    Comp_C394993:

        ' Cod_Embalagem
        sCurrComponent = "Cod_Embalagem"
        C394993DataType = 0
        C394993DataType = C386029Datatype(3)
        If C394993DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC386029(2)) Then
          C394993 = Strings.RTrim(RsC386029(2))
        Else 
          C394993 = RsC386029(2)
        End If 

        GoTo Comp_C386396

    Comp_C395618:

        ' Retorna tipo de Pedido do Acesso/Embalagem
        sCurrComponent = "Retorna tipo de Pedido do Acesso/Embalagem"
        C395618 = clsRuleDynamicallyCompiled_R9614.R9614(con, txn, IIf(Not IsDBNull(C394991), C394991, System.DBNull.Value), IIf(Not IsDBNull(C394992), C394992, System.DBNull.Value), IIf(Not IsDBNull(C394993), C394993, System.DBNull.Value), IIf(Not IsDBNull(RsC397000(0)), RsC397000(0), System.DBNull.Value))
        C395618CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C395618) Then
          iColumns = C395618.Columns.Count
        End If
        ReDim C395618DataType(iColumns)
        For iCol = 1 To iColumns
            C395618DataType(iCol) = clsRuleDynamicallyCompiled_R9614.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C395848

    Comp_C395619:

        ' Tipo_Pedido
        sCurrComponent = "Tipo_Pedido"
        C395619DataType = 0
        C395619 = C395618.rows(C395618CurrentRow - 1)(0)
        C395619DataType = C395618DataType(1)
        GoTo Comp_C386038

    Comp_C395620:

        ' Fim_Tipo_Pedido
        sCurrComponent = "Fim_Tipo_Pedido"
        C395620DataType = 4
        C395620 = (C395618CurrentRow > C395618.Rows.Count)
        GoTo Comp_C395621

    Comp_C395621:

        ' Fim_Itens_Pedido = 1?
        sCurrComponent = "Fim_Itens_Pedido = 1?"
        C395621 = (fn_ConvertValueCompiled(C395620, C395620DataType, AKB_DecimalPoint, False)  = 1)
        C395621DataType = AKBTypeConst.cAKBTypeLogical
        If C395621 Then
            GoTo Comp_C386039
        Else
            GoTo Comp_C395619
        End If

    Comp_C395848:

        ' Primeiro_Tipo_Pedido
        sCurrComponent = "Primeiro_Tipo_Pedido"
        C395848DataType = 4
        C395618CurrentRow = 1
        C395848 = True

        GoTo Comp_C395620

    Comp_C397000:

        ' SelTabelaPreçoVenda
        sCurrComponent = "SelTabelaPreçoVenda"
        QueryC397000 = con.CreateCommand()
        QueryC397000.CommandText = QueryC397000.CommandText & " " & "SELECT WF_Pedido.Tabela_Preco_Venda "
        QueryC397000.CommandText = QueryC397000.CommandText & " " & "FROM WF_Pedido"
        QueryC397000.CommandText = QueryC397000.CommandText & " " & "WHERE Wf_Pedido.Pedido = " & _ 
ConvertValue(P63327, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC397000.CommandText = QueryC397000.CommandText & " " & "AND Wf_Pedido.Tp_Produto = " & _ 
ConvertValue(P63425, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC397000.Transaction = txn
        RsC397000 = QueryC397000.ExecuteReader()
        Dim iC397000 As Short
        ReDim C397000DataType(RsC397000.FieldCount)
        For iC397000 = 0 to RsC397000.FieldCount - 1
            Select Case RsC397000.GetDataTypeName(iC397000).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C397000DataType(iC397000 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C397000DataType(iC397000 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C397000DataType(iC397000 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC397000
        RsC397000_EOF = Not RsC397000.Read()

        GoTo Comp_C386029

    Comp_C569312:

        ' SelPedEcommerce
        sCurrComponent = "SelPedEcommerce"
        QueryC569312 = con.CreateCommand()
        QueryC569312.CommandText = QueryC569312.CommandText & " " & "SELECT NVL(pedido_ecommerce ,0)"
        QueryC569312.CommandText = QueryC569312.CommandText & " " & "FROM WF_PEDIDO"
        QueryC569312.CommandText = QueryC569312.CommandText & " " & "WHERE PEDIDO = " & _ 
ConvertValue(P63327, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC569312.Transaction = txn
        RsC569312 = QueryC569312.ExecuteReader()
        Dim iC569312 As Short
        ReDim C569312DataType(RsC569312.FieldCount)
        For iC569312 = 0 to RsC569312.FieldCount - 1
            Select Case RsC569312.GetDataTypeName(iC569312).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C569312DataType(iC569312 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C569312DataType(iC569312 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C569312DataType(iC569312 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC569312
        RsC569312_EOF = Not RsC569312.Read()

        GoTo Comp_C569313

    Comp_C569313:

        ' ÉEcommerce?
        sCurrComponent = "ÉEcommerce?"
        C569313 = (fn_ConvertValueCompiled(RsC569312(0), C569312DataType(1), AKB_DecimalPoint, False) = 1)
        C569313DataType = AKBTypeConst.cAKBTypeLogical
        If C569313 Then
            GoTo Comp_C386033
        Else
            GoTo Comp_C397000
        End If

    Exit_R17504:

        Exit Function

    Err_R17504:
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
