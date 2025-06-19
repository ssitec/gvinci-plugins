Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R1533

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

    Public Shared Function R1533(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P1206 As Object, ByVal P2495 As Double, ByVal P2496 As Object, ByVal P3377 As Object, ByVal P12286 As Object) As DataTable
    ' 
    ' 1533 - Verifica validade da tabela de preço - Versão: 0
    ' 
        'On Error GoTo Err_R1533
        Dim sCurrComponent as String

        Dim QueryC9569 As New Object
        Dim RsC9569 As New Object
        Dim C9569DataType() As Short
        Dim RsC9569_EOF As Boolean
        Dim C9572 As Short
        Dim C9572DataType As Short
        Dim C9572ReturnDataType() As Short

        Dim C9574 As Boolean
        Dim C9574DataType As Short
        Dim C9578 As Object
        Dim C9578DataType As Short
        Dim C9581 As Boolean
        Dim C9581DataType As Short
        Dim C9584DataType() As Short
        Dim C9591 As Object
        Dim C9591DataType As Short
        Dim C9592 As Object
        Dim C9592DataType As Short
        Dim C12837DataType() As Short
        Dim C12838 As Boolean
        Dim C12838DataType As Short
        Dim C12839 As Short
        Dim C12839DataType As Short
        Dim C12839ReturnDataType() As Short

        Dim QueryC38276 As New Object
        Dim RsC38276 As New Object
        Dim C38276DataType() As Short
        Dim RsC38276_EOF As Boolean
        Dim C38277 As Object
        Dim C38277DataType As Short
        Dim QueryC58187 As New Object
        Dim RsC58187 As New Object
        Dim C58187DataType() As Short
        Dim RsC58187_EOF As Boolean
        Dim C58188 As Object
        Dim C58188DataType As Short
        Dim C58189 As Boolean
        Dim C58189DataType As Short
        Dim C58379 As Boolean
        Dim C58379DataType As Short
        Dim C58380 As Boolean
        Dim C58380DataType As Short
        Dim C265437 As Boolean
        Dim C265437DataType As Short
        Dim C265439DataType() As Short
        Dim C282604 As Boolean
        Dim C282604DataType As Short
        Dim C282605 As Short
        Dim C282605DataType As Short
        Dim C282605ReturnDataType() As Short

        Dim C320114 As Object
        Dim C320114DataType As Short
        P1206 = IIf(IsDBNull(P1206), 0, P1206)

        P2496 = IIf(IsDBNull(P2496), 0, P2496)

        P3377 = IIf(IsDBNull(P3377), 0, P3377)

        ReDim ReturnDatatype(0)

        GoTo Comp_C9591

    Comp_C9569:

        ' TabVenda
        sCurrComponent = "TabVenda"
        QueryC9569 = con.CreateCommand()
        QueryC9569.CommandText = QueryC9569.CommandText & " " & "SELECT WF_TABELA_PRECO_VENDA.Data_Validade_Final "
        QueryC9569.CommandText = QueryC9569.CommandText & " " & ""
        QueryC9569.CommandText = QueryC9569.CommandText & " " & "FROM AKBUSER01.WF_TABELA_PRECO_VENDA "
        QueryC9569.CommandText = QueryC9569.CommandText & " " & ""
        QueryC9569.CommandText = QueryC9569.CommandText & " " & "WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(P1206, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC9569.CommandText = QueryC9569.CommandText & " " & "AND WF_TABELA_PRECO_VENDA.Data_Validade_Final > " & _ 
ConvertValue(P12286, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC9569.CommandText = QueryC9569.CommandText & " " & ""
        QueryC9569.Transaction = txn
        RsC9569 = QueryC9569.ExecuteReader()
        Dim iC9569 As Short
        ReDim C9569DataType(RsC9569.FieldCount)
        For iC9569 = 0 to RsC9569.FieldCount - 1
            Select Case RsC9569.GetDataTypeName(iC9569).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C9569DataType(iC9569 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C9569DataType(iC9569 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C9569DataType(iC9569 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC9569
        RsC9569_EOF = Not RsC9569.Read()

        GoTo Comp_C9578

    Comp_C9572:

        ' MsgInvalid
        sCurrComponent = "MsgInvalid"
        C9572 = 1
        C9572DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C12837

    Comp_C9574:

        ' TipoVenda
        sCurrComponent = "TipoVenda"
        C9574 = (fn_ConvertValueCompiled(C38277, C38277DataType, AKB_DecimalPoint, False) = "Venda")
        C9574DataType = AKBTypeConst.cAKBTypeLogical
        If C9574 Then
            GoTo Comp_C9569
        Else
            GoTo Comp_C58379
        End If

    Comp_C9578:

        ' Eof1
        sCurrComponent = "Eof1"
        C9578DataType = 4
        C9578 = RsC9569_EOF
        GoTo Comp_C9581

    Comp_C9581:

        ' Eof1=0
        sCurrComponent = "Eof1=0"
        C9581 = (fn_ConvertValueCompiled(C9578, C9578DataType, AKB_DecimalPoint, False) =0)
        C9581DataType = AKBTypeConst.cAKBTypeLogical
        If C9581 Then
            GoTo Comp_C9584
        Else
            GoTo Comp_C9572
        End If

    Comp_C9584:

        ' RetTrue
        sCurrComponent = "RetTrue"
        Dim tb_C9584 As DataTable = New DataTable()
        tb_C9584.Columns.Add("VTrue" & "")
        Dim row_C9584 As DataRow = tb_C9584.NewRow()
        row_C9584(0) = C9592
        tb_C9584.Rows.Add(row_C9584)
        R1533 = tb_C9584
        ReDim C9584DataType(1)
        C9584DataType(1) = C9592DataType
        ReturnDataType = C9584DataType

        GoTo Exit_R1533

    Comp_C9591:

        ' VFalse
        sCurrComponent = "VFalse"
        C9591 = 0
        C9591DataType = 4
        GoTo Comp_C9592

    Comp_C9592:

        ' VTrue
        sCurrComponent = "VTrue"
        C9592 = 1
        C9592DataType = 4
        GoTo Comp_C320114

    Comp_C12837:

        ' RetFalse
        sCurrComponent = "RetFalse"
        Dim tb_C12837 As DataTable = New DataTable()
        tb_C12837.Columns.Add("VFalse" & "")
        Dim row_C12837 As DataRow = tb_C12837.NewRow()
        row_C12837(0) = C9591
        tb_C12837.Rows.Add(row_C12837)
        R1533 = tb_C12837
        ReDim C12837DataType(1)
        C12837DataType(1) = C9591DataType
        ReturnDataType = C12837DataType

        GoTo Exit_R1533

    Comp_C12838:

        ' Ped<0
        sCurrComponent = "Ped<0"
        C12838 = (fn_ConvertValueCompiled(P3377, 1, AKB_DecimalPoint, False) <0)
        C12838DataType = AKBTypeConst.cAKBTypeLogical
        If C12838 Then
            GoTo Comp_C12839
        Else
            GoTo Comp_C265437
        End If

    Comp_C12839:

        ' MSG1
        sCurrComponent = "MSG1"
        C12839 = 1
        C12839DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C12837

    Comp_C38276:

        ' SelTipoVda
        sCurrComponent = "SelTipoVda"
        QueryC38276 = con.CreateCommand()
        QueryC38276.CommandText = QueryC38276.CommandText & " " & "SELECT WF_TP_PRODUTO.Tp_Tab_Preco "
        QueryC38276.CommandText = QueryC38276.CommandText & " " & "FROM AKBUSER01.WF_TP_PRODUTO "
        QueryC38276.CommandText = QueryC38276.CommandText & " " & "WHERE WF_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P2495, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC38276.Transaction = txn
        RsC38276 = QueryC38276.ExecuteReader()
        Dim iC38276 As Short
        ReDim C38276DataType(RsC38276.FieldCount)
        For iC38276 = 0 to RsC38276.FieldCount - 1
            Select Case RsC38276.GetDataTypeName(iC38276).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C38276DataType(iC38276 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C38276DataType(iC38276 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C38276DataType(iC38276 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC38276
        RsC38276_EOF = Not RsC38276.Read()

        GoTo Comp_C38277

    Comp_C38277:

        ' TpVenda
        sCurrComponent = "TpVenda"
        C38277DataType = 0
        C38277 = RsC38276(0)
        C38277DataType = C38276Datatype(1)
        If C38277DataType = AKBTypeConst.cAKBTypeString Then
          C38277 = IIF(IsDBNull(C38277), C38277, Strings.RTrim(C38277))
        End If 

        GoTo Comp_C9574

    Comp_C58187:

        ' TabCusto
        sCurrComponent = "TabCusto"
        QueryC58187 = con.CreateCommand()
        QueryC58187.CommandText = QueryC58187.CommandText & " " & "SELECT  WF_TAB_PRECO_CUSTO.Data_Validade_Final "
        QueryC58187.CommandText = QueryC58187.CommandText & " " & ""
        QueryC58187.CommandText = QueryC58187.CommandText & " " & "FROM  AKBUSER01.WF_TAB_PRECO_CUSTO "
        QueryC58187.CommandText = QueryC58187.CommandText & " " & ""
        QueryC58187.CommandText = QueryC58187.CommandText & " " & "WHERE WF_TAB_PRECO_CUSTO.Tabela_Preco_Custo = " & _ 
ConvertValue(P2496, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC58187.CommandText = QueryC58187.CommandText & " " & "AND (WF_TAB_PRECO_CUSTO.Data_Validade_Final  > " & _ 
ConvertValue(P12286, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " OR WF_TAB_PRECO_CUSTO.Data_Validade_Final IS NULL)"
        QueryC58187.Transaction = txn
        RsC58187 = QueryC58187.ExecuteReader()
        Dim iC58187 As Short
        ReDim C58187DataType(RsC58187.FieldCount)
        For iC58187 = 0 to RsC58187.FieldCount - 1
            Select Case RsC58187.GetDataTypeName(iC58187).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C58187DataType(iC58187 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C58187DataType(iC58187 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C58187DataType(iC58187 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC58187
        RsC58187_EOF = Not RsC58187.Read()

        GoTo Comp_C58188

    Comp_C58188:

        ' Eof2
        sCurrComponent = "Eof2"
        C58188DataType = 4
        C58188 = RsC58187_EOF
        GoTo Comp_C58189

    Comp_C58189:

        ' Eof2=0
        sCurrComponent = "Eof2=0"
        C58189 = (fn_ConvertValueCompiled(C58188, C58188DataType, AKB_DecimalPoint, False) = 0)
        C58189DataType = AKBTypeConst.cAKBTypeLogical
        If C58189 Then
            GoTo Comp_C9584
        Else
            GoTo Comp_C9572
        End If

    Comp_C58379:

        ' TipoInterna
        sCurrComponent = "TipoInterna"
        C58379 = (fn_ConvertValueCompiled(C38277, C38277DataType, AKB_DecimalPoint, False) = "Interno")
        C58379DataType = AKBTypeConst.cAKBTypeLogical
        If C58379 Then
            GoTo Comp_C9584
        Else
            GoTo Comp_C58380
        End If

    Comp_C58380:

        ' TipoServiço
        sCurrComponent = "TipoServiço"
        C58380 = (fn_ConvertValueCompiled(C38277, C38277DataType, AKB_DecimalPoint, False) = "Serviço")
        C58380DataType = AKBTypeConst.cAKBTypeLogical
        If C58380 Then
            GoTo Comp_C58187
        Else
            GoTo Comp_C9584
        End If

    Comp_C265437:

        ' Tipo=200
        sCurrComponent = "Tipo=200"
        C265437 = (( fn_ConvertValueCompiled(P2495, 1, AKB_DecimalPoint, False) = 200 ) or ( fn_ConvertValueCompiled(P2495, 1, AKB_DecimalPoint, False) = 201 ) or ( fn_ConvertValueCompiled(P2495, 1, AKB_DecimalPoint, False) = 202 ) or ( fn_ConvertValueCompiled(P2495, 1, AKB_DecimalPoint, False) = 205 ))
        C265437DataType = AKBTypeConst.cAKBTypeLogical
        If C265437 Then
            GoTo Comp_C265439
        Else
            GoTo Comp_C38276
        End If

    Comp_C265439:

        ' RetOk
        sCurrComponent = "RetOk"
        Dim tb_C265439 As DataTable = New DataTable()
        tb_C265439.Columns.Add("VTrue" & "")
        Dim row_C265439 As DataRow = tb_C265439.NewRow()
        row_C265439(0) = C9592
        tb_C265439.Rows.Add(row_C265439)
        R1533 = tb_C265439
        ReDim C265439DataType(1)
        C265439DataType(1) = C9592DataType
        ReturnDataType = C265439DataType

        GoTo Exit_R1533

    Comp_C282604:

        ' DtPedido está nulo
        sCurrComponent = "DtPedido está nulo"
        C282604 = (fn_ConvertValueCompiled(C320114, C320114DataType, AKB_DecimalPoint, False) = 1)
        C282604DataType = AKBTypeConst.cAKBTypeLogical
        If C282604 Then
            GoTo Comp_C282605
        Else
            GoTo Comp_C12838
        End If

    Comp_C282605:

        ' MSG2
        sCurrComponent = "MSG2"
        C282605 = 1
        C282605DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C12837

    Comp_C320114:

        ' nuloDtPedido
        sCurrComponent = "nuloDtPedido"
        C320114DataType = 4
        C320114 = IsDBNull (P12286)
        GoTo Comp_C282604

    Exit_R1533:

        Exit Function

    Err_R1533:
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
