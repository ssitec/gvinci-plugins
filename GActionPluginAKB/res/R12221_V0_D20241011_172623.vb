Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R12221

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

    Public Shared Function R12221(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P40649 As Double, ByVal P40651 As Double, ByVal P40652 As Double, ByVal P68672 As Object) As DataTable
    ' 
    ' 12221 - Verifica Zona do Representante x Cliente - Versão: 0
    ' 
        'On Error GoTo Err_R12221
        Dim sCurrComponent as String

        Dim QueryC233525 As New Object
        Dim RsC233525 As New Object
        Dim C233525DataType() As Short
        Dim RsC233525_EOF As Boolean
        Dim C233528DataType() As Short
        Dim C233529 As Object
        Dim C233529DataType As Short
        Dim QueryC233530 As New Object
        Dim RsC233530 As New Object
        Dim C233530DataType() As Short
        Dim RsC233530_EOF As Boolean
        Dim C233531 As Object
        Dim C233531DataType As Short
        Dim C233532 As Object
        Dim C233532DataType As Short
        Dim C233533 As Boolean
        Dim C233533DataType As Short
        Dim C233534 As Short
        Dim C233534DataType As Short
        Dim C233534ReturnDataType() As Short

        Dim C233535 As Object
        Dim C233535DataType As Short
        Dim C233536 As Boolean
        Dim C233536DataType As Short
        Dim C233537 As Short
        Dim C233537DataType As Short
        Dim C233537ReturnDataType() As Short

        Dim QueryC233546 As New Object
        Dim RsC233546 As New Object
        Dim C233546DataType() As Short
        Dim RsC233546_EOF As Boolean
        Dim C233548 As Object
        Dim C233548DataType As Short
        Dim C233549 As Object
        Dim C233549DataType As Short
        Dim C233551 As Boolean
        Dim C233551DataType As Short
        Dim C233553 As Boolean
        Dim C233553DataType As Short
        Dim QueryC233997 As New Object
        Dim RsC233997 As New Object
        Dim C233997DataType() As Short
        Dim RsC233997_EOF As Boolean
        Dim C233998 As Boolean
        Dim C233998DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C233529

    Comp_C233525:

        ' SelRepres
        sCurrComponent = "SelRepres"
        QueryC233525 = con.CreateCommand()
        QueryC233525.CommandText = QueryC233525.CommandText & " " & "Select count(*)"
        QueryC233525.CommandText = QueryC233525.CommandText & " " & "from AKBUSER01.WF_REPRES_ZONA "
        QueryC233525.CommandText = QueryC233525.CommandText & " " & "where WF_REPRES_ZONA.Cod_Repres = " & _ 
ConvertValue(C233548, C233548DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233525.CommandText = QueryC233525.CommandText & " " & "and WF_REPRES_ZONA.Cod_Zona = " & _ 
ConvertValue(C233531, C233531DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233525.Transaction = txn
        RsC233525 = QueryC233525.ExecuteReader()
        Dim iC233525 As Short
        ReDim C233525DataType(RsC233525.FieldCount)
        For iC233525 = 0 to RsC233525.FieldCount - 1
            Select Case RsC233525.GetDataTypeName(iC233525).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C233525DataType(iC233525 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C233525DataType(iC233525 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C233525DataType(iC233525 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC233525
        RsC233525_EOF = Not RsC233525.Read()

        GoTo Comp_C233536

    Comp_C233528:

        ' Ret
        sCurrComponent = "Ret"
        Dim tb_C233528 As DataTable = New DataTable()
        tb_C233528.Columns.Add("vRet" & "")
        Dim row_C233528 As DataRow = tb_C233528.NewRow()
        row_C233528(0) = C233529
        tb_C233528.Rows.Add(row_C233528)
        R12221 = tb_C233528
        ReDim C233528DataType(1)
        C233528DataType(1) = C233529DataType
        ReturnDataType = C233528DataType

        GoTo Exit_R12221

    Comp_C233529:

        ' vRet
        sCurrComponent = "vRet"
        C233529 = 1
        C233529DataType = 4
        GoTo Comp_C233997

    Comp_C233530:

        ' selCli
        sCurrComponent = "selCli"
        QueryC233530 = con.CreateCommand()
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "Select WF_CLIENTES.Cod_Zona "
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "from AKBUSER01.WF_CLIENTES "
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "where WF_CLIENTES.Cod_Cliente = " & _ 
ConvertValue(P40649, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "and " & _ 
ConvertValue(P68672, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " IS NULL"
        QueryC233530.CommandText = QueryC233530.CommandText & " " & ""
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "UNION"
        QueryC233530.CommandText = QueryC233530.CommandText & " " & ""
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "Select WF_CLIENTES.Cod_Zona "
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "from AKBUSER01.WF_CLIENTES "
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "where WF_CLIENTES.Cod_Cliente = " & _ 
ConvertValue(P68672, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233530.CommandText = QueryC233530.CommandText & " " & "and " & _ 
ConvertValue(P68672, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " IS NOT NULL"
        QueryC233530.Transaction = txn
        RsC233530 = QueryC233530.ExecuteReader()
        Dim iC233530 As Short
        ReDim C233530DataType(RsC233530.FieldCount)
        For iC233530 = 0 to RsC233530.FieldCount - 1
            Select Case RsC233530.GetDataTypeName(iC233530).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C233530DataType(iC233530 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C233530DataType(iC233530 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C233530DataType(iC233530 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC233530
        RsC233530_EOF = Not RsC233530.Read()

        GoTo Comp_C233532

    Comp_C233531:

        ' ZonaCli
        sCurrComponent = "ZonaCli"
        C233531DataType = 0
        C233531 = RsC233530(0)
        C233531DataType = C233530Datatype(1)
        If C233531DataType = AKBTypeConst.cAKBTypeString Then
          C233531 = IIF(IsDBNull(C233531), C233531, Strings.RTrim(C233531))
        End If 

        GoTo Comp_C233525

    Comp_C233532:

        ' EOF
        sCurrComponent = "EOF"
        C233532DataType = 4
        C233532 = RsC233530_EOF
        GoTo Comp_C233533

    Comp_C233533:

        ' EOF=1
        sCurrComponent = "EOF=1"
        C233533 = (fn_ConvertValueCompiled(C233532, C233532DataType, AKB_DecimalPoint, False) = 1)
        C233533DataType = AKBTypeConst.cAKBTypeLogical
        If C233533 Then
            GoTo Comp_C233534
        Else
            GoTo Comp_C233531
        End If

    Comp_C233534:

        ' MsgZonaCli
        sCurrComponent = "MsgZonaCli"
        C233534 = 1
        C233534DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C233535

    Comp_C233535:

        ' atvRet=0
        sCurrComponent = "atvRet=0"
        C233535DataType = 4
        C233529 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C233535 = True
        GoTo Comp_C233528

    Comp_C233536:

        ' SelRepres>0
        sCurrComponent = "SelRepres>0"
        C233536 = (fn_ConvertValueCompiled(RsC233525(0), C233525DataType(1), AKB_DecimalPoint, False) > 0)
        C233536DataType = AKBTypeConst.cAKBTypeLogical
        If C233536 Then
            GoTo Comp_C233528
        Else
            GoTo Comp_C233537
        End If

    Comp_C233537:

        ' MsgZonaRepres
        sCurrComponent = "MsgZonaRepres"
        C233537 = 1
        C233537DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C233535

    Comp_C233546:

        ' Sel_Ped
        sCurrComponent = "Sel_Ped"
        QueryC233546 = con.CreateCommand()
        QueryC233546.CommandText = QueryC233546.CommandText & " " & "select NVL(WF_PEDIDO.Cod_Repres,0)"
        QueryC233546.CommandText = QueryC233546.CommandText & " " & "from AKBUSER01.WF_PEDIDO "
        QueryC233546.CommandText = QueryC233546.CommandText & " " & "where WF_PEDIDO.Pedido = " & _ 
ConvertValue(P40652, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233546.CommandText = QueryC233546.CommandText & " " & "and WF_PEDIDO.Tp_Produto =" & _ 
ConvertValue(P40651, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233546.Transaction = txn
        RsC233546 = QueryC233546.ExecuteReader()
        Dim iC233546 As Short
        ReDim C233546DataType(RsC233546.FieldCount)
        For iC233546 = 0 to RsC233546.FieldCount - 1
            Select Case RsC233546.GetDataTypeName(iC233546).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C233546DataType(iC233546 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C233546DataType(iC233546 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C233546DataType(iC233546 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC233546
        RsC233546_EOF = Not RsC233546.Read()

        GoTo Comp_C233549

    Comp_C233548:

        ' Repres
        sCurrComponent = "Repres"
        C233548DataType = 0
        C233548 = RsC233546(0)
        C233548DataType = C233546Datatype(1)
        If C233548DataType = AKBTypeConst.cAKBTypeString Then
          C233548 = IIF(IsDBNull(C233548), C233548, Strings.RTrim(C233548))
        End If 

        GoTo Comp_C233553

    Comp_C233549:

        ' EOF_Ped
        sCurrComponent = "EOF_Ped"
        C233549DataType = 4
        C233549 = RsC233546_EOF
        GoTo Comp_C233551

    Comp_C233551:

        ' EOF_Ped=1
        sCurrComponent = "EOF_Ped=1"
        C233551 = (fn_ConvertValueCompiled(C233549, C233549DataType, AKB_DecimalPoint, False) = 1)
        C233551DataType = AKBTypeConst.cAKBTypeLogical
        If C233551 Then
            GoTo Comp_C233528
        Else
            GoTo Comp_C233548
        End If

    Comp_C233553:

        ' Repres=nulo
        sCurrComponent = "Repres=nulo"
        C233553 = (fn_ConvertValueCompiled(C233548, C233548DataType, AKB_DecimalPoint, False) = 0)
        C233553DataType = AKBTypeConst.cAKBTypeLogical
        If C233553 Then
            GoTo Comp_C233528
        Else
            GoTo Comp_C233530
        End If

    Comp_C233997:

        ' Sel_Transf
        sCurrComponent = "Sel_Transf"
        QueryC233997 = con.CreateCommand()
        QueryC233997.CommandText = QueryC233997.CommandText & " " & "Select count(*)"
        QueryC233997.CommandText = QueryC233997.CommandText & " " & "from  AKBUSER01.WF_PEDIDO_ITENS_SUBST "
        QueryC233997.CommandText = QueryC233997.CommandText & " " & "where WF_PEDIDO_ITENS_SUBST.Pedido2 = " & _ 
ConvertValue(P40652, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233997.CommandText = QueryC233997.CommandText & " " & "and  WF_PEDIDO_ITENS_SUBST.Tp_Produto2 = " & _ 
ConvertValue(P40651, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC233997.CommandText = QueryC233997.CommandText & " " & ""
        QueryC233997.Transaction = txn
        RsC233997 = QueryC233997.ExecuteReader()
        Dim iC233997 As Short
        ReDim C233997DataType(RsC233997.FieldCount)
        For iC233997 = 0 to RsC233997.FieldCount - 1
            Select Case RsC233997.GetDataTypeName(iC233997).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C233997DataType(iC233997 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C233997DataType(iC233997 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C233997DataType(iC233997 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC233997
        RsC233997_EOF = Not RsC233997.Read()

        GoTo Comp_C233998

    Comp_C233998:

        ' Sel_Transf>0
        sCurrComponent = "Sel_Transf>0"
        C233998 = (1 = 1)
        C233998DataType = AKBTypeConst.cAKBTypeLogical
        If C233998 Then
            GoTo Comp_C233528
        Else
            GoTo Comp_C233546
        End If

    Exit_R12221:

        Exit Function

    Err_R12221:
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
