Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R14036

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

    Public Shared Function R14036(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByRef P49743() As Object, ByRef P49744() As Object) As DataTable
    ' 
    ' 14036 - Monta Lista de Departamentos - Vers�o: 0
    ' 
        'On Error GoTo Err_R14036
        Dim sCurrComponent as String

        Dim C295347DataType() As Short
        Dim C295354 As Object
        Dim C295354DataType As Short
        Dim C295356 As Object
        Dim C295356DataType As Short
        Dim C295357 As Object
        Dim C295357DataType As Short
        Dim C295366 As Boolean
        Dim C295366DataType As Short
        Dim C295368 As Object
        Dim C295368DataType As Short
        Dim C295369 As Object
        Dim C295369DataType As Short
        Dim C295370 As Object
        Dim C295370DataType As Short
        Dim C295371 As Object
        Dim C295371DataType As Short
        Dim C295372 As Object
        Dim C295372DataType As Short
        Dim C295373 As Boolean
        Dim C295373DataType As Short
        Dim C295374 As Object
        Dim C295374DataType As Short
        Dim C295375 As Object
        Dim C295375DataType As Short
        Dim C295380 As Object
        Dim C295380DataType As Short
        Dim C295381 As Object
        Dim C295381DataType As Short
        Dim C295382 As Object
        Dim C295382DataType As Short
        Dim C295383 As Object
        Dim C295383DataType As Short
        Dim C418139 As Object
        Dim C418139DataType As Short
        Dim C418140 As Double
        Dim C418140DataType As Short
        Dim C418141 As Object
        Dim C418141DataType As Short
        Dim C418142 As Boolean
        Dim C418142DataType As Short
        Dim P49743Current As Integer

        Dim P49744Current As Integer

        ReDim ReturnDatatype(0)

        GoTo Comp_C295354

    Comp_C295347:

        ' RetPedidos
        sCurrComponent = "RetPedidos"
        Dim tb_C295347 As DataTable = New DataTable()
        tb_C295347.Columns.Add("vQryDepto" & "")
        Dim row_C295347 As DataRow = tb_C295347.NewRow()
        row_C295347(0) = C295354
        tb_C295347.Rows.Add(row_C295347)
        R14036 = tb_C295347
        ReDim C295347DataType(1)
        C295347DataType(1) = C295354DataType
        ReturnDataType = C295347DataType

        GoTo Exit_R14036

    Comp_C295354:

        ' vQryDepto
        sCurrComponent = "vQryDepto"
        C295354 = "AND 1=1"
        C295354DataType = 5
        GoTo Comp_C295356

    Comp_C295356:

        ' vListaDepto
        sCurrComponent = "vListaDepto"
        C295356 = "Null"
        C295356DataType = 3
        GoTo Comp_C295357

    Comp_C295357:

        ' vQuery
        sCurrComponent = "vQuery"
        C295357 = System.DBNull.Value
        C295357DataType = 5
        GoTo Comp_C418139

    Comp_C295366:

        ' Fim Depto?
        sCurrComponent = "Fim Depto?"
        C295366 = (fn_ConvertValueCompiled(C295370, C295370DataType, AKB_DecimalPoint, False) = 1)
        C295366DataType = AKBTypeConst.cAKBTypeLogical
        If C295366 Then
            GoTo Comp_C418142
        Else
            GoTo Comp_C295371
        End If

    Comp_C295368:

        ' FirstDepto
        sCurrComponent = "FirstDepto"
        C295368DataType = 4
        P49743Current = 1
        C295368 = True
        P49744Current = 1
        C295368 = True

        GoTo Comp_C295370

    Comp_C295369:

        ' Next_Depto
        sCurrComponent = "Next_Depto"
        C295369DataType = 4
        If Not (P49743Current > UBound(P49743)) Then
            P49743Current = P49743Current + 1
            C295369 = True
        Else
            C295369 = False
        End If
        If Not (P49744Current > UBound(P49744)) Then
            P49744Current = P49744Current + 1
            C295369 = True
        Else
            C295369 = False
        End If

        GoTo Comp_C295370

    Comp_C295370:

        ' EOF_Depto
        sCurrComponent = "EOF_Depto"
        C295370DataType = 4
        C295370 = (P49743Current > UBound(P49743))
        GoTo Comp_C295366

    Comp_C295371:

        ' vDepto
        sCurrComponent = "vDepto"
        C295371DataType = 0
        C295371 = P49743(P49743Current)
        C295371DataType = 3
        GoTo Comp_C295372

    Comp_C295372:

        ' vDeptoSel
        sCurrComponent = "vDeptoSel"
        C295372DataType = 0
        C295372 = P49744(P49744Current)
        C295372DataType = 4
        GoTo Comp_C295373

    Comp_C295373:

        ' Depto Marcado?
        sCurrComponent = "Depto Marcado?"
        C295373 = (fn_ConvertValueCompiled(C295372, C295372DataType, AKB_DecimalPoint, False) = 1)
        C295373DataType = AKBTypeConst.cAKBTypeLogical
        If C295373 Then
            GoTo Comp_C295374
        Else
            GoTo Comp_C295369
        End If

    Comp_C295374:

        ' ConcatDepto
        sCurrComponent = "ConcatDepto"
        C295374DataType = 3
        C295374 = C295356  & ", '" & C295371 & "'"
        GoTo Comp_C418140

    Comp_C295375:

        ' atvConcatDepto
        sCurrComponent = "atvConcatDepto"
        C295375DataType = 4
        C295356 = fn_ConvertValueCompiled(C295374 , 3, AKB_DecimalPoint)
        C295375 = True
        GoTo Comp_C295369

    Comp_C295380:

        ' TiraAspasDepto1
        sCurrComponent = "TiraAspasDepto1"
        C295380DataType = 3
        C295380 = Replace(C295356 , "'''", "'")
        GoTo Comp_C295381

    Comp_C295381:

        ' TiraNullDepto
        sCurrComponent = "TiraNullDepto"
        C295381DataType = 3
        C295381 = Replace(C295380 , "Null,", "")
        GoTo Comp_C295382

    Comp_C295382:

        ' atvListaDepto
        sCurrComponent = "atvListaDepto"
        C295382DataType = 4
        C295356 = fn_ConvertValueCompiled("(" & C295381 & ")", 3, AKB_DecimalPoint)
        C295382 = True
        GoTo Comp_C295383

    Comp_C295383:

        ' atvQryDepto
        sCurrComponent = "atvQryDepto"
        C295383DataType = 4
        C295354 = fn_ConvertValueCompiled("AND WF_PRE_PEDIDO.DEPARTAMENTO IN " & C295356 , 5, AKB_DecimalPoint)
        C295383 = True
        GoTo Comp_C295347

    Comp_C418139:

        ' Count
        sCurrComponent = "Count"
        C418139 = 0
        C418139DataType = 1
        GoTo Comp_C295368

    Comp_C418140:

        ' Count+1
        sCurrComponent = "Count+1"
        C418140 = fn_ConvertValueCompiled(C418139, C418139DataType, AKB_DecimalPoint, True) + 1
        C418140DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C418141

    Comp_C418141:

        ' AttCount
        sCurrComponent = "AttCount"
        C418141DataType = 4
        C418139 = fn_ConvertValueCompiled(C418140 , 1, AKB_DecimalPoint)
        C418141 = True
        GoTo Comp_C295375

    Comp_C418142:

        ' Count>0?
        sCurrComponent = "Count>0?"
        C418142 = (fn_ConvertValueCompiled(C418139, C418139DataType, AKB_DecimalPoint, False) > 0)
        C418142DataType = AKBTypeConst.cAKBTypeLogical
        If C418142 Then
            GoTo Comp_C295380
        Else
            GoTo Comp_C295347
        End If

    Exit_R14036:

        Exit Function

    Err_R14036:
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
    ' Retorna o inteiro maior ou igual ao n�mero
    '
    Public Shared Function fn_Ceiling(ByVal dParam As Double) As Integer

        ' Obt�m valor
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
    ' Retorna o inteiro menor ou igual ao n�mero
    '
    Public Shared Function fn_Floor(ByVal dParam As Double) As Integer

        ' Obt�m valor
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
           Extenso = "Valor N�o Numerico"
           Exit Function
       End If

       g(1, 0) = "CENTAVO"
       g(2, 0) = ""
       g(3, 0) = "MIL"
       g(4, 0) = "MILH�O"
       g(5, 0) = "BILH�O"
       g(6, 0) = "TRILH�O"

       g(1, 1) = "CENTAVOS"
       g(2, 1) = ""
       g(3, 1) = "MIL"
       g(4, 1) = "MILH�ES"
       g(5, 1) = "BILH�ES"
       g(6, 1) = "TRILH�ES"

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

       'Retira os caracteres de formata��o
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

       'Retira os caracteres de formata��o
       CPF = Replace(CPF, ".", "")
       CPF = Replace(CPF, "/", "")
       CPF = Replace(CPF, "-", "")

       'Esta rotina foi adaptada da revista F�rum Access
       On Error GoTo Err_CPF
       Dim i As Short        'utilizada nos FOR... NEXT
       Dim strcampo As String  'armazena do CPF que ser� utilizada para o c�lculo
       Dim strCaracter As String   'armazena os d�gitos do CPF da direita para a esquerda
       Dim intNumero As Short    'armazena o digito separado para c�lculo (uma a um)
       Dim intMais As Short  'armazena o digito espec�fico multiplicado pela sua base
       Dim lngSoma As Integer     'armazena a soma dos d�gitos multiplicados pela sua base(intmais)
       Dim dblDivisao As Double    'armazena a divis�o dos d�gitos * base por 11
       Dim lngInteiro As Integer  'armazena inteiro da divis�o
       Dim intResto As Short     'armazena o resto
       Dim intDig1 As Short  'armazena o 1� digito verificador
       Dim intDig2 As Short  'armazena o 2� digito verificador
       Dim strConf As String   'armazena o digito verificador

       lngSoma = 0
       intNumero = 0
       intMais = 0
       strcampo = Left(CPF, 9)

       'Inicia c�lculos do 1� d�gito
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
       'Inicia c�lculos do 2� d�gito
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
       MsgBox("Erro ao executar a vers�o compilada GVINCI da fun��o ConvertValue ")
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
    Public Shared Function ObjTable_NewIdentity(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal TableName as String) As Long
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select SEQ_" & TableName & ".nextval from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Query.Transaction = txn
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            ObjTable_NewIdentity = Rs(0)
        End If
        Rs.Close()
    End Function

End Class
