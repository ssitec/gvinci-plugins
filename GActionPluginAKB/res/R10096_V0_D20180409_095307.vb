Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R10096

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

    Public Shared Function R10096(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P32279 As Double, ByVal P32280 As Double, ByVal P32281 As Double) As DataTable
    ' 
    ' 10096 - Marca pedido como Verificado se Cliente Coligada - Versão: 0
    ' 
        'On Error GoTo Err_R10096
        Dim sCurrComponent as String

        Dim QueryC176545 As New Object
        Dim RsC176545 As New Object
        Dim C176545DataType() As Short
        Dim RsC176545_EOF As Boolean
        Dim QueryC176547 As New Object
        Dim C176547 As Integer
        Dim C176547DataType As Short
        Dim C176548 As Boolean
        Dim C176548DataType As Short
        Dim C176549DataType() As Short
        Dim C176551 As Object
        Dim C176551DataType As Short
        Dim C176558 As Object
        Dim C176558DataType As Short
        Dim C176560 As Object
        Dim C176560DataType As Short
        Dim QueryC493688 As New Object
        Dim RsC493688 As New Object
        Dim C493688DataType() As Short
        Dim RsC493688_EOF As Boolean
        Dim C493689 As Boolean
        Dim C493689DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C176551

    Comp_C176545:

        ' Sel_Coligada
        sCurrComponent = "Sel_Coligada"
        QueryC176545 = con.CreateCommand()
        QueryC176545.CommandText = QueryC176545.CommandText & " " & "Select  count(*)"
        QueryC176545.CommandText = QueryC176545.CommandText & " " & "from AKBUSER01.WF_ESTAB_JURID_COLIGADAS , AKBUSER01.WF_PEDIDO "
        QueryC176545.CommandText = QueryC176545.CommandText & " " & "where WF_ESTAB_JURID_COLIGADAS.Cod_Pessoa = WF_PEDIDO.Cod_Cliente "
        QueryC176545.CommandText = QueryC176545.CommandText & " " & "and WF_ESTAB_JURID_COLIGADAS.Cod_Pessoa_Estab_Juridico = WF_PEDIDO.Cod_Pessoa_Estab_Juridico "
        QueryC176545.CommandText = QueryC176545.CommandText & " " & ""
        QueryC176545.CommandText = QueryC176545.CommandText & " " & "and WF_PEDIDO.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P32281, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC176545.CommandText = QueryC176545.CommandText & " " & "and WF_PEDIDO.Pedido = " & _ 
ConvertValue(P32279, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC176545.CommandText = QueryC176545.CommandText & " " & "and WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P32280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC176545.CommandText = QueryC176545.CommandText & " " & "and WF_PEDIDO.Origem_OC = 1"
        QueryC176545.CommandText = QueryC176545.CommandText & " " & ""
        QueryC176545.Transaction = txn
        RsC176545 = QueryC176545.ExecuteReader()
        Dim iC176545 As Short
        ReDim C176545DataType(RsC176545.FieldCount)
        For iC176545 = 0 to RsC176545.FieldCount - 1
            Select Case RsC176545.GetDataTypeName(iC176545).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C176545DataType(iC176545 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C176545DataType(iC176545 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C176545DataType(iC176545 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC176545
        RsC176545_EOF = Not RsC176545.Read()

        GoTo Comp_C176548

    Comp_C176547:

        ' Up_Verificado
        sCurrComponent = "Up_Verificado"
        QueryC176547 = con.CreateCommand()
        QueryC176547.CommandText = QueryC176547.CommandText & " " & "Update AKBUSER01.WF_PEDIDO set WF_PEDIDO.Verificado = 1,"
        QueryC176547.CommandText = QueryC176547.CommandText & " " & "                                          WF_PEDIDO.Usuario_Verificacao = " & _ 
ConvertValue(C176560, C176560DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC176547.CommandText = QueryC176547.CommandText & " " & "                                          WF_PEDIDO.Data_Verificacao = " & _ 
ConvertValue(C176558, C176558DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC176547.CommandText = QueryC176547.CommandText & " " & "where WF_PEDIDO.Pedido = " & _ 
ConvertValue(P32279, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC176547.CommandText = QueryC176547.CommandText & " " & "and WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P32280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC176547.CommandText = QueryC176547.CommandText & " " & "and WF_PEDIDO.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P32281, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC176547.Transaction = txn
        C176547 = QueryC176547.ExecuteNonQuery()
        C176547DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C176549

    Comp_C176548:

        ' Sel_Coligada>0
        sCurrComponent = "Sel_Coligada>0"
        C176548 = (fn_ConvertValueCompiled(RsC176545(0), C176545DataType(1), AKB_DecimalPoint, False) > 0)
        C176548DataType = AKBTypeConst.cAKBTypeLogical
        If C176548 Then
            GoTo Comp_C176558
        Else
            GoTo Comp_C493688
        End If

    Comp_C176549:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C176549 As DataTable = New DataTable()
        tb_C176549.Columns.Add("vRet" & "")
        Dim row_C176549 As DataRow = tb_C176549.NewRow()
        row_C176549(0) = C176551
        tb_C176549.Rows.Add(row_C176549)
        R10096 = tb_C176549
        ReDim C176549DataType(1)
        C176549DataType(1) = C176551DataType
        ReturnDataType = C176549DataType

        GoTo Exit_R10096

    Comp_C176551:

        ' vRet
        sCurrComponent = "vRet"
        C176551 = 1
        C176551DataType = 4
        GoTo Comp_C176545

    Comp_C176558:

        ' DataSistema
        sCurrComponent = "DataSistema"
        C176558DataType = 2
        C176558 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C176560

    Comp_C176560:

        ' Usuario
        sCurrComponent = "Usuario"
        C176560DataType = 3
        C176560 = Mid(con.ConnectionString, InStr(con.ConnectionString, "User Id=")+8, Instr(InStr(con.ConnectionString, "User Id="), con.ConnectionString, ";")-InStr(con.ConnectionString, "User Id=")-8)
        GoTo Comp_C176547

    Comp_C493688:

        ' FlagVerifTP
        sCurrComponent = "FlagVerifTP"
        QueryC493688 = con.CreateCommand()
        QueryC493688.CommandText = QueryC493688.CommandText & " " & "SELECT Nao_lib_Cred_sem_Verificacao"
        QueryC493688.CommandText = QueryC493688.CommandText & " " & "FROM WF_TP_PRODUTO"
        QueryC493688.CommandText = QueryC493688.CommandText & " " & "WHERE Tp_Produto = " & _ 
ConvertValue(P32280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC493688.Transaction = txn
        RsC493688 = QueryC493688.ExecuteReader()
        Dim iC493688 As Short
        ReDim C493688DataType(RsC493688.FieldCount)
        For iC493688 = 0 to RsC493688.FieldCount - 1
            Select Case RsC493688.GetDataTypeName(iC493688).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C493688DataType(iC493688 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C493688DataType(iC493688 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C493688DataType(iC493688 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC493688
        RsC493688_EOF = Not RsC493688.Read()

        GoTo Comp_C493689

    Comp_C493689:

        ' FlagVerifTP=1?
        sCurrComponent = "FlagVerifTP=1?"
        C493689 = (fn_ConvertValueCompiled(RsC493688(0), C493688DataType(1), AKB_DecimalPoint, False) = 1)
        C493689DataType = AKBTypeConst.cAKBTypeLogical
        If C493689 Then
            GoTo Comp_C176549
        Else
            GoTo Comp_C176558
        End If

    Exit_R10096:

        Exit Function

    Err_R10096:
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
