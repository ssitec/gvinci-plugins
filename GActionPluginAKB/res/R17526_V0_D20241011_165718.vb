Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R17526

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

    Public Shared Function R17526(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P63498 As Double, ByVal P63499 As Double, ByVal P63925 As Double) As DataTable
    ' 
    ' 17526 - Atualiza Representante de Cliente Disponível - Versão: 0
    ' 
        'On Error GoTo Err_R17526
        Dim sCurrComponent as String

        Dim QueryC386918 As New Object
        Dim RsC386918 As New Object
        Dim C386918DataType() As Short
        Dim RsC386918_EOF As Boolean
        Dim C386919 As Object
        Dim C386919DataType As Short
        Dim C386920 As Boolean
        Dim C386920DataType As Short
        Dim QueryC386921 As New Object
        Dim C386921 As Integer
        Dim C386921DataType As Short
        Dim C386923DataType() As Short
        Dim QueryC390163 As New Object
        Dim C390163 As Integer
        Dim C390163DataType As Short
        Dim QueryC391071 As New Object
        Dim C391071 As Integer
        Dim C391071DataType As Short
        Dim QueryC573153 As New Object
        Dim C573153 As Integer
        Dim C573153DataType As Short
        Dim QueryC573162 As New Object
        Dim RsC573162 As New Object
        Dim C573162DataType() As Short
        Dim RsC573162_EOF As Boolean
        Dim C573163 As Boolean
        Dim C573163DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C386918

    Comp_C386918:

        ' selClienteDisp
        sCurrComponent = "selClienteDisp"
        QueryC386918 = con.CreateCommand()
        QueryC386918.CommandText = QueryC386918.CommandText & " " & "select COUNT(*)"
        QueryC386918.CommandText = QueryC386918.CommandText & " " & "from WF_CLIENTES_DISPONIVEIS"
        QueryC386918.CommandText = QueryC386918.CommandText & " " & "where Cod_Cliente = " & _ 
ConvertValue(P63498, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC386918.Transaction = txn
        RsC386918 = QueryC386918.ExecuteReader()
        Dim iC386918 As Short
        ReDim C386918DataType(RsC386918.FieldCount)
        For iC386918 = 0 to RsC386918.FieldCount - 1
            Select Case RsC386918.GetDataTypeName(iC386918).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C386918DataType(iC386918 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C386918DataType(iC386918 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C386918DataType(iC386918 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC386918
        RsC386918_EOF = Not RsC386918.Read()

        GoTo Comp_C386919

    Comp_C386919:

        ' contar
        sCurrComponent = "contar"
        C386919DataType = 0
        C386919 = RsC386918(0)
        C386919DataType = C386918Datatype(1)
        If C386919DataType = AKBTypeConst.cAKBTypeString Then
          C386919 = IIF(IsDBNull(C386919), C386919, Strings.RTrim(C386919))
        End If 

        GoTo Comp_C386920

    Comp_C386920:

        ' clienteDisponivel
        sCurrComponent = "clienteDisponivel"
        C386920 = (fn_ConvertValueCompiled(C386919, C386919DataType, AKB_DecimalPoint, False) > 0)
        C386920DataType = AKBTypeConst.cAKBTypeLogical
        If C386920 Then
            GoTo Comp_C386921
        Else
            GoTo Comp_C386923
        End If

    Comp_C386921:

        ' upReprsCliente
        sCurrComponent = "upReprsCliente"
        QueryC386921 = con.CreateCommand()
        QueryC386921.CommandText = QueryC386921.CommandText & " " & "update WF_CLIENTES"
        QueryC386921.CommandText = QueryC386921.CommandText & " " & "set Cod_Repres = " & _ 
ConvertValue(P63499, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC386921.CommandText = QueryC386921.CommandText & " " & "where Cod_Cliente = " & _ 
ConvertValue(P63498, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC386921.Transaction = txn
        C386921 = QueryC386921.ExecuteNonQuery()
        C386921DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C573162

    Comp_C386923:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C386923 As DataTable = New DataTable()
        tb_C386923.Columns.Add("clienteDisponivel" & "")
        Dim row_C386923 As DataRow = tb_C386923.NewRow()
        row_C386923(0) = C386920
        tb_C386923.Rows.Add(row_C386923)
        R17526 = tb_C386923
        ReDim C386923DataType(1)
        C386923DataType(1) = C386920DataType
        ReturnDataType = C386923DataType

        GoTo Exit_R17526

    Comp_C390163:

        ' upMetas
        sCurrComponent = "upMetas"
        QueryC390163 = con.CreateCommand()
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "update wf_prev_vendas_prod set wf_prev_vendas_prod.cod_repres = " & _ 
ConvertValue(P63499, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "where wf_prev_vendas_prod.seq in ("
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "select wf_prev_vendas_prod.seq"
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "from wf_previsao_vendas, wf_prev_vendas_prod"
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "where wf_previsao_vendas.cod_previsao_venda = wf_prev_vendas_prod.cod_previsao_venda"
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "and wf_previsao_vendas.previsao_anual_repres = 1"
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "and sysdate between wf_previsao_vendas.data_inicial and wf_previsao_vendas.data_final"
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "and wf_previsao_vendas.cod_pessoa_estab_juridico = " & _ 
ConvertValue(P63925, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "and wf_prev_vendas_prod.cod_cliente = " & _ 
ConvertValue(P63498, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC390163.CommandText = QueryC390163.CommandText & " " & "and wf_prev_vendas_prod.Cod_Repres = (select max(cod_repres) from wf_representante where representante_interno = 1))"
        QueryC390163.Transaction = txn
        C390163 = QueryC390163.ExecuteNonQuery()
        C390163DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C391071

    Comp_C391071:

        ' delete
        sCurrComponent = "delete"
        QueryC391071 = con.CreateCommand()
        QueryC391071.CommandText = QueryC391071.CommandText & " " & "delete from WF_CLIENTES_DISPONIVEIS"
        QueryC391071.CommandText = QueryC391071.CommandText & " " & "where Cod_Cliente = " & _ 
ConvertValue(P63498, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC391071.Transaction = txn
        C391071 = QueryC391071.ExecuteNonQuery()
        C391071DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C386923

    Comp_C573153:

        ' INS1
        sCurrComponent = "INS1"
        QueryC573153 = con.CreateCommand()
        QueryC573153.CommandText = QueryC573153.CommandText & " " & "insert into wf_representante_cliente (Cod_Repres,Cod_Cliente,Representante_Master) values (" & _ 
ConvertValue(P63499, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(P63498, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", 1) "
        QueryC573153.Transaction = txn
        C573153 = QueryC573153.ExecuteNonQuery()
        C573153DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C390163

    Comp_C573162:

        ' Count
        sCurrComponent = "Count"
        QueryC573162 = con.CreateCommand()
        QueryC573162.CommandText = QueryC573162.CommandText & " " & "select nvl(count(*),0) "
        QueryC573162.CommandText = QueryC573162.CommandText & " " & "from  wf_representante_cliente r "
        QueryC573162.CommandText = QueryC573162.CommandText & " " & "  where r.Cod_Repres = " & _ 
ConvertValue(P63499, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " and  r.Cod_Cliente = " & _ 
ConvertValue(P63498, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC573162.Transaction = txn
        RsC573162 = QueryC573162.ExecuteReader()
        Dim iC573162 As Short
        ReDim C573162DataType(RsC573162.FieldCount)
        For iC573162 = 0 to RsC573162.FieldCount - 1
            Select Case RsC573162.GetDataTypeName(iC573162).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C573162DataType(iC573162 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C573162DataType(iC573162 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C573162DataType(iC573162 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC573162
        RsC573162_EOF = Not RsC573162.Read()

        GoTo Comp_C573163

    Comp_C573163:

        ' Count=0?
        sCurrComponent = "Count=0?"
        C573163 = (fn_ConvertValueCompiled(RsC573162(0), C573162DataType(1), AKB_DecimalPoint, False) = 0)
        C573163DataType = AKBTypeConst.cAKBTypeLogical
        If C573163 Then
            GoTo Comp_C573153
        Else
            GoTo Comp_C390163
        End If

    Exit_R17526:

        Exit Function

    Err_R17526:
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
