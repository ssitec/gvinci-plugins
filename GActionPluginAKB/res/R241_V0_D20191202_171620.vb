Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R241

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

    Public Shared Function R241(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable) As DataTable
    ' 
    ' 241 - Retorna C�digo Estabelecimento Default - Vers�o: 0
    ' 
        'On Error GoTo Err_R241
        Dim sCurrComponent as String

        Dim C865DataType() As Short
        Dim QueryC23984 As New Object
        Dim RsC23984 As New Object
        Dim C23984DataType() As Short
        Dim RsC23984_EOF As Boolean
        Dim C260424 As Object
        Dim C260424DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C260424

    Comp_C865:

        ' Ret1
        sCurrComponent = "Ret1"
        Dim tb_C865 As DataTable = New DataTable()
        tb_C865.Columns.Add("WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico" & "")
        tb_C865.Columns.Add("WF_PESSOAS.Nome_Pessoa" & "")
        tb_C865.Columns.Add("WF_ESTAB_JURIDICO.Contr_Pres_Fornec_TelaB" & "")
        tb_C865.Columns.Add("WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Escritorio" & "")
        tb_C865.Columns.Add("PE.Nome_Pessoa" & "")
        tb_C865.Columns.Add("WF_COMUNIC.Num_Comunicacao" & "")
        tb_C865.Columns.Add("1" & "")
        ReDim ReturnDatatype(RsC23984.FieldCount)
        Do Until RsC23984_EOF
            Dim row_C865 As DataRow = tb_C865.NewRow()
            Dim iC865 As Short
            For iC865 = 0 To 6
                row_C865(iC865) = RsC23984(iC865)
                ReturnDatatype(iC865 + 1) = C23984DataType(iC865)
            Next iC865
            tb_C865.Rows.Add(row_C865)
            RsC23984_EOF = Not RsC23984.Read()
        Loop
        R241 = tb_C865

        GoTo Exit_R241

    Comp_C23984:

        ' Sel_DadosEstab
        sCurrComponent = "Sel_DadosEstab"
        QueryC23984 = con.CreateCommand()
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "SELECT  WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_PESSOAS.Nome_Pessoa, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_ESTAB_JURIDICO.Contr_Pres_Fornec_TelaB,"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Escritorio, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "PE.Nome_Pessoa, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "NULL, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "1"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & ""
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "FROM WF_ESTAB_JURIDICO, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_PESSOAS, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_ESTAB_JURIDICO_PARAM,"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_PESSOAS PE, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_COMUNIC, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_PES_JURIDICA, WF_LOGINS"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & ""
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WHERE WF_ESTAB_JURIDICO.Estabelecimento_Default = 1 "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND WF_LOGINS.Nao_Usa_Estab_Default = 0"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_LOGINS.Login = " & _ 
ConvertValue(C260424, C260424DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_PESSOAS.Cod_Pessoa = WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_PES_JURIDICA.Cod_Pessoa (+) = WF_PESSOAS.Cod_Pessoa "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_PES_JURIDICA.Tp_Pessoa = 'J' "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  PE.Cod_Pessoa (+) = WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Escritorio"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_COMUNIC.Cod_Pessoa (+) = PE.Cod_Pessoa "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_COMUNIC.Cod_Comunic (+) = 2"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND Not Exists ("
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "	Select WF_LOGIN_X_ESTAB.Cod_Pessoa_Estab_Juridico "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "    from WF_LOGIN_X_ESTAB "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "    where WF_LOGIN_X_ESTAB.Login = " & _ 
ConvertValue(C260424, C260424DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "    and WF_LOGIN_X_ESTAB.Estab_Default = 1"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & ")"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & ""
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "Union all"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & ""
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "SELECT WF_LOGIN_X_ESTAB.Cod_Pessoa_Estab_Juridico, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_PESSOAS.Nome_Pessoa, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_ESTAB_JURIDICO.Contr_Pres_Fornec_TelaB,"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Escritorio, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "PE.Nome_Pessoa, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "NULL, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "1"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & ""
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "FROM WF_LOGIN_X_ESTAB, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_PESSOAS, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_ESTAB_JURIDICO, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_ESTAB_JURIDICO_PARAM,"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_PESSOAS PE, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_COMUNIC, "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_PES_JURIDICA,"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WF_LOGINS"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & ""
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "WHERE WF_LOGIN_X_ESTAB.Estab_Default = 1 "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND WF_LOGINS.Nao_Usa_Estab_Default = 0"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND WF_LOGIN_X_ESTAB.Login = " & _ 
ConvertValue(C260424, C260424DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND WF_LOGINS.Login = WF_LOGIN_X_ESTAB.Login"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_PESSOAS.Cod_Pessoa = WF_LOGIN_X_ESTAB.Cod_Pessoa_Estab_Juridico "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico = WF_LOGIN_X_ESTAB.Cod_Pessoa_Estab_Juridico"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_PES_JURIDICA.Cod_Pessoa (+) = WF_PESSOAS.Cod_Pessoa "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_PES_JURIDICA.Tp_Pessoa = 'J' "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  PE.Cod_Pessoa (+) = WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Escritorio"
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_COMUNIC.Cod_Pessoa (+) = PE.Cod_Pessoa "
        QueryC23984.CommandText = QueryC23984.CommandText & " " & "AND  WF_COMUNIC.Cod_Comunic (+) = 2"
        QueryC23984.Transaction = txn
        RsC23984 = QueryC23984.ExecuteReader()
        Dim iC23984 As Short
        ReDim C23984DataType(RsC23984.FieldCount)
        For iC23984 = 0 to RsC23984.FieldCount - 1
            Select Case RsC23984.GetDataTypeName(iC23984).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C23984DataType(iC23984 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C23984DataType(iC23984 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C23984DataType(iC23984 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC23984
        RsC23984_EOF = Not RsC23984.Read()

        GoTo Comp_C865

    Comp_C260424:

        ' Usuario
        sCurrComponent = "Usuario"
        C260424DataType = 3
        C260424 = Mid(con.ConnectionString, InStr(con.ConnectionString, "User Id=")+8, Instr(InStr(con.ConnectionString, "User Id="), con.ConnectionString, ";")-InStr(con.ConnectionString, "User Id=")-8)
        GoTo Comp_C23984

    Exit_R241:

        Exit Function

    Err_R241:
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
