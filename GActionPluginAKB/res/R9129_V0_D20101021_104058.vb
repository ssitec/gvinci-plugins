Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R9129

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

    Public Shared Function R9129(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P28733 As Double, ByVal P28734 As Double, ByVal P28735 As Object) As DataTable
    ' 
    ' 9129 - Exclui pré-pedido item da carga de máquina - Versão: 0
    ' 
        'On Error GoTo Err_R9129
        Dim sCurrComponent as String

        Dim C151138 As Boolean
        Dim C151138DataType As Short
        Dim C151139DataType() As Short
        Dim C151141 As Boolean
        Dim C151141DataType As Short
        Dim C151142 As Boolean
        Dim C151142DataType As Short
        Dim C151143 As Boolean
        Dim C151143DataType As Short
        Dim C151144 As Object
        Dim C151144DataType As Short
        Dim QueryC151145 As New Object
        Dim C151145 As Integer
        Dim C151145DataType As Short
        Dim QueryC151146 As New Object
        Dim C151146 As Integer
        Dim C151146DataType As Short
        Dim QueryC151153 As New Object
        Dim C151153 As Integer
        Dim C151153DataType As Short
        Dim QueryC151161 As New Object
        Dim C151161 As Integer
        Dim C151161DataType As Short
        Dim QueryC151165 As New Object
        Dim C151165 As Integer
        Dim C151165DataType As Short
        P28735 = IIf(IsDBNull(P28735), 1, P28735)

        ReDim ReturnDatatype(0)

        GoTo Comp_C151144

    Comp_C151138:

        ' Begin
        sCurrComponent = "Begin"
        txn = con.BeginTransaction
        C151138 = True
        C151138DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C151146

    Comp_C151139:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C151139 As DataTable = New DataTable()
        tb_C151139.Columns.Add("vRetorno" & "")
        Dim row_C151139 As DataRow = tb_C151139.NewRow()
        row_C151139(0) = C151144
        tb_C151139.Rows.Add(row_C151139)
        R9129 = tb_C151139
        ReDim C151139DataType(1)
        C151139DataType(1) = C151144DataType
        ReturnDataType = C151139DataType

        GoTo Exit_R9129

    Comp_C151141:

        ' TemControleTransação?=1
        sCurrComponent = "TemControleTransação?=1"
        C151141 = (fn_ConvertValueCompiled(P28735, 4, AKB_DecimalPoint, False) = 1)
        C151141DataType = AKBTypeConst.cAKBTypeLogical
        If C151141 Then
            GoTo Comp_C151138
        Else
            GoTo Comp_C151146
        End If

    Comp_C151142:

        ' TemControleTransação?=1 (2)
        sCurrComponent = "TemControleTransação?=1 (2)"
        C151142 = (fn_ConvertValueCompiled(P28735, 4, AKB_DecimalPoint, False) = 1)
        C151142DataType = AKBTypeConst.cAKBTypeLogical
        If C151142 Then
            GoTo Comp_C151143
        Else
            GoTo Comp_C151139
        End If

    Comp_C151143:

        ' Commit
        sCurrComponent = "Commit"
        txn.Commit()
        C151143 = True
        C151143DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C151139

    Comp_C151144:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C151144 = 1
        C151144DataType = 4
        GoTo Comp_C151141

    Comp_C151145:

        ' EXCCarga maq x Pré Pedido Itens
        sCurrComponent = "EXCCarga maq x Pré Pedido Itens"
        QueryC151145 = con.CreateCommand()
        QueryC151145.CommandText = QueryC151145.CommandText & " " & "DELETE FROM AKBUSER01.WF_CARG_MAQ_PRE_PED_ITENS WHERE WF_CARG_MAQ_PRE_PED_ITENS.Id_PrePedido = " & _ 
ConvertValue(P28733, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_CARG_MAQ_PRE_PED_ITENS.Seq_Item = " & _ 
ConvertValue(P28734, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151145.Transaction = txn
        C151145 = QueryC151145.ExecuteNonQuery()
        C151145DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C151153

    Comp_C151146:

        ' ATUAL Horas Res Carga de máquina
        sCurrComponent = "ATUAL Horas Res Carga de máquina"
        QueryC151146 = con.CreateCommand()
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "UPDATE AKBUSER01.WF_CARGA_MAQUINA "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & ""
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "SET WF_CARGA_MAQUINA.Total_horas_res =  WF_CARGA_MAQUINA.Total_horas_res - "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & ""
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "(SELECT NVL(SUM(WF_CARG_MAQ_PRE_PED_ITENS.Horas_reservadas), 0)"
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "FROM AKBUSER01.WF_CARG_MAQ_PRE_PED_ITENS "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "WHERE WF_CARG_MAQ_PRE_PED_ITENS.Cod_Maq_Fase = WF_CARGA_MAQUINA.Cod_Maq_Fase "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "AND  WF_CARG_MAQ_PRE_PED_ITENS.Data_inicial = WF_CARGA_MAQUINA.Data_inicial"
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "AND  WF_CARG_MAQ_PRE_PED_ITENS.Data_final = WF_CARGA_MAQUINA.Data_final"
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "AND  WF_CARG_MAQ_PRE_PED_ITENS.Id_PrePedido = " & _ 
ConvertValue(P28733, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "AND  WF_CARG_MAQ_PRE_PED_ITENS.Seq_Item = " & _ 
ConvertValue(P28734, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC151146.CommandText = QueryC151146.CommandText & " " & ""
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "WHERE EXISTS "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "(SELECT *"
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "FROM AKBUSER01.WF_CARG_MAQ_PRE_PED_ITENS "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "WHERE WF_CARG_MAQ_PRE_PED_ITENS.Cod_Maq_Fase = WF_CARGA_MAQUINA.Cod_Maq_Fase "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "AND  WF_CARG_MAQ_PRE_PED_ITENS.Data_inicial = WF_CARGA_MAQUINA.Data_inicial"
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "AND  WF_CARG_MAQ_PRE_PED_ITENS.Data_final = WF_CARGA_MAQUINA.Data_final"
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "AND  WF_CARG_MAQ_PRE_PED_ITENS.Id_PrePedido = " & _ 
ConvertValue(P28733, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151146.CommandText = QueryC151146.CommandText & " " & "AND  WF_CARG_MAQ_PRE_PED_ITENS.Seq_Item = " & _ 
ConvertValue(P28734, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC151146.CommandText = QueryC151146.CommandText & " " & ""
        QueryC151146.Transaction = txn
        C151146 = QueryC151146.ExecuteNonQuery()
        C151146DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C151145

    Comp_C151153:

        ' ATUAL Horas Res Carga maq x Prod Ped Interno
        sCurrComponent = "ATUAL Horas Res Carga maq x Prod Ped Interno"
        QueryC151153 = con.CreateCommand()
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "UPDATE AKBUSER01.WF_CARGA_MAQ_PROD_PED_INT "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & ""
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "SET WF_CARGA_MAQ_PROD_PED_INT.Total_horas_res_ped_venda =  WF_CARGA_MAQ_PROD_PED_INT.Total_horas_res_ped_venda -"
        QueryC151153.CommandText = QueryC151153.CommandText & " " & ""
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "(SELECT NVL(SUM(WF_CARG_MAQ_PEDINT_PREPED.Horas_reservadas) ,0)"
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "FROM AKBUSER01.WF_CARG_MAQ_PEDINT_PREPED "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "WHERE WF_CARG_MAQ_PEDINT_PREPED.Id_PrePedido = " & _ 
ConvertValue(P28733, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Seq_Item = " & _ 
ConvertValue(P28734, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Cod_Maq_Fase = WF_CARGA_MAQ_PROD_PED_INT.Cod_Maq_Fase "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Data_inicial =  WF_CARGA_MAQ_PROD_PED_INT.Data_inicial "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Data_final = WF_CARGA_MAQ_PROD_PED_INT.Data_final  "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Sigla_Prod = WF_CARGA_MAQ_PROD_PED_INT.Sigla_Prod "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Acesso = WF_CARGA_MAQ_PROD_PED_INT.Acesso "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Cod_Embalagem =  WF_CARGA_MAQ_PROD_PED_INT.Cod_Embalagem)"
        QueryC151153.CommandText = QueryC151153.CommandText & " " & ""
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "WHERE EXISTS"
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "(SELECT * "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "FROM AKBUSER01.WF_CARG_MAQ_PEDINT_PREPED "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "WHERE WF_CARG_MAQ_PEDINT_PREPED.Id_PrePedido = " & _ 
ConvertValue(P28733, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Seq_Item = " & _ 
ConvertValue(P28734, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Cod_Maq_Fase = WF_CARGA_MAQ_PROD_PED_INT.Cod_Maq_Fase "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Data_inicial =  WF_CARGA_MAQ_PROD_PED_INT.Data_inicial "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Data_final = WF_CARGA_MAQ_PROD_PED_INT.Data_final  "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Sigla_Prod = WF_CARGA_MAQ_PROD_PED_INT.Sigla_Prod "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Acesso = WF_CARGA_MAQ_PROD_PED_INT.Acesso "
        QueryC151153.CommandText = QueryC151153.CommandText & " " & "AND  WF_CARG_MAQ_PEDINT_PREPED.Cod_Embalagem =  WF_CARGA_MAQ_PROD_PED_INT.Cod_Embalagem)"
        QueryC151153.Transaction = txn
        C151153 = QueryC151153.ExecuteNonQuery()
        C151153DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C151161

    Comp_C151161:

        ' EXCCarga maq ped int x Pre Ped It
        sCurrComponent = "EXCCarga maq ped int x Pre Ped It"
        QueryC151161 = con.CreateCommand()
        QueryC151161.CommandText = QueryC151161.CommandText & " " & "DELETE FROM AKBUSER01.WF_CARG_MAQ_PEDINT_PREPED WHERE WF_CARG_MAQ_PEDINT_PREPED.Id_PrePedido = " & _ 
ConvertValue(P28733, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_CARG_MAQ_PEDINT_PREPED.Seq_Item = " & _ 
ConvertValue(P28734, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151161.Transaction = txn
        C151161 = QueryC151161.ExecuteNonQuery()
        C151161DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C151165

    Comp_C151165:

        ' EXCPré-Ped ItensxCarg maq inicial
        sCurrComponent = "EXCPré-Ped ItensxCarg maq inicial"
        QueryC151165 = con.CreateCommand()
        QueryC151165.CommandText = QueryC151165.CommandText & " " & "DELETE FROM AKBUSER01.WF_PRE_PED_IT_CARG_MAQ_IN WHERE WF_PRE_PED_IT_CARG_MAQ_IN.Id_PrePedido = " & _ 
ConvertValue(P28733, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PRE_PED_IT_CARG_MAQ_IN.Seq_Item = " & _ 
ConvertValue(P28734, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC151165.Transaction = txn
        C151165 = QueryC151165.ExecuteNonQuery()
        C151165DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C151142

    Exit_R9129:

        Exit Function

    Err_R9129:
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
