Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R4717

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

    Public Shared Function R4717(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P13143 As Double, ByVal P13144 As Double, ByVal P13419 As Object, ByVal P13420 As Object, ByVal P31029 As Object, ByVal P33305 As Object) As DataTable
    ' 
    ' 4717 - Grava Alterações no Pré-Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R4717
        Dim sCurrComponent as String

        Dim QueryC57966 As New Object
        Dim C57966 As Integer
        Dim C57966DataType As Short
        Dim C57970 As Short
        Dim C57970DataType As Short
        Dim C57970ReturnDataType() As Short

        Dim C59463 As Boolean
        Dim C59463DataType As Short
        Dim C59464 As Boolean
        Dim C59464DataType As Short
        Dim QueryC59465 As New Object
        Dim C59465 As Integer
        Dim C59465DataType As Short
        Dim QueryC59466 As New Object
        Dim C59466 As Integer
        Dim C59466DataType As Short
        Dim C159437 As Short
        Dim C159437DataType As Short
        Dim C159437ReturnDataType() As Short

        Dim QueryC159439 As New Object
        Dim RsC159439 As New Object
        Dim C159439DataType() As Short
        Dim RsC159439_EOF As Boolean
        Dim C159440 As Boolean
        Dim C159440DataType As Short
        Dim C159445 As Boolean
        Dim C159445DataType As Short
        Dim QueryC159447 As New Object
        Dim C159447 As Integer
        Dim C159447DataType As Short
        Dim C159600DataType() As Short
        Dim C159601 As Object
        Dim C159601DataType As Short
        Dim C168286 As Object
        Dim C168286DataType As Short
        Dim C168287 As Object
        Dim C168287DataType As Short
        Dim C168292 As Boolean
        Dim C168292DataType As Short
        Dim QueryC184577 As New Object
        Dim C184577 As Integer
        Dim C184577DataType As Short
        Dim C579548 As Object
        Dim C579548DataType As Short
        Dim C579549 As Object
        Dim C579549DataType As Short
        P13419 = IIf(IsDBNull(P13419), 0, P13419)

        ReDim ReturnDatatype(0)

        GoTo Comp_C168286

    Comp_C57966:

        ' UpPréPedidoCliente
        sCurrComponent = "UpPréPedidoCliente"
        QueryC57966 = con.CreateCommand()
        QueryC57966.CommandText = QueryC57966.CommandText & " " & "UPDATE  AKBUSER01.WF_PRE_PEDIDO "
        QueryC57966.CommandText = QueryC57966.CommandText & " " & ""
        QueryC57966.CommandText = QueryC57966.CommandText & " " & "SET  WF_PRE_PEDIDO.Cod_Cliente = " & _ 
ConvertValue(P13144, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC57966.CommandText = QueryC57966.CommandText & " " & ""
        QueryC57966.CommandText = QueryC57966.CommandText & " " & "WHERE  WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P13143, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC57966.CommandText = QueryC57966.CommandText & " " & ""
        QueryC57966.CommandText = QueryC57966.CommandText & " " & ""
        QueryC57966.Transaction = txn
        C57966 = QueryC57966.ExecuteNonQuery()
        C57966DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C579548

    Comp_C57970:

        ' MsgOK.
        sCurrComponent = "MsgOK."
        Dim row_C57970 As DataRow = msg.NewRow()
        row_C57970(0) = "Dados atualizados." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C57970)
        C57970DataType = AKBTypeConst.cAKBTypeNumeric
        C57970 = 1

        GoTo Comp_C159600

    Comp_C59463:

        ' PagtoNulo
        sCurrComponent = "PagtoNulo"
        C59463 = (fn_ConvertValueCompiled(P13419, 1, AKB_DecimalPoint, False) = 0 OR fn_ConvertValueCompiled(C579548, C579548DataType, AKB_DecimalPoint, False) = 1)
        C59463DataType = AKBTypeConst.cAKBTypeLogical
        If C59463 Then
            GoTo Comp_C59464
        Else
            GoTo Comp_C59465
        End If

    Comp_C59464:

        ' DeptoNulo
        sCurrComponent = "DeptoNulo"
        C59464 = (fn_ConvertValueCompiled(P13420, 3, AKB_DecimalPoint, False) = "" OR fn_ConvertValueCompiled(C579549, C579549DataType, AKB_DecimalPoint, False) = 1)
        C59464DataType = AKBTypeConst.cAKBTypeLogical
        If C59464 Then
            GoTo Comp_C159439
        Else
            GoTo Comp_C59466
        End If

    Comp_C59465:

        ' UpPagto
        sCurrComponent = "UpPagto"
        QueryC59465 = con.CreateCommand()
        QueryC59465.CommandText = QueryC59465.CommandText & " " & "UPDATE  AKBUSER01.WF_PRE_PEDIDO "
        QueryC59465.CommandText = QueryC59465.CommandText & " " & ""
        QueryC59465.CommandText = QueryC59465.CommandText & " " & "SET  WF_PRE_PEDIDO.Cod_Pagto = " & _ 
ConvertValue(P13419, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC59465.CommandText = QueryC59465.CommandText & " " & ""
        QueryC59465.CommandText = QueryC59465.CommandText & " " & "WHERE  WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P13143, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC59465.CommandText = QueryC59465.CommandText & " " & ""
        QueryC59465.CommandText = QueryC59465.CommandText & " " & ""
        QueryC59465.Transaction = txn
        C59465 = QueryC59465.ExecuteNonQuery()
        C59465DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C59464

    Comp_C59466:

        ' UpPréPedDepto
        sCurrComponent = "UpPréPedDepto"
        QueryC59466 = con.CreateCommand()
        QueryC59466.CommandText = QueryC59466.CommandText & " " & "UPDATE  AKBUSER01.WF_PRE_PEDIDO "
        QueryC59466.CommandText = QueryC59466.CommandText & " " & ""
        QueryC59466.CommandText = QueryC59466.CommandText & " " & "SET  WF_PRE_PEDIDO.Departamento = " & _ 
ConvertValue(P13420, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC59466.CommandText = QueryC59466.CommandText & " " & ""
        QueryC59466.CommandText = QueryC59466.CommandText & " " & "WHERE  WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P13143, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC59466.CommandText = QueryC59466.CommandText & " " & ""
        QueryC59466.CommandText = QueryC59466.CommandText & " " & ""
        QueryC59466.Transaction = txn
        C59466 = QueryC59466.ExecuteNonQuery()
        C59466DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C159439

    Comp_C159437:

        ' MsgMesmoDepto
        sCurrComponent = "MsgMesmoDepto"
        Dim row_C159437 As DataRow = msg.NewRow()
        row_C159437(0) = "Desejá gravar o mesmo Depto para a Observação ?" & Chr(13) & Chr(10) & "Sim" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C159437)
        C159437DataType = AKBTypeConst.cAKBTypeNumeric
        C159437 = 6

        GoTo Comp_C159445

    Comp_C159439:

        ' Depto_OBS_Antigo
        sCurrComponent = "Depto_OBS_Antigo"
        QueryC159439 = con.CreateCommand()
        QueryC159439.CommandText = QueryC159439.CommandText & " " & "SELECT WF_DEPARTAMENTO.Departamento , WF_DEPARTAMENTO.Desc_Departamento , "
        QueryC159439.CommandText = QueryC159439.CommandText & " " & "                  WF_PRE_PEDIDO_OBS.Obs, WF_PRE_PEDIDO_OBS.Destinatario , "
        QueryC159439.CommandText = QueryC159439.CommandText & " " & "                  WF_PRE_PEDIDO_OBS.DATA_GERACAO , WF_PRE_PEDIDO_OBS.USUARIO_GERACAO "
        QueryC159439.CommandText = QueryC159439.CommandText & " " & "FROM AKBUSER01.WF_DEPARTAMENTO , AKBUSER01.WF_PRE_PEDIDO_OBS "
        QueryC159439.CommandText = QueryC159439.CommandText & " " & "WHERE WF_PRE_PEDIDO_OBS.Id_PrePedido = " & _ 
ConvertValue(P13143, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC159439.CommandText = QueryC159439.CommandText & " " & "AND  WF_PRE_PEDIDO_OBS.Departamento = WF_DEPARTAMENTO.Departamento"
        QueryC159439.Transaction = txn
        RsC159439 = QueryC159439.ExecuteReader()
        Dim iC159439 As Short
        ReDim C159439DataType(RsC159439.FieldCount)
        For iC159439 = 0 to RsC159439.FieldCount - 1
            Select Case RsC159439.GetDataTypeName(iC159439).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C159439DataType(iC159439 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C159439DataType(iC159439 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C159439DataType(iC159439 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC159439
        RsC159439_EOF = Not RsC159439.Read()

        GoTo Comp_C168292

    Comp_C159440:

        ' Depto_OBS_Antigo = Depto
        sCurrComponent = "Depto_OBS_Antigo = Depto"
        C159440 = (fn_ConvertValueCompiled(RsC159439(0), C159439DataType(1), AKB_DecimalPoint, False) = fn_ConvertValueCompiled(P13420, 3, AKB_DecimalPoint, False))
        C159440DataType = AKBTypeConst.cAKBTypeLogical
        If C159440 Then
            GoTo Comp_C184577
        Else
            GoTo Comp_C159437
        End If

    Comp_C159445:

        ' MSG2 = 6
        sCurrComponent = "MSG2 = 6"
        C159445 = (fn_ConvertValueCompiled(C159437, C159437DataType, AKB_DecimalPoint, False) = 6)
        C159445DataType = AKBTypeConst.cAKBTypeLogical
        If C159445 Then
            GoTo Comp_C168287
        Else
            GoTo Comp_C159447
        End If

    Comp_C159447:

        ' UpPréPedObs
        sCurrComponent = "UpPréPedObs"
        QueryC159447 = con.CreateCommand()
        QueryC159447.CommandText = QueryC159447.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_OBS "
        QueryC159447.CommandText = QueryC159447.CommandText & " " & ""
        QueryC159447.CommandText = QueryC159447.CommandText & " " & "SET WF_PRE_PEDIDO_OBS.Departamento = " & _ 
ConvertValue(C168286, C168286DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC159447.CommandText = QueryC159447.CommandText & " " & "WF_PRE_PEDIDO_OBS.Obs =" & _ 
ConvertValue(P33305, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC159447.CommandText = QueryC159447.CommandText & " " & ""
        QueryC159447.CommandText = QueryC159447.CommandText & " " & "WHERE WF_PRE_PEDIDO_OBS.Id_PrePedido = " & _ 
ConvertValue(P13143, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC159447.Transaction = txn
        C159447 = QueryC159447.ExecuteNonQuery()
        C159447DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C57970

    Comp_C159600:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C159600 As DataTable = New DataTable()
        tb_C159600.Columns.Add("Atualizar" & "")
        Dim row_C159600 As DataRow = tb_C159600.NewRow()
        row_C159600(0) = C159601
        tb_C159600.Rows.Add(row_C159600)
        R4717 = tb_C159600
        ReDim C159600DataType(1)
        C159600DataType(1) = C159601DataType
        ReturnDataType = C159600DataType

        GoTo Exit_R4717

    Comp_C159601:

        ' Atualizar
        sCurrComponent = "Atualizar"
        C159601 = 1
        C159601DataType = 4
        GoTo Comp_C57966

    Comp_C168286:

        ' Depto_OBS
        sCurrComponent = "Depto_OBS"
        C168286 = P31029 & " "
        C168286DataType = 3
        GoTo Comp_C159601

    Comp_C168287:

        ' Atv_Depto
        sCurrComponent = "Atv_Depto"
        C168287DataType = 4
        C168286 = fn_ConvertValueCompiled(P13420 , 3, AKB_DecimalPoint)
        C168287 = True
        GoTo Comp_C159447

    Comp_C168292:

        ' Depto_OBS_Antigo = Depto_OBS
        sCurrComponent = "Depto_OBS_Antigo = Depto_OBS"
        C168292 = (fn_ConvertValueCompiled(RsC159439(0), C159439DataType(1), AKB_DecimalPoint, False) <> fn_ConvertValueCompiled(P31029, 3, AKB_DecimalPoint, False))
        C168292DataType = AKBTypeConst.cAKBTypeLogical
        If C168292 Then
            GoTo Comp_C159447
        Else
            GoTo Comp_C159440
        End If

    Comp_C184577:

        ' UpPréPedObs2
        sCurrComponent = "UpPréPedObs2"
        QueryC184577 = con.CreateCommand()
        QueryC184577.CommandText = QueryC184577.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_OBS "
        QueryC184577.CommandText = QueryC184577.CommandText & " " & ""
        QueryC184577.CommandText = QueryC184577.CommandText & " " & "SET  WF_PRE_PEDIDO_OBS.Obs =" & _ 
ConvertValue(P33305, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC184577.CommandText = QueryC184577.CommandText & " " & ""
        QueryC184577.CommandText = QueryC184577.CommandText & " " & "WHERE WF_PRE_PEDIDO_OBS.Id_PrePedido = " & _ 
ConvertValue(P13143, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC184577.Transaction = txn
        C184577 = QueryC184577.ExecuteNonQuery()
        C184577DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C57970

    Comp_C579548:

        ' ÉNULO_Pagto
        sCurrComponent = "ÉNULO_Pagto"
        C579548DataType = 4
        C579548 = IsDBNull (P13419)
        GoTo Comp_C579549

    Comp_C579549:

        ' ÉNULO_Depto
        sCurrComponent = "ÉNULO_Depto"
        C579549DataType = 4
        C579549 = IsDBNull (P13420)
        GoTo Comp_C59463

    Exit_R4717:

        Exit Function

    Err_R4717:
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
