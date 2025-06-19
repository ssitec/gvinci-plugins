Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R17756

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

    Public Shared Function R17756(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P64528 As Double) As DataTable
    ' 
    ' 17756 - Atualiza Coleção no Pré-Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R17756
        Dim sCurrComponent as String

        Dim QueryC394278 As New Object
        Dim RsC394278 As New Object
        Dim C394278DataType() As Short
        Dim RsC394278_EOF As Boolean
        Dim C394279 As Object
        Dim C394279DataType As Short
        Dim C394280 As Boolean
        Dim C394280DataType As Short
        Dim C394281 As Object
        Dim C394281DataType As Short
        Dim C394283DataType() As Short
        Dim C394285 As Object
        Dim C394285DataType As Short
        Dim C394286 As Boolean
        Dim C394286DataType As Short
        Dim C394287 As Object
        Dim C394287DataType As Short
        Dim QueryC394288 As New Object
        Dim RsC394288 As New Object
        Dim C394288DataType() As Short
        Dim RsC394288_EOF As Boolean
        Dim C394289 As Object
        Dim C394289DataType As Short
        Dim C394290 As Object
        Dim C394290DataType As Short
        Dim C394291 As Object
        Dim C394291DataType As Short
        Dim C394293 As Boolean
        Dim C394293DataType As Short
        Dim C394295 As Object
        Dim C394295DataType As Short
        Dim QueryC394296 As New Object
        Dim C394296 As Integer
        Dim C394296DataType As Short
        Dim C394298 As Object
        Dim C394298DataType As Short
        Dim C394299 As Object
        Dim C394299DataType As Short
        Dim C394300 As Object
        Dim C394300DataType As Short
        Dim C394301 As Object
        Dim C394301DataType As Short
        Dim C394302 As Object
        Dim C394302DataType As Short
        Dim C394303 As Object
        Dim C394303DataType As Short
        Dim QueryC394314 As New Object
        Dim C394314 As Integer
        Dim C394314DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C394281

    Comp_C394278:

        ' Sel_Itens_Pre_Pedido
        sCurrComponent = "Sel_Itens_Pre_Pedido"
        QueryC394278 = con.CreateCommand()
        QueryC394278.CommandText = QueryC394278.CommandText & " " & "SELECT NVL(WF_PRE_PEDIDO_ITENS.Id_Colecao, 0),"
        QueryC394278.CommandText = QueryC394278.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Seq_Item, "
        QueryC394278.CommandText = QueryC394278.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod, "
        QueryC394278.CommandText = QueryC394278.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso, "
        QueryC394278.CommandText = QueryC394278.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem "
        QueryC394278.CommandText = QueryC394278.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC394278.CommandText = QueryC394278.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P64528, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394278.Transaction = txn
        RsC394278 = QueryC394278.ExecuteReader()
        Dim iC394278 As Short
        ReDim C394278DataType(RsC394278.FieldCount)
        For iC394278 = 0 to RsC394278.FieldCount - 1
            Select Case RsC394278.GetDataTypeName(iC394278).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C394278DataType(iC394278 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C394278DataType(iC394278 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C394278DataType(iC394278 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC394278
        RsC394278_EOF = Not RsC394278.Read()

        GoTo Comp_C394279

    Comp_C394279:

        ' Fim_Itens
        sCurrComponent = "Fim_Itens"
        C394279DataType = 4
        C394279 = RsC394278_EOF
        GoTo Comp_C394280

    Comp_C394280:

        ' Fim_Itens = 1?
        sCurrComponent = "Fim_Itens = 1?"
        C394280 = (fn_ConvertValueCompiled(C394279, C394279DataType, AKB_DecimalPoint, False) = 1)
        C394280DataType = AKBTypeConst.cAKBTypeLogical
        If C394280 Then
            GoTo Comp_C394283
        Else
            GoTo Comp_C394285
        End If

    Comp_C394281:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C394281 = 1
        C394281DataType = 4
        GoTo Comp_C394298

    Comp_C394283:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C394283 As DataTable = New DataTable()
        tb_C394283.Columns.Add("vRetorno" & "")
        Dim row_C394283 As DataRow = tb_C394283.NewRow()
        row_C394283(0) = C394281
        tb_C394283.Rows.Add(row_C394283)
        R17756 = tb_C394283
        ReDim C394283DataType(1)
        C394283DataType(1) = C394281DataType
        ReturnDataType = C394283DataType

        GoTo Exit_R17756

    Comp_C394285:

        ' Id_Colecao
        sCurrComponent = "Id_Colecao"
        C394285DataType = 0
        C394285 = RsC394278(0)
        C394285DataType = C394278Datatype(1)
        If C394285DataType = AKBTypeConst.cAKBTypeString Then
          C394285 = IIF(IsDBNull(C394285), C394285, Strings.RTrim(C394285))
        End If 

        GoTo Comp_C394286

    Comp_C394286:

        ' Id_Colecao = 0?
        sCurrComponent = "Id_Colecao = 0?"
        C394286 = (fn_ConvertValueCompiled(C394285, C394285DataType, AKB_DecimalPoint, False) = 0)
        C394286DataType = AKBTypeConst.cAKBTypeLogical
        If C394286 Then
            GoTo Comp_C394287
        Else
            GoTo Comp_C394295
        End If

    Comp_C394287:

        ' Próximo_Item
        sCurrComponent = "Próximo_Item"
        C394287DataType = 4
        RsC394278_EOF = Not RsC394278.Read()
        C394287 = True

        GoTo Comp_C394279

    Comp_C394288:

        ' Verifica_Se_Produto_Tem_Coleção
        sCurrComponent = "Verifica_Se_Produto_Tem_Coleção"
        QueryC394288 = con.CreateCommand()
        QueryC394288.CommandText = QueryC394288.CommandText & " " & "SELECT COUNT(*) FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC394288.CommandText = QueryC394288.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = " & _ 
ConvertValue(C394289, C394289DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394288.CommandText = QueryC394288.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = " & _ 
ConvertValue(C394290, C394290DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " 	"
        QueryC394288.CommandText = QueryC394288.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = " & _ 
ConvertValue(C394291, C394291DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394288.CommandText = QueryC394288.CommandText & " " & "AND (WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL  "
        QueryC394288.CommandText = QueryC394288.CommandText & " " & "OR  WF_COLECAO_PRODUTOS.Data_Validade_Final >= SYSDATE)"
        QueryC394288.CommandText = QueryC394288.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= SYSDATE"
        QueryC394288.Transaction = txn
        RsC394288 = QueryC394288.ExecuteReader()
        Dim iC394288 As Short
        ReDim C394288DataType(RsC394288.FieldCount)
        For iC394288 = 0 to RsC394288.FieldCount - 1
            Select Case RsC394288.GetDataTypeName(iC394288).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C394288DataType(iC394288 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C394288DataType(iC394288 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C394288DataType(iC394288 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC394288
        RsC394288_EOF = Not RsC394288.Read()

        GoTo Comp_C394293

    Comp_C394289:

        ' Sigla_Prod
        sCurrComponent = "Sigla_Prod"
        C394289DataType = 0
        C394289DataType = C394278Datatype(3)
        If C394289DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394278(2)) Then
          C394289 = Strings.RTrim(RsC394278(2))
        Else 
          C394289 = RsC394278(2)
        End If 

        GoTo Comp_C394290

    Comp_C394290:

        ' Acesso
        sCurrComponent = "Acesso"
        C394290DataType = 0
        C394290DataType = C394278Datatype(4)
        If C394290DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394278(3)) Then
          C394290 = Strings.RTrim(RsC394278(3))
        Else 
          C394290 = RsC394278(3)
        End If 

        GoTo Comp_C394291

    Comp_C394291:

        ' Cod_Emb
        sCurrComponent = "Cod_Emb"
        C394291DataType = 0
        C394291DataType = C394278Datatype(5)
        If C394291DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394278(4)) Then
          C394291 = Strings.RTrim(RsC394278(4))
        Else 
          C394291 = RsC394278(4)
        End If 

        GoTo Comp_C394288

    Comp_C394293:

        ' Verifica_Se_Produto_Tem_Coleção = 0?
        sCurrComponent = "Verifica_Se_Produto_Tem_Coleção = 0?"
        C394293 = (fn_ConvertValueCompiled(RsC394288(0), C394288DataType(1), AKB_DecimalPoint, False) = 0)
        C394293DataType = AKBTypeConst.cAKBTypeLogical
        If C394293 Then
            GoTo Comp_C394314
        Else
            GoTo Comp_C394301
        End If

    Comp_C394295:

        ' Seq_Item
        sCurrComponent = "Seq_Item"
        C394295DataType = 0
        C394295DataType = C394278Datatype(2)
        If C394295DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394278(1)) Then
          C394295 = Strings.RTrim(RsC394278(1))
        Else 
          C394295 = RsC394278(1)
        End If 

        GoTo Comp_C394289

    Comp_C394296:

        ' Up_Pre_Pedido_Itens
        sCurrComponent = "Up_Pre_Pedido_Itens"
        QueryC394296 = con.CreateCommand()
        QueryC394296.CommandText = QueryC394296.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC394296.CommandText = QueryC394296.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.Sigla_Prod2 = " & _ 
ConvertValue(C394298, C394298DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  ,"
        QueryC394296.CommandText = QueryC394296.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso2 = " & _ 
ConvertValue(C394299, C394299DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  ,"
        QueryC394296.CommandText = QueryC394296.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem2 = " & _ 
ConvertValue(C394300, C394300DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC394296.CommandText = QueryC394296.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P64528, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394296.CommandText = QueryC394296.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C394295, C394295DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394296.Transaction = txn
        C394296 = QueryC394296.ExecuteNonQuery()
        C394296DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C394287

    Comp_C394298:

        ' vSigla_Prod
        sCurrComponent = "vSigla_Prod"
        C394298 = System.DBNull.Value
        C394298DataType = 3
        GoTo Comp_C394299

    Comp_C394299:

        ' vAcesso
        sCurrComponent = "vAcesso"
        C394299 = System.DBNull.Value
        C394299DataType = 1
        GoTo Comp_C394300

    Comp_C394300:

        ' vCod_Emb
        sCurrComponent = "vCod_Emb"
        C394300 = System.DBNull.Value
        C394300DataType = 1
        GoTo Comp_C394278

    Comp_C394301:

        ' AtrSiglaProd
        sCurrComponent = "AtrSiglaProd"
        C394301DataType = 4
        C394298 = fn_ConvertValueCompiled(C394289 , 3, AKB_DecimalPoint)
        C394301 = True
        GoTo Comp_C394302

    Comp_C394302:

        ' AtrAcesso
        sCurrComponent = "AtrAcesso"
        C394302DataType = 4
        C394299 = fn_ConvertValueCompiled(C394290 , 1, AKB_DecimalPoint)
        C394302 = True
        GoTo Comp_C394303

    Comp_C394303:

        ' AtrCodEmb
        sCurrComponent = "AtrCodEmb"
        C394303DataType = 4
        C394300 = fn_ConvertValueCompiled(C394291 , 1, AKB_DecimalPoint)
        C394303 = True
        GoTo Comp_C394296

    Comp_C394314:

        ' Up_Limpa_Coleção
        sCurrComponent = "Up_Limpa_Coleção"
        QueryC394314 = con.CreateCommand()
        QueryC394314.CommandText = QueryC394314.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC394314.CommandText = QueryC394314.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.Id_Colecao = NULL ,"
        QueryC394314.CommandText = QueryC394314.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod2 = NULL ,"
        QueryC394314.CommandText = QueryC394314.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso2 = NULL ,"
        QueryC394314.CommandText = QueryC394314.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem2 = NULL"
        QueryC394314.CommandText = QueryC394314.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P64528, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394314.CommandText = QueryC394314.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C394295, C394295DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394314.Transaction = txn
        C394314 = QueryC394314.ExecuteNonQuery()
        C394314DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C394287

    Exit_R17756:

        Exit Function

    Err_R17756:
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
