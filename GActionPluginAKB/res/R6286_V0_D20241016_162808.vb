Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R6286

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

    Public Shared Function R6286(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P18430 As Object, ByVal P18432 As Object) As DataTable
    ' 
    ' 6286 - Validação do Acesso do Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R6286
        Dim sCurrComponent as String

        Dim C85288 As Object
        Dim C85288DataType As Short
        Dim QueryC85289 As New Object
        Dim RsC85289 As New Object
        Dim C85289DataType() As Short
        Dim RsC85289_EOF As Boolean
        Dim C85290 As Boolean
        Dim C85290DataType As Short
        Dim C85292 As Boolean
        Dim C85292DataType As Short
        Dim C85293 As Short
        Dim C85293DataType As Short
        Dim C85293ReturnDataType() As Short

        Dim C85294 As Short
        Dim C85294DataType As Short
        Dim C85294ReturnDataType() As Short

        Dim C85305DataType() As Short
        Dim C85309 As Object
        Dim C85309DataType As Short
        Dim C85310 As Boolean
        Dim C85310DataType As Short
        Dim QueryC85342 As New Object
        Dim RsC85342 As New Object
        Dim C85342DataType() As Short
        Dim RsC85342_EOF As Boolean
        Dim C85344 As Object
        Dim C85344DataType As Short
        Dim C85347 As Object
        Dim C85347DataType As Short
        Dim C85348 As Object
        Dim C85348DataType As Short
        Dim QueryC85356 As New Object
        Dim RsC85356 As New Object
        Dim C85356DataType() As Short
        Dim RsC85356_EOF As Boolean
        Dim C85357 As Boolean
        Dim C85357DataType As Short
        Dim C85359 As Short
        Dim C85359DataType As Short
        Dim C85359ReturnDataType() As Short

        Dim C85361 As Boolean
        Dim C85361DataType As Short
        Dim C85373 As Boolean
        Dim C85373DataType As Short
        Dim C85376 As Short
        Dim C85376DataType As Short
        Dim C85376ReturnDataType() As Short

        Dim C85380 As Boolean
        Dim C85380DataType As Short
        Dim C85382 As Object
        Dim C85382DataType As Short
        Dim C85383 As Object
        Dim C85383DataType As Short
        Dim C85384 As Object
        Dim C85384DataType As Short
        Dim C85385 As Boolean
        Dim C85385DataType As Short
        Dim C85388DataType() As Short
        Dim C85394 As Object
        Dim C85394DataType As Short
        Dim C85395 As Object
        Dim C85395DataType As Short
        Dim C85396 As Object
        Dim C85396DataType As Short
        Dim C191250 As Boolean
        Dim C191250DataType As Short
        Dim C191251 As Boolean
        Dim C191251DataType As Short
        Dim C191285 As DataTable
        Dim C191285CurrentRow As Integer
        Dim C191285DataType() As Short
        Dim C191286 As DataTable
        Dim C191286CurrentRow As Integer
        Dim C191286DataType() As Short
        Dim C191330 As Object
        Dim C191330DataType As Short
        Dim C191332 As Object
        Dim C191332DataType As Short
        Dim C191335 As Object
        Dim C191335DataType As Short
        Dim C296998 As Object
        Dim C296998DataType As Short
        Dim C297000 As Boolean
        Dim C297000DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C85288

    Comp_C85288:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C85288 = 1
        C85288DataType = 4
        GoTo Comp_C191330

    Comp_C85289:

        ' Sel_TpProduto
        sCurrComponent = "Sel_TpProduto"
        QueryC85289 = con.CreateCommand()
        QueryC85289.CommandText = QueryC85289.CommandText & " " & "Select NVL(Barra_Ac_Inativo,0)"
        QueryC85289.CommandText = QueryC85289.CommandText & " " & "from AkbUser01.WF_Tp_Produto"
        QueryC85289.CommandText = QueryC85289.CommandText & " " & "where Tp_Produto = " & _ 
ConvertValue(P18432, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC85289.Transaction = txn
        RsC85289 = QueryC85289.ExecuteReader()
        Dim iC85289 As Short
        ReDim C85289DataType(RsC85289.FieldCount)
        For iC85289 = 0 to RsC85289.FieldCount - 1
            Select Case RsC85289.GetDataTypeName(iC85289).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C85289DataType(iC85289 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C85289DataType(iC85289 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C85289DataType(iC85289 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC85289
        RsC85289_EOF = Not RsC85289.Read()

        GoTo Comp_C85357

    Comp_C85290:

        ' Sel_Inativo=1
        sCurrComponent = "Sel_Inativo=1"
        C85290 = (fn_ConvertValueCompiled(RsC85342(0), C85342DataType(1), AKB_DecimalPoint, False) = 1)
        C85290DataType = AKBTypeConst.cAKBTypeLogical
        If C85290 Then
            GoTo Comp_C191335
        Else
            GoTo Comp_C191250
        End If

    Comp_C85292:

        ' Sel_TpProduto=1
        sCurrComponent = "Sel_TpProduto=1"
        C85292 = (fn_ConvertValueCompiled(RsC85289(0), C85289DataType(1), AKB_DecimalPoint, False) = 1)
        C85292DataType = AKBTypeConst.cAKBTypeLogical
        If C85292 Then
            GoTo Comp_C85293
        Else
            GoTo Comp_C85294
        End If

    Comp_C85293:

        ' MsgBarra
        sCurrComponent = "MsgBarra"
        C85293 = 1
        C85293DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C85309

    Comp_C85294:

        ' MsgPergunta
        sCurrComponent = "MsgPergunta"
        C85294 = 6
        C85294DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C85310

    Comp_C85305:

        ' Fim
        sCurrComponent = "Fim"
        Dim tb_C85305 As DataTable = New DataTable()
        tb_C85305.Columns.Add("vRetorno" & "")
        Dim row_C85305 As DataRow = tb_C85305.NewRow()
        row_C85305(0) = C85288
        tb_C85305.Rows.Add(row_C85305)
        R6286 = tb_C85305
        ReDim C85305DataType(1)
        C85305DataType(1) = C85288DataType
        ReturnDataType = C85305DataType

        GoTo Exit_R6286

    Comp_C85309:

        ' AtribBarra
        sCurrComponent = "AtribBarra"
        C85309DataType = 4
        C85288 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C85309 = True
        GoTo Comp_C191250

    Comp_C85310:

        ' MgsPerg=Sim
        sCurrComponent = "MgsPerg=Sim"
        C85310 = (fn_ConvertValueCompiled(C85294, C85294DataType, AKB_DecimalPoint, False) = 6)
        C85310DataType = AKBTypeConst.cAKBTypeLogical
        If C85310 Then
            GoTo Comp_C191250
        Else
            GoTo Comp_C85309
        End If

    Comp_C85342:

        ' Sel_Inativo
        sCurrComponent = "Sel_Inativo"
        QueryC85342 = con.CreateCommand()
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "Select NVL( WF_Acessos.Produto_Inativo,0), WF_Pedido_Itens.Sigla_Prod, WF_Pedido_Itens.Acesso,                                        WF_Pedido_Itens.Cod_Embalagem"
        QueryC85342.CommandText = QueryC85342.CommandText & " " & ""
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "from AkbUser01.WF_Acessos, AkbUser01.WF_Pedido_Itens, AkbUser01.WF_Tipo_Ped"
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "where WF_Pedido_Itens.Pedido = " & _ 
ConvertValue(P18430, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "and WF_Pedido_Itens.Tp_Produto = " & _ 
ConvertValue(P18432, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC85342.CommandText = QueryC85342.CommandText & " " & ""
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "and WF_Tipo_Ped.Tipo_Ped = WF_Pedido_Itens.Tipo_Ped"
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "and WF_Tipo_Ped.Permite_Prod_Inativo = 0"
        QueryC85342.CommandText = QueryC85342.CommandText & " " & ""
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "and WF_Acessos.Acesso = WF_Pedido_Itens.Acesso"
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "and WF_Acessos.Sigla_Prod = WF_Pedido_Itens.Sigla_Prod"
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "and NVL(Produto_Inativo,0) = 1"
        QueryC85342.CommandText = QueryC85342.CommandText & " " & "and WF_Pedido_Itens.Flag_Pos_Ped not in (8, 299, 10, 12, 13)  "
        QueryC85342.Transaction = txn
        RsC85342 = QueryC85342.ExecuteReader()
        Dim iC85342 As Short
        ReDim C85342DataType(RsC85342.FieldCount)
        For iC85342 = 0 to RsC85342.FieldCount - 1
            Select Case RsC85342.GetDataTypeName(iC85342).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C85342DataType(iC85342 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C85342DataType(iC85342 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C85342DataType(iC85342 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC85342
        RsC85342_EOF = Not RsC85342.Read()

        GoTo Comp_C296998

    Comp_C85344:

        ' Sigla
        sCurrComponent = "Sigla"
        C85344DataType = 0
        C85344DataType = C85342Datatype(2)
        If C85344DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC85342(1)) Then
          C85344 = Strings.RTrim(RsC85342(1))
        Else 
          C85344 = RsC85342(1)
        End If 

        GoTo Comp_C85347

    Comp_C85347:

        ' Acesso
        sCurrComponent = "Acesso"
        C85347DataType = 0
        C85347DataType = C85342Datatype(3)
        If C85347DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC85342(2)) Then
          C85347 = Strings.RTrim(RsC85342(2))
        Else 
          C85347 = RsC85342(2)
        End If 

        GoTo Comp_C85348

    Comp_C85348:

        ' Emb
        sCurrComponent = "Emb"
        C85348DataType = 0
        C85348DataType = C85342Datatype(4)
        If C85348DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC85342(3)) Then
          C85348 = Strings.RTrim(RsC85342(3))
        Else 
          C85348 = RsC85342(3)
        End If 

        GoTo Comp_C85373

    Comp_C85356:

        ' Sel_Count
        sCurrComponent = "Sel_Count"
        QueryC85356 = con.CreateCommand()
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "Select count(*)"
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "from AkbUser01.WF_Acessos, AkbUser01.WF_Pedido_Itens, AkbUser01.WF_Tipo_Ped"
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "where WF_Pedido_Itens.Pedido = " & _ 
ConvertValue(P18430, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "and WF_Pedido_Itens.Tp_Produto = " & _ 
ConvertValue(P18432, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "and WF_Acessos.Acesso = WF_Pedido_Itens.Acesso"
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "and WF_Acessos.Sigla_Prod = WF_Pedido_Itens.Sigla_Prod"
        QueryC85356.CommandText = QueryC85356.CommandText & " " & ""
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "and WF_Tipo_Ped.Tipo_Ped = WF_Pedido_Itens.Tipo_Ped"
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "and WF_Tipo_Ped.Permite_Prod_Inativo = 0"
        QueryC85356.CommandText = QueryC85356.CommandText & " " & ""
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "and NVL(Produto_Inativo,0) = 1 "
        QueryC85356.CommandText = QueryC85356.CommandText & " " & "and WF_Pedido_Itens.Flag_Pos_Ped not in (8, 299, 10, 12, 13) "
        QueryC85356.Transaction = txn
        RsC85356 = QueryC85356.ExecuteReader()
        Dim iC85356 As Short
        ReDim C85356DataType(RsC85356.FieldCount)
        For iC85356 = 0 to RsC85356.FieldCount - 1
            Select Case RsC85356.GetDataTypeName(iC85356).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C85356DataType(iC85356 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C85356DataType(iC85356 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C85356DataType(iC85356 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC85356
        RsC85356_EOF = Not RsC85356.Read()

        GoTo Comp_C85342

    Comp_C85357:

        ' Sel_Count>1
        sCurrComponent = "Sel_Count>1"
        C85357 = (fn_ConvertValueCompiled(RsC85356(0), C85356DataType(1), AKB_DecimalPoint, False) > 1)
        C85357DataType = AKBTypeConst.cAKBTypeLogical
        If C85357 Then
            GoTo Comp_C191332
        Else
            GoTo Comp_C297000
        End If

    Comp_C85359:

        ' MSG1
        sCurrComponent = "MSG1"
        C85359 = 6
        C85359DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C85361

    Comp_C85361:

        ' MSG=Sim
        sCurrComponent = "MSG=Sim"
        C85361 = (fn_ConvertValueCompiled(C85359, C85359DataType, AKB_DecimalPoint, False) = 6)
        C85361DataType = AKBTypeConst.cAKBTypeLogical
        If C85361 Then
            GoTo Comp_C85383
        Else
            GoTo Comp_C297000
        End If

    Comp_C85373:

        ' TpProd=1
        sCurrComponent = "TpProd=1"
        C85373 = (fn_ConvertValueCompiled(RsC85289(0), C85289DataType(1), AKB_DecimalPoint, False) = 1)
        C85373DataType = AKBTypeConst.cAKBTypeLogical
        If C85373 Then
            GoTo Comp_C85293
        Else
            GoTo Comp_C85376
        End If

    Comp_C85376:

        ' MsgPerg
        sCurrComponent = "MsgPerg"
        C85376 = 7
        C85376DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C85380

    Comp_C85380:

        ' MsgPerg=Sim
        sCurrComponent = "MsgPerg=Sim"
        C85380 = (fn_ConvertValueCompiled(C85376, C85376DataType, AKB_DecimalPoint, False) = 6)
        C85380DataType = AKBTypeConst.cAKBTypeLogical
        If C85380 Then
            GoTo Comp_C85382
        Else
            GoTo Comp_C85309
        End If

    Comp_C85382:

        ' Next
        sCurrComponent = "Next"
        C85382DataType = 4
        RsC85342_EOF = Not RsC85342.Read()
        C85382 = True

        GoTo Comp_C85384

    Comp_C85383:

        ' BOF
        sCurrComponent = "BOF"
        C85383DataType = 4

        GoTo Comp_C85384

    Comp_C85384:

        ' EOF
        sCurrComponent = "EOF"
        C85384DataType = 4
        C85384 = RsC85342_EOF
        GoTo Comp_C85385

    Comp_C85385:

        ' EOF=1
        sCurrComponent = "EOF=1"
        C85385 = (fn_ConvertValueCompiled(C85384, C85384DataType, AKB_DecimalPoint, False) = 1)
        C85385DataType = AKBTypeConst.cAKBTypeLogical
        If C85385 Then
            GoTo Comp_C191251
        Else
            GoTo Comp_C85344
        End If

    Comp_C85388:

        ' Fim2
        sCurrComponent = "Fim2"
        Dim tb_C85388 As DataTable = New DataTable()
        tb_C85388.Columns.Add("vRetorno" & "")
        Dim row_C85388 As DataRow = tb_C85388.NewRow()
        row_C85388(0) = C85288
        tb_C85388.Rows.Add(row_C85388)
        R6286 = tb_C85388
        ReDim C85388DataType(1)
        C85388DataType(1) = C85288DataType
        ReturnDataType = C85388DataType

        GoTo Exit_R6286

    Comp_C85394:

        ' Embalagem
        sCurrComponent = "Embalagem"
        C85394DataType = 0
        C85394DataType = C85342Datatype(4)
        If C85394DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC85342(3)) Then
          C85394 = Strings.RTrim(RsC85342(3))
        Else 
          C85394 = RsC85342(3)
        End If 

        GoTo Comp_C85292

    Comp_C85395:

        ' Sigla2
        sCurrComponent = "Sigla2"
        C85395DataType = 0
        C85395DataType = C85342Datatype(2)
        If C85395DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC85342(1)) Then
          C85395 = Strings.RTrim(RsC85342(1))
        Else 
          C85395 = RsC85342(1)
        End If 

        GoTo Comp_C85396

    Comp_C85396:

        ' Acesso2
        sCurrComponent = "Acesso2"
        C85396DataType = 0
        C85396DataType = C85342Datatype(3)
        If C85396DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC85342(2)) Then
          C85396 = Strings.RTrim(RsC85342(2))
        Else 
          C85396 = RsC85342(2)
        End If 

        GoTo Comp_C85394

    Comp_C191250:

        ' Gravar Obs Prod Inativo?
        sCurrComponent = "Gravar Obs Prod Inativo?"
        C191250 = (fn_ConvertValueCompiled(C191330, C191330DataType, AKB_DecimalPoint, False) = 1 AND fn_ConvertValueCompiled(C85288, C85288DataType, AKB_DecimalPoint, False) = 1)
        C191250DataType = AKBTypeConst.cAKBTypeLogical
        If C191250 Then
            GoTo Comp_C191285
        Else
            GoTo Comp_C85305
        End If

    Comp_C191251:

        ' Gravar Obs Prod Inat?
        sCurrComponent = "Gravar Obs Prod Inat?"
        C191251 = (fn_ConvertValueCompiled(C191330, C191330DataType, AKB_DecimalPoint, False) =  1 AND fn_ConvertValueCompiled(C85288, C85288DataType, AKB_DecimalPoint, False) = 1)
        C191251DataType = AKBTypeConst.cAKBTypeLogical
        If C191251 Then
            GoTo Comp_C191286
        Else
            GoTo Comp_C85388
        End If

    Comp_C191285:

        ' Gravar Obs
        sCurrComponent = "Gravar Obs"
        C191285 = clsRuleDynamicallyCompiled_R10618.R10618(con, txn, IIf(Not IsDBNull(P18432), P18432, System.DBNull.Value), IIf(Not IsDBNull(P18430), P18430, System.DBNull.Value))
        C191285CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C191285) Then
          iColumns = C191285.Columns.Count
        End If
        ReDim C191285DataType(iColumns)
        For iCol = 1 To iColumns
            C191285DataType(iCol) = clsRuleDynamicallyCompiled_R10618.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C85305

    Comp_C191286:

        ' Gravar Obs Prod Inat
        sCurrComponent = "Gravar Obs Prod Inat"
        C191286 = clsRuleDynamicallyCompiled_R10618.R10618(con, txn, IIf(Not IsDBNull(P18432), P18432, System.DBNull.Value), IIf(Not IsDBNull(P18430), P18430, System.DBNull.Value))
        C191286CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C191286) Then
          iColumns = C191286.Columns.Count
        End If
        ReDim C191286DataType(iColumns)
        For iCol = 1 To iColumns
            C191286DataType(iCol) = clsRuleDynamicallyCompiled_R10618.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C85388

    Comp_C191330:

        ' vControlInat
        sCurrComponent = "vControlInat"
        C191330 = 0
        C191330DataType = 1
        GoTo Comp_C85356

    Comp_C191332:

        ' vControlInat=1
        sCurrComponent = "vControlInat=1"
        C191332DataType = 4
        C191330 = fn_ConvertValueCompiled(1, 1, AKB_DecimalPoint)
        C191332 = True
        GoTo Comp_C85359

    Comp_C191335:

        ' vControlInativo=1
        sCurrComponent = "vControlInativo=1"
        C191335DataType = 4
        C191330 = fn_ConvertValueCompiled(1, 1, AKB_DecimalPoint)
        C191335 = True
        GoTo Comp_C85395

    Comp_C296998:

        ' EOF Inativo
        sCurrComponent = "EOF Inativo"
        C296998DataType = 4
        C296998 = RsC85342_EOF
        GoTo Comp_C85289

    Comp_C297000:

        ' Tem Inativo?
        sCurrComponent = "Tem Inativo?"
        C297000 = (fn_ConvertValueCompiled(C296998, C296998DataType, AKB_DecimalPoint, False) = 0)
        C297000DataType = AKBTypeConst.cAKBTypeLogical
        If C297000 Then
            GoTo Comp_C85290
        Else
            GoTo Comp_C191250
        End If

    Exit_R6286:

        Exit Function

    Err_R6286:
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
