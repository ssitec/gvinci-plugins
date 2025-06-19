Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R17223

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

    Public Shared Function R17223(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P62199 As Double, ByVal P62200 As Double, ByVal P62201 As Double) As DataTable
    ' 
    ' 17223 - Atualiza desconto Pedido Seq - PUD - Versão: 0
    ' 
        'On Error GoTo Err_R17223
        Dim sCurrComponent as String

        Dim QueryC378323 As New Object
        Dim RsC378323 As New Object
        Dim C378323DataType() As Short
        Dim RsC378323_EOF As Boolean
        Dim C378324 As Object
        Dim C378324DataType As Short
        Dim C378325 As Object
        Dim C378325DataType As Short
        Dim C378326 As Boolean
        Dim C378326DataType As Short
        Dim C378327 As Object
        Dim C378327DataType As Short
        Dim C378328 As Object
        Dim C378328DataType As Short
        Dim C378330DataType() As Short
        Dim C378331 As Object
        Dim C378331DataType As Short
        Dim C378333 As Object
        Dim C378333DataType As Short
        Dim QueryC378334 As New Object
        Dim RsC378334 As New Object
        Dim C378334DataType() As Short
        Dim RsC378334_EOF As Boolean
        Dim C378335 As Object
        Dim C378335DataType As Short
        Dim C378336 As Boolean
        Dim C378336DataType As Short
        Dim C378337 As Object
        Dim C378337DataType As Short
        Dim QueryC378338 As New Object
        Dim C378338 As Integer
        Dim C378338DataType As Short
        Dim C378339 As Object
        Dim C378339DataType As Short
        Dim C378340 As Object
        Dim C378340DataType As Short
        Dim C378341 As Object
        Dim C378341DataType As Short
        Dim C378342 As Boolean
        Dim C378342DataType As Short
        Dim C378343 As Object
        Dim C378343DataType As Short
        Dim C378344 As Object
        Dim C378344DataType As Short
        Dim C378345 As Object
        Dim C378345DataType As Short
        Dim C378346 As Object
        Dim C378346DataType As Short
        Dim C378347 As Boolean
        Dim C378347DataType As Short
        Dim C378348 As Object
        Dim C378348DataType As Short
        Dim C378349 As Boolean
        Dim C378349DataType As Short
        Dim C378350 As Object
        Dim C378350DataType As Short
        Dim C378351 As Boolean
        Dim C378351DataType As Short
        Dim C378352 As Boolean
        Dim C378352DataType As Short
        Dim C378353DataType() As Short
        Dim QueryC378354 As New Object
        Dim RsC378354 As New Object
        Dim C378354DataType() As Short
        Dim RsC378354_EOF As Boolean
        Dim C378355 As Object
        Dim C378355DataType As Short
        Dim C378356 As Object
        Dim C378356DataType As Short
        Dim QueryC378357 As New Object
        Dim RsC378357 As New Object
        Dim C378357DataType() As Short
        Dim RsC378357_EOF As Boolean
        Dim QueryC378358 As New Object
        Dim RsC378358 As New Object
        Dim C378358DataType() As Short
        Dim RsC378358_EOF As Boolean
        Dim QueryC378359 As New Object
        Dim RsC378359 As New Object
        Dim C378359DataType() As Short
        Dim RsC378359_EOF As Boolean
        Dim QueryC378360 As New Object
        Dim RsC378360 As New Object
        Dim C378360DataType() As Short
        Dim RsC378360_EOF As Boolean
        Dim QueryC378361 As New Object
        Dim RsC378361 As New Object
        Dim C378361DataType() As Short
        Dim RsC378361_EOF As Boolean
        Dim C378525 As Object
        Dim C378525DataType As Short
        Dim QueryC379984 As New Object
        Dim C379984 As Integer
        Dim C379984DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C378525

    Comp_C378323:

        ' SelDesc
        sCurrComponent = "SelDesc"
        QueryC378323 = con.CreateCommand()
        QueryC378323.CommandText = QueryC378323.CommandText & " " & "SELECT WF_PED_SEQ_DESCONTO.Porc_Desconto , WF_PED_SEQ_DESCONTO.Tipo_Desconto "
        QueryC378323.CommandText = QueryC378323.CommandText & " " & ""
        QueryC378323.CommandText = QueryC378323.CommandText & " " & ""
        QueryC378323.CommandText = QueryC378323.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_DESCONTO "
        QueryC378323.CommandText = QueryC378323.CommandText & " " & ""
        QueryC378323.CommandText = QueryC378323.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P62199, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378323.CommandText = QueryC378323.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P62200, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378323.CommandText = QueryC378323.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Seq = " & _ 
ConvertValue(P62201, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC378323.CommandText = QueryC378323.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Tipo_Desconto <> 280"
        QueryC378323.CommandText = QueryC378323.CommandText & " " & " "
        QueryC378323.Transaction = txn
        RsC378323 = QueryC378323.ExecuteReader()
        Dim iC378323 As Short
        ReDim C378323DataType(RsC378323.FieldCount)
        For iC378323 = 0 to RsC378323.FieldCount - 1
            Select Case RsC378323.GetDataTypeName(iC378323).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C378323DataType(iC378323 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C378323DataType(iC378323 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C378323DataType(iC378323 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC378323
        RsC378323_EOF = Not RsC378323.Read()

        GoTo Comp_C378324

    Comp_C378324:

        ' Eof
        sCurrComponent = "Eof"
        C378324DataType = 4
        C378324 = RsC378323_EOF
        GoTo Comp_C378326

    Comp_C378325:

        ' Valor
        sCurrComponent = "Valor"
        C378325 = 100
        C378325DataType = 1
        GoTo Comp_C378341

    Comp_C378326:

        ' Eof=F
        sCurrComponent = "Eof=F"
        C378326 = (fn_ConvertValueCompiled(C378324, C378324DataType, AKB_DecimalPoint, False) = 0)
        C378326DataType = AKBTypeConst.cAKBTypeLogical
        If C378326 Then
            GoTo Comp_C378331
        Else
            GoTo Comp_C378358
        End If

    Comp_C378327:

        ' Next
        sCurrComponent = "Next"
        C378327DataType = 4
        RsC378323_EOF = Not RsC378323.Read()
        C378327 = True

        GoTo Comp_C378324

    Comp_C378328:

        ' GuardaValor
        sCurrComponent = "GuardaValor"
        C378328DataType = 4
        C378325 = fn_ConvertValueCompiled(RsC378357(0) , 1, AKB_DecimalPoint)
        C378328 = True
        GoTo Comp_C378327

    Comp_C378330:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C378330 As DataTable = New DataTable()
        tb_C378330.Columns.Add("Ret" & "")
        Dim row_C378330 As DataRow = tb_C378330.NewRow()
        row_C378330(0) = C378525
        tb_C378330.Rows.Add(row_C378330)
        R17223 = tb_C378330
        ReDim C378330DataType(1)
        C378330DataType(1) = C378525DataType
        ReturnDataType = C378330DataType

        GoTo Exit_R17223

    Comp_C378331:

        ' Desc
        sCurrComponent = "Desc"
        C378331DataType = 0
        C378331 = RsC378323(0)
        C378331DataType = C378323Datatype(1)
        If C378331DataType = AKBTypeConst.cAKBTypeString Then
          C378331 = IIF(IsDBNull(C378331), C378331, Strings.RTrim(C378331))
        End If 

        GoTo Comp_C378340

    Comp_C378333:

        ' ZeraValor
        sCurrComponent = "ZeraValor"
        C378333DataType = 4
        C378325 = fn_ConvertValueCompiled(100, 1, AKB_DecimalPoint)
        C378333 = True
        GoTo Comp_C378361

    Comp_C378334:

        ' SelDescItem
        sCurrComponent = "SelDescItem"
        QueryC378334 = con.CreateCommand()
        QueryC378334.CommandText = QueryC378334.CommandText & " " & "SELECT WF_PED_ITENS_DESC.Porc_Desconto , WF_PED_ITENS_DESC.Seq_Item "
        QueryC378334.CommandText = QueryC378334.CommandText & " " & "FROM AKBUSER01.WF_PED_ITENS_DESC "
        QueryC378334.CommandText = QueryC378334.CommandText & " " & "WHERE WF_PED_ITENS_DESC.Tp_Produto = " & _ 
ConvertValue(P62199, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378334.CommandText = QueryC378334.CommandText & " " & "AND  WF_PED_ITENS_DESC.Pedido = " & _ 
ConvertValue(P62200, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378334.CommandText = QueryC378334.CommandText & " " & "AND  WF_PED_ITENS_DESC.Tipo_Desconto <> 280 "
        QueryC378334.CommandText = QueryC378334.CommandText & " " & "ORDER BY WF_PED_ITENS_DESC.Seq_Item "
        QueryC378334.Transaction = txn
        RsC378334 = QueryC378334.ExecuteReader()
        Dim iC378334 As Short
        ReDim C378334DataType(RsC378334.FieldCount)
        For iC378334 = 0 to RsC378334.FieldCount - 1
            Select Case RsC378334.GetDataTypeName(iC378334).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C378334DataType(iC378334 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C378334DataType(iC378334 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C378334DataType(iC378334 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC378334
        RsC378334_EOF = Not RsC378334.Read()

        GoTo Comp_C378335

    Comp_C378335:

        ' Eof DescItem
        sCurrComponent = "Eof DescItem"
        C378335DataType = 4
        C378335 = RsC378334_EOF
        GoTo Comp_C378336

    Comp_C378336:

        ' Eof DescItem=T
        sCurrComponent = "Eof DescItem=T"
        C378336 = (fn_ConvertValueCompiled(C378335, C378335DataType, AKB_DecimalPoint, False) =1)
        C378336DataType = AKBTypeConst.cAKBTypeLogical
        If C378336 Then
            GoTo Comp_C378349
        Else
            GoTo Comp_C378344
        End If

    Comp_C378337:

        ' Next DescItem
        sCurrComponent = "Next DescItem"
        C378337DataType = 4
        RsC378334_EOF = Not RsC378334.Read()
        C378337 = True

        GoTo Comp_C378335

    Comp_C378338:

        ' UpDescItem2
        sCurrComponent = "UpDescItem2"
        QueryC378338 = con.CreateCommand()
        QueryC378338.CommandText = QueryC378338.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC378338.CommandText = QueryC378338.CommandText & " " & ""
        QueryC378338.CommandText = QueryC378338.CommandText & " " & "SET WF_PEDIDO_ITENS.DPUD = " & _ 
ConvertValue(RsC378360(0), C378360DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378338.CommandText = QueryC378338.CommandText & " " & ""
        QueryC378338.CommandText = QueryC378338.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P62199, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378338.CommandText = QueryC378338.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P62200, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378338.CommandText = QueryC378338.CommandText & " " & "AND  WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C378346, C378346DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378338.CommandText = QueryC378338.CommandText & " " & ""
        QueryC378338.Transaction = txn
        C378338 = QueryC378338.ExecuteNonQuery()
        C378338DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C378351

    Comp_C378339:

        ' GuardaValor2
        sCurrComponent = "GuardaValor2"
        C378339DataType = 4
        C378325 = fn_ConvertValueCompiled(RsC378361(0) , 1, AKB_DecimalPoint)
        C378339 = True
        GoTo Comp_C378337

    Comp_C378340:

        ' TpDesc
        sCurrComponent = "TpDesc"
        C378340DataType = 0
        C378340DataType = C378323Datatype(2)
        If C378340DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC378323(1)) Then
          C378340 = Strings.RTrim(RsC378323(1))
        Else 
          C378340 = RsC378323(1)
        End If 

        GoTo Comp_C378342

    Comp_C378341:

        ' VDesc2
        sCurrComponent = "VDesc2"
        C378341 = 0
        C378341DataType = 1
        GoTo Comp_C378323

    Comp_C378342:

        ' Desc=280
        sCurrComponent = "Desc=280"
        C378342 = (fn_ConvertValueCompiled(C378340, C378340DataType, AKB_DecimalPoint, False) =280 OR fn_ConvertValueCompiled(C378340, C378340DataType, AKB_DecimalPoint, False) =281)
        C378342DataType = AKBTypeConst.cAKBTypeLogical
        If C378342 Then
            GoTo Comp_C378343
        Else
            GoTo Comp_C378357
        End If

    Comp_C378343:

        ' AtDesc2
        sCurrComponent = "AtDesc2"
        C378343DataType = 4
        C378341 = fn_ConvertValueCompiled(C378331 , 1, AKB_DecimalPoint)
        C378343 = True
        GoTo Comp_C378357

    Comp_C378344:

        ' DescItem
        sCurrComponent = "DescItem"
        C378344DataType = 0
        C378344 = RsC378334(0)
        C378344DataType = C378334Datatype(1)
        If C378344DataType = AKBTypeConst.cAKBTypeString Then
          C378344 = IIF(IsDBNull(C378344), C378344, Strings.RTrim(C378344))
        End If 

        GoTo Comp_C378345

    Comp_C378345:

        ' SeqItem
        sCurrComponent = "SeqItem"
        C378345DataType = 0
        C378345DataType = C378334Datatype(2)
        If C378345DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC378334(1)) Then
          C378345 = Strings.RTrim(RsC378334(1))
        Else 
          C378345 = RsC378334(1)
        End If 

        GoTo Comp_C378347

    Comp_C378346:

        ' VSeqItem
        sCurrComponent = "VSeqItem"
        C378346 = 0
        C378346DataType = 1
        GoTo Comp_C379984

    Comp_C378347:

        ' Item<>VItem
        sCurrComponent = "Item<>VItem"
        C378347 = (fn_ConvertValueCompiled(C378345, C378345DataType, AKB_DecimalPoint, False) <>fn_ConvertValueCompiled(C378346, C378346DataType, AKB_DecimalPoint, False))
        C378347DataType = AKBTypeConst.cAKBTypeLogical
        If C378347 Then
            GoTo Comp_C378349
        Else
            GoTo Comp_C378361
        End If

    Comp_C378348:

        ' NovoItem
        sCurrComponent = "NovoItem"
        C378348DataType = 4
        C378346 = fn_ConvertValueCompiled(C378345 , 1, AKB_DecimalPoint)
        C378348 = True
        GoTo Comp_C378333

    Comp_C378349:

        ' VItem>0
        sCurrComponent = "VItem>0"
        C378349 = (fn_ConvertValueCompiled(C378346, C378346DataType, AKB_DecimalPoint, False)  >0)
        C378349DataType = AKBTypeConst.cAKBTypeLogical
        If C378349 Then
            GoTo Comp_C378359
        Else
            GoTo Comp_C378352
        End If

    Comp_C378350:

        ' GValor2
        sCurrComponent = "GValor2"
        C378350DataType = 4
        C378325 = fn_ConvertValueCompiled(RsC378359(0) , 1, AKB_DecimalPoint)
        C378350 = True
        GoTo Comp_C378360

    Comp_C378351:

        ' EOF2=T
        sCurrComponent = "EOF2=T"
        C378351 = (fn_ConvertValueCompiled(C378335, C378335DataType, AKB_DecimalPoint, False) =1)
        C378351DataType = AKBTypeConst.cAKBTypeLogical
        If C378351 Then
            GoTo Comp_C378330
        Else
            GoTo Comp_C378348
        End If

    Comp_C378352:

        ' EOF2
        sCurrComponent = "EOF2"
        C378352 = (fn_ConvertValueCompiled(C378335, C378335DataType, AKB_DecimalPoint, False) =1)
        C378352DataType = AKBTypeConst.cAKBTypeLogical
        If C378352 Then
            GoTo Comp_C378353
        Else
            GoTo Comp_C378348
        End If

    Comp_C378353:

        ' RET2
        sCurrComponent = "RET2"
        Dim tb_C378353 As DataTable = New DataTable()
        tb_C378353.Columns.Add("Ret" & "")
        Dim row_C378353 As DataRow = tb_C378353.NewRow()
        row_C378353(0) = C378525
        tb_C378353.Rows.Add(row_C378353)
        R17223 = tb_C378353
        ReDim C378353DataType(1)
        C378353DataType(1) = C378525DataType
        ReturnDataType = C378353DataType

        GoTo Exit_R17223

    Comp_C378354:

        ' SelParâmetros
        sCurrComponent = "SelParâmetros"
        QueryC378354 = con.CreateCommand()
        QueryC378354.CommandText = QueryC378354.CommandText & " " & "SELECT WF_TP_PRODUTO.Desconto_Decimal , WF_TP_PRODUTO.Trunca_Desconto FROM AKBUSER01.WF_TP_PRODUTO WHERE WF_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P62199, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC378354.Transaction = txn
        RsC378354 = QueryC378354.ExecuteReader()
        Dim iC378354 As Short
        ReDim C378354DataType(RsC378354.FieldCount)
        For iC378354 = 0 to RsC378354.FieldCount - 1
            Select Case RsC378354.GetDataTypeName(iC378354).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C378354DataType(iC378354 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C378354DataType(iC378354 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C378354DataType(iC378354 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC378354
        RsC378354_EOF = Not RsC378354.Read()

        GoTo Comp_C378355

    Comp_C378355:

        ' TruncaDesconto
        sCurrComponent = "TruncaDesconto"
        C378355DataType = 0
        C378355DataType = C378354Datatype(2)
        If C378355DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC378354(1)) Then
          C378355 = Strings.RTrim(RsC378354(1))
        Else 
          C378355 = RsC378354(1)
        End If 

        GoTo Comp_C378356

    Comp_C378356:

        ' DecimaisDesconto
        sCurrComponent = "DecimaisDesconto"
        C378356DataType = 0
        C378356 = RsC378354(0)
        C378356DataType = C378354Datatype(1)
        If C378356DataType = AKBTypeConst.cAKBTypeString Then
          C378356 = IIF(IsDBNull(C378356), C378356, Strings.RTrim(C378356))
        End If 

        GoTo Comp_C378325

    Comp_C378357:

        ' CalcDesc
        sCurrComponent = "CalcDesc"
        QueryC378357 = con.CreateCommand()
        QueryC378357.CommandText = QueryC378357.CommandText & " " & "SELECT  "
        QueryC378357.CommandText = QueryC378357.CommandText & " " & ""
        QueryC378357.CommandText = QueryC378357.CommandText & " " & "DECODE( " & _ 
ConvertValue(C378355, C378355DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC378357.CommandText = QueryC378357.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C378331, C378331DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC378357.CommandText = QueryC378357.CommandText & " " & "ROUND( " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C378331, C378331DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC378357.CommandText = QueryC378357.CommandText & " " & ""
        QueryC378357.CommandText = QueryC378357.CommandText & " " & "FROM DUAL"
        QueryC378357.CommandText = QueryC378357.CommandText & " " & ""
        QueryC378357.Transaction = txn
        RsC378357 = QueryC378357.ExecuteReader()
        Dim iC378357 As Short
        ReDim C378357DataType(RsC378357.FieldCount)
        For iC378357 = 0 to RsC378357.FieldCount - 1
            Select Case RsC378357.GetDataTypeName(iC378357).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C378357DataType(iC378357 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C378357DataType(iC378357 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C378357DataType(iC378357 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC378357
        RsC378357_EOF = Not RsC378357.Read()

        GoTo Comp_C378328

    Comp_C378358:

        ' DescTotal
        sCurrComponent = "DescTotal"
        QueryC378358 = con.CreateCommand()
        QueryC378358.CommandText = QueryC378358.CommandText & " " & "SELECT  "
        QueryC378358.CommandText = QueryC378358.CommandText & " " & ""
        QueryC378358.CommandText = QueryC378358.CommandText & " " & "DECODE( " & _ 
ConvertValue(C378355, C378355DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC378358.CommandText = QueryC378358.CommandText & " " & "TRUNC( 100-" & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC378358.CommandText = QueryC378358.CommandText & " " & "ROUND( 100-" & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC378358.CommandText = QueryC378358.CommandText & " " & ""
        QueryC378358.CommandText = QueryC378358.CommandText & " " & "FROM DUAL"
        QueryC378358.CommandText = QueryC378358.CommandText & " " & ""
        QueryC378358.Transaction = txn
        RsC378358 = QueryC378358.ExecuteReader()
        Dim iC378358 As Short
        ReDim C378358DataType(RsC378358.FieldCount)
        For iC378358 = 0 to RsC378358.FieldCount - 1
            Select Case RsC378358.GetDataTypeName(iC378358).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C378358DataType(iC378358 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C378358DataType(iC378358 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C378358DataType(iC378358 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC378358
        RsC378358_EOF = Not RsC378358.Read()

        GoTo Comp_C378346

    Comp_C378359:

        ' CALC DESCITEM2
        sCurrComponent = "CALC DESCITEM2"
        QueryC378359 = con.CreateCommand()
        QueryC378359.CommandText = QueryC378359.CommandText & " " & "SELECT  "
        QueryC378359.CommandText = QueryC378359.CommandText & " " & ""
        QueryC378359.CommandText = QueryC378359.CommandText & " " & "DECODE( " & _ 
ConvertValue(C378355, C378355DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC378359.CommandText = QueryC378359.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C378341, C378341DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC378359.CommandText = QueryC378359.CommandText & " " & "ROUND( " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C378341, C378341DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC378359.CommandText = QueryC378359.CommandText & " " & ""
        QueryC378359.CommandText = QueryC378359.CommandText & " " & "FROM DUAL"
        QueryC378359.CommandText = QueryC378359.CommandText & " " & ""
        QueryC378359.Transaction = txn
        RsC378359 = QueryC378359.ExecuteReader()
        Dim iC378359 As Short
        ReDim C378359DataType(RsC378359.FieldCount)
        For iC378359 = 0 to RsC378359.FieldCount - 1
            Select Case RsC378359.GetDataTypeName(iC378359).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C378359DataType(iC378359 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C378359DataType(iC378359 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C378359DataType(iC378359 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC378359
        RsC378359_EOF = Not RsC378359.Read()

        GoTo Comp_C378350

    Comp_C378360:

        ' Desc Total Item
        sCurrComponent = "Desc Total Item"
        QueryC378360 = con.CreateCommand()
        QueryC378360.CommandText = QueryC378360.CommandText & " " & "SELECT  "
        QueryC378360.CommandText = QueryC378360.CommandText & " " & ""
        QueryC378360.CommandText = QueryC378360.CommandText & " " & "DECODE( " & _ 
ConvertValue(C378355, C378355DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC378360.CommandText = QueryC378360.CommandText & " " & "TRUNC( 100 - " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC378360.CommandText = QueryC378360.CommandText & " " & "ROUND( 100 - " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC378360.CommandText = QueryC378360.CommandText & " " & ""
        QueryC378360.CommandText = QueryC378360.CommandText & " " & "FROM DUAL"
        QueryC378360.CommandText = QueryC378360.CommandText & " " & ""
        QueryC378360.Transaction = txn
        RsC378360 = QueryC378360.ExecuteReader()
        Dim iC378360 As Short
        ReDim C378360DataType(RsC378360.FieldCount)
        For iC378360 = 0 to RsC378360.FieldCount - 1
            Select Case RsC378360.GetDataTypeName(iC378360).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C378360DataType(iC378360 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C378360DataType(iC378360 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C378360DataType(iC378360 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC378360
        RsC378360_EOF = Not RsC378360.Read()

        GoTo Comp_C378338

    Comp_C378361:

        ' CalcDescItem
        sCurrComponent = "CalcDescItem"
        QueryC378361 = con.CreateCommand()
        QueryC378361.CommandText = QueryC378361.CommandText & " " & "SELECT  "
        QueryC378361.CommandText = QueryC378361.CommandText & " " & ""
        QueryC378361.CommandText = QueryC378361.CommandText & " " & "DECODE( " & _ 
ConvertValue(C378355, C378355DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC378361.CommandText = QueryC378361.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C378344, C378344DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC378361.CommandText = QueryC378361.CommandText & " " & "ROUND( " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C378344, C378344DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C378325, C378325DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C378356, C378356DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC378361.CommandText = QueryC378361.CommandText & " " & ""
        QueryC378361.CommandText = QueryC378361.CommandText & " " & "FROM DUAL"
        QueryC378361.CommandText = QueryC378361.CommandText & " " & ""
        QueryC378361.Transaction = txn
        RsC378361 = QueryC378361.ExecuteReader()
        Dim iC378361 As Short
        ReDim C378361DataType(RsC378361.FieldCount)
        For iC378361 = 0 to RsC378361.FieldCount - 1
            Select Case RsC378361.GetDataTypeName(iC378361).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C378361DataType(iC378361 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C378361DataType(iC378361 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C378361DataType(iC378361 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC378361
        RsC378361_EOF = Not RsC378361.Read()

        GoTo Comp_C378339

    Comp_C378525:

        ' Ret
        sCurrComponent = "Ret"
        C378525 = 1
        C378525DataType = 4
        GoTo Comp_C378354

    Comp_C379984:

        ' UpDescItem
        sCurrComponent = "UpDescItem"
        QueryC379984 = con.CreateCommand()
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "UPDATE WF_PEDIDO_ITENS"
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "SET WF_PEDIDO_ITENS.DPUD = " & _ 
ConvertValue(RsC378358(0), C378358DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P62199, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P62200, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "AND NOT EXISTS (SELECT * FROM wf_tabpreco_vda_produto, WF_PEDIDO"
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "WHERE WF_PEDIDO.PEDIDO = WF_PEDIDO_ITENS.PEDIDO"
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "AND WF_PEDIDO.TABELA_PRECO_VENDA = wf_tabpreco_vda_produto.TABELA_PRECO_VENDA"
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "AND WF_PEDIDO_ITENS.ACESSO = wf_tabpreco_vda_produto.ACESSO"
        QueryC379984.CommandText = QueryC379984.CommandText & " " & "AND nvl(wf_tabpreco_vda_produto.ACESSO_SERVICO,0) <> 0)"
        QueryC379984.Transaction = txn
        C379984 = QueryC379984.ExecuteNonQuery()
        C379984DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C378334

    Exit_R17223:

        Exit Function

    Err_R17223:
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
