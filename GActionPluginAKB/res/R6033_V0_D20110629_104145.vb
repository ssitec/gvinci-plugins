Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R6033

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

    Public Shared Function R6033(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P17841 As Double, ByVal P17842 As Double, ByVal P17844 As Double, ByVal P17845 As Object, ByVal P17846 As String, ByVal P56041 As Object) As DataTable
    ' 
    ' 6033 - Cálcula Prêmio do Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R6033
        Dim sCurrComponent as String

        Dim QueryC82453 As New Object
        Dim RsC82453 As New Object
        Dim C82453DataType() As Short
        Dim RsC82453_EOF As Boolean
        Dim C82454 As Object
        Dim C82454DataType As Short
        Dim C82455 As Object
        Dim C82455DataType As Short
        Dim C82456 As Boolean
        Dim C82456DataType As Short
        Dim C82457 As Object
        Dim C82457DataType As Short
        Dim C82458 As Object
        Dim C82458DataType As Short
        Dim C82461 As Object
        Dim C82461DataType As Short
        Dim C82463 As Object
        Dim C82463DataType As Short
        Dim QueryC82464 As New Object
        Dim RsC82464 As New Object
        Dim C82464DataType() As Short
        Dim RsC82464_EOF As Boolean
        Dim C82465 As Object
        Dim C82465DataType As Short
        Dim C82466 As Boolean
        Dim C82466DataType As Short
        Dim C82467 As Object
        Dim C82467DataType As Short
        Dim QueryC82468 As New Object
        Dim C82468 As Integer
        Dim C82468DataType As Short
        Dim C82469 As Object
        Dim C82469DataType As Short
        Dim C82470 As Object
        Dim C82470DataType As Short
        Dim C82471 As Object
        Dim C82471DataType As Short
        Dim C82474 As Object
        Dim C82474DataType As Short
        Dim C82475 As Object
        Dim C82475DataType As Short
        Dim C82476 As Object
        Dim C82476DataType As Short
        Dim C82477 As Boolean
        Dim C82477DataType As Short
        Dim C82478 As Object
        Dim C82478DataType As Short
        Dim C82479 As Boolean
        Dim C82479DataType As Short
        Dim C82480 As Object
        Dim C82480DataType As Short
        Dim C82481 As Boolean
        Dim C82481DataType As Short
        Dim C82482 As Boolean
        Dim C82482DataType As Short
        Dim C82483DataType() As Short
        Dim QueryC82484 As New Object
        Dim RsC82484 As New Object
        Dim C82484DataType() As Short
        Dim RsC82484_EOF As Boolean
        Dim C82485 As Object
        Dim C82485DataType As Short
        Dim C82486 As Object
        Dim C82486DataType As Short
        Dim QueryC82487 As New Object
        Dim RsC82487 As New Object
        Dim C82487DataType() As Short
        Dim RsC82487_EOF As Boolean
        Dim QueryC82488 As New Object
        Dim RsC82488 As New Object
        Dim C82488DataType() As Short
        Dim RsC82488_EOF As Boolean
        Dim QueryC82489 As New Object
        Dim RsC82489 As New Object
        Dim C82489DataType() As Short
        Dim RsC82489_EOF As Boolean
        Dim QueryC82490 As New Object
        Dim RsC82490 As New Object
        Dim C82490DataType() As Short
        Dim RsC82490_EOF As Boolean
        Dim QueryC82491 As New Object
        Dim RsC82491 As New Object
        Dim C82491DataType() As Short
        Dim RsC82491_EOF As Boolean
        Dim C82495 As Object
        Dim C82495DataType As Short
        Dim QueryC82496 As New Object
        Dim RsC82496 As New Object
        Dim C82496DataType() As Short
        Dim RsC82496_EOF As Boolean
        Dim QueryC82497 As New Object
        Dim RsC82497 As New Object
        Dim C82497DataType() As Short
        Dim RsC82497_EOF As Boolean
        Dim QueryC82718 As New Object
        Dim RsC82718 As New Object
        Dim C82718DataType() As Short
        Dim RsC82718_EOF As Boolean
        Dim C82719 As Object
        Dim C82719DataType As Short
        Dim QueryC82887 As New Object
        Dim RsC82887 As New Object
        Dim C82887DataType() As Short
        Dim RsC82887_EOF As Boolean
        Dim C82892DataType() As Short
        Dim QueryC82900 As New Object
        Dim C82900 As Integer
        Dim C82900DataType As Short
        Dim C201716 As Object
        Dim C201716DataType As Short
        Dim C201717 As Object
        Dim C201717DataType As Short
        Dim C201718 As Boolean
        Dim C201718DataType As Short
        Dim QueryC201720 As New Object
        Dim C201720 As Integer
        Dim C201720DataType As Short
        Dim C201729 As Object
        Dim C201729DataType As Short
        Dim C201730 As Object
        Dim C201730DataType As Short
        Dim C201732 As Boolean
        Dim C201732DataType As Short
        Dim QueryC201735 As New Object
        Dim C201735 As Integer
        Dim C201735DataType As Short
        Dim C285569 As Boolean
        Dim C285569DataType As Short
        Dim C285570 As Object
        Dim C285570DataType As Short
        Dim C285571 As Object
        Dim C285571DataType As Short
        Dim C285573DataType() As Short
        Dim QueryC285574 As New Object
        Dim RsC285574 As New Object
        Dim C285574DataType() As Short
        Dim RsC285574_EOF As Boolean
        Dim C285575 As Boolean
        Dim C285575DataType As Short
        Dim C285576 As Object
        Dim C285576DataType As Short
        Dim C285586 As Object
        Dim C285586DataType As Short
        Dim C285587 As Object
        Dim C285587DataType As Short
        Dim C285588 As Boolean
        Dim C285588DataType As Short
        Dim QueryC285817 As New Object
        Dim C285817 As Integer
        Dim C285817DataType As Short
        Dim C285840 As Boolean
        Dim C285840DataType As Short
        Dim QueryC285842 As New Object
        Dim C285842 As Integer
        Dim C285842DataType As Short
        Dim C285845 As Object
        Dim C285845DataType As Short
        Dim C285846 As Boolean
        Dim C285846DataType As Short
        Dim QueryC285975 As New Object
        Dim C285975 As Integer
        Dim C285975DataType As Short
        Dim QueryC321206 As New Object
        Dim RsC321206 As New Object
        Dim C321206DataType() As Short
        Dim RsC321206_EOF As Boolean
        Dim C321207 As Boolean
        Dim C321207DataType As Short
        Dim QueryC321208 As New Object
        Dim C321208 As Integer
        Dim C321208DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C82495

    Comp_C82453:

        ' SelDesc
        sCurrComponent = "SelDesc"
        QueryC82453 = con.CreateCommand()
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "SELECT WF_PED_SEQ_DESCONTO.Porc_Desconto , WF_PED_SEQ_DESCONTO.Tipo_Desconto "
        QueryC82453.CommandText = QueryC82453.CommandText & " " & ""
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_DESCONTO "
        QueryC82453.CommandText = QueryC82453.CommandText & " " & ""
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Seq = 1"
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "and WF_PED_SEQ_DESCONTO.Tipo_Desconto in "
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "("
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "SELECT WF_TP_DESCONTO_PREMIO.Tipo_Desconto FROM AKBUSER01.WF_TP_DESCONTO_PREMIO "
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "WHERE"
        QueryC82453.CommandText = QueryC82453.CommandText & " " & "WF_TP_DESCONTO_PREMIO.Politica = " & _ 
ConvertValue(C82719, C82719DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82453.CommandText = QueryC82453.CommandText & " " & ")"
        QueryC82453.CommandText = QueryC82453.CommandText & " " & ""
        QueryC82453.Transaction = txn
        RsC82453 = QueryC82453.ExecuteReader()
        Dim iC82453 As Short
        ReDim C82453DataType(RsC82453.FieldCount)
        For iC82453 = 0 to RsC82453.FieldCount - 1
            Select Case RsC82453.GetDataTypeName(iC82453).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82453DataType(iC82453 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82453DataType(iC82453 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82453DataType(iC82453 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82453
        RsC82453_EOF = Not RsC82453.Read()

        GoTo Comp_C82454

    Comp_C82454:

        ' Eof
        sCurrComponent = "Eof"
        C82454DataType = 4
        C82454 = RsC82453_EOF
        GoTo Comp_C82456

    Comp_C82455:

        ' Valor
        sCurrComponent = "Valor"
        C82455 = 100
        C82455DataType = 1
        GoTo Comp_C82471

    Comp_C82456:

        ' Eof=F
        sCurrComponent = "Eof=F"
        C82456 = (fn_ConvertValueCompiled(C82454, C82454DataType, AKB_DecimalPoint, False) = 0)
        C82456DataType = AKBTypeConst.cAKBTypeLogical
        If C82456 Then
            GoTo Comp_C82461
        Else
            GoTo Comp_C82488
        End If

    Comp_C82457:

        ' Next
        sCurrComponent = "Next"
        C82457DataType = 4
        RsC82453_EOF = Not RsC82453.Read()
        C82457 = True

        GoTo Comp_C82454

    Comp_C82458:

        ' GuardaValor
        sCurrComponent = "GuardaValor"
        C82458DataType = 4
        C82455 = fn_ConvertValueCompiled(RsC82487(0) , 1, AKB_DecimalPoint)
        C82458 = True
        GoTo Comp_C82457

    Comp_C82461:

        ' Desc
        sCurrComponent = "Desc"
        C82461DataType = 0
        C82461 = RsC82453(0)
        C82461DataType = C82453Datatype(1)
        If C82461DataType = AKBTypeConst.cAKBTypeString Then
          C82461 = IIF(IsDBNull(C82461), C82461, Strings.RTrim(C82461))
        End If 

        GoTo Comp_C82470

    Comp_C82463:

        ' ZeraValor
        sCurrComponent = "ZeraValor"
        C82463DataType = 4
        C82455 = fn_ConvertValueCompiled(100, 1, AKB_DecimalPoint)
        C82463 = True
        GoTo Comp_C82491

    Comp_C82464:

        ' SelDescItem
        sCurrComponent = "SelDescItem"
        QueryC82464 = con.CreateCommand()
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "SELECT WF_PED_ITENS_DESC.Porc_Desconto , WF_PED_ITENS_DESC.Seq_Item "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "FROM AKBUSER01.WF_PED_ITENS_DESC "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "WHERE "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "WF_PED_ITENS_DESC.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "AND  WF_PED_ITENS_DESC.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "AND  WF_PED_ITENS_DESC.Tipo_Desconto in "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "("
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "SELECT WF_TP_DESCONTO_PREMIO.Tipo_Desconto FROM AKBUSER01.WF_TP_DESCONTO_PREMIO "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "WHERE"
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "WF_TP_DESCONTO_PREMIO.Politica = " & _ 
ConvertValue(C82719, C82719DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & ")"
        QueryC82464.CommandText = QueryC82464.CommandText & " " & ""
        QueryC82464.CommandText = QueryC82464.CommandText & " " & "ORDER BY WF_PED_ITENS_DESC.Seq_Item "
        QueryC82464.CommandText = QueryC82464.CommandText & " " & ""
        QueryC82464.Transaction = txn
        RsC82464 = QueryC82464.ExecuteReader()
        Dim iC82464 As Short
        ReDim C82464DataType(RsC82464.FieldCount)
        For iC82464 = 0 to RsC82464.FieldCount - 1
            Select Case RsC82464.GetDataTypeName(iC82464).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82464DataType(iC82464 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82464DataType(iC82464 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82464DataType(iC82464 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82464
        RsC82464_EOF = Not RsC82464.Read()

        GoTo Comp_C82465

    Comp_C82465:

        ' Eof DescItem
        sCurrComponent = "Eof DescItem"
        C82465DataType = 4
        C82465 = RsC82464_EOF
        GoTo Comp_C82466

    Comp_C82466:

        ' Eof DescItem=T
        sCurrComponent = "Eof DescItem=T"
        C82466 = (fn_ConvertValueCompiled(C82465, C82465DataType, AKB_DecimalPoint, False) =1)
        C82466DataType = AKBTypeConst.cAKBTypeLogical
        If C82466 Then
            GoTo Comp_C82479
        Else
            GoTo Comp_C82474
        End If

    Comp_C82467:

        ' Next DescItem
        sCurrComponent = "Next DescItem"
        C82467DataType = 4
        RsC82464_EOF = Not RsC82464.Read()
        C82467 = True

        GoTo Comp_C82465

    Comp_C82468:

        ' UpDescItem2
        sCurrComponent = "UpDescItem2"
        QueryC82468 = con.CreateCommand()
        QueryC82468.CommandText = QueryC82468.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC82468.CommandText = QueryC82468.CommandText & " " & ""
        QueryC82468.CommandText = QueryC82468.CommandText & " " & "SET WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO = round("
        QueryC82468.CommandText = QueryC82468.CommandText & " " & "((((((WF_PEDIDO_ITENS.Preco_Unit * WF_PEDIDO_ITENS.Qtde_Pedido) * " & _ 
ConvertValue(RsC82497(0), C82497DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) / 100) * " & _ 
ConvertValue(C201729, C201729DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") / 100) / WF_PEDIDO_ITENS.Qtde_Pedido),4)"
        QueryC82468.CommandText = QueryC82468.CommandText & " " & ""
        QueryC82468.CommandText = QueryC82468.CommandText & " " & ""
        QueryC82468.CommandText = QueryC82468.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82468.CommandText = QueryC82468.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82468.CommandText = QueryC82468.CommandText & " " & "AND  WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C82476, C82476DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82468.CommandText = QueryC82468.CommandText & " " & ""
        QueryC82468.Transaction = txn
        C82468 = QueryC82468.ExecuteNonQuery()
        C82468DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C82481

    Comp_C82469:

        ' GuardaValor2
        sCurrComponent = "GuardaValor2"
        C82469DataType = 4
        C82455 = fn_ConvertValueCompiled(RsC82491(0) , 1, AKB_DecimalPoint)
        C82469 = True
        GoTo Comp_C82467

    Comp_C82470:

        ' TpDesc
        sCurrComponent = "TpDesc"
        C82470DataType = 0
        C82470DataType = C82453Datatype(2)
        If C82470DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82453(1)) Then
          C82470 = Strings.RTrim(RsC82453(1))
        Else 
          C82470 = RsC82453(1)
        End If 

        GoTo Comp_C82487

    Comp_C82471:

        ' VDesc2
        sCurrComponent = "VDesc2"
        C82471 = 0
        C82471DataType = 1
        GoTo Comp_C82453

    Comp_C82474:

        ' DescItem
        sCurrComponent = "DescItem"
        C82474DataType = 0
        C82474 = RsC82464(0)
        C82474DataType = C82464Datatype(1)
        If C82474DataType = AKBTypeConst.cAKBTypeString Then
          C82474 = IIF(IsDBNull(C82474), C82474, Strings.RTrim(C82474))
        End If 

        GoTo Comp_C82475

    Comp_C82475:

        ' SeqItem
        sCurrComponent = "SeqItem"
        C82475DataType = 0
        C82475DataType = C82464Datatype(2)
        If C82475DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82464(1)) Then
          C82475 = Strings.RTrim(RsC82464(1))
        Else 
          C82475 = RsC82464(1)
        End If 

        GoTo Comp_C82477

    Comp_C82476:

        ' VSeqItem
        sCurrComponent = "VSeqItem"
        C82476 = 0
        C82476DataType = 1
        GoTo Comp_C82464

    Comp_C82477:

        ' Item<>VItem
        sCurrComponent = "Item<>VItem"
        C82477 = (fn_ConvertValueCompiled(C82475, C82475DataType, AKB_DecimalPoint, False) <>fn_ConvertValueCompiled(C82476, C82476DataType, AKB_DecimalPoint, False))
        C82477DataType = AKBTypeConst.cAKBTypeLogical
        If C82477 Then
            GoTo Comp_C82479
        Else
            GoTo Comp_C82491
        End If

    Comp_C82478:

        ' NovoItem
        sCurrComponent = "NovoItem"
        C82478DataType = 4
        C82476 = fn_ConvertValueCompiled(C82475 , 1, AKB_DecimalPoint)
        C82478 = True
        GoTo Comp_C82463

    Comp_C82479:

        ' VItem>0
        sCurrComponent = "VItem>0"
        C82479 = (fn_ConvertValueCompiled(C82476, C82476DataType, AKB_DecimalPoint, False)  >0)
        C82479DataType = AKBTypeConst.cAKBTypeLogical
        If C82479 Then
            GoTo Comp_C82489
        Else
            GoTo Comp_C82482
        End If

    Comp_C82480:

        ' GValor2
        sCurrComponent = "GValor2"
        C82480DataType = 4
        C82455 = fn_ConvertValueCompiled(RsC82489(0) , 1, AKB_DecimalPoint)
        C82480 = True
        GoTo Comp_C82490

    Comp_C82481:

        ' EOF2=T
        sCurrComponent = "EOF2=T"
        C82481 = (fn_ConvertValueCompiled(C82465, C82465DataType, AKB_DecimalPoint, False) =1)
        C82481DataType = AKBTypeConst.cAKBTypeLogical
        If C82481 Then
            GoTo Comp_C82892
        Else
            GoTo Comp_C82478
        End If

    Comp_C82482:

        ' Fim_dos_Descontos_dos_Itens
        sCurrComponent = "Fim_dos_Descontos_dos_Itens"
        C82482 = (fn_ConvertValueCompiled(C82465, C82465DataType, AKB_DecimalPoint, False) =1)
        C82482DataType = AKBTypeConst.cAKBTypeLogical
        If C82482 Then
            GoTo Comp_C82483
        Else
            GoTo Comp_C82478
        End If

    Comp_C82483:

        ' Fim-1
        sCurrComponent = "Fim-1"
        Dim tb_C82483 As DataTable = New DataTable()
        tb_C82483.Columns.Add("vTrue" & "")
        Dim row_C82483 As DataRow = tb_C82483.NewRow()
        row_C82483(0) = C82495
        tb_C82483.Rows.Add(row_C82483)
        R6033 = tb_C82483
        ReDim C82483DataType(1)
        C82483DataType(1) = C82495DataType
        ReturnDataType = C82483DataType

        GoTo Exit_R6033

    Comp_C82484:

        ' SelParâmetros
        sCurrComponent = "SelParâmetros"
        QueryC82484 = con.CreateCommand()
        QueryC82484.CommandText = QueryC82484.CommandText & " " & "SELECT WF_TP_PRODUTO.Desconto_Decimal , WF_TP_PRODUTO.Trunca_Desconto FROM AKBUSER01.WF_TP_PRODUTO "
        QueryC82484.CommandText = QueryC82484.CommandText & " " & "WHERE WF_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC82484.Transaction = txn
        RsC82484 = QueryC82484.ExecuteReader()
        Dim iC82484 As Short
        ReDim C82484DataType(RsC82484.FieldCount)
        For iC82484 = 0 to RsC82484.FieldCount - 1
            Select Case RsC82484.GetDataTypeName(iC82484).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82484DataType(iC82484 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82484DataType(iC82484 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82484DataType(iC82484 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82484
        RsC82484_EOF = Not RsC82484.Read()

        GoTo Comp_C82485

    Comp_C82485:

        ' TruncaDesconto
        sCurrComponent = "TruncaDesconto"
        C82485DataType = 0
        C82485DataType = C82484Datatype(2)
        If C82485DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82484(1)) Then
          C82485 = Strings.RTrim(RsC82484(1))
        Else 
          C82485 = RsC82484(1)
        End If 

        GoTo Comp_C82486

    Comp_C82486:

        ' DecimaisDesconto
        sCurrComponent = "DecimaisDesconto"
        C82486DataType = 0
        C82486 = RsC82484(0)
        C82486DataType = C82484Datatype(1)
        If C82486DataType = AKBTypeConst.cAKBTypeString Then
          C82486 = IIF(IsDBNull(C82486), C82486, Strings.RTrim(C82486))
        End If 

        GoTo Comp_C82455

    Comp_C82487:

        ' CalcDesc
        sCurrComponent = "CalcDesc"
        QueryC82487 = con.CreateCommand()
        QueryC82487.CommandText = QueryC82487.CommandText & " " & "SELECT  "
        QueryC82487.CommandText = QueryC82487.CommandText & " " & ""
        QueryC82487.CommandText = QueryC82487.CommandText & " " & "DECODE( " & _ 
ConvertValue(C82485, C82485DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC82487.CommandText = QueryC82487.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C82461, C82461DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC82487.CommandText = QueryC82487.CommandText & " " & "ROUND( " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C82461, C82461DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC82487.CommandText = QueryC82487.CommandText & " " & ""
        QueryC82487.CommandText = QueryC82487.CommandText & " " & "FROM DUAL"
        QueryC82487.CommandText = QueryC82487.CommandText & " " & ""
        QueryC82487.Transaction = txn
        RsC82487 = QueryC82487.ExecuteReader()
        Dim iC82487 As Short
        ReDim C82487DataType(RsC82487.FieldCount)
        For iC82487 = 0 to RsC82487.FieldCount - 1
            Select Case RsC82487.GetDataTypeName(iC82487).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82487DataType(iC82487 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82487DataType(iC82487 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82487DataType(iC82487 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82487
        RsC82487_EOF = Not RsC82487.Read()

        GoTo Comp_C82458

    Comp_C82488:

        ' DescTotal
        sCurrComponent = "DescTotal"
        QueryC82488 = con.CreateCommand()
        QueryC82488.CommandText = QueryC82488.CommandText & " " & "SELECT  "
        QueryC82488.CommandText = QueryC82488.CommandText & " " & ""
        QueryC82488.CommandText = QueryC82488.CommandText & " " & "DECODE( " & _ 
ConvertValue(C82485, C82485DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC82488.CommandText = QueryC82488.CommandText & " " & "TRUNC( 100-" & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC82488.CommandText = QueryC82488.CommandText & " " & "ROUND( 100-" & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC82488.CommandText = QueryC82488.CommandText & " " & ""
        QueryC82488.CommandText = QueryC82488.CommandText & " " & "FROM DUAL"
        QueryC82488.CommandText = QueryC82488.CommandText & " " & ""
        QueryC82488.Transaction = txn
        RsC82488 = QueryC82488.ExecuteReader()
        Dim iC82488 As Short
        ReDim C82488DataType(RsC82488.FieldCount)
        For iC82488 = 0 to RsC82488.FieldCount - 1
            Select Case RsC82488.GetDataTypeName(iC82488).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82488DataType(iC82488 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82488DataType(iC82488 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82488DataType(iC82488 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82488
        RsC82488_EOF = Not RsC82488.Read()

        GoTo Comp_C285840

    Comp_C82489:

        ' Calc_Desc_Item_2
        sCurrComponent = "Calc_Desc_Item_2"
        QueryC82489 = con.CreateCommand()
        QueryC82489.CommandText = QueryC82489.CommandText & " " & "SELECT  "
        QueryC82489.CommandText = QueryC82489.CommandText & " " & ""
        QueryC82489.CommandText = QueryC82489.CommandText & " " & "DECODE( " & _ 
ConvertValue(C82485, C82485DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC82489.CommandText = QueryC82489.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C82471, C82471DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC82489.CommandText = QueryC82489.CommandText & " " & "ROUND( " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C82471, C82471DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC82489.CommandText = QueryC82489.CommandText & " " & ""
        QueryC82489.CommandText = QueryC82489.CommandText & " " & "FROM DUAL"
        QueryC82489.CommandText = QueryC82489.CommandText & " " & ""
        QueryC82489.Transaction = txn
        RsC82489 = QueryC82489.ExecuteReader()
        Dim iC82489 As Short
        ReDim C82489DataType(RsC82489.FieldCount)
        For iC82489 = 0 to RsC82489.FieldCount - 1
            Select Case RsC82489.GetDataTypeName(iC82489).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82489DataType(iC82489 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82489DataType(iC82489 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82489DataType(iC82489 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82489
        RsC82489_EOF = Not RsC82489.Read()

        GoTo Comp_C82480

    Comp_C82490:

        ' Desc Total Item
        sCurrComponent = "Desc Total Item"
        QueryC82490 = con.CreateCommand()
        QueryC82490.CommandText = QueryC82490.CommandText & " " & "SELECT  "
        QueryC82490.CommandText = QueryC82490.CommandText & " " & ""
        QueryC82490.CommandText = QueryC82490.CommandText & " " & "DECODE( " & _ 
ConvertValue(C82485, C82485DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC82490.CommandText = QueryC82490.CommandText & " " & "TRUNC( 100 - " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC82490.CommandText = QueryC82490.CommandText & " " & "ROUND( 100 - " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC82490.CommandText = QueryC82490.CommandText & " " & ""
        QueryC82490.CommandText = QueryC82490.CommandText & " " & "FROM DUAL"
        QueryC82490.CommandText = QueryC82490.CommandText & " " & ""
        QueryC82490.Transaction = txn
        RsC82490 = QueryC82490.ExecuteReader()
        Dim iC82490 As Short
        ReDim C82490DataType(RsC82490.FieldCount)
        For iC82490 = 0 to RsC82490.FieldCount - 1
            Select Case RsC82490.GetDataTypeName(iC82490).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82490DataType(iC82490 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82490DataType(iC82490 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82490DataType(iC82490 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82490
        RsC82490_EOF = Not RsC82490.Read()

        GoTo Comp_C82496

    Comp_C82491:

        ' CalcDescItem
        sCurrComponent = "CalcDescItem"
        QueryC82491 = con.CreateCommand()
        QueryC82491.CommandText = QueryC82491.CommandText & " " & "SELECT  "
        QueryC82491.CommandText = QueryC82491.CommandText & " " & ""
        QueryC82491.CommandText = QueryC82491.CommandText & " " & "DECODE( " & _ 
ConvertValue(C82485, C82485DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 1, "
        QueryC82491.CommandText = QueryC82491.CommandText & " " & "TRUNC( " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C82474, C82474DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) , "
        QueryC82491.CommandText = QueryC82491.CommandText & " " & "ROUND( " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " - " & _ 
ConvertValue(C82474, C82474DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  / 100 * " & _ 
ConvertValue(C82455, C82455DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C82486, C82486DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) )"
        QueryC82491.CommandText = QueryC82491.CommandText & " " & ""
        QueryC82491.CommandText = QueryC82491.CommandText & " " & "FROM DUAL"
        QueryC82491.CommandText = QueryC82491.CommandText & " " & ""
        QueryC82491.Transaction = txn
        RsC82491 = QueryC82491.ExecuteReader()
        Dim iC82491 As Short
        ReDim C82491DataType(RsC82491.FieldCount)
        For iC82491 = 0 to RsC82491.FieldCount - 1
            Select Case RsC82491.GetDataTypeName(iC82491).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82491DataType(iC82491 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82491DataType(iC82491 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82491DataType(iC82491 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82491
        RsC82491_EOF = Not RsC82491.Read()

        GoTo Comp_C82469

    Comp_C82495:

        ' vTrue
        sCurrComponent = "vTrue"
        C82495 = 1
        C82495DataType = 4
        GoTo Comp_C82718

    Comp_C82496:

        ' Qry_02
        sCurrComponent = "Qry_02"
        QueryC82496 = con.CreateCommand()
        QueryC82496.CommandText = QueryC82496.CommandText & " " & "select NVL(WF_COMISSAO_PREMIO.Premio,0) ,"
        QueryC82496.CommandText = QueryC82496.CommandText & " " & "             WF_COMISSAO_PREMIO.Porc_Comissao_base_Premio "
        QueryC82496.CommandText = QueryC82496.CommandText & " " & "from AKBUSER01.WF_COMISSAO_PREMIO "
        QueryC82496.CommandText = QueryC82496.CommandText & " " & ""
        QueryC82496.CommandText = QueryC82496.CommandText & " " & "where WF_COMISSAO_PREMIO.Politica = " & _ 
ConvertValue(C82719, C82719DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC82496.CommandText = QueryC82496.CommandText & " " & "and WF_COMISSAO_PREMIO.Departamento = " & _ 
ConvertValue(P17846, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82496.CommandText = QueryC82496.CommandText & " " & "and WF_COMISSAO_PREMIO.Faixa_Inicial <= " & _ 
ConvertValue(RsC82490(0), C82490DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82496.CommandText = QueryC82496.CommandText & " " & "and WF_COMISSAO_PREMIO.Faixa_Final >= " & _ 
ConvertValue(RsC82490(0), C82490DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82496.Transaction = txn
        RsC82496 = QueryC82496.ExecuteReader()
        Dim iC82496 As Short
        ReDim C82496DataType(RsC82496.FieldCount)
        For iC82496 = 0 to RsC82496.FieldCount - 1
            Select Case RsC82496.GetDataTypeName(iC82496).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82496DataType(iC82496 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82496DataType(iC82496 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82496DataType(iC82496 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82496
        RsC82496_EOF = Not RsC82496.Read()

        GoTo Comp_C285845

    Comp_C82497:

        ' Qry_03
        sCurrComponent = "Qry_03"
        QueryC82497 = con.CreateCommand()
        QueryC82497.CommandText = QueryC82497.CommandText & " " & "select NVL(WF_PEDIDO_SEQ.Porc_Comissao,0) "
        QueryC82497.CommandText = QueryC82497.CommandText & " " & ""
        QueryC82497.CommandText = QueryC82497.CommandText & " " & "from AKBUSER01.WF_PEDIDO_SEQ "
        QueryC82497.CommandText = QueryC82497.CommandText & " " & "where WF_PEDIDO_SEQ.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82497.CommandText = QueryC82497.CommandText & " " & "and WF_PEDIDO_SEQ.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82497.CommandText = QueryC82497.CommandText & " " & "and WF_PEDIDO_SEQ.Seq = 1"
        QueryC82497.Transaction = txn
        RsC82497 = QueryC82497.ExecuteReader()
        Dim iC82497 As Short
        ReDim C82497DataType(RsC82497.FieldCount)
        For iC82497 = 0 to RsC82497.FieldCount - 1
            Select Case RsC82497.GetDataTypeName(iC82497).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82497DataType(iC82497 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82497DataType(iC82497 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82497DataType(iC82497 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82497
        RsC82497_EOF = Not RsC82497.Read()

        GoTo Comp_C82468

    Comp_C82718:

        ' Politíca
        sCurrComponent = "Politíca"
        QueryC82718 = con.CreateCommand()
        QueryC82718.CommandText = QueryC82718.CommandText & " " & "SELECT WF_POLITICA_COMISSAO.Politica , WF_POLITICA_COMISSAO.Porc_Premio_Default "
        QueryC82718.CommandText = QueryC82718.CommandText & " " & ""
        QueryC82718.CommandText = QueryC82718.CommandText & " " & "FROM AKBUSER01.WF_POLITICA_COMISSAO , AKBUSER01.WF_POL_COMISSAO_REPRES "
        QueryC82718.CommandText = QueryC82718.CommandText & " " & ""
        QueryC82718.CommandText = QueryC82718.CommandText & " " & "WHERE WF_POLITICA_COMISSAO.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P17844, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82718.CommandText = QueryC82718.CommandText & " " & "AND WF_POLITICA_COMISSAO.Validade_Inicial <= " & _ 
ConvertValue(P17845, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82718.CommandText = QueryC82718.CommandText & " " & "AND WF_POLITICA_COMISSAO.Validade_Final >= " & _ 
ConvertValue(P17845, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82718.CommandText = QueryC82718.CommandText & " " & "AND WF_POLITICA_COMISSAO.Politica = WF_POL_COMISSAO_REPRES.Politica "
        QueryC82718.CommandText = QueryC82718.CommandText & " " & "AND WF_POL_COMISSAO_REPRES.Cod_Repres  = " & _ 
ConvertValue(P56041, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82718.Transaction = txn
        RsC82718 = QueryC82718.ExecuteReader()
        Dim iC82718 As Short
        ReDim C82718DataType(RsC82718.FieldCount)
        For iC82718 = 0 to RsC82718.FieldCount - 1
            Select Case RsC82718.GetDataTypeName(iC82718).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82718DataType(iC82718 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82718DataType(iC82718 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82718DataType(iC82718 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82718
        RsC82718_EOF = Not RsC82718.Read()

        GoTo Comp_C285570

    Comp_C82719:

        ' Cod_Politíca
        sCurrComponent = "Cod_Politíca"
        C82719DataType = 0
        C82719 = RsC82718(0)
        C82719DataType = C82718Datatype(1)
        If C82719DataType = AKBTypeConst.cAKBTypeString Then
          C82719 = IIF(IsDBNull(C82719), C82719, Strings.RTrim(C82719))
        End If 

        GoTo Comp_C285576

    Comp_C82887:

        ' Qry_04
        sCurrComponent = "Qry_04"
        QueryC82887 = con.CreateCommand()
        QueryC82887.CommandText = QueryC82887.CommandText & " " & "select NVL(WF_COMISSAO_PREMIO.Premio ,0), WF_COMISSAO_PREMIO.Porc_Comissao_base_Premio "
        QueryC82887.CommandText = QueryC82887.CommandText & " " & "from AKBUSER01.WF_COMISSAO_PREMIO "
        QueryC82887.CommandText = QueryC82887.CommandText & " " & "where WF_COMISSAO_PREMIO.Politica = " & _ 
ConvertValue(C82719, C82719DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82887.CommandText = QueryC82887.CommandText & " " & "and WF_COMISSAO_PREMIO.Departamento = " & _ 
ConvertValue(P17846, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82887.CommandText = QueryC82887.CommandText & " " & "and WF_COMISSAO_PREMIO.Faixa_Inicial <= " & _ 
ConvertValue(RsC82488(0), C82488DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82887.CommandText = QueryC82887.CommandText & " " & "and WF_COMISSAO_PREMIO.Faixa_Final >= " & _ 
ConvertValue(RsC82488(0), C82488DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82887.Transaction = txn
        RsC82887 = QueryC82887.ExecuteReader()
        Dim iC82887 As Short
        ReDim C82887DataType(RsC82887.FieldCount)
        For iC82887 = 0 to RsC82887.FieldCount - 1
            Select Case RsC82887.GetDataTypeName(iC82887).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C82887DataType(iC82887 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C82887DataType(iC82887 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C82887DataType(iC82887 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC82887
        RsC82887_EOF = Not RsC82887.Read()

        GoTo Comp_C285586

    Comp_C82892:

        ' Fim-2
        sCurrComponent = "Fim-2"
        Dim tb_C82892 As DataTable = New DataTable()
        tb_C82892.Columns.Add("vTrue" & "")
        Dim row_C82892 As DataRow = tb_C82892.NewRow()
        row_C82892(0) = C82495
        tb_C82892.Rows.Add(row_C82892)
        R6033 = tb_C82892
        ReDim C82892DataType(1)
        C82892DataType(1) = C82495DataType
        ReturnDataType = C82892DataType

        GoTo Exit_R6033

    Comp_C82900:

        ' Up_02
        sCurrComponent = "Up_02"
        QueryC82900 = con.CreateCommand()
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "update AKBUSER01.WF_PEDIDO_ITENS "
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "set WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO = "
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "( select round("
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "((((((WF_PEDIDO_ITENS.Preco_Unit * WF_PEDIDO_ITENS.Qtde_Pedido) * WF_PEDIDO_SEQ.Porc_Comissao) / 100) * " & _ 
ConvertValue(C201716, C201716DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) / 100) / WF_PEDIDO_ITENS.Qtde_Pedido),4)"
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "from  AKBUSER01.WF_PEDIDO_SEQ "
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "where WF_PEDIDO_ITENS.Tp_Produto = WF_PEDIDO_SEQ.Tp_Produto"
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "and WF_PEDIDO_ITENS.Pedido = WF_PEDIDO_SEQ.Pedido "
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "and WF_PEDIDO_SEQ.Seq = 1)"
        QueryC82900.CommandText = QueryC82900.CommandText & " " & ""
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "where WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82900.CommandText = QueryC82900.CommandText & " " & "and WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC82900.Transaction = txn
        C82900 = QueryC82900.ExecuteNonQuery()
        C82900DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C82476

    Comp_C201716:

        ' %Premio
        sCurrComponent = "%Premio"
        C201716DataType = 0
        C201716 = RsC82887(0)
        C201716DataType = C82887Datatype(1)
        If C201716DataType = AKBTypeConst.cAKBTypeString Then
          C201716 = IIF(IsDBNull(C201716), C201716, Strings.RTrim(C201716))
        End If 

        GoTo Comp_C201717

    Comp_C201717:

        ' Comissão é Base para o Prêmio?
        sCurrComponent = "Comissão é Base para o Prêmio?"
        C201717DataType = 0
        C201717DataType = C82887Datatype(2)
        If C201717DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82887(1)) Then
          C201717 = Strings.RTrim(RsC82887(1))
        Else 
          C201717 = RsC82887(1)
        End If 

        GoTo Comp_C201718

    Comp_C201718:

        ' ComissaoBasePremio?
        sCurrComponent = "ComissaoBasePremio?"
        C201718 = (fn_ConvertValueCompiled(C201717, C201717DataType, AKB_DecimalPoint, False) = 1)
        C201718DataType = AKBTypeConst.cAKBTypeLogical
        If C201718 Then
            GoTo Comp_C82900
        Else
            GoTo Comp_C201720
        End If

    Comp_C201720:

        ' Up_02_1
        sCurrComponent = "Up_02_1"
        QueryC201720 = con.CreateCommand()
        QueryC201720.CommandText = QueryC201720.CommandText & " " & "update AKBUSER01.WF_PEDIDO_ITENS "
        QueryC201720.CommandText = QueryC201720.CommandText & " " & "set WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO = round("
        QueryC201720.CommandText = QueryC201720.CommandText & " " & "((((WF_PEDIDO_ITENS.Preco_Unit * WF_PEDIDO_ITENS.Qtde_Pedido) * " & _ 
ConvertValue(C201716, C201716DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) / 100) / WF_PEDIDO_ITENS.Qtde_Pedido),4)"
        QueryC201720.CommandText = QueryC201720.CommandText & " " & ""
        QueryC201720.CommandText = QueryC201720.CommandText & " " & "where WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC201720.CommandText = QueryC201720.CommandText & " " & "and WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC201720.Transaction = txn
        C201720 = QueryC201720.ExecuteNonQuery()
        C201720DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C82476

    Comp_C201729:

        ' Porc_Premio
        sCurrComponent = "Porc_Premio"
        C201729DataType = 0
        C201729 = RsC82496(0)
        C201729DataType = C82496Datatype(1)
        If C201729DataType = AKBTypeConst.cAKBTypeString Then
          C201729 = IIF(IsDBNull(C201729), C201729, Strings.RTrim(C201729))
        End If 

        GoTo Comp_C201730

    Comp_C201730:

        ' ComissaoBasePremio
        sCurrComponent = "ComissaoBasePremio"
        C201730DataType = 0
        C201730DataType = C82496Datatype(2)
        If C201730DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82496(1)) Then
          C201730 = Strings.RTrim(RsC82496(1))
        Else 
          C201730 = RsC82496(1)
        End If 

        GoTo Comp_C201732

    Comp_C201732:

        ' ComissaoBasePremio=1
        sCurrComponent = "ComissaoBasePremio=1"
        C201732 = (fn_ConvertValueCompiled(C201730, C201730DataType, AKB_DecimalPoint, False) = 1)
        C201732DataType = AKBTypeConst.cAKBTypeLogical
        If C201732 Then
            GoTo Comp_C82497
        Else
            GoTo Comp_C201735
        End If

    Comp_C201735:

        ' UpDescItem
        sCurrComponent = "UpDescItem"
        QueryC201735 = con.CreateCommand()
        QueryC201735.CommandText = QueryC201735.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC201735.CommandText = QueryC201735.CommandText & " " & ""
        QueryC201735.CommandText = QueryC201735.CommandText & " " & "SET WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO = round("
        QueryC201735.CommandText = QueryC201735.CommandText & " " & "((((WF_PEDIDO_ITENS.Preco_Unit * WF_PEDIDO_ITENS.Qtde_Pedido) * " & _ 
ConvertValue(C201729, C201729DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") / 100) / WF_PEDIDO_ITENS.Qtde_Pedido),4)"
        QueryC201735.CommandText = QueryC201735.CommandText & " " & ""
        QueryC201735.CommandText = QueryC201735.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC201735.CommandText = QueryC201735.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC201735.CommandText = QueryC201735.CommandText & " " & "AND  WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C82476, C82476DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC201735.CommandText = QueryC201735.CommandText & " " & ""
        QueryC201735.Transaction = txn
        C201735 = QueryC201735.ExecuteNonQuery()
        C201735DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C82481

    Comp_C285569:

        ' EOF_Pol=1
        sCurrComponent = "EOF_Pol=1"
        C285569 = (fn_ConvertValueCompiled(C285571, C285571DataType, AKB_DecimalPoint, False) = 1)
        C285569DataType = AKBTypeConst.cAKBTypeLogical
        If C285569 Then
            GoTo Comp_C285573
        Else
            GoTo Comp_C82719
        End If

    Comp_C285570:

        ' BOF_Pol
        sCurrComponent = "BOF_Pol"
        C285570DataType = 4

        GoTo Comp_C285571

    Comp_C285571:

        ' EOF_Pol
        sCurrComponent = "EOF_Pol"
        C285571DataType = 4
        C285571 = RsC82718_EOF
        GoTo Comp_C285569

    Comp_C285573:

        ' Ret
        sCurrComponent = "Ret"
        Dim tb_C285573 As DataTable = New DataTable()
        tb_C285573.Columns.Add("vTrue" & "")
        Dim row_C285573 As DataRow = tb_C285573.NewRow()
        row_C285573(0) = C82495
        tb_C285573.Rows.Add(row_C285573)
        R6033 = tb_C285573
        ReDim C285573DataType(1)
        C285573DataType(1) = C82495DataType
        ReturnDataType = C285573DataType

        GoTo Exit_R6033

    Comp_C285574:

        ' CountDepto
        sCurrComponent = "CountDepto"
        QueryC285574 = con.CreateCommand()
        QueryC285574.CommandText = QueryC285574.CommandText & " " & "Select count(*)"
        QueryC285574.CommandText = QueryC285574.CommandText & " " & "from AKBUSER01.WF_COMISSAO_PREMIO "
        QueryC285574.CommandText = QueryC285574.CommandText & " " & "where WF_COMISSAO_PREMIO.Politica = " & _ 
ConvertValue(C82719, C82719DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC285574.CommandText = QueryC285574.CommandText & " " & "and WF_COMISSAO_PREMIO.Departamento = " & _ 
ConvertValue(P17846, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC285574.Transaction = txn
        RsC285574 = QueryC285574.ExecuteReader()
        Dim iC285574 As Short
        ReDim C285574DataType(RsC285574.FieldCount)
        For iC285574 = 0 to RsC285574.FieldCount - 1
            Select Case RsC285574.GetDataTypeName(iC285574).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C285574DataType(iC285574 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C285574DataType(iC285574 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C285574DataType(iC285574 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC285574
        RsC285574_EOF = Not RsC285574.Read()

        GoTo Comp_C285575

    Comp_C285575:

        ' Depto>0
        sCurrComponent = "Depto>0"
        C285575 = (fn_ConvertValueCompiled(RsC285574(0), C285574DataType(1), AKB_DecimalPoint, False) > 0)
        C285575DataType = AKBTypeConst.cAKBTypeLogical
        If C285575 Then
            GoTo Comp_C82484
        Else
            GoTo Comp_C321206
        End If

    Comp_C285576:

        ' PremioDefault
        sCurrComponent = "PremioDefault"
        C285576DataType = 0
        C285576DataType = C82718Datatype(2)
        If C285576DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC82718(1)) Then
          C285576 = Strings.RTrim(RsC82718(1))
        Else 
          C285576 = RsC82718(1)
        End If 

        GoTo Comp_C285574

    Comp_C285586:

        ' BOF_Premio
        sCurrComponent = "BOF_Premio"
        C285586DataType = 4

        GoTo Comp_C285587

    Comp_C285587:

        ' EOF_Premio
        sCurrComponent = "EOF_Premio"
        C285587DataType = 4
        C285587 = RsC82887_EOF
        GoTo Comp_C285588

    Comp_C285588:

        ' EOF_Premio=1
        sCurrComponent = "EOF_Premio=1"
        C285588 = (fn_ConvertValueCompiled(C285587, C285587DataType, AKB_DecimalPoint, False) = 1)
        C285588DataType = AKBTypeConst.cAKBTypeLogical
        If C285588 Then
            GoTo Comp_C285975
        Else
            GoTo Comp_C201716
        End If

    Comp_C285817:

        ' UpPremioDefault
        sCurrComponent = "UpPremioDefault"
        QueryC285817 = con.CreateCommand()
        QueryC285817.CommandText = QueryC285817.CommandText & " " & "update AKBUSER01.WF_PEDIDO_ITENS "
        QueryC285817.CommandText = QueryC285817.CommandText & " " & "set WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO = round("
        QueryC285817.CommandText = QueryC285817.CommandText & " " & "((((WF_PEDIDO_ITENS.Preco_Unit * WF_PEDIDO_ITENS.Qtde_Pedido) * " & _ 
ConvertValue(C285576, C285576DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) / 100) / WF_PEDIDO_ITENS.Qtde_Pedido),4)"
        QueryC285817.CommandText = QueryC285817.CommandText & " " & ""
        QueryC285817.CommandText = QueryC285817.CommandText & " " & "where WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC285817.CommandText = QueryC285817.CommandText & " " & "and WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC285817.Transaction = txn
        C285817 = QueryC285817.ExecuteNonQuery()
        C285817DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C82476

    Comp_C285840:

        ' DescTotal=0
        sCurrComponent = "DescTotal=0"
        C285840 = (fn_ConvertValueCompiled(RsC82488(0), C82488DataType(1), AKB_DecimalPoint, False) = 0)
        C285840DataType = AKBTypeConst.cAKBTypeLogical
        If C285840 Then
            GoTo Comp_C285817
        Else
            GoTo Comp_C82887
        End If

    Comp_C285842:

        ' UpPremio=0
        sCurrComponent = "UpPremio=0"
        QueryC285842 = con.CreateCommand()
        QueryC285842.CommandText = QueryC285842.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC285842.CommandText = QueryC285842.CommandText & " " & ""
        QueryC285842.CommandText = QueryC285842.CommandText & " " & "SET WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO = 0"
        QueryC285842.CommandText = QueryC285842.CommandText & " " & ""
        QueryC285842.CommandText = QueryC285842.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC285842.CommandText = QueryC285842.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC285842.CommandText = QueryC285842.CommandText & " " & "AND  WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C82476, C82476DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC285842.Transaction = txn
        C285842 = QueryC285842.ExecuteNonQuery()
        C285842DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C82481

    Comp_C285845:

        ' EOF_PremioItens
        sCurrComponent = "EOF_PremioItens"
        C285845DataType = 4
        C285845 = RsC82496_EOF
        GoTo Comp_C285846

    Comp_C285846:

        ' EOF_PremioItens=1
        sCurrComponent = "EOF_PremioItens=1"
        C285846 = (fn_ConvertValueCompiled(C285845, C285845DataType, AKB_DecimalPoint, False) = 1)
        C285846DataType = AKBTypeConst.cAKBTypeLogical
        If C285846 Then
            GoTo Comp_C285842
        Else
            GoTo Comp_C201729
        End If

    Comp_C285975:

        ' UpPremioPed=0
        sCurrComponent = "UpPremioPed=0"
        QueryC285975 = con.CreateCommand()
        QueryC285975.CommandText = QueryC285975.CommandText & " " & "update AKBUSER01.WF_PEDIDO_ITENS "
        QueryC285975.CommandText = QueryC285975.CommandText & " " & "set WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO = 0"
        QueryC285975.CommandText = QueryC285975.CommandText & " " & ""
        QueryC285975.CommandText = QueryC285975.CommandText & " " & "where WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC285975.CommandText = QueryC285975.CommandText & " " & "and WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC285975.Transaction = txn
        C285975 = QueryC285975.ExecuteNonQuery()
        C285975DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C82476

    Comp_C321206:

        ' temPrêmioPreenchido?
        sCurrComponent = "temPrêmioPreenchido?"
        QueryC321206 = con.CreateCommand()
        QueryC321206.CommandText = QueryC321206.CommandText & " " & "SELECT COUNT(WF_PEDIDO_ITENS.Pedido) "
        QueryC321206.CommandText = QueryC321206.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS"
        QueryC321206.CommandText = QueryC321206.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC321206.CommandText = QueryC321206.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC321206.CommandText = QueryC321206.CommandText & " " & "AND WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO <> 0"
        QueryC321206.CommandText = QueryC321206.CommandText & " " & "AND WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO IS NOT NULL "
        QueryC321206.Transaction = txn
        RsC321206 = QueryC321206.ExecuteReader()
        Dim iC321206 As Short
        ReDim C321206DataType(RsC321206.FieldCount)
        For iC321206 = 0 to RsC321206.FieldCount - 1
            Select Case RsC321206.GetDataTypeName(iC321206).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C321206DataType(iC321206 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C321206DataType(iC321206 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C321206DataType(iC321206 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC321206
        RsC321206_EOF = Not RsC321206.Read()

        GoTo Comp_C321207

    Comp_C321207:

        ' temPrêmioPreenchido? > 0
        sCurrComponent = "temPrêmioPreenchido? > 0"
        C321207 = (fn_ConvertValueCompiled(RsC321206(0), C321206DataType(1), AKB_DecimalPoint, False) > 0)
        C321207DataType = AKBTypeConst.cAKBTypeLogical
        If C321207 Then
            GoTo Comp_C321208
        Else
            GoTo Comp_C285573
        End If

    Comp_C321208:

        ' UPPREMIOPED=01
        sCurrComponent = "UPPREMIOPED=01"
        QueryC321208 = con.CreateCommand()
        QueryC321208.CommandText = QueryC321208.CommandText & " " & "update AKBUSER01.WF_PEDIDO_ITENS "
        QueryC321208.CommandText = QueryC321208.CommandText & " " & "set WF_PEDIDO_ITENS.VALOR_PREMIO_UNITARIO = 0"
        QueryC321208.CommandText = QueryC321208.CommandText & " " & ""
        QueryC321208.CommandText = QueryC321208.CommandText & " " & "where WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P17841, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC321208.CommandText = QueryC321208.CommandText & " " & "and WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P17842, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC321208.Transaction = txn
        C321208 = QueryC321208.ExecuteNonQuery()
        C321208DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C285573

    Exit_R6033:

        Exit Function

    Err_R6033:
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
