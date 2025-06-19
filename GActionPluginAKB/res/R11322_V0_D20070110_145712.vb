Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R11322

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

    Public Shared Function R11322(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P37343 As Double, ByVal P37344 As Double, ByVal P37345 As Double) As DataTable
    ' 
    ' 11322 - Desconto Itens do Pedido - PPR - Versão: 0
    ' 
        'On Error GoTo Err_R11322
        Dim sCurrComponent as String

        Dim C213007 As Object
        Dim C213007DataType As Short
        Dim C213008 As Boolean
        Dim C213008DataType As Short
        Dim C213009 As Object
        Dim C213009DataType As Short
        Dim C213011 As Object
        Dim C213011DataType As Short
        Dim C213012 As Object
        Dim C213012DataType As Short
        Dim C213013 As Object
        Dim C213013DataType As Short
        Dim C213014 As Object
        Dim C213014DataType As Short
        Dim C213015 As Object
        Dim C213015DataType As Short
        Dim C213016 As Object
        Dim C213016DataType As Short
        Dim C213019 As Double
        Dim C213019DataType As Short
        Dim C213020 As Object
        Dim C213020DataType As Short
        Dim QueryC213022 As New Object
        Dim RsC213022 As New Object
        Dim C213022DataType() As Short
        Dim RsC213022_EOF As Boolean
        Dim C213023 As Object
        Dim C213023DataType As Short
        Dim QueryC213024 As New Object
        Dim RsC213024 As New Object
        Dim C213024DataType() As Short
        Dim RsC213024_EOF As Boolean
        Dim C213025 As Object
        Dim C213025DataType As Short
        Dim C213026 As Boolean
        Dim C213026DataType As Short
        Dim C213027 As Object
        Dim C213027DataType As Short
        Dim C213028 As Object
        Dim C213028DataType As Short
        Dim C213029 As Object
        Dim C213029DataType As Short
        Dim C213030 As Object
        Dim C213030DataType As Short
        Dim C213031 As Boolean
        Dim C213031DataType As Short
        Dim C213032 As Object
        Dim C213032DataType As Short
        Dim C213033 As Double
        Dim C213033DataType As Short
        Dim QueryC213034 As New Object
        Dim RsC213034 As New Object
        Dim C213034DataType() As Short
        Dim RsC213034_EOF As Boolean
        Dim C213035 As Object
        Dim C213035DataType As Short
        Dim C213036 As Object
        Dim C213036DataType As Short
        Dim C213037 As Object
        Dim C213037DataType As Short
        Dim C213038 As Boolean
        Dim C213038DataType As Short
        Dim C213039 As Object
        Dim C213039DataType As Short
        Dim C213040 As Double
        Dim C213040DataType As Short
        Dim C213041 As Object
        Dim C213041DataType As Short
        Dim C213042 As Object
        Dim C213042DataType As Short
        Dim C213044 As Double
        Dim C213044DataType As Short
        Dim C213045 As Object
        Dim C213045DataType As Short
        Dim C213050DataType() As Short
        Dim QueryC213052 As New Object
        Dim C213052 As Integer
        Dim C213052DataType As Short
        Dim C213574 As Object
        Dim C213574DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C213574

    Comp_C213007:

        ' EofSelPedidos
        sCurrComponent = "EofSelPedidos"
        C213007DataType = 4
        C213007 = RsC213022_EOF
        GoTo Comp_C213008

    Comp_C213008:

        ' IfEofPedidos
        sCurrComponent = "IfEofPedidos"
        C213008 = (fn_ConvertValueCompiled(C213007, C213007DataType, AKB_DecimalPoint, False) =1)
        C213008DataType = AKBTypeConst.cAKBTypeLogical
        If C213008 Then
            GoTo Comp_C213050
        Else
            GoTo Comp_C213011
        End If

    Comp_C213009:

        ' Next
        sCurrComponent = "Next"
        C213009DataType = 4
        RsC213022_EOF = Not RsC213022.Read()
        C213009 = True

        GoTo Comp_C213007

    Comp_C213011:

        ' TipoProd
        sCurrComponent = "TipoProd"
        C213011DataType = 0
        C213011 = RsC213022(0)
        C213011DataType = C213022Datatype(1)
        If C213011DataType = AKBTypeConst.cAKBTypeString Then
          C213011 = IIF(IsDBNull(C213011), C213011, Strings.RTrim(C213011))
        End If 

        GoTo Comp_C213014

    Comp_C213012:

        ' PrecoUnit
        sCurrComponent = "PrecoUnit"
        C213012DataType = 0
        C213012DataType = C213022Datatype(4)
        If C213012DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC213022(3)) Then
          C213012 = Strings.RTrim(RsC213022(3))
        Else 
          C213012 = RsC213022(3)
        End If 

        GoTo Comp_C213023

    Comp_C213013:

        ' SeqItem
        sCurrComponent = "SeqItem"
        C213013DataType = 0
        C213013DataType = C213022Datatype(3)
        If C213013DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC213022(2)) Then
          C213013 = Strings.RTrim(RsC213022(2))
        Else 
          C213013 = RsC213022(2)
        End If 

        GoTo Comp_C213012

    Comp_C213014:

        ' Pedido
        sCurrComponent = "Pedido"
        C213014DataType = 0
        C213014DataType = C213022Datatype(2)
        If C213014DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC213022(1)) Then
          C213014 = Strings.RTrim(RsC213022(1))
        Else 
          C213014 = RsC213022(1)
        End If 

        GoTo Comp_C213013

    Comp_C213015:

        ' TotalComDesconto
        sCurrComponent = "TotalComDesconto"
        C213015 = 0
        C213015DataType = 1
        GoTo Comp_C213016

    Comp_C213016:

        ' TotalSemDesconto
        sCurrComponent = "TotalSemDesconto"
        C213016 = 0
        C213016DataType = 1
        GoTo Comp_C213022

    Comp_C213019:

        ' CalcTotalComDescontoItem
        sCurrComponent = "CalcTotalComDescontoItem"
        C213019 = ( fn_ConvertValueCompiled(C213023, C213023DataType, AKB_DecimalPoint, True)  * fn_ConvertValueCompiled(C213012, C213012DataType, AKB_DecimalPoint, True) * ( 100 - fn_ConvertValueCompiled(C213028, C213028DataType, AKB_DecimalPoint, True) ) / 100 )
        C213019DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C213020

    Comp_C213020:

        ' AtribTotalComDescontoItem
        sCurrComponent = "AtribTotalComDescontoItem"
        C213020DataType = 4
        C213015 = fn_ConvertValueCompiled(C213019 , 1, AKB_DecimalPoint)
        C213020 = True
        GoTo Comp_C213052

    Comp_C213022:

        ' SelPedidos
        sCurrComponent = "SelPedidos"
        QueryC213022 = con.CreateCommand()
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "SELECT "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WF_PEDIDO_ITENS.Tp_Produto , "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WF_PEDIDO_ITENS.Pedido , "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WF_PEDIDO_ITENS.Seq_Item , "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WF_PEDIDO_ITENS.PRECO_SEM_DESCONTO, "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Pedido "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "FROM "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WF_PEDIDO , "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WF_PEDIDO_ITENS "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WHERE "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO.Pedido > 0 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO.Flag_Pos_Ped <> 8 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO.Flag_Pos_Ped <> 9 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO.Flag_Pos_Ped <> 12 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO.Flag_Pos_Ped <> 299 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO_ITENS.Flag_Pos_Ped <> 8 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO_ITENS.Flag_Pos_Ped <> 9 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO_ITENS.Flag_Pos_Ped <> 12 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO_ITENS.Flag_Pos_Ped <> 13 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO_ITENS.Flag_Pos_Ped <> 299 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO_ITENS.Tipo_Ped <> 51 "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND WF_PEDIDO.PEDIDO = " & _ 
ConvertValue(P37345, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P37344, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213022.CommandText = QueryC213022.CommandText & " " & "AND  WF_PEDIDO.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P37343, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213022.Transaction = txn
        RsC213022 = QueryC213022.ExecuteReader()
        Dim iC213022 As Short
        ReDim C213022DataType(RsC213022.FieldCount)
        For iC213022 = 0 to RsC213022.FieldCount - 1
            Select Case RsC213022.GetDataTypeName(iC213022).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C213022DataType(iC213022 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C213022DataType(iC213022 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C213022DataType(iC213022 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC213022
        RsC213022_EOF = Not RsC213022.Read()

        GoTo Comp_C213007

    Comp_C213023:

        ' QtdePedido
        sCurrComponent = "QtdePedido"
        C213023DataType = 0
        C213023DataType = C213022Datatype(5)
        If C213023DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC213022(4)) Then
          C213023 = Strings.RTrim(RsC213022(4))
        Else 
          C213023 = RsC213022(4)
        End If 

        GoTo Comp_C213029

    Comp_C213024:

        ' SelPedItensDesc
        sCurrComponent = "SelPedItensDesc"
        QueryC213024 = con.CreateCommand()
        QueryC213024.CommandText = QueryC213024.CommandText & " " & "SELECT WF_PED_ITENS_DESC.Porc_Desconto "
        QueryC213024.CommandText = QueryC213024.CommandText & " " & "FROM  WF_PED_ITENS_DESC, WF_TIPO_DESC"
        QueryC213024.CommandText = QueryC213024.CommandText & " " & "WHERE WF_PED_ITENS_DESC.Tp_Produto =  " & _ 
ConvertValue(C213011, C213011DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213024.CommandText = QueryC213024.CommandText & " " & "AND  WF_PED_ITENS_DESC.Pedido =  " & _ 
ConvertValue(C213014, C213014DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213024.CommandText = QueryC213024.CommandText & " " & "AND  WF_PED_ITENS_DESC.Seq_Item =  " & _ 
ConvertValue(C213013, C213013DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213024.CommandText = QueryC213024.CommandText & " " & "AND  WF_PED_ITENS_DESC.Tipo_Desconto = WF_TIPO_DESC.Tipo_Desconto "
        QueryC213024.CommandText = QueryC213024.CommandText & " " & "AND WF_TIPO_DESC.Ent_Calc_Desc_Medio = 1"
        QueryC213024.Transaction = txn
        RsC213024 = QueryC213024.ExecuteReader()
        Dim iC213024 As Short
        ReDim C213024DataType(RsC213024.FieldCount)
        For iC213024 = 0 to RsC213024.FieldCount - 1
            Select Case RsC213024.GetDataTypeName(iC213024).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C213024DataType(iC213024 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C213024DataType(iC213024 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C213024DataType(iC213024 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC213024
        RsC213024_EOF = Not RsC213024.Read()

        GoTo Comp_C213025

    Comp_C213025:

        ' EofPedItensDesc
        sCurrComponent = "EofPedItensDesc"
        C213025DataType = 4
        C213025 = RsC213024_EOF
        GoTo Comp_C213026

    Comp_C213026:

        ' IfEofPedItensDesc
        sCurrComponent = "IfEofPedItensDesc"
        C213026 = (fn_ConvertValueCompiled(C213025, C213025DataType, AKB_DecimalPoint, False) =1)
        C213026DataType = AKBTypeConst.cAKBTypeLogical
        If C213026 Then
            GoTo Comp_C213031
        Else
            GoTo Comp_C213027
        End If

    Comp_C213027:

        ' DescItem
        sCurrComponent = "DescItem"
        C213027DataType = 0
        C213027 = RsC213024(0)
        C213027DataType = C213024Datatype(1)
        If C213027DataType = AKBTypeConst.cAKBTypeString Then
          C213027 = IIF(IsDBNull(C213027), C213027, Strings.RTrim(C213027))
        End If 

        GoTo Comp_C213033

    Comp_C213028:

        ' vDescItem
        sCurrComponent = "vDescItem"
        C213028 = 0
        C213028DataType = 1
        GoTo Comp_C213015

    Comp_C213029:

        ' Atr100VDescItem
        sCurrComponent = "Atr100VDescItem"
        C213029DataType = 4
        C213028 = fn_ConvertValueCompiled(100, 1, AKB_DecimalPoint)
        C213029 = True
        GoTo Comp_C213024

    Comp_C213030:

        ' AtribuiDescItem
        sCurrComponent = "AtribuiDescItem"
        C213030DataType = 4
        C213028 = fn_ConvertValueCompiled(C213033 , 1, AKB_DecimalPoint)
        C213030 = True
        GoTo Comp_C213032

    Comp_C213031:

        ' IfTemDescontoItem
        sCurrComponent = "IfTemDescontoItem"
        C213031 = (fn_ConvertValueCompiled(C213028, C213028DataType, AKB_DecimalPoint, False)<>100)
        C213031DataType = AKBTypeConst.cAKBTypeLogical
        If C213031 Then
            GoTo Comp_C213019
        Else
            GoTo Comp_C213036
        End If

    Comp_C213032:

        ' NextPedItensDesc
        sCurrComponent = "NextPedItensDesc"
        C213032DataType = 4
        RsC213024_EOF = Not RsC213024.Read()
        C213032 = True

        GoTo Comp_C213025

    Comp_C213033:

        ' CalcDescItem
        sCurrComponent = "CalcDescItem"
        C213033 = fn_ConvertValueCompiled(C213028, C213028DataType, AKB_DecimalPoint, True)  - ( fn_ConvertValueCompiled(C213028, C213028DataType, AKB_DecimalPoint, True)  * ( fn_ConvertValueCompiled(C213027, C213027DataType, AKB_DecimalPoint, True)  / 100 ))  
        C213033DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C213030

    Comp_C213034:

        ' SelDescPedido
        sCurrComponent = "SelDescPedido"
        QueryC213034 = con.CreateCommand()
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "SELECT WF_PED_SEQ_DESCONTO.Porc_Desconto "
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "FROM WF_PED_SEQ_DESCONTO , WF_PEDIDO_SEQ , WF_PEDIDO, WF_TIPO_DESC"
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "WHERE WF_PEDIDO_SEQ.Tp_Produto = WF_PED_SEQ_DESCONTO.Tp_Produto "
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "AND  WF_PEDIDO_SEQ.Pedido = WF_PED_SEQ_DESCONTO.Pedido "
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "AND  WF_PEDIDO_SEQ.Seq = WF_PED_SEQ_DESCONTO.Seq "
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = WF_PEDIDO_SEQ.Tp_Produto "
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "AND  WF_PEDIDO.Pedido = WF_PEDIDO_SEQ.Pedido "
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "AND  WF_PEDIDO.Pedido =  " & _ 
ConvertValue(C213014, C213014DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "AND  WF_PED_SEQ_DESCONTO.Tp_Produto =1"
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "AND  WF_PED_SEQ_DESCONTO.Tipo_Desconto = WF_TIPO_DESC.Tipo_Desconto "
        QueryC213034.CommandText = QueryC213034.CommandText & " " & "AND  WF_TIPO_DESC.Ent_Calc_Desc_Medio = 1"
        QueryC213034.Transaction = txn
        RsC213034 = QueryC213034.ExecuteReader()
        Dim iC213034 As Short
        ReDim C213034DataType(RsC213034.FieldCount)
        For iC213034 = 0 to RsC213034.FieldCount - 1
            Select Case RsC213034.GetDataTypeName(iC213034).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C213034DataType(iC213034 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C213034DataType(iC213034 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C213034DataType(iC213034 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC213034
        RsC213034_EOF = Not RsC213034.Read()

        GoTo Comp_C213037

    Comp_C213035:

        ' vDescPed
        sCurrComponent = "vDescPed"
        C213035 = 0
        C213035DataType = 1
        GoTo Comp_C213028

    Comp_C213036:

        ' Atr100DescPed
        sCurrComponent = "Atr100DescPed"
        C213036DataType = 4
        C213035 = fn_ConvertValueCompiled(100, 1, AKB_DecimalPoint)
        C213036 = True
        GoTo Comp_C213034

    Comp_C213037:

        ' EOFDescPed
        sCurrComponent = "EOFDescPed"
        C213037DataType = 4
        C213037 = RsC213034_EOF
        GoTo Comp_C213038

    Comp_C213038:

        ' IfEofDescPed
        sCurrComponent = "IfEofDescPed"
        C213038 = (fn_ConvertValueCompiled(C213037, C213037DataType, AKB_DecimalPoint, False) =1)
        C213038DataType = AKBTypeConst.cAKBTypeLogical
        If C213038 Then
            GoTo Comp_C213044
        Else
            GoTo Comp_C213039
        End If

    Comp_C213039:

        ' DescPed
        sCurrComponent = "DescPed"
        C213039DataType = 0
        C213039 = RsC213034(0)
        C213039DataType = C213034Datatype(1)
        If C213039DataType = AKBTypeConst.cAKBTypeString Then
          C213039 = IIF(IsDBNull(C213039), C213039, Strings.RTrim(C213039))
        End If 

        GoTo Comp_C213040

    Comp_C213040:

        ' CalcDescPed
        sCurrComponent = "CalcDescPed"
        C213040 = fn_ConvertValueCompiled(C213035, C213035DataType, AKB_DecimalPoint, True)  - ( fn_ConvertValueCompiled(C213035, C213035DataType, AKB_DecimalPoint, True)  *  ( fn_ConvertValueCompiled(C213039, C213039DataType, AKB_DecimalPoint, True)  / 100 ))
        C213040DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C213041

    Comp_C213041:

        ' AtrDescPed
        sCurrComponent = "AtrDescPed"
        C213041DataType = 4
        C213035 = fn_ConvertValueCompiled(C213040 , 1, AKB_DecimalPoint)
        C213041 = True
        GoTo Comp_C213042

    Comp_C213042:

        ' NextDescPedido
        sCurrComponent = "NextDescPedido"
        C213042DataType = 4
        RsC213034_EOF = Not RsC213034.Read()
        C213042 = True

        GoTo Comp_C213037

    Comp_C213044:

        ' CalcTotalComDescontoPed
        sCurrComponent = "CalcTotalComDescontoPed"
        C213044 = ( fn_ConvertValueCompiled(C213023, C213023DataType, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(C213012, C213012DataType, AKB_DecimalPoint, True) * ( 100 - fn_ConvertValueCompiled(C213035, C213035DataType, AKB_DecimalPoint, True) ) / 100 )
        C213044DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C213045

    Comp_C213045:

        ' AtribTotalComDescontoPed
        sCurrComponent = "AtribTotalComDescontoPed"
        C213045DataType = 4
        C213015 = fn_ConvertValueCompiled(C213044 , 1, AKB_DecimalPoint)
        C213045 = True
        GoTo Comp_C213052

    Comp_C213050:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C213050 As DataTable = New DataTable()
        tb_C213050.Columns.Add("vRet" & "")
        Dim row_C213050 As DataRow = tb_C213050.NewRow()
        row_C213050(0) = C213574
        tb_C213050.Rows.Add(row_C213050)
        R11322 = tb_C213050
        ReDim C213050DataType(1)
        C213050DataType(1) = C213574DataType
        ReturnDataType = C213050DataType

        GoTo Exit_R11322

    Comp_C213052:

        ' Atualiza Desc Médio do Item
        sCurrComponent = "Atualiza Desc Médio do Item"
        QueryC213052 = con.CreateCommand()
        QueryC213052.CommandText = QueryC213052.CommandText & " " & "UPDATE WF_PEDIDO_ITENS"
        QueryC213052.CommandText = QueryC213052.CommandText & " " & "SET DESC_PPR = " & _ 
ConvertValue(C213015, C213015DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213052.CommandText = QueryC213052.CommandText & " " & "WHERE PEDIDO = " & _ 
ConvertValue(C213014, C213014DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213052.CommandText = QueryC213052.CommandText & " " & "AND TP_PRODUTO = " & _ 
ConvertValue(C213011, C213011DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC213052.CommandText = QueryC213052.CommandText & " " & "AND WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C213013, C213013DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC213052.Transaction = txn
        C213052 = QueryC213052.ExecuteNonQuery()
        C213052DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C213009

    Comp_C213574:

        ' vRet
        sCurrComponent = "vRet"
        C213574 = 1
        C213574DataType = 4
        GoTo Comp_C213035

    Exit_R11322:

        Exit Function

    Err_R11322:
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
