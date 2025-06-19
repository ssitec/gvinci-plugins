Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R7681

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

    Public Shared Function R7681(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P23407 As Object, ByVal P23408 As Object, ByVal P28520 As Object, ByVal P37792 As Object) As DataTable
    ' 
    ' 7681 - Verifica se é Pedido antecipado - Versão: 0
    ' 
        'On Error GoTo Err_R7681
        Dim sCurrComponent as String

        Dim C115020 As Boolean
        Dim C115020DataType As Short
        Dim QueryC115021 As New Object
        Dim RsC115021 As New Object
        Dim C115021DataType() As Short
        Dim RsC115021_EOF As Boolean
        Dim C115142 As Boolean
        Dim C115142DataType As Short
        Dim QueryC115144 As New Object
        Dim RsC115144 As New Object
        Dim C115144DataType() As Short
        Dim RsC115144_EOF As Boolean
        Dim C115151DataType() As Short
        Dim C115152 As Object
        Dim C115152DataType As Short
        Dim QueryC128476 As New Object
        Dim RsC128476 As New Object
        Dim C128476DataType() As Short
        Dim RsC128476_EOF As Boolean
        Dim C128477 As Object
        Dim C128477DataType As Short
        Dim C128478 As Boolean
        Dim C128478DataType As Short
        Dim C128481 As Boolean
        Dim C128481DataType As Short
        Dim QueryC128505 As New Object
        Dim C128505 As Integer
        Dim C128505DataType As Short
        Dim C148902 As Boolean
        Dim C148902DataType As Short
        Dim C149124 As Short
        Dim C149124DataType As Short
        Dim C149124ReturnDataType() As Short

        Dim QueryC172993 As New Object
        Dim RsC172993 As New Object
        Dim C172993DataType() As Short
        Dim RsC172993_EOF As Boolean
        Dim C173005 As Boolean
        Dim C173005DataType As Short
        Dim QueryC203444 As New Object
        Dim RsC203444 As New Object
        Dim C203444DataType() As Short
        Dim RsC203444_EOF As Boolean
        Dim QueryC203446 As New Object
        Dim RsC203446 As New Object
        Dim C203446DataType() As Short
        Dim RsC203446_EOF As Boolean
        Dim C203447 As Double
        Dim C203447DataType As Short
        Dim QueryC203448 As New Object
        Dim C203448 As Integer
        Dim C203448DataType As Short
        Dim C203460 As Object
        Dim C203460DataType As Short
        P23407 = IIf(IsDBNull(P23407), 0, P23407)

        P23408 = IIf(IsDBNull(P23408), 0, P23408)

        P37792 = IIf(IsDBNull(P37792), 0, P37792)

        ReDim ReturnDatatype(0)

        GoTo Comp_C115152

    Comp_C115020:

        ' Pedido>0
        sCurrComponent = "Pedido>0"
        C115020 = (fn_ConvertValueCompiled(P23407, 1, AKB_DecimalPoint, False)  > 0)
        C115020DataType = AKBTypeConst.cAKBTypeLogical
        If C115020 Then
            GoTo Comp_C115021
        Else
            GoTo Comp_C115151
        End If

    Comp_C115021:

        ' Pagto_Antecipado
        sCurrComponent = "Pagto_Antecipado"
        QueryC115021 = con.CreateCommand()
        QueryC115021.CommandText = QueryC115021.CommandText & " " & "Select  NVL( WF_PEDIDO.Pagto_Antecipado , 0)"
        QueryC115021.CommandText = QueryC115021.CommandText & " " & "from  AKBUSER01.WF_PEDIDO "
        QueryC115021.CommandText = QueryC115021.CommandText & " " & "where WF_PEDIDO.Pedido  = " & _ 
ConvertValue(P23407, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC115021.CommandText = QueryC115021.CommandText & " " & "and WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P23408, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC115021.Transaction = txn
        RsC115021 = QueryC115021.ExecuteReader()
        Dim iC115021 As Short
        ReDim C115021DataType(RsC115021.FieldCount)
        For iC115021 = 0 to RsC115021.FieldCount - 1
            Select Case RsC115021.GetDataTypeName(iC115021).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C115021DataType(iC115021 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C115021DataType(iC115021 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C115021DataType(iC115021 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC115021
        RsC115021_EOF = Not RsC115021.Read()

        GoTo Comp_C115142

    Comp_C115142:

        ' Pagto_Antec=1
        sCurrComponent = "Pagto_Antec=1"
        C115142 = (fn_ConvertValueCompiled(RsC115021(0), C115021DataType(1), AKB_DecimalPoint, False) = 1)
        C115142DataType = AKBTypeConst.cAKBTypeLogical
        If C115142 Then
            GoTo Comp_C148902
        Else
            GoTo Comp_C115151
        End If

    Comp_C115144:

        ' Sel_Baixas
        sCurrComponent = "Sel_Baixas"
        QueryC115144 = con.CreateCommand()
        QueryC115144.CommandText = QueryC115144.CommandText & " " & "Select  count(*)"
        QueryC115144.CommandText = QueryC115144.CommandText & " " & "from AKBUSER01.WF_DUPL_RECEBER_BAIXAS "
        QueryC115144.CommandText = QueryC115144.CommandText & " " & "where WF_DUPL_RECEBER_BAIXAS.Ident_Duplicata = " & _ 
ConvertValue(RsC128476(0), C128476DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC115144.CommandText = QueryC115144.CommandText & " " & "and WF_DUPL_RECEBER_BAIXAS.Estornada = 0"
        QueryC115144.Transaction = txn
        RsC115144 = QueryC115144.ExecuteReader()
        Dim iC115144 As Short
        ReDim C115144DataType(RsC115144.FieldCount)
        For iC115144 = 0 to RsC115144.FieldCount - 1
            Select Case RsC115144.GetDataTypeName(iC115144).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C115144DataType(iC115144 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C115144DataType(iC115144 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C115144DataType(iC115144 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC115144
        RsC115144_EOF = Not RsC115144.Read()

        GoTo Comp_C128481

    Comp_C115151:

        ' Fim
        sCurrComponent = "Fim"
        Dim tb_C115151 As DataTable = New DataTable()
        tb_C115151.Columns.Add("vRetorno" & "")
        Dim row_C115151 As DataRow = tb_C115151.NewRow()
        row_C115151(0) = C115152
        tb_C115151.Rows.Add(row_C115151)
        R7681 = tb_C115151
        ReDim C115151DataType(1)
        C115151DataType(1) = C115152DataType
        ReturnDataType = C115151DataType

        GoTo Exit_R7681

    Comp_C115152:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C115152 = 1
        C115152DataType = 1
        GoTo Comp_C115020

    Comp_C128476:

        ' Sel_Duplicata
        sCurrComponent = "Sel_Duplicata"
        QueryC128476 = con.CreateCommand()
        QueryC128476.CommandText = QueryC128476.CommandText & " " & "Select WF_DUPLIC_RECEBER.Ident_Duplicata "
        QueryC128476.CommandText = QueryC128476.CommandText & " " & "from AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC128476.CommandText = QueryC128476.CommandText & " " & "where WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P23407, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128476.CommandText = QueryC128476.CommandText & " " & "and WF_DUPLIC_RECEBER.Tp_Produto = " & _ 
ConvertValue(P23408, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128476.CommandText = QueryC128476.CommandText & " " & "and WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC128476.CommandText = QueryC128476.CommandText & " " & "and WF_DUPLIC_RECEBER.Ped_Antecipado_Prod = 1"
        QueryC128476.CommandText = QueryC128476.CommandText & " " & ""
        QueryC128476.CommandText = QueryC128476.CommandText & " " & "and WF_DUPLIC_RECEBER.Nota_Fiscal IS NULL"
        QueryC128476.CommandText = QueryC128476.CommandText & " " & ""
        QueryC128476.Transaction = txn
        RsC128476 = QueryC128476.ExecuteReader()
        Dim iC128476 As Short
        ReDim C128476DataType(RsC128476.FieldCount)
        For iC128476 = 0 to RsC128476.FieldCount - 1
            Select Case RsC128476.GetDataTypeName(iC128476).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C128476DataType(iC128476 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C128476DataType(iC128476 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C128476DataType(iC128476 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC128476
        RsC128476_EOF = Not RsC128476.Read()

        GoTo Comp_C128477

    Comp_C128477:

        ' Eof_Duplic
        sCurrComponent = "Eof_Duplic"
        C128477DataType = 4
        C128477 = RsC128476_EOF
        GoTo Comp_C128478

    Comp_C128478:

        ' Eof_Duplic=1
        sCurrComponent = "Eof_Duplic=1"
        C128478 = (fn_ConvertValueCompiled(C128477, C128477DataType, AKB_DecimalPoint, False) = 1)
        C128478DataType = AKBTypeConst.cAKBTypeLogical
        If C128478 Then
            GoTo Comp_C115151
        Else
            GoTo Comp_C115144
        End If

    Comp_C128481:

        ' Sel_Baixas=0
        sCurrComponent = "Sel_Baixas=0"
        C128481 = (fn_ConvertValueCompiled(RsC115144(0), C115144DataType(1), AKB_DecimalPoint, False) = 0)
        C128481DataType = AKBTypeConst.cAKBTypeLogical
        If C128481 Then
            GoTo Comp_C128505
        Else
            GoTo Comp_C203446
        End If

    Comp_C128505:

        ' Up_Duplic
        sCurrComponent = "Up_Duplic"
        QueryC128505 = con.CreateCommand()
        QueryC128505.CommandText = QueryC128505.CommandText & " " & "Update  AKBUSER01.WF_DUPLIC_RECEBER  set  WF_DUPLIC_RECEBER.Duplic_Cancelada = 1"
        QueryC128505.CommandText = QueryC128505.CommandText & " " & "where WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P23407, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128505.CommandText = QueryC128505.CommandText & " " & "and WF_DUPLIC_RECEBER.Tp_Produto = " & _ 
ConvertValue(P23408, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128505.CommandText = QueryC128505.CommandText & " " & "and WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC128505.Transaction = txn
        C128505 = QueryC128505.ExecuteNonQuery()
        C128505DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C115151

    Comp_C148902:

        ' Cancelar Saldo?
        sCurrComponent = "Cancelar Saldo?"
        C148902 = (fn_ConvertValueCompiled(P37792, 4, AKB_DecimalPoint, False) = 1)
        C148902DataType = AKBTypeConst.cAKBTypeLogical
        If C148902 Then
            GoTo Comp_C172993
        Else
            GoTo Comp_C128476
        End If

    Comp_C149124:

        ' MsgPagto
        sCurrComponent = "MsgPagto"
        C149124 = 1
        C149124DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C203460

    Comp_C172993:

        ' Sel_VrDuplicSaldo
        sCurrComponent = "Sel_VrDuplicSaldo"
        QueryC172993 = con.CreateCommand()
        QueryC172993.CommandText = QueryC172993.CommandText & " " & "SELECT NVL ( WF_DUPLIC_RECEBER.Valor_Duplicata,0 )  "
        QueryC172993.CommandText = QueryC172993.CommandText & " " & ""
        QueryC172993.CommandText = QueryC172993.CommandText & " " & "FROM AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC172993.CommandText = QueryC172993.CommandText & " " & ""
        QueryC172993.CommandText = QueryC172993.CommandText & " " & "WHERE WF_DUPLIC_RECEBER.Tp_Produto = " & _ 
ConvertValue(P23408, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC172993.CommandText = QueryC172993.CommandText & " " & "AND        WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P23407, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC172993.CommandText = QueryC172993.CommandText & " " & "AND        WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC172993.CommandText = QueryC172993.CommandText & " " & "AND        WF_DUPLIC_RECEBER.Ped_Antecipado_Prod = 1 "
        QueryC172993.CommandText = QueryC172993.CommandText & " " & "AND        WF_DUPLIC_RECEBER.Nota_Fiscal IS Null "
        QueryC172993.Transaction = txn
        RsC172993 = QueryC172993.ExecuteReader()
        Dim iC172993 As Short
        ReDim C172993DataType(RsC172993.FieldCount)
        For iC172993 = 0 to RsC172993.FieldCount - 1
            Select Case RsC172993.GetDataTypeName(iC172993).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C172993DataType(iC172993 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C172993DataType(iC172993 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C172993DataType(iC172993 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC172993
        RsC172993_EOF = Not RsC172993.Read()

        GoTo Comp_C173005

    Comp_C173005:

        ' Não tem saldo?
        sCurrComponent = "Não tem saldo?"
        C173005 = (fn_ConvertValueCompiled(RsC172993(0), C172993DataType(1), AKB_DecimalPoint, False) = 0)
        C173005DataType = AKBTypeConst.cAKBTypeLogical
        If C173005 Then
            GoTo Comp_C203460
        Else
            GoTo Comp_C149124
        End If

    Comp_C203444:

        ' Sel_TotItens
        sCurrComponent = "Sel_TotItens"
        QueryC203444 = con.CreateCommand()
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "Select  "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "Round(SUM( WF_PED_SEQ_ITENS.Quantidade_Pedida_Original * WF_PEDIDO_ITENS.Preco_Unit * WF_PEDIDO.COTACAO_MOEDA) , 3)"
        QueryC203444.CommandText = QueryC203444.CommandText & " " & ""
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS ,AKBUSER01.WF_PEDIDO_SEQ, "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "              AKBUSER01.WF_PEDIDO , AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & ""
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P23408, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P23407, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO_ITENS.Flag_Pos_Ped NOT IN (8 , 10 , 12 , 299)"
        QueryC203444.CommandText = QueryC203444.CommandText & " " & ""
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido =WF_PEDIDO_SEQ.Pedido "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND  WF_PEDIDO_ITENS.Tp_Produto = WF_PEDIDO_SEQ.Tp_Produto "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & ""
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & ""
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO.Tp_Produto = WF_PED_SEQ_ITENS.Tp_Produto "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO.Pedido = WF_PED_SEQ_ITENS.Pedido "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO.Sequencia_Atual = WF_PED_SEQ_ITENS.Seq "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & ""
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = WF_PED_SEQ_ITENS.Tp_Produto "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = WF_PED_SEQ_ITENS.Pedido "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO_ITENS.Seq_Item = WF_PED_SEQ_ITENS.Seq_Item "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & ""
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "AND WF_PEDIDO_SEQ.Seq = ( Select Max(Pseq.Seq)"
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "                                                                  From AKBUSER01.WF_PEDIDO_SEQ Pseq"
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "                                                                  Where Pseq.Tp_Produto = WF_PEDIDO_SEQ.Tp_Produto "
        QueryC203444.CommandText = QueryC203444.CommandText & " " & "                                                                  And      Pseq.Pedido = WF_PEDIDO_SEQ.Pedido )"
        QueryC203444.Transaction = txn
        RsC203444 = QueryC203444.ExecuteReader()
        Dim iC203444 As Short
        ReDim C203444DataType(RsC203444.FieldCount)
        For iC203444 = 0 to RsC203444.FieldCount - 1
            Select Case RsC203444.GetDataTypeName(iC203444).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C203444DataType(iC203444 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C203444DataType(iC203444 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C203444DataType(iC203444 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC203444
        RsC203444_EOF = Not RsC203444.Read()

        GoTo Comp_C203447

    Comp_C203446:

        ' Sel_VlrDuplic
        sCurrComponent = "Sel_VlrDuplic"
        QueryC203446 = con.CreateCommand()
        QueryC203446.CommandText = QueryC203446.CommandText & " " & "Select WF_DUPLIC_RECEBER.Valor_Duplicata"
        QueryC203446.CommandText = QueryC203446.CommandText & " " & "From AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC203446.CommandText = QueryC203446.CommandText & " " & "Where WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P23407, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC203446.CommandText = QueryC203446.CommandText & " " & "and WF_DUPLIC_RECEBER.Tp_Produto = " & _ 
ConvertValue(P23408, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC203446.CommandText = QueryC203446.CommandText & " " & "and WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC203446.CommandText = QueryC203446.CommandText & " " & "and WF_DUPLIC_RECEBER.Ped_Antecipado_Prod = 1"
        QueryC203446.CommandText = QueryC203446.CommandText & " " & "and WF_DUPLIC_RECEBER.Nota_Fiscal IS Null"
        QueryC203446.Transaction = txn
        RsC203446 = QueryC203446.ExecuteReader()
        Dim iC203446 As Short
        ReDim C203446DataType(RsC203446.FieldCount)
        For iC203446 = 0 to RsC203446.FieldCount - 1
            Select Case RsC203446.GetDataTypeName(iC203446).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C203446DataType(iC203446 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C203446DataType(iC203446 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C203446DataType(iC203446 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC203446
        RsC203446_EOF = Not RsC203446.Read()

        GoTo Comp_C203444

    Comp_C203447:

        ' Calc_AntecProd
        sCurrComponent = "Calc_AntecProd"
        C203447 = ( fn_ConvertValueCompiled(RsC203446(0), C203446DataType(1), AKB_DecimalPoint, True) *100 ) / fn_ConvertValueCompiled(RsC203444(0), C203444DataType(1), AKB_DecimalPoint, True)
        C203447DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C203448

    Comp_C203448:

        ' Up_%AntecProd
        sCurrComponent = "Up_%AntecProd"
        QueryC203448 = con.CreateCommand()
        QueryC203448.CommandText = QueryC203448.CommandText & " " & "Update AKBUSER01.WF_PEDIDO "
        QueryC203448.CommandText = QueryC203448.CommandText & " " & "Set WF_PEDIDO.Porc_Pg_Antec_Prod = " & _ 
ConvertValue(C203447, C203447DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC203448.CommandText = QueryC203448.CommandText & " " & "WF_PEDIDO.Valor_Antec_Prod = " & _ 
ConvertValue(RsC203446(0), C203446DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC203448.CommandText = QueryC203448.CommandText & " " & ""
        QueryC203448.CommandText = QueryC203448.CommandText & " " & "Where WF_PEDIDO.Tp_Produto  = " & _ 
ConvertValue(P23408, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC203448.CommandText = QueryC203448.CommandText & " " & "And      WF_PEDIDO.Pedido = " & _ 
ConvertValue(P23407, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC203448.CommandText = QueryC203448.CommandText & " " & ""
        QueryC203448.Transaction = txn
        C203448 = QueryC203448.ExecuteNonQuery()
        C203448DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C203460

    Comp_C203460:

        ' vRet=-1
        sCurrComponent = "vRet=-1"
        C203460DataType = 4
        C115152 = fn_ConvertValueCompiled(-1, 1, AKB_DecimalPoint)
        C203460 = True
        GoTo Comp_C115151

    Exit_R7681:

        Exit Function

    Err_R7681:
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
