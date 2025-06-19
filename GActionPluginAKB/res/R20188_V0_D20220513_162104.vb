Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R20188

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

    Public Shared Function R20188(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P76391 As Double, ByVal P76392 As Double, ByVal P76393 As Double, ByVal P76394 As Object) As DataTable
    ' 
    ' 20188 - Geração de Pedido/Seq-Itens (Tela de Pedido) - New - Versão: 0
    ' 
        'On Error GoTo Err_R20188
        Dim sCurrComponent as String

        Dim QueryC468039 As New Object
        Dim C468039 As Integer
        Dim C468039DataType As Short
        Dim C468040 As Double
        Dim C468040DataType As Short
        Dim QueryC468041 As New Object
        Dim C468041 As Integer
        Dim C468041DataType As Short
        Dim C468042DataType() As Short
        Dim QueryC468043 As New Object
        Dim RsC468043 As New Object
        Dim C468043DataType() As Short
        Dim RsC468043_EOF As Boolean
        Dim C468044 As Object
        Dim C468044DataType As Short
        Dim C468045 As Object
        Dim C468045DataType As Short
        Dim C468046 As Object
        Dim C468046DataType As Short
        Dim C468047 As Boolean
        Dim C468047DataType As Short
        Dim C468048 As Object
        Dim C468048DataType As Short
        Dim C468049 As Object
        Dim C468049DataType As Short
        Dim QueryC468050 As New Object
        Dim C468050 As Integer
        Dim C468050DataType As Short
        Dim QueryC468051 As New Object
        Dim RsC468051 As New Object
        Dim C468051DataType() As Short
        Dim RsC468051_EOF As Boolean
        Dim C468052 As Object
        Dim C468052DataType As Short
        Dim C468053 As Boolean
        Dim C468053DataType As Short
        Dim QueryC468054 As New Object
        Dim C468054 As Integer
        Dim C468054DataType As Short
        Dim C468055 As Object
        Dim C468055DataType As Short
        Dim C468056 As Object
        Dim C468056DataType As Short
        Dim QueryC468057 As New Object
        Dim C468057 As Integer
        Dim C468057DataType As Short
        Dim QueryC468058 As New Object
        Dim RsC468058 As New Object
        Dim C468058DataType() As Short
        Dim RsC468058_EOF As Boolean
        Dim C468059 As DataTable
        Dim C468059CurrentRow As Integer
        Dim C468059DataType() As Short
        Dim QueryC468060 As New Object
        Dim C468060 As Integer
        Dim C468060DataType As Short
        Dim C468061 As Object
        Dim C468061DataType As Short
        Dim C468062 As Object
        Dim C468062DataType As Short
        Dim QueryC468063 As New Object
        Dim C468063 As Integer
        Dim C468063DataType As Short
        Dim QueryC543265 As New Object
        Dim RsC543265 As New Object
        Dim C543265DataType() As Short
        Dim RsC543265_EOF As Boolean
        Dim QueryC543267 As New Object
        Dim C543267 As Integer
        Dim C543267DataType As Short
        Dim C543340 As Object
        Dim C543340DataType As Short
        Dim C543341 As Boolean
        Dim C543341DataType As Short
        Dim C552640 As Object
        Dim C552640DataType As Short
        Dim C552641 As Object
        Dim C552641DataType As Short
        P76394 = IIf(IsDBNull(P76394), 0, P76394)

        ReDim ReturnDatatype(0)

        GoTo Comp_C468040

    Comp_C468039:

        ' Del_Seq-1
        sCurrComponent = "Del_Seq-1"
        QueryC468039 = con.CreateCommand()
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "DELETE FROM AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "WHERE WF_PED_SEQ_ITENS.Tp_Produto = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "AND WF_PED_SEQ_ITENS.Pedido = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "AND WF_PED_SEQ_ITENS.Qtde_Reservada = 0 "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "AND WF_PED_SEQ_ITENS.Seq IN (" & _ 
ConvertValue(P76393, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C468040, C468040DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC468039.CommandText = QueryC468039.CommandText & " " & ""
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "AND NOT EXISTS "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "(SELECT * "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "FROM AKBUSER01.WF_PED_ITEM_ESTOQUE "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "WHERE WF_PED_ITEM_ESTOQUE.Item_do_Pedido = WF_PED_SEQ_ITENS.Item_do_Pedido "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "AND WF_PED_ITEM_ESTOQUE.Complementado = 0 )"
        QueryC468039.CommandText = QueryC468039.CommandText & " " & ""
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "AND NOT EXISTS "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "(SELECT * "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "FROM AKBUSER01.WF_NFS_Itens_Agrupamento "
        QueryC468039.CommandText = QueryC468039.CommandText & " " & "WHERE WF_NFS_Itens_Agrupamento.Item_do_Pedido = WF_PED_SEQ_ITENS.Item_do_Pedido )"
        QueryC468039.CommandText = QueryC468039.CommandText & " " & ""
        QueryC468039.Transaction = txn
        C468039 = QueryC468039.ExecuteNonQuery()
        C468039DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C468041

    Comp_C468040:

        ' Seq - 1
        sCurrComponent = "Seq - 1"
        C468040 = fn_ConvertValueCompiled(P76393, 1, AKB_DecimalPoint, True) - 1
        C468040DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C543265

    Comp_C468041:

        ' Seq-Itens
        sCurrComponent = "Seq-Itens"
        QueryC468041 = con.CreateCommand()
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "( WF_PED_SEQ_ITENS.Item_do_Pedido , WF_PED_SEQ_ITENS.Tp_Produto , WF_PED_SEQ_ITENS.Pedido , "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "WF_PED_SEQ_ITENS.Seq , WF_PED_SEQ_ITENS.Tp_Produto2 , WF_PED_SEQ_ITENS.Pedido2 ,"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & " WF_PED_SEQ_ITENS.Seq_Item , "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "WF_PED_SEQ_ITENS.Qtde_Reservada, WF_PED_SEQ_ITENS.Quantidade_Pedida_Original  ) "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "SELECT "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "SEQ_WF_PED_SEQ_ITENS.NEXTVAL , " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P76393, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "" & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , WF_PEDIDO_ITENS.Seq_Item , 0 , "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Pedido - WF_PEDIDO_ITENS.Qtde_Atendida"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & " "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS ,  AKBUSER01.WF_SIGLA_PROD_AGRUP "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PEDIDO_ITENS.Flag_Pos_Ped  NOT IN (7, 8, 9, 10, 11, 12, 13, 299)"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_SIGLA_PROD_AGRUP.Sigla_Prod =WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND NOT EXISTS"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "(SELECT * FROM AKBUSER01.WF_PED_SEQ_ITENS"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "WHERE  WF_PED_SEQ_ITENS.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND   WF_PED_SEQ_ITENS.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND   WF_PED_SEQ_ITENS.Seq = " & _ 
ConvertValue(P76393, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND   WF_PED_SEQ_ITENS.Tp_Produto2 = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND   WF_PED_SEQ_ITENS.Pedido2 = WF_PEDIDO_ITENS.Pedido  "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND   WF_PED_SEQ_ITENS.Seq_Item = WF_PEDIDO_ITENS.Seq_Item)"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND NOT EXISTS"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "(SELECT * FROM AKBUSER01.WF_PED_ITEM_ESTOQUE, AKBUSER01.WF_PED_SEQ_ITENS"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "WHERE  WF_PED_SEQ_ITENS.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_SEQ_ITENS.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_SEQ_ITENS.Seq = " & _ 
ConvertValue(C468040, C468040DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_SEQ_ITENS.Tp_Produto2 = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_SEQ_ITENS.Pedido2 = WF_PEDIDO_ITENS.Pedido  "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_SEQ_ITENS.Seq_Item = WF_PEDIDO_ITENS.Seq_Item"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_ITEM_ESTOQUE.Item_do_Pedido = WF_PED_SEQ_ITENS.Item_do_Pedido "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_ITEM_ESTOQUE.Complementado = 0 )"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND ((" & _ 
ConvertValue(P76394, 4, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 1 "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND (WF_PEDIDO_ITENS.Qtde_Atendida < ((WF_PEDIDO_ITENS.Qtde_Pedido ) - "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "(WF_PEDIDO_ITENS.Qtde_Pedido * WF_SIGLA_PROD_AGRUP.Tolerancia_Atendimento  / 100))"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "OR WF_PEDIDO_ITENS.Qtde_Atendida is null"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "OR EXISTS"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "(SELECT *"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "FROM AKBUSER01.WF_PED_ITEM_OP_ABERTO"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "WHERE WF_PED_ITEM_OP_ABERTO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_ITEM_OP_ABERTO.Pedido = WF_PEDIDO_ITENS.Pedido  "
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "AND WF_PED_ITEM_OP_ABERTO.Seq_Item = WF_PEDIDO_ITENS.Seq_Item)"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "))"
        QueryC468041.CommandText = QueryC468041.CommandText & " " & ""
        QueryC468041.CommandText = QueryC468041.CommandText & " " & "OR " & _ 
ConvertValue(P76394, 4, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " =0)"
        QueryC468041.Transaction = txn
        C468041 = QueryC468041.ExecuteNonQuery()
        C468041DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C468043

    Comp_C468042:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C468042 As DataTable = New DataTable()
        tb_C468042.Columns.Add("EofItem2=T" & "")
        Dim row_C468042 As DataRow = tb_C468042.NewRow()
        row_C468042(0) = C468047
        tb_C468042.Rows.Add(row_C468042)
        R20188 = tb_C468042
        ReDim C468042DataType(1)
        C468042DataType(1) = C468047DataType
        ReturnDataType = C468042DataType

        GoTo Exit_R20188

    Comp_C468043:

        ' Itens2
        sCurrComponent = "Itens2"
        QueryC468043 = con.CreateCommand()
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "SELECT "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "WF_PED_SEQ_ITENS.Seq_Item , WF_PED_SEQ_ITENS.Item_do_Pedido ,"
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Pedido "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & ""
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_ITENS , AKBUSER01.WF_PED_ITEM_ESTOQUE, AKBUSER01.WF_PEDIDO_ITENS"
        QueryC468043.CommandText = QueryC468043.CommandText & " " & ""
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & ""
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "AND WF_PED_SEQ_ITENS.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "AND WF_PED_SEQ_ITENS.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "AND WF_PED_SEQ_ITENS.Seq_Item = WF_PEDIDO_ITENS.Seq_Item "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & ""
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "AND WF_PED_SEQ_ITENS.Seq IN (" & _ 
ConvertValue(P76393, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C468040, C468040DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ) "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & ""
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "AND WF_PED_ITEM_ESTOQUE.Item_do_Pedido = WF_PED_SEQ_ITENS.Item_do_Pedido "
        QueryC468043.CommandText = QueryC468043.CommandText & " " & "AND WF_PED_ITEM_ESTOQUE.Complementado = 0 "
        QueryC468043.Transaction = txn
        RsC468043 = QueryC468043.ExecuteReader()
        Dim iC468043 As Short
        ReDim C468043DataType(RsC468043.FieldCount)
        For iC468043 = 0 to RsC468043.FieldCount - 1
            Select Case RsC468043.GetDataTypeName(iC468043).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C468043DataType(iC468043 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C468043DataType(iC468043 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C468043DataType(iC468043 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC468043
        RsC468043_EOF = Not RsC468043.Read()

        GoTo Comp_C468046

    Comp_C468044:

        ' ItemPedido2
        sCurrComponent = "ItemPedido2"
        C468044DataType = 0
        C468044DataType = C468043Datatype(2)
        If C468044DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC468043(1)) Then
          C468044 = Strings.RTrim(RsC468043(1))
        Else 
          C468044 = RsC468043(1)
        End If 

        GoTo Comp_C468055

    Comp_C468045:

        ' SeqItem2
        sCurrComponent = "SeqItem2"
        C468045DataType = 0
        C468045 = RsC468043(0)
        C468045DataType = C468043Datatype(1)
        If C468045DataType = AKBTypeConst.cAKBTypeString Then
          C468045 = IIF(IsDBNull(C468045), C468045, Strings.RTrim(C468045))
        End If 

        GoTo Comp_C468063

    Comp_C468046:

        ' EofItem2
        sCurrComponent = "EofItem2"
        C468046DataType = 4
        C468046 = RsC468043_EOF
        GoTo Comp_C468047

    Comp_C468047:

        ' EofItem2=T
        sCurrComponent = "EofItem2=T"
        C468047 = (fn_ConvertValueCompiled(C468046, C468046DataType, AKB_DecimalPoint, False) =1)
        C468047DataType = AKBTypeConst.cAKBTypeLogical
        If C468047 Then
            GoTo Comp_C468042
        Else
            GoTo Comp_C468044
        End If

    Comp_C468048:

        ' NextItem2
        sCurrComponent = "NextItem2"
        C468048DataType = 4
        RsC468043_EOF = Not RsC468043.Read()
        C468048 = True

        GoTo Comp_C468046

    Comp_C468049:

        ' Pedido2
        sCurrComponent = "Pedido2"
        C468049DataType = 0
        C468049DataType = C468043Datatype(3)
        If C468049DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC468043(2)) Then
          C468049 = Strings.RTrim(RsC468043(2))
        Else 
          C468049 = RsC468043(2)
        End If 

        GoTo Comp_C468059

    Comp_C468050:

        ' UpQtdeReservada
        sCurrComponent = "UpQtdeReservada"
        QueryC468050 = con.CreateCommand()
        QueryC468050.CommandText = QueryC468050.CommandText & " " & "UPDATE AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC468050.CommandText = QueryC468050.CommandText & " " & ""
        QueryC468050.CommandText = QueryC468050.CommandText & " " & "SET WF_PED_SEQ_ITENS.Quantidade_Pedida_Original = " & _ 
ConvertValue(RsC468058(0), C468058DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , "
        QueryC468050.CommandText = QueryC468050.CommandText & " " & "WF_PED_SEQ_ITENS.Qtde_Reservada = "
        QueryC468050.CommandText = QueryC468050.CommandText & " " & "(SELECT nvl(SUM(WF_PED_ITEM_ESTOQUE.Quant_Mov_Interna_1 ),0) "
        QueryC468050.CommandText = QueryC468050.CommandText & " " & "FROM AKBUSER01.WF_PED_ITEM_ESTOQUE "
        QueryC468050.CommandText = QueryC468050.CommandText & " " & "WHERE WF_PED_ITEM_ESTOQUE.Item_do_Pedido =  " & _ 
ConvertValue(C468055, C468055DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC468050.CommandText = QueryC468050.CommandText & " " & ""
        QueryC468050.CommandText = QueryC468050.CommandText & " " & "WHERE WF_PED_SEQ_ITENS.Item_do_Pedido = " & _ 
ConvertValue(C468055, C468055DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468050.Transaction = txn
        C468050 = QueryC468050.ExecuteNonQuery()
        C468050DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C468048

    Comp_C468051:

        ' SelItemPedido
        sCurrComponent = "SelItemPedido"
        QueryC468051 = con.CreateCommand()
        QueryC468051.CommandText = QueryC468051.CommandText & " " & "SELECT WF_PED_SEQ_ITENS.Item_do_Pedido"
        QueryC468051.CommandText = QueryC468051.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC468051.CommandText = QueryC468051.CommandText & " " & "WHERE WF_PED_SEQ_ITENS.Tp_Produto = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468051.CommandText = QueryC468051.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Pedido = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468051.CommandText = QueryC468051.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Seq = " & _ 
ConvertValue(P76393, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468051.CommandText = QueryC468051.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Tp_Produto2 = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468051.CommandText = QueryC468051.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Pedido2 = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468051.CommandText = QueryC468051.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Seq_Item = " & _ 
ConvertValue(C468045, C468045DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468051.Transaction = txn
        RsC468051 = QueryC468051.ExecuteReader()
        Dim iC468051 As Short
        ReDim C468051DataType(RsC468051.FieldCount)
        For iC468051 = 0 to RsC468051.FieldCount - 1
            Select Case RsC468051.GetDataTypeName(iC468051).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C468051DataType(iC468051 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C468051DataType(iC468051 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C468051DataType(iC468051 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC468051
        RsC468051_EOF = Not RsC468051.Read()

        GoTo Comp_C468052

    Comp_C468052:

        ' EofItemPedido
        sCurrComponent = "EofItemPedido"
        C468052DataType = 4
        C468052 = RsC468051_EOF
        GoTo Comp_C468058

    Comp_C468053:

        ' EofItemPedido=1
        sCurrComponent = "EofItemPedido=1"
        C468053 = (fn_ConvertValueCompiled(C468052, C468052DataType, AKB_DecimalPoint, False) = 1)
        C468053DataType = AKBTypeConst.cAKBTypeLogical
        If C468053 Then
            GoTo Comp_C468061
        Else
            GoTo Comp_C468056
        End If

    Comp_C468054:

        ' UpPedEstoque
        sCurrComponent = "UpPedEstoque"
        QueryC468054 = con.CreateCommand()
        QueryC468054.CommandText = QueryC468054.CommandText & " " & "UPDATE AKBUSER01.WF_PED_ITEM_ESTOQUE "
        QueryC468054.CommandText = QueryC468054.CommandText & " " & "SET WF_PED_ITEM_ESTOQUE.Item_do_Pedido = " & _ 
ConvertValue(C468055, C468055DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC468054.CommandText = QueryC468054.CommandText & " " & "WHERE WF_PED_ITEM_ESTOQUE.Item_do_Pedido = " & _ 
ConvertValue(C468044, C468044DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468054.CommandText = QueryC468054.CommandText & " " & "AND  WF_PED_ITEM_ESTOQUE.Complementado = 0 "
        QueryC468054.Transaction = txn
        C468054 = QueryC468054.ExecuteNonQuery()
        C468054DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C468050

    Comp_C468055:

        ' VItemPedido
        sCurrComponent = "VItemPedido"
        C468055 = C468044 
        C468055DataType = 1
        GoTo Comp_C468045

    Comp_C468056:

        ' AttVItemPedido
        sCurrComponent = "AttVItemPedido"
        C468056DataType = 4
        C468055 = fn_ConvertValueCompiled(RsC468051(0) , 1, AKB_DecimalPoint)
        C468056 = True
        GoTo Comp_C468057

    Comp_C468057:

        ' UpReservaAnterior
        sCurrComponent = "UpReservaAnterior"
        QueryC468057 = con.CreateCommand()
        QueryC468057.CommandText = QueryC468057.CommandText & " " & "UPDATE AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC468057.CommandText = QueryC468057.CommandText & " " & ""
        QueryC468057.CommandText = QueryC468057.CommandText & " " & "SET WF_PED_SEQ_ITENS.Qtde_Reservada = "
        QueryC468057.CommandText = QueryC468057.CommandText & " " & "(SELECT nvl( SUM( WF_PED_ITEM_ESTOQUE.Quant_Mov_Interna_1 ),0) "
        QueryC468057.CommandText = QueryC468057.CommandText & " " & "FROM AKBUSER01.WF_PED_ITEM_ESTOQUE "
        QueryC468057.CommandText = QueryC468057.CommandText & " " & "WHERE WF_PED_ITEM_ESTOQUE.Item_do_Pedido =  " & _ 
ConvertValue(C468044, C468044DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC468057.CommandText = QueryC468057.CommandText & " " & ""
        QueryC468057.CommandText = QueryC468057.CommandText & " " & "WHERE WF_PED_SEQ_ITENS.Item_do_Pedido = " & _ 
ConvertValue(C468044, C468044DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468057.Transaction = txn
        C468057 = QueryC468057.ExecuteNonQuery()
        C468057DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C468054

    Comp_C468058:

        ' SelSaldo
        sCurrComponent = "SelSaldo"
        QueryC468058 = con.CreateCommand()
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "SELECT WF_PEDIDO_ITENS.Qtde_Pedido - NVL( AUX.QTDE , 0 )"
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS , "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "(SELECT SUM( WF_PED_ITEM_ESTOQUE.Quant_Mov_Interna_1 ) QTDE"
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "FROM  AKBUSER01.WF_PED_SEQ_ITENS , AKBUSER01.WF_PED_ITEM_ESTOQUE "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "WHERE WF_PED_SEQ_ITENS.Seq < " & _ 
ConvertValue(P76393, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Tp_Produto2 = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Pedido2 = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Seq_Item = " & _ 
ConvertValue(C468045, C468045DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Item_do_Pedido = WF_PED_ITEM_ESTOQUE.Item_do_Pedido "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "AND  WF_PED_ITEM_ESTOQUE.Complementado = 1 ) AUX"
        QueryC468058.CommandText = QueryC468058.CommandText & " " & ""
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468058.CommandText = QueryC468058.CommandText & " " & "AND  WF_PEDIDO_ITENS.Seq_Item = " & _ 
ConvertValue(C468045, C468045DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468058.Transaction = txn
        RsC468058 = QueryC468058.ExecuteReader()
        Dim iC468058 As Short
        ReDim C468058DataType(RsC468058.FieldCount)
        For iC468058 = 0 to RsC468058.FieldCount - 1
            Select Case RsC468058.GetDataTypeName(iC468058).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C468058DataType(iC468058 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C468058DataType(iC468058 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C468058DataType(iC468058 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC468058
        RsC468058_EOF = Not RsC468058.Read()

        GoTo Comp_C468053

    Comp_C468059:

        ' FatorConversão
        sCurrComponent = "FatorConversão"
        C468059 = clsRuleDynamicallyCompiled_R4584.R4584(con, txn, IIf(Not IsDBNull(P76391), P76391, System.DBNull.Value), IIf(Not IsDBNull(C468045), C468045, System.DBNull.Value), IIf(Not IsDBNull(P76392), P76392, System.DBNull.Value))
        C468059CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C468059) Then
          iColumns = C468059.Columns.Count
        End If
        ReDim C468059DataType(iColumns)
        For iCol = 1 To iColumns
            C468059DataType(iCol) = clsRuleDynamicallyCompiled_R4584.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C468051

    Comp_C468060:

        ' InSeqItens
        sCurrComponent = "InSeqItens"
        QueryC468060 = con.CreateCommand()
        QueryC468060.CommandText = QueryC468060.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_SEQ_ITENS ( WF_PED_SEQ_ITENS.Item_do_Pedido , WF_PED_SEQ_ITENS.Tp_Produto , WF_PED_SEQ_ITENS.Pedido , WF_PED_SEQ_ITENS.Seq , WF_PED_SEQ_ITENS.Tp_Produto2 , WF_PED_SEQ_ITENS.Pedido2 , WF_PED_SEQ_ITENS.Seq_Item , WF_PED_SEQ_ITENS.Qtde_Reservada , WF_PED_SEQ_ITENS.Quantidade_Pedida_Original ) "
        QueryC468060.CommandText = QueryC468060.CommandText & " " & "VALUES( " & _ 
ConvertValue(C468061, C468061DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P76393, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C468045, C468045DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 0 , " & _ 
ConvertValue(RsC468058(0), C468058DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC468060.Transaction = txn
        C468060 = QueryC468060.ExecuteNonQuery()
        C468060DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C468054

    Comp_C468061:

        ' NewItemPedido
        sCurrComponent = "NewItemPedido"
        C468061DataType = 1
        C468061 = ObjTable_NewIdentity (con, "WF_PED_SEQ_ITENS")
        GoTo Comp_C468062

    Comp_C468062:

        ' AttVItemPedido2
        sCurrComponent = "AttVItemPedido2"
        C468062DataType = 4
        C468055 = fn_ConvertValueCompiled(C468061 , 1, AKB_DecimalPoint)
        C468062 = True
        GoTo Comp_C468060

    Comp_C468063:

        ' Up_QtdeReservada
        sCurrComponent = "Up_QtdeReservada"
        QueryC468063 = con.CreateCommand()
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "Update AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "set WF_PED_SEQ_ITENS.Qtde_Reservada = "
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "NVL(( Select NVL(sum(WF_PED_ITEM_ESTOQUE.Quant_Mov_Interna_1),0)"
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "from AKBUSER01.WF_PED_ITEM_ESTOQUE "
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "where WF_PED_ITEM_ESTOQUE.Item_do_Pedido = WF_PED_SEQ_ITENS.Item_do_Pedido"
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "and WF_PED_ITEM_ESTOQUE.Complementado = 1 ), 0 )"
        QueryC468063.CommandText = QueryC468063.CommandText & " " & ""
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "where WF_PED_SEQ_ITENS.Pedido = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "and WF_PED_SEQ_ITENS.Tp_Produto = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "and WF_PED_SEQ_ITENS.Seq_Item = " & _ 
ConvertValue(C468045, C468045DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468063.CommandText = QueryC468063.CommandText & " " & "and WF_PED_SEQ_ITENS.Item_do_Pedido = " & _ 
ConvertValue(C468044, C468044DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC468063.Transaction = txn
        C468063 = QueryC468063.ExecuteNonQuery()
        C468063DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C468049

    Comp_C543265:

        ' selItemDoPedido
        sCurrComponent = "selItemDoPedido"
        QueryC543265 = con.CreateCommand()
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "SELECT WF_PED_SEQ_ITENS.Item_do_Pedido"
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "WHERE WF_PED_SEQ_ITENS.Tp_Produto = " & _ 
ConvertValue(P76392, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "AND WF_PED_SEQ_ITENS.Pedido = " & _ 
ConvertValue(P76391, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "AND WF_PED_SEQ_ITENS.Qtde_Reservada = 0 "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "AND WF_PED_SEQ_ITENS.Seq IN (" & _ 
ConvertValue(P76393, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C468040, C468040DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC543265.CommandText = QueryC543265.CommandText & " " & ""
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "AND NOT EXISTS "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "(SELECT * "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "FROM AKBUSER01.WF_PED_ITEM_ESTOQUE "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "WHERE WF_PED_ITEM_ESTOQUE.Item_do_Pedido = WF_PED_SEQ_ITENS.Item_do_Pedido "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "AND WF_PED_ITEM_ESTOQUE.Complementado = 0 )"
        QueryC543265.CommandText = QueryC543265.CommandText & " " & ""
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "AND NOT EXISTS "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "(SELECT * "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "FROM AKBUSER01.WF_NFS_Itens_Agrupamento "
        QueryC543265.CommandText = QueryC543265.CommandText & " " & "WHERE WF_NFS_Itens_Agrupamento.Item_do_Pedido = WF_PED_SEQ_ITENS.Item_do_Pedido )"
        QueryC543265.Transaction = txn
        RsC543265 = QueryC543265.ExecuteReader()
        Dim iC543265 As Short
        ReDim C543265DataType(RsC543265.FieldCount)
        For iC543265 = 0 to RsC543265.FieldCount - 1
            Select Case RsC543265.GetDataTypeName(iC543265).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C543265DataType(iC543265 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C543265DataType(iC543265 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C543265DataType(iC543265 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC543265
        RsC543265_EOF = Not RsC543265.Read()

        GoTo Comp_C543340

    Comp_C543267:

        ' delNotaEntrada
        sCurrComponent = "delNotaEntrada"
        QueryC543267 = con.CreateCommand()
        QueryC543267.CommandText = QueryC543267.CommandText & " " & "DELETE"
        QueryC543267.CommandText = QueryC543267.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_NFE_ITENS"
        QueryC543267.CommandText = QueryC543267.CommandText & " " & "WHERE WF_PED_SEQ_NFE_ITENS.Item_do_Pedido = " & _ 
ConvertValue(C552640, C552640DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC543267.Transaction = txn
        C543267 = QueryC543267.ExecuteNonQuery()
        C543267DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C552641

    Comp_C543340:

        ' EOF ItemPedidoNota
        sCurrComponent = "EOF ItemPedidoNota"
        C543340DataType = 4
        C543340 = RsC543265_EOF
        GoTo Comp_C543341

    Comp_C543341:

        ' Tem ItemPedidoNota?
        sCurrComponent = "Tem ItemPedidoNota?"
        C543341 = (fn_ConvertValueCompiled(C543340, C543340DataType, AKB_DecimalPoint, False) = 0)
        C543341DataType = AKBTypeConst.cAKBTypeLogical
        If C543341 Then
            GoTo Comp_C552640
        Else
            GoTo Comp_C468039
        End If

    Comp_C552640:

        ' vl_item_pedido
        sCurrComponent = "vl_item_pedido"
        C552640DataType = 0
        C552640 = RsC543265(0)
        C552640DataType = C543265Datatype(1)
        If C552640DataType = AKBTypeConst.cAKBTypeString Then
          C552640 = IIF(IsDBNull(C552640), C552640, Strings.RTrim(C552640))
        End If 

        GoTo Comp_C543267

    Comp_C552641:

        ' next
        sCurrComponent = "next"
        C552641DataType = 4
        RsC543265_EOF = Not RsC543265.Read()
        C552641 = True

        GoTo Comp_C543340

    Exit_R20188:

        Exit Function

    Err_R20188:
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
