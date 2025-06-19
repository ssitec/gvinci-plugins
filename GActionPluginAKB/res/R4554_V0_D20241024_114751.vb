Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R4554

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

    Public Shared Function R4554(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P12594 As Double, ByVal P20448 As Object, ByVal P24035 As Double, ByVal P24036 As Object, ByVal P24037 As Double, ByVal P92047 As Object) As DataTable
    ' 
    ' 4554 - Apaga Pré Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R4554
        Dim sCurrComponent as String

        Dim C54073 As Boolean
        Dim C54073DataType As Short
        Dim QueryC54075 As New Object
        Dim C54075 As Integer
        Dim C54075DataType As Short
        Dim C54076 As Boolean
        Dim C54076DataType As Short
        Dim C54077 As Object
        Dim C54077DataType As Short
        Dim C54078DataType() As Short
        Dim QueryC56889 As New Object
        Dim C56889 As Integer
        Dim C56889DataType As Short
        Dim QueryC56890 As New Object
        Dim C56890 As Integer
        Dim C56890DataType As Short
        Dim QueryC56891 As New Object
        Dim C56891 As Integer
        Dim C56891DataType As Short
        Dim C56894 As DataTable
        Dim C56894CurrentRow As Integer
        Dim C56894DataType() As Short
        Dim QueryC56895 As New Object
        Dim RsC56895 As New Object
        Dim C56895DataType() As Short
        Dim RsC56895_EOF As Boolean
        Dim C56896 As Object
        Dim C56896DataType As Short
        Dim C56897 As Object
        Dim C56897DataType As Short
        Dim C56898 As Boolean
        Dim C56898DataType As Short
        Dim C96650 As Boolean
        Dim C96650DataType As Short
        Dim C96651 As Boolean
        Dim C96651DataType As Short
        Dim C119216 As Short
        Dim C119216DataType As Short
        Dim C119216ReturnDataType() As Short

        Dim C119217 As Boolean
        Dim C119217DataType As Short
        Dim QueryC119218 As New Object
        Dim C119218 As Integer
        Dim C119218DataType As Short
        Dim C119219  As Object
        Dim C119219DataType As Short
        Dim C119219ReturnDataType() As Short

        Dim C119220 As Boolean
        Dim C119220DataType As Short
        Dim C119221 As Short
        Dim C119221DataType As Short
        Dim C119221ReturnDataType() As Short

        Dim C119228 As Object
        Dim C119228DataType As Short
        Dim C119291 As Short
        Dim C119291DataType As Short
        Dim C119291ReturnDataType() As Short

        Dim QueryC128460 As New Object
        Dim RsC128460 As New Object
        Dim C128460DataType() As Short
        Dim RsC128460_EOF As Boolean
        Dim C128461 As Object
        Dim C128461DataType As Short
        Dim C128462 As Boolean
        Dim C128462DataType As Short
        Dim C128463 As Short
        Dim C128463DataType As Short
        Dim C128463ReturnDataType() As Short

        Dim QueryC556489 As New Object
        Dim RsC556489 As New Object
        Dim C556489DataType() As Short
        Dim RsC556489_EOF As Boolean
        Dim C556490 As Boolean
        Dim C556490DataType As Short
        Dim C556491 As Object
        Dim C556491DataType As Short
        Dim C556492 As Object
        Dim C556492DataType As Short
        Dim C556493 As Object
        Dim C556493DataType As Short
        Dim C556494 As Object
        Dim C556494DataType As Short
        Dim QueryC556497 As New Object
        Dim C556497 As Integer
        Dim C556497DataType As Short
        Dim QueryC556498 As New Object
        Dim C556498 As Integer
        Dim C556498DataType As Short
        Dim QueryC556499 As New Object
        Dim RsC556499 As New Object
        Dim C556499DataType() As Short
        Dim RsC556499_EOF As Boolean
        Dim C556500 As Object
        Dim C556500DataType As Short
        Dim QueryC556501 As New Object
        Dim C556501 As Integer
        Dim C556501DataType As Short
        Dim C556543 As Boolean
        Dim C556543DataType As Short
        Dim C561354 As Object
        Dim C561354DataType As Short
        Dim QueryC567753 As New Object
        Dim C567753 As Integer
        Dim C567753DataType As Short
        Dim C571830 As Boolean
        Dim C571830DataType As Short
        Dim C571831 As Short
        Dim C571831DataType As Short
        Dim C571831ReturnDataType() As Short

        Dim C579496 As Object
        Dim C579496DataType As Short
        Dim C579497 As Object
        Dim C579497DataType As Short
        Dim C579498 As Boolean
        Dim C579498DataType As Short
        Dim C579499 As Object
        Dim C579499DataType As Short
        Dim C579500 As Object
        Dim C579500DataType As Short
        Dim C579501 As Object
        Dim C579501DataType As Short
        P20448 = IIf(IsDBNull(P20448), 0, P20448)

        ReDim ReturnDatatype(0)

        GoTo Comp_C54077

    Comp_C54073:

        ' Begin
        sCurrComponent = "Begin"
        txn = con.BeginTransaction
        C54073 = True
        C54073DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C119228

    Comp_C54075:

        ' Pedido
        sCurrComponent = "Pedido"
        QueryC54075 = con.CreateCommand()
        QueryC54075.CommandText = QueryC54075.CommandText & " " & "DELETE FROM AKBUSER01.WF_PRE_PEDIDO WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC54075.Transaction = txn
        C54075 = QueryC54075.ExecuteNonQuery()
        C54075DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C96651

    Comp_C54076:

        ' Commit
        sCurrComponent = "Commit"
        txn.Commit()
        C54076 = True
        C54076DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C119291

    Comp_C54077:

        ' VTrue
        sCurrComponent = "VTrue"
        C54077 = 1
        C54077DataType = 4
        GoTo Comp_C579497

    Comp_C54078:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C54078 As DataTable = New DataTable()
        tb_C54078.Columns.Add("VTrue" & "")
        Dim row_C54078 As DataRow = tb_C54078.NewRow()
        row_C54078(0) = C54077
        tb_C54078.Rows.Add(row_C54078)
        R4554 = tb_C54078
        ReDim C54078DataType(1)
        C54078DataType(1) = C54077DataType
        ReturnDataType = C54078DataType

        GoTo Exit_R4554

    Comp_C56889:

        ' PedCND
        sCurrComponent = "PedCND"
        QueryC56889 = con.CreateCommand()
        QueryC56889.CommandText = QueryC56889.CommandText & " " & "DELETE FROM AKBUSER01.WF_PRE_PEDIDO_CND WHERE WF_PRE_PEDIDO_CND.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56889.Transaction = txn
        C56889 = QueryC56889.ExecuteNonQuery()
        C56889DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C567753

    Comp_C56890:

        ' PedDesc
        sCurrComponent = "PedDesc"
        QueryC56890 = con.CreateCommand()
        QueryC56890.CommandText = QueryC56890.CommandText & " " & "DELETE FROM AKBUSER01.WF_PRE_PEDIDO_DESCONTO WHERE WF_PRE_PEDIDO_DESCONTO.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56890.Transaction = txn
        C56890 = QueryC56890.ExecuteNonQuery()
        C56890DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56889

    Comp_C56891:

        ' PedObs
        sCurrComponent = "PedObs"
        QueryC56891 = con.CreateCommand()
        QueryC56891.CommandText = QueryC56891.CommandText & " " & "DELETE FROM AKBUSER01.WF_PRE_PEDIDO_OBS WHERE WF_PRE_PEDIDO_OBS.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56891.Transaction = txn
        C56891 = QueryC56891.ExecuteNonQuery()
        C56891DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56890

    Comp_C56894:

        ' ApagaItens
        sCurrComponent = "ApagaItens"
        C56894 = clsRuleDynamicallyCompiled_R4559.R4559(con, txn, msg, IIf(Not IsDBNull(P12594), P12594, System.DBNull.Value), IIf(Not IsDBNull(RsC56895(0)), RsC56895(0), System.DBNull.Value), fn_ConvertValueCompiled( "0", 4, AKB_DecimalPoint))
        C56894CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C56894) Then
          iColumns = C56894.Columns.Count
        End If
        ReDim C56894DataType(iColumns)
        For iCol = 1 To iColumns
            C56894DataType(iCol) = clsRuleDynamicallyCompiled_R4559.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C56897

    Comp_C56895:

        ' Sel_Itens
        sCurrComponent = "Sel_Itens"
        QueryC56895 = con.CreateCommand()
        QueryC56895.CommandText = QueryC56895.CommandText & " " & "select "
        QueryC56895.CommandText = QueryC56895.CommandText & " " & "Seq_Item, "
        QueryC56895.CommandText = QueryC56895.CommandText & " " & "Sigla_Prod, "
        QueryC56895.CommandText = QueryC56895.CommandText & " " & "Acesso, "
        QueryC56895.CommandText = QueryC56895.CommandText & " " & "Cod_Embalagem"
        QueryC56895.CommandText = QueryC56895.CommandText & " " & "from WF_PRE_PEDIDO_ITENS"
        QueryC56895.CommandText = QueryC56895.CommandText & " " & "where Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC56895.Transaction = txn
        RsC56895 = QueryC56895.ExecuteReader()
        Dim iC56895 As Short
        ReDim C56895DataType(RsC56895.FieldCount)
        For iC56895 = 0 to RsC56895.FieldCount - 1
            Select Case RsC56895.GetDataTypeName(iC56895).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56895DataType(iC56895 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56895DataType(iC56895 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56895DataType(iC56895 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56895
        RsC56895_EOF = Not RsC56895.Read()

        GoTo Comp_C556491

    Comp_C56896:

        ' Eof
        sCurrComponent = "Eof"
        C56896DataType = 4
        C56896 = RsC56895_EOF
        GoTo Comp_C56898

    Comp_C56897:

        ' Next
        sCurrComponent = "Next"
        C56897DataType = 4
        RsC56895_EOF = Not RsC56895.Read()
        C56897 = True

        GoTo Comp_C56896

    Comp_C56898:

        ' Eof=1
        sCurrComponent = "Eof=1"
        C56898 = (fn_ConvertValueCompiled(C56896, C56896DataType, AKB_DecimalPoint, False) = 1)
        C56898DataType = AKBTypeConst.cAKBTypeLogical
        If C56898 Then
            GoTo Comp_C56891
        Else
            GoTo Comp_C556489
        End If

    Comp_C96650:

        ' Transaction
        sCurrComponent = "Transaction"
        C96650 = (fn_ConvertValueCompiled(P20448, 4, AKB_DecimalPoint, False) = 1)
        C96650DataType = AKBTypeConst.cAKBTypeLogical
        If C96650 Then
            GoTo Comp_C119228
        Else
            GoTo Comp_C54073
        End If

    Comp_C96651:

        ' Transaction2
        sCurrComponent = "Transaction2"
        C96651 = (fn_ConvertValueCompiled(P20448, 4, AKB_DecimalPoint, False) = 1)
        C96651DataType = AKBTypeConst.cAKBTypeLogical
        If C96651 Then
            GoTo Comp_C119291
        Else
            GoTo Comp_C54076
        End If

    Comp_C119216:

        ' MsgConf
        sCurrComponent = "MsgConf"
        Dim row_C119216 As DataRow = msg.NewRow()
        row_C119216(0) = "Confirma exclusão do Pré-Pedido " & _ 
P12594 & " ?" & Chr(13) & Chr(10) & "Sim" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C119216)
        C119216DataType = AKBTypeConst.cAKBTypeNumeric
        C119216 = 6

        GoTo Comp_C119217

    Comp_C119217:

        ' MsgConf=S
        sCurrComponent = "MsgConf=S"
        C119217 = (fn_ConvertValueCompiled(C119216, C119216DataType, AKB_DecimalPoint, False) = 6)
        C119217DataType = AKBTypeConst.cAKBTypeLogical
        If C119217 Then
            GoTo Comp_C579496
        Else
            GoTo Comp_C54078
        End If

    Comp_C119218:

        ' Ins_Log
        sCurrComponent = "Ins_Log"
        QueryC119218 = con.CreateCommand()
        QueryC119218.CommandText = QueryC119218.CommandText & " " & "Insert into AKBUSER01.WF_Log_Excluir_Pre_Ped  ( WF_Log_Excluir_Pre_Ped.Seq , "
        QueryC119218.CommandText = QueryC119218.CommandText & " " & "                                                                           WF_Log_Excluir_Pre_Ped.Pre_Pedido ,"
        QueryC119218.CommandText = QueryC119218.CommandText & " " & "                                                                           WF_Log_Excluir_Pre_Ped.Motivo, "
        QueryC119218.CommandText = QueryC119218.CommandText & " " & "                                                                           WF_Log_Excluir_Pre_Ped.Dt_Pedido ,"
        QueryC119218.CommandText = QueryC119218.CommandText & " " & "                                                                           WF_Log_Excluir_Pre_Ped.Cod_Cliente ,"
        QueryC119218.CommandText = QueryC119218.CommandText & " " & "                                                                           WF_Log_Excluir_Pre_Ped.Cod_Repres )"
        QueryC119218.CommandText = QueryC119218.CommandText & " " & "Values ( " & _ 
ConvertValue(C119228, C119228DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ," & _ 
ConvertValue(C579497, C579497DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P24036, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P24035, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(P24037, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC119218.CommandText = QueryC119218.CommandText & " " & ""
        QueryC119218.Transaction = txn
        C119218 = QueryC119218.ExecuteNonQuery()
        C119218DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56895

    Comp_C119219:

        ' Motivo
        sCurrComponent = "Motivo"
        C119219DataType = AKBTypeConst.cAKBTypeString
        C119219 = ""

        GoTo Comp_C579500

    Comp_C119220:

        ' Motivo=Nulo
        sCurrComponent = "Motivo=Nulo"
        C119220 = (fn_ConvertValueCompiled(C579497, C579497DataType, AKB_DecimalPoint, False) = "" OR fn_ConvertValueCompiled(C579501, C579501DataType, AKB_DecimalPoint, False) = 1)
        C119220DataType = AKBTypeConst.cAKBTypeLogical
        If C119220 Then
            GoTo Comp_C119221
        Else
            GoTo Comp_C96650
        End If

    Comp_C119221:

        ' MsgMotivo
        sCurrComponent = "MsgMotivo"
        Dim row_C119221 As DataRow = msg.NewRow()
        row_C119221(0) = "O motivo da exclusão deve ser preenchido." & Chr(13) & "" & Chr(10) & "O Pré-Pedido não pode ser excluído." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C119221)
        C119221DataType = AKBTypeConst.cAKBTypeNumeric
        C119221 = 1

        GoTo Comp_C54078

    Comp_C119228:

        ' ID
        sCurrComponent = "ID"
        C119228DataType = 1
        C119228 = ObjTable_NewIdentity (con, txn, "WF_Log_Excluir_Pre_Ped")
        GoTo Comp_C119218

    Comp_C119291:

        ' MsgOk
        sCurrComponent = "MsgOk"
        Dim row_C119291 As DataRow = msg.NewRow()
        row_C119291(0) = "O Pré-Pedido foi excluído com sucesso." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C119291)
        C119291DataType = AKBTypeConst.cAKBTypeNumeric
        C119291 = 1

        GoTo Comp_C54078

    Comp_C128460:

        ' Sel_Pré Pedido x Proforma
        sCurrComponent = "Sel_Pré Pedido x Proforma"
        QueryC128460 = con.CreateCommand()
        QueryC128460.CommandText = QueryC128460.CommandText & " " & "SELECT "
        QueryC128460.CommandText = QueryC128460.CommandText & " " & "WF_PRE_PEDIDO.Id_PrePedido,"
        QueryC128460.CommandText = QueryC128460.CommandText & " " & "WF_PRE_PEDIDO.Pedido, "
        QueryC128460.CommandText = QueryC128460.CommandText & " " & "WF_PRE_PEDIDO.Proforma,"
        QueryC128460.CommandText = QueryC128460.CommandText & " " & "nvl(Pedido_Ecommerce, 0)"
        QueryC128460.CommandText = QueryC128460.CommandText & " " & ""
        QueryC128460.CommandText = QueryC128460.CommandText & " " & "FROM WF_PRE_PEDIDO "
        QueryC128460.CommandText = QueryC128460.CommandText & " " & ""
        QueryC128460.CommandText = QueryC128460.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido =  " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128460.Transaction = txn
        RsC128460 = QueryC128460.ExecuteReader()
        Dim iC128460 As Short
        ReDim C128460DataType(RsC128460.FieldCount)
        For iC128460 = 0 to RsC128460.FieldCount - 1
            Select Case RsC128460.GetDataTypeName(iC128460).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C128460DataType(iC128460 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C128460DataType(iC128460 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C128460DataType(iC128460 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC128460
        RsC128460_EOF = Not RsC128460.Read()

        GoTo Comp_C128461

    Comp_C128461:

        ' Proforma
        sCurrComponent = "Proforma"
        C128461DataType = 0
        C128461DataType = C128460Datatype(3)
        If C128461DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC128460(2)) Then
          C128461 = Strings.RTrim(RsC128460(2))
        Else 
          C128461 = RsC128460(2)
        End If 

        GoTo Comp_C561354

    Comp_C128462:

        ' É proforma?
        sCurrComponent = "É proforma?"
        C128462 = (fn_ConvertValueCompiled(C128461, C128461DataType, AKB_DecimalPoint, False) = 1)
        C128462DataType = AKBTypeConst.cAKBTypeLogical
        If C128462 Then
            GoTo Comp_C128463
        Else
            GoTo Comp_C571830
        End If

    Comp_C128463:

        ' MSG1
        sCurrComponent = "MSG1"
        Dim row_C128463 As DataRow = msg.NewRow()
        row_C128463(0) = "O Pré pedido: " & _ 
P12594 & " veio de uma Proforma e não poderá ser excluido." & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "Para cancelar vá para a tela de Cancelamento de Proforma." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C128463)
        C128463DataType = AKBTypeConst.cAKBTypeNumeric
        C128463 = 1

        GoTo Comp_C54078

    Comp_C556489:

        ' Pedido_Ecommerce
        sCurrComponent = "Pedido_Ecommerce"
        QueryC556489 = con.CreateCommand()
        QueryC556489.CommandText = QueryC556489.CommandText & " " & "select nvl(Pedido_Ecommerce, 0)"
        QueryC556489.CommandText = QueryC556489.CommandText & " " & "from WF_PRE_PEDIDO"
        QueryC556489.CommandText = QueryC556489.CommandText & " " & "where Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556489.Transaction = txn
        RsC556489 = QueryC556489.ExecuteReader()
        Dim iC556489 As Short
        ReDim C556489DataType(RsC556489.FieldCount)
        For iC556489 = 0 to RsC556489.FieldCount - 1
            Select Case RsC556489.GetDataTypeName(iC556489).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556489DataType(iC556489 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556489DataType(iC556489 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556489DataType(iC556489 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556489
        RsC556489_EOF = Not RsC556489.Read()

        GoTo Comp_C556490

    Comp_C556490:

        ' Pedido_Ecommerce = 0
        sCurrComponent = "Pedido_Ecommerce = 0"
        C556490 = (fn_ConvertValueCompiled(RsC556489(0), C556489DataType(1), AKB_DecimalPoint, False) = 0)
        C556490DataType = AKBTypeConst.cAKBTypeLogical
        If C556490 Then
            GoTo Comp_C56894
        Else
            GoTo Comp_C556497
        End If

    Comp_C556491:

        ' seqItem
        sCurrComponent = "seqItem"
        C556491DataType = 0
        C556491 = RsC56895(0)
        C556491DataType = C56895Datatype(1)
        If C556491DataType = AKBTypeConst.cAKBTypeString Then
          C556491 = IIF(IsDBNull(C556491), C556491, Strings.RTrim(C556491))
        End If 

        GoTo Comp_C556492

    Comp_C556492:

        ' siglaProd
        sCurrComponent = "siglaProd"
        C556492DataType = 0
        C556492DataType = C56895Datatype(2)
        If C556492DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56895(1)) Then
          C556492 = Strings.RTrim(RsC56895(1))
        Else 
          C556492 = RsC56895(1)
        End If 

        GoTo Comp_C556494

    Comp_C556493:

        ' emb
        sCurrComponent = "emb"
        C556493DataType = 0
        C556493DataType = C56895Datatype(4)
        If C556493DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56895(3)) Then
          C556493 = Strings.RTrim(RsC56895(3))
        Else 
          C556493 = RsC56895(3)
        End If 

        GoTo Comp_C56896

    Comp_C556494:

        ' acesso
        sCurrComponent = "acesso"
        C556494DataType = 0
        C556494DataType = C56895Datatype(3)
        If C556494DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56895(2)) Then
          C556494 = Strings.RTrim(RsC56895(2))
        Else 
          C556494 = RsC56895(2)
        End If 

        GoTo Comp_C556493

    Comp_C556497:

        ' UpQtdeDispNacionalEmbInt
        sCurrComponent = "UpQtdeDispNacionalEmbInt"
        QueryC556497 = con.CreateCommand()
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "UPDATE wf_estoque_ecommerce_disp SET"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ""
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "QTDE_DISP_NACIONAL = QTDE_DISP_NACIONAL + ("
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ""
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	SELECT DECODE(wf_emb_comp_vda_prod.cod_embalagem_int_emb,NULL,WF_PRE_PEDIDO_ITENS.Qtde_Reservada,wf_emb_comp_vda_prod.fator_conv_cpr_inter*WF_PRE_PEDIDO_ITENS.Qtde_Reservada)"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ""
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	FROM WF_PRE_PEDIDO_ITENS, "
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	wf_emb_comp_vda_prod"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ""
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	WHERE WF_PRE_PEDIDO_ITENS.ACESSO = wf_estoque_ecommerce_disp.ACESSO"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.sigla_prod = wf_estoque_ecommerce_disp.SIGLA_PROD"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	and WF_PRE_PEDIDO_ITENS.ACESSO = wf_emb_comp_vda_prod.ACESSO"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.sigla_prod = wf_emb_comp_vda_prod.SIGLA_PROD"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	and wf_emb_comp_vda_prod.ACESSO = wf_estoque_ecommerce_disp.ACESSO"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND wf_emb_comp_vda_prod.sigla_prod = wf_estoque_ecommerce_disp.SIGLA_PROD"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.cod_embalagem = wf_emb_comp_vda_prod.cod_embalagem"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND wf_emb_comp_vda_prod.cod_embalagem_int_emb IS NOT NULL"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND wf_emb_comp_vda_prod.cod_embalagem_int_emb = wf_estoque_ecommerce_disp.cod_emb_minima"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ""
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ")"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ""
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "WHERE EXISTS ("
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ""
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	SELECT *"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	FROM WF_PRE_PEDIDO_ITENS, "
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	wf_emb_comp_vda_prod"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	WHERE WF_PRE_PEDIDO_ITENS.ACESSO = wf_estoque_ecommerce_disp.ACESSO"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.sigla_prod = wf_estoque_ecommerce_disp.SIGLA_PROD"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	and WF_PRE_PEDIDO_ITENS.ACESSO = wf_emb_comp_vda_prod.ACESSO"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.sigla_prod = wf_emb_comp_vda_prod.SIGLA_PROD"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	and wf_emb_comp_vda_prod.ACESSO = wf_estoque_ecommerce_disp.ACESSO"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND wf_emb_comp_vda_prod.sigla_prod = wf_estoque_ecommerce_disp.SIGLA_PROD"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.cod_embalagem = wf_emb_comp_vda_prod.cod_embalagem"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND wf_emb_comp_vda_prod.cod_embalagem_int_emb IS NOT NULL"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND wf_emb_comp_vda_prod.cod_embalagem_int_emb = wf_estoque_ecommerce_disp.cod_emb_minima"
        QueryC556497.CommandText = QueryC556497.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ""
        QueryC556497.CommandText = QueryC556497.CommandText & " " & ")"
        QueryC556497.Transaction = txn
        C556497 = QueryC556497.ExecuteNonQuery()
        C556497DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556498

    Comp_C556498:

        ' UpQtdeDispNacionalEmb
        sCurrComponent = "UpQtdeDispNacionalEmb"
        QueryC556498 = con.CreateCommand()
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "UPDATE wf_estoque_ecommerce_disp SET"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ""
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "QTDE_DISP_NACIONAL = QTDE_DISP_NACIONAL + ("
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ""
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	SELECT DECODE(wf_emb_comp_vda_prod.cod_embalagem_int_emb,NULL,WF_PRE_PEDIDO_ITENS.Qtde_Reservada,wf_emb_comp_vda_prod.fator_conv_cpr_inter*WF_PRE_PEDIDO_ITENS.Qtde_Reservada)"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ""
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	FROM WF_PRE_PEDIDO_ITENS, wf_emb_comp_vda_prod"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ""
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	WHERE WF_PRE_PEDIDO_ITENS.ACESSO = wf_estoque_ecommerce_disp.ACESSO"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.sigla_prod = wf_estoque_ecommerce_disp.SIGLA_PROD"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	and WF_PRE_PEDIDO_ITENS.ACESSO = wf_emb_comp_vda_prod.ACESSO"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.sigla_prod = wf_emb_comp_vda_prod.SIGLA_PROD"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	and wf_emb_comp_vda_prod.ACESSO = wf_estoque_ecommerce_disp.ACESSO"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND wf_emb_comp_vda_prod.sigla_prod = wf_estoque_ecommerce_disp.SIGLA_PROD"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.cod_embalagem = wf_emb_comp_vda_prod.cod_embalagem"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND wf_emb_comp_vda_prod.cod_embalagem_int_emb IS NULL"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.cod_embalagem = wf_estoque_ecommerce_disp.cod_emb_minima"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ""
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ")"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ""
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "WHERE EXISTS ("
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ""
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	SELECT *"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	FROM WF_PRE_PEDIDO_ITENS, wf_emb_comp_vda_prod"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	WHERE WF_PRE_PEDIDO_ITENS.ACESSO = wf_estoque_ecommerce_disp.ACESSO"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.sigla_prod = wf_estoque_ecommerce_disp.SIGLA_PROD"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	and WF_PRE_PEDIDO_ITENS.ACESSO = wf_emb_comp_vda_prod.ACESSO"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.sigla_prod = wf_emb_comp_vda_prod.SIGLA_PROD"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	and wf_emb_comp_vda_prod.ACESSO = wf_estoque_ecommerce_disp.ACESSO"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND wf_emb_comp_vda_prod.sigla_prod = wf_estoque_ecommerce_disp.SIGLA_PROD"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.cod_embalagem = wf_emb_comp_vda_prod.cod_embalagem"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND wf_emb_comp_vda_prod.cod_embalagem_int_emb IS NULL"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.cod_embalagem = wf_estoque_ecommerce_disp.cod_emb_minima"
        QueryC556498.CommandText = QueryC556498.CommandText & " " & "	AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ""
        QueryC556498.CommandText = QueryC556498.CommandText & " " & ")"
        QueryC556498.Transaction = txn
        C556498 = QueryC556498.ExecuteNonQuery()
        C556498DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56894

    Comp_C556499:

        ' pedido_site
        sCurrComponent = "pedido_site"
        QueryC556499 = con.CreateCommand()
        QueryC556499.CommandText = QueryC556499.CommandText & " " & "select b2b_pedido.ped_id_site"
        QueryC556499.CommandText = QueryC556499.CommandText & " " & ""
        QueryC556499.CommandText = QueryC556499.CommandText & " " & "from b2b_pedido, WF_PRE_PEDIDO"
        QueryC556499.CommandText = QueryC556499.CommandText & " " & ""
        QueryC556499.CommandText = QueryC556499.CommandText & " " & "where WF_PRE_PEDIDO.ID_Pedido_Ecomm = b2b_pedido.identificador"
        QueryC556499.CommandText = QueryC556499.CommandText & " " & "and WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556499.Transaction = txn
        RsC556499 = QueryC556499.ExecuteReader()
        Dim iC556499 As Short
        ReDim C556499DataType(RsC556499.FieldCount)
        For iC556499 = 0 to RsC556499.FieldCount - 1
            Select Case RsC556499.GetDataTypeName(iC556499).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556499DataType(iC556499 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556499DataType(iC556499 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556499DataType(iC556499 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556499
        RsC556499_EOF = Not RsC556499.Read()

        GoTo Comp_C556500

    Comp_C556500:

        ' id_cancelamento
        sCurrComponent = "id_cancelamento"
        C556500DataType = 1
        C556500 = ObjTable_NewIdentity (con, txn, "wf_canc_ped_ecomm")
        GoTo Comp_C556501

    Comp_C556501:

        ' ins_canc_pedido_ecomm
        sCurrComponent = "ins_canc_pedido_ecomm"
        QueryC556501 = con.CreateCommand()
        QueryC556501.CommandText = QueryC556501.CommandText & " " & "insert into wf_canc_ped_ecomm("
        QueryC556501.CommandText = QueryC556501.CommandText & " " & "identificador,"
        QueryC556501.CommandText = QueryC556501.CommandText & " " & "cod_ped_site,"
        QueryC556501.CommandText = QueryC556501.CommandText & " " & "cancelado"
        QueryC556501.CommandText = QueryC556501.CommandText & " " & ")values("
        QueryC556501.CommandText = QueryC556501.CommandText & " " & "" & _ 
ConvertValue(C556500, C556500DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC556501.CommandText = QueryC556501.CommandText & " " & "" & _ 
ConvertValue(RsC556499(0), C556499DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC556501.CommandText = QueryC556501.CommandText & " " & "0"
        QueryC556501.CommandText = QueryC556501.CommandText & " " & ")"
        QueryC556501.Transaction = txn
        C556501 = QueryC556501.ExecuteNonQuery()
        C556501DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C54075

    Comp_C556543:

        ' PEDIDO_ECOMMERCE = 01
        sCurrComponent = "PEDIDO_ECOMMERCE = 01"
        C556543 = (fn_ConvertValueCompiled(C561354, C561354DataType, AKB_DecimalPoint, False) = 1)
        C556543DataType = AKBTypeConst.cAKBTypeLogical
        If C556543 Then
            GoTo Comp_C556499
        Else
            GoTo Comp_C54075
        End If

    Comp_C561354:

        ' Ecommerce?
        sCurrComponent = "Ecommerce?"
        C561354DataType = 0
        C561354DataType = C128460Datatype(4)
        If C561354DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC128460(3)) Then
          C561354 = Strings.RTrim(RsC128460(3))
        Else 
          C561354 = RsC128460(3)
        End If 

        GoTo Comp_C128462

    Comp_C567753:

        ' PlanCsPRePedidos
        sCurrComponent = "PlanCsPRePedidos"
        QueryC567753 = con.CreateCommand()
        QueryC567753.CommandText = QueryC567753.CommandText & " " & "DELETE FROM AKBUSER01.wf_plan_cs_pre_ped WHERE wf_plan_cs_pre_ped.Id_PrePedido = " & _ 
ConvertValue(P12594, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC567753.Transaction = txn
        C567753 = QueryC567753.ExecuteNonQuery()
        C567753DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556543

    Comp_C571830:

        ' ecommerce
        sCurrComponent = "ecommerce"
        C571830 = (fn_ConvertValueCompiled(C561354, C561354DataType, AKB_DecimalPoint, False) = 1)
        C571830DataType = AKBTypeConst.cAKBTypeLogical
        If C571830 Then
            GoTo Comp_C571831
        Else
            GoTo Comp_C119216
        End If

    Comp_C571831:

        ' MSG2
        sCurrComponent = "MSG2"
        Dim row_C571831 As DataRow = msg.NewRow()
        row_C571831(0) = "Pedido veio do e-commerce, não poderá ser excluído." & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "Para cancelar vá para a tela de Cancelamento de Proforma." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C571831)
        C571831DataType = AKBTypeConst.cAKBTypeNumeric
        C571831 = 1

        GoTo Comp_C54078

    Comp_C579496:

        ' Compr_PMotivo
        sCurrComponent = "Compr_PMotivo"
        C579496DataType = 1
        C579496 = Len(P92047 & "")
        GoTo Comp_C579498

    Comp_C579497:

        ' vMotivo
        sCurrComponent = "vMotivo"
        C579497 = P92047 & " "
        C579497DataType = 3
        GoTo Comp_C128460

    Comp_C579498:

        ' Compr_PMotivo > 0
        sCurrComponent = "Compr_PMotivo > 0"
        C579498 = (fn_ConvertValueCompiled(C579496, C579496DataType, AKB_DecimalPoint, False) > 0)
        C579498DataType = AKBTypeConst.cAKBTypeLogical
        If C579498 Then
            GoTo Comp_C119220
        Else
            GoTo Comp_C119219
        End If

    Comp_C579499:

        ' vMotivo := SuprDir_Motivo
        sCurrComponent = "vMotivo := SuprDir_Motivo"
        C579499DataType = 4
        C579497 = fn_ConvertValueCompiled(C579500 , 3, AKB_DecimalPoint)
        C579499 = True
        GoTo Comp_C579501

    Comp_C579500:

        ' SuprDir_Motivo
        sCurrComponent = "SuprDir_Motivo"
        C579500DataType = 3
        C579500 = RTrim(C119219)
        GoTo Comp_C579499

    Comp_C579501:

        ' ÉNULO_vMotivo
        sCurrComponent = "ÉNULO_vMotivo"
        C579501DataType = 4
        C579501 = IsDBNull (C579497)
        GoTo Comp_C119220

    Exit_R4554:

        Exit Function

    Err_R4554:
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
