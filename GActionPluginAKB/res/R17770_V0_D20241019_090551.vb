Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R17770

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

    Public Shared Function R17770(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P64626 As Double, ByVal P64627 As Double) As DataTable
    ' 
    ' 17770 - Atualiza Prazo Entrega - Tela Geração de Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R17770
        Dim sCurrComponent as String

        Dim QueryC394876 As New Object
        Dim RsC394876 As New Object
        Dim C394876DataType() As Short
        Dim RsC394876_EOF As Boolean
        Dim C394877 As Object
        Dim C394877DataType As Short
        Dim C394878DataType() As Short
        Dim QueryC394879 As New Object
        Dim C394879 As Integer
        Dim C394879DataType As Short
        Dim C394880 As Object
        Dim C394880DataType As Short
        Dim QueryC394881 As New Object
        Dim C394881 As Integer
        Dim C394881DataType As Short
        Dim QueryC394882 As New Object
        Dim RsC394882 As New Object
        Dim C394882DataType() As Short
        Dim RsC394882_EOF As Boolean
        Dim C394883 As Boolean
        Dim C394883DataType As Short
        Dim C394885 As Object
        Dim C394885DataType As Short
        Dim QueryC394895 As New Object
        Dim C394895 As Integer
        Dim C394895DataType As Short
        Dim QueryC394897 As New Object
        Dim RsC394897 As New Object
        Dim C394897DataType() As Short
        Dim RsC394897_EOF As Boolean
        Dim C394900 As Boolean
        Dim C394900DataType As Short
        Dim C394901 As Object
        Dim C394901DataType As Short
        Dim C394902 As Object
        Dim C394902DataType As Short
        Dim QueryC408167 As New Object
        Dim C408167 As Integer
        Dim C408167DataType As Short
        Dim QueryC408168 As New Object
        Dim C408168 As Integer
        Dim C408168DataType As Short
        Dim C446950 As Object
        Dim C446950DataType As Short
        Dim C446951 As Object
        Dim C446951DataType As Short
        Dim C446953 As Boolean
        Dim C446953DataType As Short
        Dim QueryC446954 As New Object
        Dim RsC446954 As New Object
        Dim C446954DataType() As Short
        Dim RsC446954_EOF As Boolean
        Dim C446966 As Object
        Dim C446966DataType As Short
        Dim C446967 As Object
        Dim C446967DataType As Short
        Dim C446968 As Object
        Dim C446968DataType As Short
        Dim C446969 As Object
        Dim C446969DataType As Short
        Dim C446970 As Object
        Dim C446970DataType As Short
        Dim C446971 As Boolean
        Dim C446971DataType As Short
        Dim QueryC446974 As New Object
        Dim RsC446974 As New Object
        Dim C446974DataType() As Short
        Dim RsC446974_EOF As Boolean
        Dim C446975 As Object
        Dim C446975DataType As Short
        Dim QueryC446976 As New Object
        Dim C446976 As Integer
        Dim C446976DataType As Short
        Dim C446977 As Object
        Dim C446977DataType As Short
        Dim C446993 As Object
        Dim C446993DataType As Short
        Dim QueryC447162 As New Object
        Dim RsC447162 As New Object
        Dim C447162DataType() As Short
        Dim RsC447162_EOF As Boolean
        Dim C467430 As Object
        Dim C467430DataType As Short
        Dim QueryC467431 As New Object
        Dim RsC467431 As New Object
        Dim C467431DataType() As Short
        Dim RsC467431_EOF As Boolean
        Dim C467432 As Boolean
        Dim C467432DataType As Short
        Dim C467434 As Short
        Dim C467434DataType As Short
        Dim C467434ReturnDataType() As Short

        Dim C467435 As Object
        Dim C467435DataType As Short
        Dim QueryC482594 As New Object
        Dim C482594 As Integer
        Dim C482594DataType As Short
        Dim QueryC482595 As New Object
        Dim RsC482595 As New Object
        Dim C482595DataType() As Short
        Dim RsC482595_EOF As Boolean
        Dim C482596 As Boolean
        Dim C482596DataType As Short
        Dim C482597 As Short
        Dim C482597DataType As Short
        Dim C482597ReturnDataType() As Short

        Dim QueryC482598 As New Object
        Dim C482598 As Integer
        Dim C482598DataType As Short
        Dim QueryC482633 As New Object
        Dim C482633 As Integer
        Dim C482633DataType As Short
        Dim QueryC482634 As New Object
        Dim C482634 As Integer
        Dim C482634DataType As Short
        Dim C482982 As Object
        Dim C482982DataType As Short
        Dim QueryC501545 As New Object
        Dim RsC501545 As New Object
        Dim C501545DataType() As Short
        Dim RsC501545_EOF As Boolean
        Dim C501546 As Object
        Dim C501546DataType As Short
        Dim C501547 As Object
        Dim C501547DataType As Short
        Dim C501548 As Object
        Dim C501548DataType As Short
        Dim C501550 As Object
        Dim C501550DataType As Short
        Dim C501551 As Object
        Dim C501551DataType As Short
        Dim C501552 As Object
        Dim C501552DataType As Short
        Dim C501553 As Object
        Dim C501553DataType As Short
        Dim C501554 As Object
        Dim C501554DataType As Short
        Dim C501555 As Object
        Dim C501555DataType As Short
        Dim C501556 As Object
        Dim C501556DataType As Short
        Dim C501557 As Object
        Dim C501557DataType As Short
        Dim C501558 As Object
        Dim C501558DataType As Short
        Dim C501559 As Object
        Dim C501559DataType As Short
        Dim C501560 As Object
        Dim C501560DataType As Short
        Dim QueryC501561 As New Object
        Dim RsC501561 As New Object
        Dim C501561DataType() As Short
        Dim RsC501561_EOF As Boolean
        Dim C501562 As Object
        Dim C501562DataType As Short
        Dim C501563 As Object
        Dim C501563DataType As Short
        Dim C501564 As Boolean
        Dim C501564DataType As Short
        Dim C501565 As Object
        Dim C501565DataType As Short
        Dim C501566 As Object
        Dim C501566DataType As Short
        Dim C501567 As Boolean
        Dim C501567DataType As Short
        Dim C501568 As Object
        Dim C501568DataType As Short
        Dim C501569 As Object
        Dim C501569DataType As Short
        Dim C501570 As Boolean
        Dim C501570DataType As Short
        Dim C501571 As Object
        Dim C501571DataType As Short
        Dim C501572 As Object
        Dim C501572DataType As Short
        Dim C501573 As Boolean
        Dim C501573DataType As Short
        Dim C501575 As Object
        Dim C501575DataType As Short
        Dim C501576 As Object
        Dim C501576DataType As Short
        Dim C501577 As Boolean
        Dim C501577DataType As Short
        Dim C501579 As Object
        Dim C501579DataType As Short
        Dim C501580 As Object
        Dim C501580DataType As Short
        Dim C501581 As Boolean
        Dim C501581DataType As Short
        Dim C501582 As Object
        Dim C501582DataType As Short
        Dim C501583 As Object
        Dim C501583DataType As Short
        Dim C501584 As Boolean
        Dim C501584DataType As Short
        Dim C501585 As Object
        Dim C501585DataType As Short
        Dim C501586 As Object
        Dim C501586DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C394877

    Comp_C394876:

        ' SelTipoPedido
        sCurrComponent = "SelTipoPedido"
        QueryC394876 = con.CreateCommand()
        QueryC394876.CommandText = QueryC394876.CommandText & " " & "SELECT TRUNC(WF_PEDIDO.Dt_Pedido)"
        QueryC394876.CommandText = QueryC394876.CommandText & " " & "FROM WF_PEDIDO"
        QueryC394876.CommandText = QueryC394876.CommandText & " " & "WHERE WF_PEDIDO.pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394876.CommandText = QueryC394876.CommandText & " " & "AND WF_PEDIDO.tp_produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394876.Transaction = txn
        RsC394876 = QueryC394876.ExecuteReader()
        Dim iC394876 As Short
        ReDim C394876DataType(RsC394876.FieldCount)
        For iC394876 = 0 to RsC394876.FieldCount - 1
            Select Case RsC394876.GetDataTypeName(iC394876).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C394876DataType(iC394876 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C394876DataType(iC394876 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C394876DataType(iC394876 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC394876
        RsC394876_EOF = Not RsC394876.Read()

        GoTo Comp_C394880

    Comp_C394877:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C394877 = 1
        C394877DataType = 4
        GoTo Comp_C394901

    Comp_C394878:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C394878 As DataTable = New DataTable()
        tb_C394878.Columns.Add("vRetorno" & "")
        Dim row_C394878 As DataRow = tb_C394878.NewRow()
        row_C394878(0) = C394877
        tb_C394878.Rows.Add(row_C394878)
        R17770 = tb_C394878
        ReDim C394878DataType(1)
        C394878DataType(1) = C394877DataType
        ReturnDataType = C394878DataType

        GoTo Exit_R17770

    Comp_C394879:

        ' Up_Ped_Itens_Pronta_Entrega
        sCurrComponent = "Up_Ped_Itens_Pronta_Entrega"
        QueryC394879 = con.CreateCommand()
        QueryC394879.CommandText = QueryC394879.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS"
        QueryC394879.CommandText = QueryC394879.CommandText & " " & "SET WF_PEDIDO_ITENS.Prazo_Entrega = " & _ 
ConvertValue(C394901, C394901DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC394879.CommandText = QueryC394879.CommandText & " " & "Prazo_Entrega_Comercial = " & _ 
ConvertValue(C467430, C467430DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394879.CommandText = QueryC394879.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394879.CommandText = QueryC394879.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394879.CommandText = QueryC394879.CommandText & " " & "AND WF_PEDIDO_ITENS.Tipo_Ped = 1"
        QueryC394879.Transaction = txn
        C394879 = QueryC394879.ExecuteNonQuery()
        C394879DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C482633

    Comp_C394880:

        ' Data_Pedido
        sCurrComponent = "Data_Pedido"
        C394880DataType = 0
        C394880 = RsC394876(0)
        C394880DataType = C394876Datatype(1)
        If C394880DataType = AKBTypeConst.cAKBTypeString Then
          C394880 = IIF(IsDBNull(C394880), C394880, Strings.RTrim(C394880))
        End If 

        GoTo Comp_C482982

    Comp_C394881:

        ' Up_Ped_Itens_Programados
        sCurrComponent = "Up_Ped_Itens_Programados"
        QueryC394881 = con.CreateCommand()
        QueryC394881.CommandText = QueryC394881.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS"
        QueryC394881.CommandText = QueryC394881.CommandText & " " & "SET WF_PEDIDO_ITENS.Prazo_Entrega = " & _ 
ConvertValue(C394901, C394901DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ","
        QueryC394881.CommandText = QueryC394881.CommandText & " " & "Prazo_Entrega_Comercial = null"
        QueryC394881.CommandText = QueryC394881.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394881.CommandText = QueryC394881.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394881.CommandText = QueryC394881.CommandText & " " & "AND WF_PEDIDO_ITENS.Tipo_Ped = 2"
        QueryC394881.Transaction = txn
        C394881 = QueryC394881.ExecuteNonQuery()
        C394881DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C408167

    Comp_C394882:

        ' Sel_Param_Estab
        sCurrComponent = "Sel_Param_Estab"
        QueryC394882 = con.CreateCommand()
        QueryC394882.CommandText = QueryC394882.CommandText & " " & "SELECT NVL(WF_ESTAB_JURIDICO_PARAM.Aplica_Calc_Prazo_Ent_Fab, 0), NVL(WF_ESTAB_JURIDICO_PARAM.Calc_Prazo_Ent_Ind ,0)"
        QueryC394882.CommandText = QueryC394882.CommandText & " " & "FROM AKBUSER01.WF_ESTAB_JURIDICO_PARAM , AKBUSER01.WF_PEDIDO"
        QueryC394882.CommandText = QueryC394882.CommandText & " " & "WHERE WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = WF_PEDIDO.Cod_Pessoa_Estab_Juridico"
        QueryC394882.CommandText = QueryC394882.CommandText & " " & "AND WF_PEDIDO.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394882.CommandText = QueryC394882.CommandText & " " & "AND WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394882.Transaction = txn
        RsC394882 = QueryC394882.ExecuteReader()
        Dim iC394882 As Short
        ReDim C394882DataType(RsC394882.FieldCount)
        For iC394882 = 0 to RsC394882.FieldCount - 1
            Select Case RsC394882.GetDataTypeName(iC394882).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C394882DataType(iC394882 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C394882DataType(iC394882 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C394882DataType(iC394882 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC394882
        RsC394882_EOF = Not RsC394882.Read()

        GoTo Comp_C446951

    Comp_C394883:

        ' Aplica Cálculo de Prazo Entrega Fábrica?
        sCurrComponent = "Aplica Cálculo de Prazo Entrega Fábrica?"
        C394883 = (fn_ConvertValueCompiled(C446951, C446951DataType, AKB_DecimalPoint, False)  = 1)
        C394883DataType = AKBTypeConst.cAKBTypeLogical
        If C394883 Then
            GoTo Comp_C467431
        Else
            GoTo Comp_C446953
        End If

    Comp_C394885:

        ' Prazo_Entrega
        sCurrComponent = "Prazo_Entrega"
        C394885DataType = 2
        C394885 = DateAdd("d", C501563, C482982)
        GoTo Comp_C394902

    Comp_C394895:

        ' Up_Prazo_Entrega_Pedido
        sCurrComponent = "Up_Prazo_Entrega_Pedido"
        QueryC394895 = con.CreateCommand()
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "UPDATE WF_Pedido_Seq"
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "SET WF_Pedido_Seq.Prazo_Entrega_Fabrica =  " & _ 
ConvertValue(C394901, C394901DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "WHERE WF_Pedido_Seq.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "AND WF_Pedido_Seq.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "AND (SELECT COUNT(*)"
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "  FROM Wf_Pedido_Itens"
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "  WHERE Wf_Pedido_Itens.Pedido   = WF_Pedido_Seq.Pedido"
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "  AND Wf_Pedido_Itens.Tp_Produto = WF_Pedido_Seq.Tp_Produto"
        QueryC394895.CommandText = QueryC394895.CommandText & " " & "  AND Wf_Pedido_Itens.Tipo_Ped   = 1) > 0"
        QueryC394895.Transaction = txn
        C394895 = QueryC394895.ExecuteNonQuery()
        C394895DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C446953

    Comp_C394897:

        ' ContaItensProntaEntrega
        sCurrComponent = "ContaItensProntaEntrega"
        QueryC394897 = con.CreateCommand()
        QueryC394897.CommandText = QueryC394897.CommandText & " " & "SELECT COUNT(*)"
        QueryC394897.CommandText = QueryC394897.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS"
        QueryC394897.CommandText = QueryC394897.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394897.CommandText = QueryC394897.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394897.CommandText = QueryC394897.CommandText & " " & "AND WF_PEDIDO_ITENS.Tipo_Ped = 1"
        QueryC394897.Transaction = txn
        RsC394897 = QueryC394897.ExecuteReader()
        Dim iC394897 As Short
        ReDim C394897DataType(RsC394897.FieldCount)
        For iC394897 = 0 to RsC394897.FieldCount - 1
            Select Case RsC394897.GetDataTypeName(iC394897).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C394897DataType(iC394897 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C394897DataType(iC394897 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C394897DataType(iC394897 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC394897
        RsC394897_EOF = Not RsC394897.Read()

        GoTo Comp_C394900

    Comp_C394900:

        ' ContaItensProntaEntrega = 0?
        sCurrComponent = "ContaItensProntaEntrega = 0?"
        C394900 = (fn_ConvertValueCompiled(RsC394897(0), C394897DataType(1), AKB_DecimalPoint, False) = 0)
        C394900DataType = AKBTypeConst.cAKBTypeLogical
        If C394900 Then
            GoTo Comp_C394895
        Else
            GoTo Comp_C501545
        End If

    Comp_C394901:

        ' vPrazo_Entrega
        sCurrComponent = "vPrazo_Entrega"
        C394901 = System.DBNull.Value
        C394901DataType = 2
        GoTo Comp_C394876

    Comp_C394902:

        ' AtrPrazoEntrega
        sCurrComponent = "AtrPrazoEntrega"
        C394902DataType = 4
        C394901 = fn_ConvertValueCompiled(C394885 , 2, AKB_DecimalPoint)
        C394902 = True
        GoTo Comp_C467430

    Comp_C408167:

        ' Ins_Log_Prazo_Entrega_Ped_Prog
        sCurrComponent = "Ins_Log_Prazo_Entrega_Ped_Prog"
        QueryC408167 = con.CreateCommand()
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "INSERT INTO AKBUSER01.WF_Log_Prazo_Ent_Ped_Prog"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "(WF_Log_Prazo_Ent_Ped_Prog.Sequencia,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Tp_Produto,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Pedido,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Seq_Item,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Prazo_Entrega,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Qtde,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Tipo_Pedido,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Usuario_Atualizacao,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Data_Atualizacao)"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "SELECT Seq_Wf_Log_Prazo_Ent_Ped_Prog.Nextval ,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_PEDIDO_ITENS.Tp_Produto ,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_PEDIDO_ITENS.Pedido ,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_PEDIDO_ITENS.Seq_Item ,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "" & _ 
ConvertValue(C394901, C394901DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Pedido ,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "2 ,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "USER ,"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "SYSDATE"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS"
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC408167.CommandText = QueryC408167.CommandText & " " & "AND WF_PEDIDO_ITENS.Tipo_Ped = 2"
        QueryC408167.Transaction = txn
        C408167 = QueryC408167.ExecuteNonQuery()
        C408167DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C394897

    Comp_C408168:

        ' Ins_Log_Prazo_Entrega_Ped_Pronta_Entrega
        sCurrComponent = "Ins_Log_Prazo_Entrega_Ped_Pronta_Entrega"
        QueryC408168 = con.CreateCommand()
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "INSERT INTO AKBUSER01.WF_Log_Prazo_Ent_Ped_Prog"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "(WF_Log_Prazo_Ent_Ped_Prog.Sequencia,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Tp_Produto,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Pedido,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Seq_Item,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Prazo_Entrega,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Qtde,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Tipo_Pedido,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Usuario_Atualizacao,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_Log_Prazo_Ent_Ped_Prog.Data_Atualizacao)"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "SELECT Seq_Wf_Log_Prazo_Ent_Ped_Prog.Nextval ,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_PEDIDO_ITENS.Tp_Produto ,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_PEDIDO_ITENS.Pedido ,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_PEDIDO_ITENS.Seq_Item ,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "" & _ 
ConvertValue(C394901, C394901DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Pedido ,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "1 ,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "USER ,"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "SYSDATE"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS"
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC408168.CommandText = QueryC408168.CommandText & " " & "AND WF_PEDIDO_ITENS.Tipo_Ped = 1"
        QueryC408168.Transaction = txn
        C408168 = QueryC408168.ExecuteNonQuery()
        C408168DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C482594

    Comp_C446950:

        ' vCalcPrazo
        sCurrComponent = "vCalcPrazo"
        C446950DataType = 0
        C446950DataType = C394882Datatype(2)
        If C446950DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394882(1)) Then
          C446950 = Strings.RTrim(RsC394882(1))
        Else 
          C446950 = RsC394882(1)
        End If 

        GoTo Comp_C394883

    Comp_C446951:

        ' vAplicValor
        sCurrComponent = "vAplicValor"
        C446951DataType = 0
        C446951 = RsC394882(0)
        C446951DataType = C394882Datatype(1)
        If C446951DataType = AKBTypeConst.cAKBTypeString Then
          C446951 = IIF(IsDBNull(C446951), C446951, Strings.RTrim(C446951))
        End If 

        GoTo Comp_C446950

    Comp_C446953:

        ' Calcula Prazo Ent Individual?
        sCurrComponent = "Calcula Prazo Ent Individual?"
        C446953 = (fn_ConvertValueCompiled(C446950, C446950DataType, AKB_DecimalPoint, False) = 1)
        C446953DataType = AKBTypeConst.cAKBTypeLogical
        If C446953 Then
            GoTo Comp_C446954
        Else
            GoTo Comp_C394878
        End If

    Comp_C446954:

        ' SelDadosPed
        sCurrComponent = "SelDadosPed"
        QueryC446954 = con.CreateCommand()
        QueryC446954.CommandText = QueryC446954.CommandText & " " & "SELECT WF_PEDIDO_ITENS.SIGLA_PROD, WF_PEDIDO_ITENS.ACESSO, "
        QueryC446954.CommandText = QueryC446954.CommandText & " " & "WF_PEDIDO_ITENS.SEQ_ITEM, WF_PEDIDO.TABELA_PRECO_VENDA  "
        QueryC446954.CommandText = QueryC446954.CommandText & " " & "FROM WF_PEDIDO_ITENS, WF_PEDIDO"
        QueryC446954.CommandText = QueryC446954.CommandText & " " & "WHERE  WF_PEDIDO_ITENS.PEDIDO = WF_PEDIDO.PEDIDO "
        QueryC446954.CommandText = QueryC446954.CommandText & " " & "AND  WF_PEDIDO_ITENS.TP_PRODUTO   =  WF_PEDIDO.TP_PRODUTO   "
        QueryC446954.CommandText = QueryC446954.CommandText & " " & "AND WF_PEDIDO.TP_PRODUTO = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446954.CommandText = QueryC446954.CommandText & " " & "AND WF_PEDIDO.PEDIDO = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446954.Transaction = txn
        RsC446954 = QueryC446954.ExecuteReader()
        Dim iC446954 As Short
        ReDim C446954DataType(RsC446954.FieldCount)
        For iC446954 = 0 to RsC446954.FieldCount - 1
            Select Case RsC446954.GetDataTypeName(iC446954).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C446954DataType(iC446954 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C446954DataType(iC446954 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C446954DataType(iC446954 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC446954
        RsC446954_EOF = Not RsC446954.Read()

        GoTo Comp_C446969

    Comp_C446966:

        ' vAcesso
        sCurrComponent = "vAcesso"
        C446966DataType = 0
        C446966DataType = C446954Datatype(2)
        If C446966DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC446954(1)) Then
          C446966 = Strings.RTrim(RsC446954(1))
        Else 
          C446966 = RsC446954(1)
        End If 

        GoTo Comp_C446967

    Comp_C446967:

        ' vSeqItem
        sCurrComponent = "vSeqItem"
        C446967DataType = 0
        C446967DataType = C446954Datatype(3)
        If C446967DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC446954(2)) Then
          C446967 = Strings.RTrim(RsC446954(2))
        Else 
          C446967 = RsC446954(2)
        End If 

        GoTo Comp_C446968

    Comp_C446968:

        ' vTabPrecVenda
        sCurrComponent = "vTabPrecVenda"
        C446968DataType = 0
        C446968DataType = C446954Datatype(4)
        If C446968DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC446954(3)) Then
          C446968 = Strings.RTrim(RsC446954(3))
        Else 
          C446968 = RsC446954(3)
        End If 

        GoTo Comp_C446974

    Comp_C446969:

        ' éFim?
        sCurrComponent = "éFim?"
        C446969DataType = 4
        C446969 = RsC446954_EOF
        GoTo Comp_C446971

    Comp_C446970:

        ' vSiglaProd
        sCurrComponent = "vSiglaProd"
        C446970DataType = 0
        C446970 = RsC446954(0)
        C446970DataType = C446954Datatype(1)
        If C446970DataType = AKBTypeConst.cAKBTypeString Then
          C446970 = IIF(IsDBNull(C446970), C446970, Strings.RTrim(C446970))
        End If 

        GoTo Comp_C446966

    Comp_C446971:

        ' Fim?
        sCurrComponent = "Fim?"
        C446971 = (fn_ConvertValueCompiled(C446969, C446969DataType, AKB_DecimalPoint, False) = 1)
        C446971DataType = AKBTypeConst.cAKBTypeLogical
        If C446971 Then
            GoTo Comp_C394878
        Else
            GoTo Comp_C446970
        End If

    Comp_C446974:

        ' SelParamPrazo
        sCurrComponent = "SelParamPrazo"
        QueryC446974 = con.CreateCommand()
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "SELECT DECODE(DIAS_IND,0, DIAS_DEFAULT, DIAS_IND)"
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "FROM ("
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "SELECT WF_TP_PRODUTO.DIAS_DEFAULT_PRZ_ENTREGA_IND DIAS_DEFAULT,"
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "NVL((SELECT MAX(WF_TBPR_TP_GRP_PROG.DIAS_PE_INDUSTRIAL)"
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "FROM WF_TBPR_TP_GRP_PROG, WF_GRUPO_PRODUTO"
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "WHERE WF_TBPR_TP_GRP_PROG.COD_GRUPO = WF_GRUPO_PRODUTO.COD_GRUPO "
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "AND WF_TBPR_TP_GRP_PROG.TABELA_PRECO_VENDA = " & _ 
ConvertValue(C446968, C446968DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "AND WF_TBPR_TP_GRP_PROG.TP_PRODUTO = WF_TP_PRODUTO.TP_PRODUTO "
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "AND WF_GRUPO_PRODUTO.SIGLA_PROD = " & _ 
ConvertValue(C446970, C446970DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "AND WF_GRUPO_PRODUTO.ACESSO = " & _ 
ConvertValue(C446966, C446966DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ),0) DIAS_IND"
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "FROM WF_TP_PRODUTO"
        QueryC446974.CommandText = QueryC446974.CommandText & " " & "WHERE WF_TP_PRODUTO.TP_PRODUTO =  " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446974.CommandText = QueryC446974.CommandText & " " & ")"
        QueryC446974.Transaction = txn
        RsC446974 = QueryC446974.ExecuteReader()
        Dim iC446974 As Short
        ReDim C446974DataType(RsC446974.FieldCount)
        For iC446974 = 0 to RsC446974.FieldCount - 1
            Select Case RsC446974.GetDataTypeName(iC446974).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C446974DataType(iC446974 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C446974DataType(iC446974 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C446974DataType(iC446974 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC446974
        RsC446974_EOF = Not RsC446974.Read()

        GoTo Comp_C446993

    Comp_C446975:

        ' attPrazoEntrega
        sCurrComponent = "attPrazoEntrega"
        C446975DataType = 4
        C394901 = fn_ConvertValueCompiled(RsC447162(0) , 2, AKB_DecimalPoint)
        C446975 = True
        GoTo Comp_C446976

    Comp_C446976:

        ' UpPedItens
        sCurrComponent = "UpPedItens"
        QueryC446976 = con.CreateCommand()
        QueryC446976.CommandText = QueryC446976.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS"
        QueryC446976.CommandText = QueryC446976.CommandText & " " & "SET WF_PEDIDO_ITENS.Prazo_Entrega = " & _ 
ConvertValue(C394901, C394901DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446976.CommandText = QueryC446976.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446976.CommandText = QueryC446976.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446976.CommandText = QueryC446976.CommandText & " " & "AND WF_PEDIDO_ITENS.Seq_Item  = " & _ 
ConvertValue(C446967, C446967DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC446976.Transaction = txn
        C446976 = QueryC446976.ExecuteNonQuery()
        C446976DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C446977

    Comp_C446977:

        ' nextVal
        sCurrComponent = "nextVal"
        C446977DataType = 4
        RsC446954_EOF = Not RsC446954.Read()
        C446977 = True

        GoTo Comp_C446969

    Comp_C446993:

        ' SomPrazoEntrega
        sCurrComponent = "SomPrazoEntrega"
        C446993DataType = 2
        C446993 = DateAdd("d", RsC446974(0), C394880)
        GoTo Comp_C447162

    Comp_C447162:

        ' SelDiaSemana
        sCurrComponent = "SelDiaSemana"
        QueryC447162 = con.CreateCommand()
        QueryC447162.CommandText = QueryC447162.CommandText & " " & "SELECT CASE WHEN TO_NUMBER(TO_CHAR(" & _ 
ConvertValue(C446993, C446993DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",'D')) IN (1,7) THEN"
        QueryC447162.CommandText = QueryC447162.CommandText & " " & "NEXT_DAY(" & _ 
ConvertValue(C446993, C446993DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",2)"
        QueryC447162.CommandText = QueryC447162.CommandText & " " & "ELSE"
        QueryC447162.CommandText = QueryC447162.CommandText & " " & "" & _ 
ConvertValue(C446993, C446993DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC447162.CommandText = QueryC447162.CommandText & " " & "END"
        QueryC447162.CommandText = QueryC447162.CommandText & " " & "FROM DUAL"
        QueryC447162.Transaction = txn
        RsC447162 = QueryC447162.ExecuteReader()
        Dim iC447162 As Short
        ReDim C447162DataType(RsC447162.FieldCount)
        For iC447162 = 0 to RsC447162.FieldCount - 1
            Select Case RsC447162.GetDataTypeName(iC447162).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C447162DataType(iC447162 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C447162DataType(iC447162 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C447162DataType(iC447162 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC447162
        RsC447162_EOF = Not RsC447162.Read()

        GoTo Comp_C446975

    Comp_C467430:

        ' Prazo Entrega Cliente Tipo 1
        sCurrComponent = "Prazo Entrega Cliente Tipo 1"
        C467430DataType = 2
        C467430 = DateAdd("d", C501562, C482982)
        GoTo Comp_C482595

    Comp_C467431:

        ' selProgSemPrazoEntregaCliente
        sCurrComponent = "selProgSemPrazoEntregaCliente"
        QueryC467431 = con.CreateCommand()
        QueryC467431.CommandText = QueryC467431.CommandText & " " & "SELECT COUNT(*)"
        QueryC467431.CommandText = QueryC467431.CommandText & " " & "FROM wf_pedido_itens"
        QueryC467431.CommandText = QueryC467431.CommandText & " " & "WHERE Tp_Produto        = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC467431.CommandText = QueryC467431.CommandText & " " & "AND pedido              = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC467431.CommandText = QueryC467431.CommandText & " " & "AND Tipo_Ped            = 2"
        QueryC467431.CommandText = QueryC467431.CommandText & " " & "AND Prazo_Entr_Cliente IS NULL"
        QueryC467431.Transaction = txn
        RsC467431 = QueryC467431.ExecuteReader()
        Dim iC467431 As Short
        ReDim C467431DataType(RsC467431.FieldCount)
        For iC467431 = 0 to RsC467431.FieldCount - 1
            Select Case RsC467431.GetDataTypeName(iC467431).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C467431DataType(iC467431 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C467431DataType(iC467431 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C467431DataType(iC467431 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC467431
        RsC467431_EOF = Not RsC467431.Read()

        GoTo Comp_C467432

    Comp_C467432:

        ' Pedido Prog Sem Prazo Entrega?
        sCurrComponent = "Pedido Prog Sem Prazo Entrega?"
        C467432 = (1=0)
        C467432DataType = AKBTypeConst.cAKBTypeLogical
        If C467432 Then
            GoTo Comp_C467434
        Else
            GoTo Comp_C394881
        End If

    Comp_C467434:

        ' msgPrazoEntregaCliente
        sCurrComponent = "msgPrazoEntregaCliente"
        C467434 = 1
        C467434DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C467435

    Comp_C467435:

        ' atFalse
        sCurrComponent = "atFalse"
        C467435DataType = 4
        C394877 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C467435 = True
        GoTo Comp_C394878

    Comp_C482594:

        ' upPrazoCliente
        sCurrComponent = "upPrazoCliente"
        QueryC482594 = con.CreateCommand()
        QueryC482594.CommandText = QueryC482594.CommandText & " " & "UPDATE WF_PEDIDO_ITENS"
        QueryC482594.CommandText = QueryC482594.CommandText & " " & "SET Prazo_Entr_Cliente = " & _ 
ConvertValue(C467430, C467430DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482594.CommandText = QueryC482594.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482594.CommandText = QueryC482594.CommandText & " " & "AND WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482594.CommandText = QueryC482594.CommandText & " " & "AND WF_PEDIDO_ITENS.Tipo_Ped = 1"
        QueryC482594.Transaction = txn
        C482594 = QueryC482594.ExecuteNonQuery()
        C482594DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C482634

    Comp_C482595:

        ' selDataMenorSistema
        sCurrComponent = "selDataMenorSistema"
        QueryC482595 = con.CreateCommand()
        QueryC482595.CommandText = QueryC482595.CommandText & " " & "select count(*)"
        QueryC482595.CommandText = QueryC482595.CommandText & " " & "from WF_PEDIDO_ITENS"
        QueryC482595.CommandText = QueryC482595.CommandText & " " & "where Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482595.CommandText = QueryC482595.CommandText & " " & "and pedido = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482595.CommandText = QueryC482595.CommandText & " " & "and Tipo_Ped = 1"
        QueryC482595.CommandText = QueryC482595.CommandText & " " & "and Prazo_Entr_Cliente < sysdate"
        QueryC482595.Transaction = txn
        RsC482595 = QueryC482595.ExecuteReader()
        Dim iC482595 As Short
        ReDim C482595DataType(RsC482595.FieldCount)
        For iC482595 = 0 to RsC482595.FieldCount - 1
            Select Case RsC482595.GetDataTypeName(iC482595).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C482595DataType(iC482595 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C482595DataType(iC482595 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C482595DataType(iC482595 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC482595
        RsC482595_EOF = Not RsC482595.Read()

        GoTo Comp_C482596

    Comp_C482596:

        ' Data Menor Que Sitema?
        sCurrComponent = "Data Menor Que Sitema?"
        C482596 = (1=0)
        C482596DataType = AKBTypeConst.cAKBTypeLogical
        If C482596 Then
            GoTo Comp_C482597
        Else
            GoTo Comp_C394879
        End If

    Comp_C482597:

        ' msgPrazoMenorSistema
        sCurrComponent = "msgPrazoMenorSistema"
        C482597 = 1
        C482597DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C467435

    Comp_C482598:

        ' upPrazoClienteComercialMax
        sCurrComponent = "upPrazoClienteComercialMax"
        QueryC482598 = con.CreateCommand()
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "UPDATE wf_pedido p"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "SET P.Prazo_Faturamento_Comercial ="
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  (SELECT max(Pi.Prazo_Entrega_Comercial)"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  FROM wf_pedido_itens pi"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  WHERE pi.Tp_Produto = p.Tp_Produto"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  AND pi.pedido       = p.pedido"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  ),"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "p.Prazo_Entr_Cliente ="
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  (SELECT max(pi.Prazo_Entr_Cliente)"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  FROM wf_pedido_itens pi"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  WHERE pi.Tp_Produto = p.Tp_Produto"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  AND pi.pedido       = p.pedido"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "  )"
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "WHERE p.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482598.CommandText = QueryC482598.CommandText & " " & "AND p.pedido       = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482598.Transaction = txn
        C482598 = QueryC482598.ExecuteNonQuery()
        C482598DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C394895

    Comp_C482633:

        ' insLogPrazoComercial
        sCurrComponent = "insLogPrazoComercial"
        QueryC482633 = con.CreateCommand()
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "INSERT"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "INTO Wf_Log_Prazo_Comercial"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "  ("
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "    id,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "    Tp_Produto,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "    pedido,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "    Seq_Item,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "    Prazo_Comercial,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "    login,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "    data"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "  )"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "SELECT seq_Wf_Log_Prazo_Comercial.nextval,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "  Pi.Tp_Produto,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "  pi.pedido,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "  Pi.Seq_Item,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "  Pi.Prazo_Entrega_Comercial,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "  USER,"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "  sysdate"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "FROM wf_pedido_itens pi"
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "WHERE Pi.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "AND pi.pedido       = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482633.CommandText = QueryC482633.CommandText & " " & "AND pi.Tipo_Ped = 1"
        QueryC482633.Transaction = txn
        C482633 = QueryC482633.ExecuteNonQuery()
        C482633DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C408168

    Comp_C482634:

        ' insLogPrazoCliente
        sCurrComponent = "insLogPrazoCliente"
        QueryC482634 = con.CreateCommand()
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "INSERT"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "INTO WF_LOG_PRAZO_ITEM_CLIENTE"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "  ("
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "    id,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "    Tp_Produto,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "    pedido,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "    Seq_Item,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "    Prazo_Cliente,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "    login,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "    data"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "  )"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "SELECT seq_WF_LOG_PRAZO_ITEM_CLIENTE.nextval,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "  Pi.Tp_Produto,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "  pi.pedido,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "  Pi.Seq_Item,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "  Pi.Prazo_Entr_Cliente,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "  USER,"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "  sysdate"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "FROM wf_pedido_itens pi"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "WHERE Pi.Tp_Produto = " & _ 
ConvertValue(P64627, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "AND pi.pedido       = " & _ 
ConvertValue(P64626, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "AND pi.Tipo_Ped = 1"
        QueryC482634.CommandText = QueryC482634.CommandText & " " & "AND Prazo_Entr_Cliente is null"
        QueryC482634.Transaction = txn
        C482634 = QueryC482634.ExecuteNonQuery()
        C482634DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C482598

    Comp_C482982:

        ' Data Sistema
        sCurrComponent = "Data Sistema"
        C482982DataType = 2
        C482982 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C394882

    Comp_C501545:

        ' DiasPE
        sCurrComponent = "DiasPE"
        QueryC501545 = con.CreateCommand()
        QueryC501545.CommandText = QueryC501545.CommandText & " " & "SELECT NVL(WF_TIPO_PED.Dias_Acres_PE_Cliente_Segunda,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Cliente_Terca,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Cliente_Quarta,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Cliente_Quinta,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Cliente_Sexta,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Cliente_Sabado,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Cliente_Domingo,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Segunda,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Terca,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Quarta,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Quinta,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Sexta,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Sabado,0) , NVL(WF_TIPO_PED.Dias_Acres_PE_Domingo,0) FROM AKBUSER01.WF_TIPO_PED WHERE WF_TIPO_PED.Tipo_Ped = 1 "
        QueryC501545.Transaction = txn
        RsC501545 = QueryC501545.ExecuteReader()
        Dim iC501545 As Short
        ReDim C501545DataType(RsC501545.FieldCount)
        For iC501545 = 0 to RsC501545.FieldCount - 1
            Select Case RsC501545.GetDataTypeName(iC501545).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C501545DataType(iC501545 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C501545DataType(iC501545 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C501545DataType(iC501545 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC501545
        RsC501545_EOF = Not RsC501545.Read()

        GoTo Comp_C501546

    Comp_C501546:

        ' DiasPEcliSeg
        sCurrComponent = "DiasPEcliSeg"
        C501546DataType = 0
        C501546 = RsC501545(0)
        C501546DataType = C501545Datatype(1)
        If C501546DataType = AKBTypeConst.cAKBTypeString Then
          C501546 = IIF(IsDBNull(C501546), C501546, Strings.RTrim(C501546))
        End If 

        GoTo Comp_C501547

    Comp_C501547:

        ' DiasPEcliTer
        sCurrComponent = "DiasPEcliTer"
        C501547DataType = 0
        C501547DataType = C501545Datatype(2)
        If C501547DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(1)) Then
          C501547 = Strings.RTrim(RsC501545(1))
        Else 
          C501547 = RsC501545(1)
        End If 

        GoTo Comp_C501548

    Comp_C501548:

        ' DiasPEcliQua
        sCurrComponent = "DiasPEcliQua"
        C501548DataType = 0
        C501548DataType = C501545Datatype(3)
        If C501548DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(2)) Then
          C501548 = Strings.RTrim(RsC501545(2))
        Else 
          C501548 = RsC501545(2)
        End If 

        GoTo Comp_C501550

    Comp_C501550:

        ' DiasPEcliQui
        sCurrComponent = "DiasPEcliQui"
        C501550DataType = 0
        C501550DataType = C501545Datatype(4)
        If C501550DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(3)) Then
          C501550 = Strings.RTrim(RsC501545(3))
        Else 
          C501550 = RsC501545(3)
        End If 

        GoTo Comp_C501551

    Comp_C501551:

        ' DiasPEcliSex
        sCurrComponent = "DiasPEcliSex"
        C501551DataType = 0
        C501551DataType = C501545Datatype(5)
        If C501551DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(4)) Then
          C501551 = Strings.RTrim(RsC501545(4))
        Else 
          C501551 = RsC501545(4)
        End If 

        GoTo Comp_C501552

    Comp_C501552:

        ' DiasPEcliSab
        sCurrComponent = "DiasPEcliSab"
        C501552DataType = 0
        C501552DataType = C501545Datatype(6)
        If C501552DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(5)) Then
          C501552 = Strings.RTrim(RsC501545(5))
        Else 
          C501552 = RsC501545(5)
        End If 

        GoTo Comp_C501553

    Comp_C501553:

        ' DiasPEcliDom
        sCurrComponent = "DiasPEcliDom"
        C501553DataType = 0
        C501553DataType = C501545Datatype(7)
        If C501553DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(6)) Then
          C501553 = Strings.RTrim(RsC501545(6))
        Else 
          C501553 = RsC501545(6)
        End If 

        GoTo Comp_C501554

    Comp_C501554:

        ' DiasPEcomSeg
        sCurrComponent = "DiasPEcomSeg"
        C501554DataType = 0
        C501554DataType = C501545Datatype(8)
        If C501554DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(7)) Then
          C501554 = Strings.RTrim(RsC501545(7))
        Else 
          C501554 = RsC501545(7)
        End If 

        GoTo Comp_C501555

    Comp_C501555:

        ' DiasPEcomTer
        sCurrComponent = "DiasPEcomTer"
        C501555DataType = 0
        C501555DataType = C501545Datatype(9)
        If C501555DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(8)) Then
          C501555 = Strings.RTrim(RsC501545(8))
        Else 
          C501555 = RsC501545(8)
        End If 

        GoTo Comp_C501556

    Comp_C501556:

        ' DiasPEcomQua
        sCurrComponent = "DiasPEcomQua"
        C501556DataType = 0
        C501556DataType = C501545Datatype(10)
        If C501556DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(9)) Then
          C501556 = Strings.RTrim(RsC501545(9))
        Else 
          C501556 = RsC501545(9)
        End If 

        GoTo Comp_C501557

    Comp_C501557:

        ' DiasPEcomQui
        sCurrComponent = "DiasPEcomQui"
        C501557DataType = 0
        C501557DataType = C501545Datatype(11)
        If C501557DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(10)) Then
          C501557 = Strings.RTrim(RsC501545(10))
        Else 
          C501557 = RsC501545(10)
        End If 

        GoTo Comp_C501558

    Comp_C501558:

        ' DiasPEcomSex
        sCurrComponent = "DiasPEcomSex"
        C501558DataType = 0
        C501558DataType = C501545Datatype(12)
        If C501558DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(11)) Then
          C501558 = Strings.RTrim(RsC501545(11))
        Else 
          C501558 = RsC501545(11)
        End If 

        GoTo Comp_C501559

    Comp_C501559:

        ' DiasPEcomSab
        sCurrComponent = "DiasPEcomSab"
        C501559DataType = 0
        C501559DataType = C501545Datatype(13)
        If C501559DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(12)) Then
          C501559 = Strings.RTrim(RsC501545(12))
        Else 
          C501559 = RsC501545(12)
        End If 

        GoTo Comp_C501560

    Comp_C501560:

        ' DiasPEcomDom
        sCurrComponent = "DiasPEcomDom"
        C501560DataType = 0
        C501560DataType = C501545Datatype(14)
        If C501560DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC501545(13)) Then
          C501560 = Strings.RTrim(RsC501545(13))
        Else 
          C501560 = RsC501545(13)
        End If 

        GoTo Comp_C501561

    Comp_C501561:

        ' DiaSemana
        sCurrComponent = "DiaSemana"
        QueryC501561 = con.CreateCommand()
        QueryC501561.CommandText = QueryC501561.CommandText & " " & "SELECT TO_CHAR(To_Date(" & _ 
ConvertValue(C482982, C482982DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "), 'D') FROM DUAL"
        QueryC501561.Transaction = txn
        RsC501561 = QueryC501561.ExecuteReader()
        Dim iC501561 As Short
        ReDim C501561DataType(RsC501561.FieldCount)
        For iC501561 = 0 to RsC501561.FieldCount - 1
            Select Case RsC501561.GetDataTypeName(iC501561).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C501561DataType(iC501561 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C501561DataType(iC501561 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C501561DataType(iC501561 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC501561
        RsC501561_EOF = Not RsC501561.Read()

        GoTo Comp_C501562

    Comp_C501562:

        ' vDiasPEcli
        sCurrComponent = "vDiasPEcli"
        C501562 = 0
        C501562DataType = 1
        GoTo Comp_C501563

    Comp_C501563:

        ' vDiasPE
        sCurrComponent = "vDiasPE"
        C501563 = 0
        C501563DataType = 1
        GoTo Comp_C501564

    Comp_C501564:

        ' D=2seg?
        sCurrComponent = "D=2seg?"
        C501564 = (fn_ConvertValueCompiled(RsC501561(0), C501561DataType(1), AKB_DecimalPoint, False) = 2)
        C501564DataType = AKBTypeConst.cAKBTypeLogical
        If C501564 Then
            GoTo Comp_C501565
        Else
            GoTo Comp_C501567
        End If

    Comp_C501565:

        ' ATRIBUIVA1
        sCurrComponent = "ATRIBUIVA1"
        C501565DataType = 4
        C501562 = fn_ConvertValueCompiled(C501546 , 1, AKB_DecimalPoint)
        C501565 = True
        GoTo Comp_C501566

    Comp_C501566:

        ' ATRIBUIVA2
        sCurrComponent = "ATRIBUIVA2"
        C501566DataType = 4
        C501563 = fn_ConvertValueCompiled(C501554 , 1, AKB_DecimalPoint)
        C501566 = True
        GoTo Comp_C394885

    Comp_C501567:

        ' D=3ter?
        sCurrComponent = "D=3ter?"
        C501567 = (fn_ConvertValueCompiled(RsC501561(0), C501561DataType(1), AKB_DecimalPoint, False) = 3)
        C501567DataType = AKBTypeConst.cAKBTypeLogical
        If C501567 Then
            GoTo Comp_C501568
        Else
            GoTo Comp_C501570
        End If

    Comp_C501568:

        ' ATRIBUIVA11
        sCurrComponent = "ATRIBUIVA11"
        C501568DataType = 4
        C501562 = fn_ConvertValueCompiled(C501547 , 1, AKB_DecimalPoint)
        C501568 = True
        GoTo Comp_C501569

    Comp_C501569:

        ' ATRIBUIVA21
        sCurrComponent = "ATRIBUIVA21"
        C501569DataType = 4
        C501563 = fn_ConvertValueCompiled(C501555 , 1, AKB_DecimalPoint)
        C501569 = True
        GoTo Comp_C394885

    Comp_C501570:

        ' D=4qua?
        sCurrComponent = "D=4qua?"
        C501570 = (fn_ConvertValueCompiled(RsC501561(0), C501561DataType(1), AKB_DecimalPoint, False)  = 4)
        C501570DataType = AKBTypeConst.cAKBTypeLogical
        If C501570 Then
            GoTo Comp_C501571
        Else
            GoTo Comp_C501573
        End If

    Comp_C501571:

        ' ATRIBUIVA111
        sCurrComponent = "ATRIBUIVA111"
        C501571DataType = 4
        C501562 = fn_ConvertValueCompiled(C501548 , 1, AKB_DecimalPoint)
        C501571 = True
        GoTo Comp_C501572

    Comp_C501572:

        ' ATRIBUIVA211
        sCurrComponent = "ATRIBUIVA211"
        C501572DataType = 4
        C501563 = fn_ConvertValueCompiled(C501556 , 1, AKB_DecimalPoint)
        C501572 = True
        GoTo Comp_C394885

    Comp_C501573:

        ' D= 5qui?
        sCurrComponent = "D= 5qui?"
        C501573 = (fn_ConvertValueCompiled(RsC501561(0), C501561DataType(1), AKB_DecimalPoint, False) = 5)
        C501573DataType = AKBTypeConst.cAKBTypeLogical
        If C501573 Then
            GoTo Comp_C501575
        Else
            GoTo Comp_C501577
        End If

    Comp_C501575:

        ' ATRIBUIVA1111
        sCurrComponent = "ATRIBUIVA1111"
        C501575DataType = 4
        C501562 = fn_ConvertValueCompiled(C501550 , 1, AKB_DecimalPoint)
        C501575 = True
        GoTo Comp_C501576

    Comp_C501576:

        ' ATRIBUIVA2111
        sCurrComponent = "ATRIBUIVA2111"
        C501576DataType = 4
        C501563 = fn_ConvertValueCompiled(C501557 , 1, AKB_DecimalPoint)
        C501576 = True
        GoTo Comp_C394885

    Comp_C501577:

        ' D = 6sex?
        sCurrComponent = "D = 6sex?"
        C501577 = (fn_ConvertValueCompiled(RsC501561(0), C501561DataType(1), AKB_DecimalPoint, False) = 6)
        C501577DataType = AKBTypeConst.cAKBTypeLogical
        If C501577 Then
            GoTo Comp_C501579
        Else
            GoTo Comp_C501581
        End If

    Comp_C501579:

        ' ATRIBUIVA11111
        sCurrComponent = "ATRIBUIVA11111"
        C501579DataType = 4
        C501562 = fn_ConvertValueCompiled(C501551 , 1, AKB_DecimalPoint)
        C501579 = True
        GoTo Comp_C501580

    Comp_C501580:

        ' ATRIBUIVA21111
        sCurrComponent = "ATRIBUIVA21111"
        C501580DataType = 4
        C501563 = fn_ConvertValueCompiled(C501558 , 1, AKB_DecimalPoint)
        C501580 = True
        GoTo Comp_C394885

    Comp_C501581:

        ' D = 7sab?
        sCurrComponent = "D = 7sab?"
        C501581 = (fn_ConvertValueCompiled(RsC501561(0), C501561DataType(1), AKB_DecimalPoint, False) = 7)
        C501581DataType = AKBTypeConst.cAKBTypeLogical
        If C501581 Then
            GoTo Comp_C501582
        Else
            GoTo Comp_C501584
        End If

    Comp_C501582:

        ' ATRIBUIVA111111
        sCurrComponent = "ATRIBUIVA111111"
        C501582DataType = 4
        C501562 = fn_ConvertValueCompiled(C501552 , 1, AKB_DecimalPoint)
        C501582 = True
        GoTo Comp_C501583

    Comp_C501583:

        ' ATRIBUIVA211111
        sCurrComponent = "ATRIBUIVA211111"
        C501583DataType = 4
        C501563 = fn_ConvertValueCompiled(C501559 , 1, AKB_DecimalPoint)
        C501583 = True
        GoTo Comp_C394885

    Comp_C501584:

        ' D = 1dom?
        sCurrComponent = "D = 1dom?"
        C501584 = (fn_ConvertValueCompiled(RsC501561(0), C501561DataType(1), AKB_DecimalPoint, False) = 1)
        C501584DataType = AKBTypeConst.cAKBTypeLogical
        If C501584 Then
            GoTo Comp_C501585
        Else
            GoTo Comp_C394885
        End If

    Comp_C501585:

        ' ATRIBUIVA1111111
        sCurrComponent = "ATRIBUIVA1111111"
        C501585DataType = 4
        C501562 = fn_ConvertValueCompiled(C501553 , 1, AKB_DecimalPoint)
        C501585 = True
        GoTo Comp_C501586

    Comp_C501586:

        ' ATRIBUIVA2111111
        sCurrComponent = "ATRIBUIVA2111111"
        C501586DataType = 4
        C501563 = fn_ConvertValueCompiled(C501560 , 1, AKB_DecimalPoint)
        C501586 = True
        GoTo Comp_C394885

    Exit_R17770:

        Exit Function

    Err_R17770:
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
