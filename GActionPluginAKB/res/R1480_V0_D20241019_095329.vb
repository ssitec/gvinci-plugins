Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R1480

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

    Public Shared Function R1480(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P2280 As Double, ByVal P2286 As Double, ByVal P2289 As Double, ByVal P2305 As Double, ByVal P20449 As Object, ByVal P90926 As Object) As DataTable
    ' 
    ' 1480 - Liberação de crédito automático - 1 pedido - Versão: 0
    ' 
        'On Error GoTo Err_R1480
        Dim sCurrComponent as String

        Dim QueryC8580 As New Object
        Dim RsC8580 As New Object
        Dim C8580DataType() As Short
        Dim RsC8580_EOF As Boolean
        Dim C8584DataType() As Short
        Dim C8585 As Object
        Dim C8585DataType As Short
        Dim C8586 As Boolean
        Dim C8586DataType As Short
        Dim C8589 As Object
        Dim C8589DataType As Short
        Dim C8590 As Object
        Dim C8590DataType As Short
        Dim QueryC8592 As New Object
        Dim RsC8592 As New Object
        Dim C8592DataType() As Short
        Dim RsC8592_EOF As Boolean
        Dim C8594 As Object
        Dim C8594DataType As Short
        Dim C8596 As Object
        Dim C8596DataType As Short
        Dim C8597 As Object
        Dim C8597DataType As Short
        Dim C8599 As Boolean
        Dim C8599DataType As Short
        Dim C8602 As Boolean
        Dim C8602DataType As Short
        Dim C8603 As Double
        Dim C8603DataType As Short
        Dim C8604 As Short
        Dim C8604DataType As Short
        Dim C8604ReturnDataType() As Short

        Dim QueryC8606 As New Object
        Dim C8606 As Integer
        Dim C8606DataType As Short
        Dim QueryC8607 As New Object
        Dim C8607 As Integer
        Dim C8607DataType As Short
        Dim C8618 As Object
        Dim C8618DataType As Short
        Dim QueryC8687 As New Object
        Dim RsC8687 As New Object
        Dim C8687DataType() As Short
        Dim RsC8687_EOF As Boolean
        Dim C8688 As Object
        Dim C8688DataType As Short
        Dim C8690 As Object
        Dim C8690DataType As Short
        Dim C8692 As Boolean
        Dim C8692DataType As Short
        Dim QueryC8708 As New Object
        Dim RsC8708 As New Object
        Dim C8708DataType() As Short
        Dim RsC8708_EOF As Boolean
        Dim C8709 As Object
        Dim C8709DataType As Short
        Dim C8710 As Boolean
        Dim C8710DataType As Short
        Dim C8711 As Object
        Dim C8711DataType As Short
        Dim C8712 As Boolean
        Dim C8712DataType As Short
        Dim QueryC8824 As New Object
        Dim RsC8824 As New Object
        Dim C8824DataType() As Short
        Dim RsC8824_EOF As Boolean
        Dim C8887 As Object
        Dim C8887DataType As Short
        Dim C8889 As Boolean
        Dim C8889DataType As Short
        Dim C8891 As Short
        Dim C8891DataType As Short
        Dim C8891ReturnDataType() As Short

        Dim C8892 As Short
        Dim C8892DataType As Short
        Dim C8892ReturnDataType() As Short

        Dim C8893 As Object
        Dim C8893DataType As Short
        Dim C8894 As Short
        Dim C8894DataType As Short
        Dim C8894ReturnDataType() As Short

        Dim C8895 As Boolean
        Dim C8895DataType As Short
        Dim C8896 As Boolean
        Dim C8896DataType As Short
        Dim QueryC12958 As New Object
        Dim RsC12958 As New Object
        Dim C12958DataType() As Short
        Dim RsC12958_EOF As Boolean
        Dim C12959 As Object
        Dim C12959DataType As Short
        Dim QueryC53541 As New Object
        Dim RsC53541 As New Object
        Dim C53541DataType() As Short
        Dim RsC53541_EOF As Boolean
        Dim C53542 As Object
        Dim C53542DataType As Short
        Dim C53543 As Object
        Dim C53543DataType As Short
        Dim C53544 As Boolean
        Dim C53544DataType As Short
        Dim C96652 As Boolean
        Dim C96652DataType As Short
        Dim C96653 As Boolean
        Dim C96653DataType As Short
        Dim C97797 As Object
        Dim C97797DataType As Short
        Dim C97798 As Boolean
        Dim C97798DataType As Short
        Dim QueryC111652 As New Object
        Dim RsC111652 As New Object
        Dim C111652DataType() As Short
        Dim RsC111652_EOF As Boolean
        Dim C111653 As Boolean
        Dim C111653DataType As Short
        Dim QueryC111654 As New Object
        Dim C111654 As Integer
        Dim C111654DataType As Short
        Dim QueryC113272 As New Object
        Dim C113272 As Integer
        Dim C113272DataType As Short
        Dim QueryC125373 As New Object
        Dim RsC125373 As New Object
        Dim C125373DataType() As Short
        Dim RsC125373_EOF As Boolean
        Dim QueryC125377 As New Object
        Dim RsC125377 As New Object
        Dim C125377DataType() As Short
        Dim RsC125377_EOF As Boolean
        Dim C125379 As Boolean
        Dim C125379DataType As Short
        Dim QueryC126475 As New Object
        Dim RsC126475 As New Object
        Dim C126475DataType() As Short
        Dim RsC126475_EOF As Boolean
        Dim C126476 As Boolean
        Dim C126476DataType As Short
        Dim QueryC128605 As New Object
        Dim RsC128605 As New Object
        Dim C128605DataType() As Short
        Dim RsC128605_EOF As Boolean
        Dim C128606 As Object
        Dim C128606DataType As Short
        Dim C128607 As Object
        Dim C128607DataType As Short
        Dim C128608 As Object
        Dim C128608DataType As Short
        Dim C128609 As Boolean
        Dim C128609DataType As Short
        Dim C128610 As Boolean
        Dim C128610DataType As Short
        Dim C128611 As Boolean
        Dim C128611DataType As Short
        Dim C128612 As Short
        Dim C128612DataType As Short
        Dim C128612ReturnDataType() As Short

        Dim C139381 As DataTable
        Dim C139381CurrentRow As Integer
        Dim C139381DataType() As Short
        Dim C139383 As Boolean
        Dim C139383DataType As Short
        Dim QueryC161492 As New Object
        Dim RsC161492 As New Object
        Dim C161492DataType() As Short
        Dim RsC161492_EOF As Boolean
        Dim C161494 As Boolean
        Dim C161494DataType As Short
        Dim C161495 As Short
        Dim C161495DataType As Short
        Dim C161495ReturnDataType() As Short

        Dim QueryC163125 As New Object
        Dim RsC163125 As New Object
        Dim C163125DataType() As Short
        Dim RsC163125_EOF As Boolean
        Dim C163126 As Boolean
        Dim C163126DataType As Short
        Dim C163127 As Object
        Dim C163127DataType As Short
        Dim C163128 As Object
        Dim C163128DataType As Short
        Dim C165638 As DataTable
        Dim C165638CurrentRow As Integer
        Dim C165638DataType() As Short
        Dim C165639 As Boolean
        Dim C165639DataType As Short
        Dim C165641 As Boolean
        Dim C165641DataType As Short
        Dim C165642 As Boolean
        Dim C165642DataType As Short
        Dim QueryC236085 As New Object
        Dim RsC236085 As New Object
        Dim C236085DataType() As Short
        Dim RsC236085_EOF As Boolean
        Dim QueryC317472 As New Object
        Dim RsC317472 As New Object
        Dim C317472DataType() As Short
        Dim RsC317472_EOF As Boolean
        Dim C317473 As Boolean
        Dim C317473DataType As Short
        Dim C320568 As Boolean
        Dim C320568DataType As Short
        Dim C331780 As Boolean
        Dim C331780DataType As Short
        Dim C332854 As DataTable
        Dim C332854CurrentRow As Integer
        Dim C332854DataType() As Short
        Dim C502251 As DataTable
        Dim C502251CurrentRow As Integer
        Dim C502251DataType() As Short
        Dim C571768 As Boolean
        Dim C571768DataType As Short
        Dim QueryC572052 As New Object
        Dim RsC572052 As New Object
        Dim C572052DataType() As Short
        Dim RsC572052_EOF As Boolean
        Dim C572053 As Object
        Dim C572053DataType As Short
        Dim C572054 As Boolean
        Dim C572054DataType As Short
        Dim C572056 As Object
        Dim C572056DataType As Short
        Dim QueryC572057 As New Object
        Dim RsC572057 As New Object
        Dim C572057DataType() As Short
        Dim RsC572057_EOF As Boolean
        Dim C572058 As Boolean
        Dim C572058DataType As Short
        Dim C572059 As Short
        Dim C572059DataType As Short
        Dim C572059ReturnDataType() As Short

        Dim QueryC572060 As New Object
        Dim RsC572060 As New Object
        Dim C572060DataType() As Short
        Dim RsC572060_EOF As Boolean
        Dim C572061 As Boolean
        Dim C572061DataType As Short
        Dim C572062 As Short
        Dim C572062DataType As Short
        Dim C572062ReturnDataType() As Short

        Dim QueryC572063 As New Object
        Dim RsC572063 As New Object
        Dim C572063DataType() As Short
        Dim RsC572063_EOF As Boolean
        Dim C572064 As Boolean
        Dim C572064DataType As Short
        Dim C572065 As Short
        Dim C572065DataType As Short
        Dim C572065ReturnDataType() As Short

        Dim C572425 As Boolean
        Dim C572425DataType As Short
        Dim C572427 As Boolean
        Dim C572427DataType As Short
        Dim C572428 As Boolean
        Dim C572428DataType As Short
        Dim C572430 As Boolean
        Dim C572430DataType As Short
        Dim C572431 As Boolean
        Dim C572431DataType As Short
        Dim C572435 As Boolean
        Dim C572435DataType As Short
        Dim C572437 As Boolean
        Dim C572437DataType As Short
        Dim C572438 As Boolean
        Dim C572438DataType As Short
        Dim C572440 As Boolean
        Dim C572440DataType As Short
        Dim C572441 As Boolean
        Dim C572441DataType As Short
        Dim C572443 As Boolean
        Dim C572443DataType As Short
        P20449 = IIf(IsDBNull(P20449), 1, P20449)

        P90926 = IIf(IsDBNull(P90926), 0, P90926)

        ReDim ReturnDatatype(0)

        GoTo Comp_C8893

    Comp_C8580:

        ' SelCred
        sCurrComponent = "SelCred"
        QueryC8580 = con.CreateCommand()
        QueryC8580.CommandText = QueryC8580.CommandText & " " & "SELECT "
        QueryC8580.CommandText = QueryC8580.CommandText & " " & "WF_LIM_CRED.Limite_Credito , "
        QueryC8580.CommandText = QueryC8580.CommandText & " " & "WF_LIM_CRED.Data_Vencimento , "
        QueryC8580.CommandText = QueryC8580.CommandText & " " & "WF_LIM_CRED.Cod_Cliente "
        QueryC8580.CommandText = QueryC8580.CommandText & " " & ""
        QueryC8580.CommandText = QueryC8580.CommandText & " " & "FROM AKBUSER01.WF_LIM_CRED "
        QueryC8580.CommandText = QueryC8580.CommandText & " " & ""
        QueryC8580.CommandText = QueryC8580.CommandText & " " & "WHERE WF_LIM_CRED.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC8580.CommandText = QueryC8580.CommandText & " " & ""
        QueryC8580.CommandText = QueryC8580.CommandText & " " & "AND  WF_LIM_CRED.Sequencia  = ( SELECT MAX(WF_LIM_CRED.Sequencia ) FROM AKBUSER01.WF_LIM_CRED  WHERE WF_LIM_CRED.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC8580.CommandText = QueryC8580.CommandText & " " & ""
        QueryC8580.Transaction = txn
        RsC8580 = QueryC8580.ExecuteReader()
        Dim iC8580 As Short
        ReDim C8580DataType(RsC8580.FieldCount)
        For iC8580 = 0 to RsC8580.FieldCount - 1
            Select Case RsC8580.GetDataTypeName(iC8580).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C8580DataType(iC8580 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C8580DataType(iC8580 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C8580DataType(iC8580 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC8580
        RsC8580_EOF = Not RsC8580.Read()

        GoTo Comp_C8585

    Comp_C8584:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C8584 As DataTable = New DataTable()
        tb_C8584.Columns.Add("RetVar" & "")
        Dim row_C8584 As DataRow = tb_C8584.NewRow()
        row_C8584(0) = C8893
        tb_C8584.Rows.Add(row_C8584)
        R1480 = tb_C8584
        ReDim C8584DataType(1)
        C8584DataType(1) = C8893DataType
        ReturnDataType = C8584DataType

        GoTo Exit_R1480

    Comp_C8585:

        ' Eof2
        sCurrComponent = "Eof2"
        C8585DataType = 4
        C8585 = RsC8580_EOF
        GoTo Comp_C572431

    Comp_C8586:

        ' Eof2=True
        sCurrComponent = "Eof2=True"
        C8586 = (fn_ConvertValueCompiled(C8585, C8585DataType, AKB_DecimalPoint, False) =1)
        C8586DataType = AKBTypeConst.cAKBTypeLogical
        If C8586 Then
            GoTo Comp_C8584
        Else
            GoTo Comp_C8589
        End If

    Comp_C8589:

        ' Limite
        sCurrComponent = "Limite"
        C8589DataType = 0
        C8589 = RsC8580(0)
        C8589DataType = C8580Datatype(1)
        If C8589DataType = AKBTypeConst.cAKBTypeString Then
          C8589 = IIF(IsDBNull(C8589), C8589, Strings.RTrim(C8589))
        End If 

        GoTo Comp_C8590

    Comp_C8590:

        ' Validade
        sCurrComponent = "Validade"
        C8590DataType = 0
        C8590DataType = C8580Datatype(2)
        If C8590DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC8580(1)) Then
          C8590 = Strings.RTrim(RsC8580(1))
        Else 
          C8590 = RsC8580(1)
        End If 

        GoTo Comp_C8596

    Comp_C8592:

        ' SelLiberad
        sCurrComponent = "SelLiberad"
        QueryC8592 = con.CreateCommand()
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "SELECT WF_PEDIDO.Cod_Cliente,"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "       NVL(ROUND(SUM((WF_PED_SEQ_ITENS.Quantidade_Pedida_Original * WF_PEDIDO_ITENS.Preco_Unit) * NVL(M.MINIMO , 1)*100/(100-nvl(D.porc_desconto,0))), 2), 0) "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "       "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO, AKBUSER01.WF_PEDIDO_ITENS , AKBUSER01.WF_PED_SEQ_ITENS,"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "     WF_MOEDAS_ULTIMA_COTACAO M, WF_PED_SEQ_DESCONTO D, WF_TP_PRODUTO"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & ""
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "WHERE WF_PEDIDO.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO.Pedido > 0 "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO.Flag_Pos_Ped IN ( 2, 4, 5, 6, 9, 453 )"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO_ITENS.Flag_Pos_Ped NOT IN (7,8,10,11,12,13,299)"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Tp_Produto = WF_PEDIDO.Tp_Produto "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Pedido = WF_PEDIDO.Pedido "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Seq = WF_PEDIDO.Sequencia_Atual "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO_ITENS.Tp_Produto = WF_PED_SEQ_ITENS.Tp_Produto "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = WF_PED_SEQ_ITENS.Pedido "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO_ITENS.Seq_Item = WF_PED_SEQ_ITENS.Seq_Item "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PED_SEQ_ITENS.Quantidade_Pedida_Original > 0"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO.pedido = D.pedido (+)"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & ""
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  D.Tp_Produto (+) = WF_PEDIDO.Tp_Produto"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  D.tipo_desconto (+) = 280    "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  D.seq (+) = 1"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  " & _ 
ConvertValue(RsC236085(0), C236085DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = M.Codigo_Moeda_Valor  (+) "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO.Codigo_Moeda = M.Codigo_Moeda"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & ""
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_TP_PRODUTO.Tp_Produto = WF_PEDIDO.Tp_Produto"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_TP_PRODUTO.Tp_Tab_Preco <> 'Interno'"
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "AND  WF_PEDIDO.Pedido <> " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC8592.CommandText = QueryC8592.CommandText & " " & ""
        QueryC8592.CommandText = QueryC8592.CommandText & " " & "GROUP BY WF_PEDIDO.Cod_Cliente"
        QueryC8592.Transaction = txn
        RsC8592 = QueryC8592.ExecuteReader()
        Dim iC8592 As Short
        ReDim C8592DataType(RsC8592.FieldCount)
        For iC8592 = 0 to RsC8592.FieldCount - 1
            Select Case RsC8592.GetDataTypeName(iC8592).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C8592DataType(iC8592 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C8592DataType(iC8592 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C8592DataType(iC8592 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC8592
        RsC8592_EOF = Not RsC8592.Read()

        GoTo Comp_C8594

    Comp_C8594:

        ' VLiberado
        sCurrComponent = "VLiberado"
        C8594DataType = 0
        C8594DataType = C8592Datatype(2)
        If C8594DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC8592(1)) Then
          C8594 = Strings.RTrim(RsC8592(1))
        Else 
          C8594 = RsC8592(1)
        End If 

        GoTo Comp_C12958

    Comp_C8596:

        ' DateDiff
        sCurrComponent = "DateDiff"
        C8596DataType = 1
        C8596 = DateDiff("d", C8597, C8590)
        GoTo Comp_C8599

    Comp_C8597:

        ' SysDate
        sCurrComponent = "SysDate"
        C8597DataType = 2
        C8597 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy HH:mm:ss"))
        GoTo Comp_C125373

    Comp_C8599:

        ' Date<=0
        sCurrComponent = "Date<=0"
        C8599 = (fn_ConvertValueCompiled(C8596, C8596DataType, AKB_DecimalPoint, False) <=0)
        C8599DataType = AKBTypeConst.cAKBTypeLogical
        If C8599 Then
            GoTo Comp_C8584
        Else
            GoTo Comp_C161492
        End If

    Comp_C8602:

        ' TotLib>Lim
        sCurrComponent = "TotLib>Lim"
        C8602 = (fn_ConvertValueCompiled(C8603, C8603DataType, AKB_DecimalPoint, False) >fn_ConvertValueCompiled(C8589, C8589DataType, AKB_DecimalPoint, False))
        C8602DataType = AKBTypeConst.cAKBTypeLogical
        If C8602 Then
            GoTo Comp_C8604
        Else
            GoTo Comp_C96652
        End If

    Comp_C8603:

        ' Tot_Lib
        sCurrComponent = "Tot_Lib"
        C8603 = fn_ConvertValueCompiled(C8594, C8594DataType, AKB_DecimalPoint, True) + fn_ConvertValueCompiled(C8688, C8688DataType, AKB_DecimalPoint, True) + fn_ConvertValueCompiled(C12959, C12959DataType, AKB_DecimalPoint, True)
        C8603DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C8602

    Comp_C8604:

        ' MSG2
        sCurrComponent = "MSG2"
        Dim row_C8604 As DataRow = msg.NewRow()
        row_C8604(0) = "O pedido " & _ 
P2280 & "  não pode ser liberado porque excede o limite de crédito do cliente.  " & Chr(13) & "" & Chr(10) & "Pedidos em aberto: " & _ 
C8594 & " " & Chr(13) & "" & Chr(10) & "Duplic.   : " & _ 
C12959 & " " & Chr(13) & "" & Chr(10) & "Pedido  atual : " & _ 
C8688 & "" & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C8604)
        C8604DataType = AKBTypeConst.cAKBTypeNumeric
        C8604 = 1

        GoTo Comp_C8584

    Comp_C8606:

        ' Libera
        sCurrComponent = "Libera"
        QueryC8606 = con.CreateCommand()
        QueryC8606.CommandText = QueryC8606.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO SET WF_PEDIDO.Flag_Pos_Ped = 2 , "
        QueryC8606.CommandText = QueryC8606.CommandText & " " & "                                                 WF_PEDIDO.Data_Manut_Credito = " & _ 
ConvertValue(C8597, C8597DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC8606.CommandText = QueryC8606.CommandText & " " & "                                                 WF_PEDIDO.Liberado_Credito = 'Liberado' , "
        QueryC8606.CommandText = QueryC8606.CommandText & " " & "                                                 WF_PEDIDO.Usuario_Manut_Credito = " & _ 
ConvertValue(C8618, C8618DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC8606.CommandText = QueryC8606.CommandText & " " & "WHERE WF_PEDIDO.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC8606.Transaction = txn
        C8606 = QueryC8606.ExecuteNonQuery()
        C8606DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C8607

    Comp_C8607:

        ' LibItem
        sCurrComponent = "LibItem"
        QueryC8607 = con.CreateCommand()
        QueryC8607.CommandText = QueryC8607.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC8607.CommandText = QueryC8607.CommandText & " " & ""
        QueryC8607.CommandText = QueryC8607.CommandText & " " & "SET WF_PEDIDO_ITENS.Flag_Pos_Ped = 2 "
        QueryC8607.CommandText = QueryC8607.CommandText & " " & ""
        QueryC8607.CommandText = QueryC8607.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC8607.CommandText = QueryC8607.CommandText & " " & "AND WF_PEDIDO_ITENS.Flag_Pos_Ped = 1"
        QueryC8607.CommandText = QueryC8607.CommandText & " " & ""
        QueryC8607.Transaction = txn
        C8607 = QueryC8607.ExecuteNonQuery()
        C8607DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C332854

    Comp_C8618:

        ' UserName
        sCurrComponent = "UserName"
        C8618DataType = 3
        C8618 = Mid(con.ConnectionString, InStr(con.ConnectionString, "User Id=")+8, Instr(InStr(con.ConnectionString, "User Id="), con.ConnectionString, ";")-InStr(con.ConnectionString, "User Id=")-8)
        GoTo Comp_C8597

    Comp_C8687:

        ' SelValor
        sCurrComponent = "SelValor"
        QueryC8687 = con.CreateCommand()
        QueryC8687.CommandText = QueryC8687.CommandText & " " & "SELECT  WF_PEDIDO.Pedido, SUM( WF_PEDIDO.Total_Itens )"
        QueryC8687.CommandText = QueryC8687.CommandText & " " & ""
        QueryC8687.CommandText = QueryC8687.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO "
        QueryC8687.CommandText = QueryC8687.CommandText & " " & ""
        QueryC8687.CommandText = QueryC8687.CommandText & " " & "WHERE  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC8687.CommandText = QueryC8687.CommandText & " " & "AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC8687.CommandText = QueryC8687.CommandText & " " & ""
        QueryC8687.CommandText = QueryC8687.CommandText & " " & "GROUP BY  WF_PEDIDO.Pedido"
        QueryC8687.CommandText = QueryC8687.CommandText & " " & ""
        QueryC8687.Transaction = txn
        RsC8687 = QueryC8687.ExecuteReader()
        Dim iC8687 As Short
        ReDim C8687DataType(RsC8687.FieldCount)
        For iC8687 = 0 to RsC8687.FieldCount - 1
            Select Case RsC8687.GetDataTypeName(iC8687).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C8687DataType(iC8687 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C8687DataType(iC8687 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C8687DataType(iC8687 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC8687
        RsC8687_EOF = Not RsC8687.Read()

        GoTo Comp_C8690

    Comp_C8688:

        ' TotalPed
        sCurrComponent = "TotalPed"
        C8688DataType = 0
        C8688DataType = C8687Datatype(2)
        If C8688DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC8687(1)) Then
          C8688 = Strings.RTrim(RsC8687(1))
        Else 
          C8688 = RsC8687(1)
        End If 

        GoTo Comp_C8580

    Comp_C8690:

        ' EofVal
        sCurrComponent = "EofVal"
        C8690DataType = 4
        C8690 = RsC8687_EOF
        GoTo Comp_C8692

    Comp_C8692:

        ' EofVal=T
        sCurrComponent = "EofVal=T"
        C8692 = (fn_ConvertValueCompiled(C8690, C8690DataType, AKB_DecimalPoint, False) =1)
        C8692DataType = AKBTypeConst.cAKBTypeLogical
        If C8692 Then
            GoTo Comp_C8894
        Else
            GoTo Comp_C8688
        End If

    Comp_C8708:

        ' SelRest
        sCurrComponent = "SelRest"
        QueryC8708 = con.CreateCommand()
        QueryC8708.CommandText = QueryC8708.CommandText & " " & "SELECT WF_CLIENTE_RESTICAO.Tem_Restricao FROM AKBUSER01.WF_CLIENTE_RESTICAO "
        QueryC8708.CommandText = QueryC8708.CommandText & " " & ""
        QueryC8708.CommandText = QueryC8708.CommandText & " " & "WHERE WF_CLIENTE_RESTICAO.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC8708.CommandText = QueryC8708.CommandText & " " & ""
        QueryC8708.CommandText = QueryC8708.CommandText & " " & "AND WF_CLIENTE_RESTICAO.Sequencia  = (SELECT MAX(WF_CLIENTE_RESTICAO.Sequencia ) FROM AKBUSER01.WF_CLIENTE_RESTICAO "
        QueryC8708.CommandText = QueryC8708.CommandText & " " & "WHERE WF_CLIENTE_RESTICAO.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC8708.Transaction = txn
        RsC8708 = QueryC8708.ExecuteReader()
        Dim iC8708 As Short
        ReDim C8708DataType(RsC8708.FieldCount)
        For iC8708 = 0 to RsC8708.FieldCount - 1
            Select Case RsC8708.GetDataTypeName(iC8708).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C8708DataType(iC8708 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C8708DataType(iC8708 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C8708DataType(iC8708 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC8708
        RsC8708_EOF = Not RsC8708.Read()

        GoTo Comp_C8709

    Comp_C8709:

        ' EofRest
        sCurrComponent = "EofRest"
        C8709DataType = 4
        C8709 = RsC8708_EOF
        GoTo Comp_C8710

    Comp_C8710:

        ' EofRest=T
        sCurrComponent = "EofRest=T"
        C8710 = (fn_ConvertValueCompiled(C8709, C8709DataType, AKB_DecimalPoint, False) =1)
        C8710DataType = AKBTypeConst.cAKBTypeLogical
        If C8710 Then
            GoTo Comp_C8687
        Else
            GoTo Comp_C8711
        End If

    Comp_C8711:

        ' VRest
        sCurrComponent = "VRest"
        C8711DataType = 0
        C8711 = RsC8708(0)
        C8711DataType = C8708Datatype(1)
        If C8711DataType = AKBTypeConst.cAKBTypeString Then
          C8711 = IIF(IsDBNull(C8711), C8711, Strings.RTrim(C8711))
        End If 

        GoTo Comp_C8712

    Comp_C8712:

        ' VRest=T
        sCurrComponent = "VRest=T"
        C8712 = (fn_ConvertValueCompiled(C8711, C8711DataType, AKB_DecimalPoint, False) =1)
        C8712DataType = AKBTypeConst.cAKBTypeLogical
        If C8712 Then
            GoTo Comp_C572428
        Else
            GoTo Comp_C8687
        End If

    Comp_C8824:

        ' SelDupl
        sCurrComponent = "SelDupl"
        QueryC8824 = con.CreateCommand()
        QueryC8824.CommandText = QueryC8824.CommandText & " " & "SELECT  DR.Data_Vencimento "
        QueryC8824.CommandText = QueryC8824.CommandText & " " & ""
        QueryC8824.CommandText = QueryC8824.CommandText & " " & "FROM  AKBUSER01.WF_DUPLIC_RECEBER DR , AKBUSER01.WF_ESTAB_JURIDICO  E"
        QueryC8824.CommandText = QueryC8824.CommandText & " " & ""
        QueryC8824.CommandText = QueryC8824.CommandText & " " & "WHERE   DR.Valor_Aberto > 0 "
        QueryC8824.CommandText = QueryC8824.CommandText & " " & "AND  DR.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC8824.CommandText = QueryC8824.CommandText & " " & "AND  DR.Duplic_Cancelada = 0"
        QueryC8824.CommandText = QueryC8824.CommandText & " " & ""
        QueryC8824.CommandText = QueryC8824.CommandText & " " & "AND  DR.Cod_Pessoa_Estab_Juridico = E.Cod_Pessoa_Estab_Juridico "
        QueryC8824.CommandText = QueryC8824.CommandText & " " & ""
        QueryC8824.CommandText = QueryC8824.CommandText & " " & "AND  DR.Data_Vencimento + NVL(E.DIA_BLOQUEIO_DUPLICATA , 0)  < SYSDATE"
        QueryC8824.CommandText = QueryC8824.CommandText & " " & ""
        QueryC8824.Transaction = txn
        RsC8824 = QueryC8824.ExecuteReader()
        Dim iC8824 As Short
        ReDim C8824DataType(RsC8824.FieldCount)
        For iC8824 = 0 to RsC8824.FieldCount - 1
            Select Case RsC8824.GetDataTypeName(iC8824).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C8824DataType(iC8824 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C8824DataType(iC8824 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C8824DataType(iC8824 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC8824
        RsC8824_EOF = Not RsC8824.Read()

        GoTo Comp_C8887

    Comp_C8887:

        ' Eof1
        sCurrComponent = "Eof1"
        C8887DataType = 4
        C8887 = RsC8824_EOF
        GoTo Comp_C8889

    Comp_C8889:

        ' Eof1=T
        sCurrComponent = "Eof1=T"
        C8889 = (fn_ConvertValueCompiled(C8887, C8887DataType, AKB_DecimalPoint, False) =1)
        C8889DataType = AKBTypeConst.cAKBTypeLogical
        If C8889 Then
            GoTo Comp_C8708
        Else
            GoTo Comp_C572425
        End If

    Comp_C8891:

        ' MSG13
        sCurrComponent = "MSG13"
        Dim row_C8891 As DataRow = msg.NewRow()
        row_C8891(0) = "O pedido " & _ 
P2280 & "  não pode ser liberado porque o cliente tem duplicatas atrasadas." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C8891)
        C8891DataType = AKBTypeConst.cAKBTypeNumeric
        C8891 = 1

        GoTo Comp_C8584

    Comp_C8892:

        ' MSG14
        sCurrComponent = "MSG14"
        Dim row_C8892 As DataRow = msg.NewRow()
        row_C8892(0) = "O pedido " & _ 
P2280 & "  não pode ser liberado devido a uma restrição ao cliente" & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C8892)
        C8892DataType = AKBTypeConst.cAKBTypeNumeric
        C8892 = 1

        GoTo Comp_C8584

    Comp_C8893:

        ' RetVar
        sCurrComponent = "RetVar"
        C8893 = 1
        C8893DataType = 4
        GoTo Comp_C163127

    Comp_C8894:

        ' MSG15
        sCurrComponent = "MSG15"
        Dim row_C8894 As DataRow = msg.NewRow()
        row_C8894(0) = "Erro ao calcular o valor do pedido" & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C8894)
        C8894DataType = AKBTypeConst.cAKBTypeNumeric
        C8894 = 1

        GoTo Comp_C8584

    Comp_C8895:

        ' Begin
        sCurrComponent = "Begin"
        txn = con.BeginTransaction
        C8895 = True
        C8895DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C111652

    Comp_C8896:

        ' Commit
        sCurrComponent = "Commit"
        txn.Commit()
        C8896 = True
        C8896DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C8584

    Comp_C12958:

        ' SelDuplic
        sCurrComponent = "SelDuplic"
        QueryC12958 = con.CreateCommand()
        QueryC12958.CommandText = QueryC12958.CommandText & " " & "SELECT NVL(SUM(WF_DUPLIC_RECEBER.Valor_Aberto) ,0)"
        QueryC12958.CommandText = QueryC12958.CommandText & " " & ""
        QueryC12958.CommandText = QueryC12958.CommandText & " " & "FROM AKBUSER01.WF_DUPLIC_RECEBER , AKBUSER01.WF_NF_SAIDA "
        QueryC12958.CommandText = QueryC12958.CommandText & " " & ""
        QueryC12958.CommandText = QueryC12958.CommandText & " " & "WHERE WF_NF_SAIDA.Cod_Pessoa_Estab_Juridico = WF_DUPLIC_RECEBER.Cod_Pessoa_Estab_Juridico "
        QueryC12958.CommandText = QueryC12958.CommandText & " " & "AND  WF_NF_SAIDA.Nota_Fiscal = WF_DUPLIC_RECEBER.Nota_Fiscal "
        QueryC12958.CommandText = QueryC12958.CommandText & " " & "AND  WF_NF_SAIDA.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC12958.CommandText = QueryC12958.CommandText & " " & "AND  WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC12958.CommandText = QueryC12958.CommandText & " " & "AND  WF_DUPLIC_RECEBER.Valor_Aberto > 0 "
        QueryC12958.CommandText = QueryC12958.CommandText & " " & ""
        QueryC12958.Transaction = txn
        RsC12958 = QueryC12958.ExecuteReader()
        Dim iC12958 As Short
        ReDim C12958DataType(RsC12958.FieldCount)
        For iC12958 = 0 to RsC12958.FieldCount - 1
            Select Case RsC12958.GetDataTypeName(iC12958).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C12958DataType(iC12958 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C12958DataType(iC12958 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C12958DataType(iC12958 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC12958
        RsC12958_EOF = Not RsC12958.Read()

        GoTo Comp_C12959

    Comp_C12959:

        ' VDupl
        sCurrComponent = "VDupl"
        C12959DataType = 0
        C12959 = RsC12958(0)
        C12959DataType = C12958Datatype(1)
        If C12959DataType = AKBTypeConst.cAKBTypeString Then
          C12959 = IIF(IsDBNull(C12959), C12959, Strings.RTrim(C12959))
        End If 

        GoTo Comp_C8603

    Comp_C53541:

        ' VerLiberação
        sCurrComponent = "VerLiberação"
        QueryC53541 = con.CreateCommand()
        QueryC53541.CommandText = QueryC53541.CommandText & " " & "SELECT WF_TP_PRODUTO.Nao_lib_Cred_sem_Verificacao , WF_PEDIDO.Verificado, "
        QueryC53541.CommandText = QueryC53541.CommandText & " " & "                  WF_TP_PRODUTO.Tp_Tab_Preco "
        QueryC53541.CommandText = QueryC53541.CommandText & " " & "FROM AKBUSER01.WF_TP_PRODUTO , AKBUSER01.WF_PEDIDO "
        QueryC53541.CommandText = QueryC53541.CommandText & " " & "WHERE WF_TP_PRODUTO.Tp_Produto = WF_PEDIDO.Tp_Produto "
        QueryC53541.CommandText = QueryC53541.CommandText & " " & "AND  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC53541.CommandText = QueryC53541.CommandText & " " & "AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC53541.Transaction = txn
        RsC53541 = QueryC53541.ExecuteReader()
        Dim iC53541 As Short
        ReDim C53541DataType(RsC53541.FieldCount)
        For iC53541 = 0 to RsC53541.FieldCount - 1
            Select Case RsC53541.GetDataTypeName(iC53541).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C53541DataType(iC53541 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C53541DataType(iC53541 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C53541DataType(iC53541 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC53541
        RsC53541_EOF = Not RsC53541.Read()

        GoTo Comp_C97797

    Comp_C53542:

        ' NãoLibera
        sCurrComponent = "NãoLibera"
        C53542DataType = 0
        C53542 = RsC53541(0)
        C53542DataType = C53541Datatype(1)
        If C53542DataType = AKBTypeConst.cAKBTypeString Then
          C53542 = IIF(IsDBNull(C53542), C53542, Strings.RTrim(C53542))
        End If 

        GoTo Comp_C53543

    Comp_C53543:

        ' Verificado
        sCurrComponent = "Verificado"
        C53543DataType = 0
        C53543DataType = C53541Datatype(2)
        If C53543DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC53541(1)) Then
          C53543 = Strings.RTrim(RsC53541(1))
        Else 
          C53543 = RsC53541(1)
        End If 

        GoTo Comp_C317472

    Comp_C53544:

        ' ErroLiberar
        sCurrComponent = "ErroLiberar"
        C53544 = (fn_ConvertValueCompiled(C53542, C53542DataType, AKB_DecimalPoint, False) = 1 AND fn_ConvertValueCompiled(C53543, C53543DataType, AKB_DecimalPoint, False) = 0)
        C53544DataType = AKBTypeConst.cAKBTypeLogical
        If C53544 Then
            GoTo Comp_C8584
        Else
            GoTo Comp_C126476
        End If

    Comp_C96652:

        ' Transaction
        sCurrComponent = "Transaction"
        C96652 = (fn_ConvertValueCompiled(P20449, 4, AKB_DecimalPoint, False) = 1)
        C96652DataType = AKBTypeConst.cAKBTypeLogical
        If C96652 Then
            GoTo Comp_C111652
        Else
            GoTo Comp_C8895
        End If

    Comp_C96653:

        ' Transaction2
        sCurrComponent = "Transaction2"
        C96653 = (fn_ConvertValueCompiled(P20449, 4, AKB_DecimalPoint, False) = 1)
        C96653DataType = AKBTypeConst.cAKBTypeLogical
        If C96653 Then
            GoTo Comp_C8584
        Else
            GoTo Comp_C8896
        End If

    Comp_C97797:

        ' Tp_Fat
        sCurrComponent = "Tp_Fat"
        C97797DataType = 0
        C97797DataType = C53541Datatype(3)
        If C97797DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC53541(2)) Then
          C97797 = Strings.RTrim(RsC53541(2))
        Else 
          C97797 = RsC53541(2)
        End If 

        GoTo Comp_C53542

    Comp_C97798:

        ' Tp_Fat=Interno
        sCurrComponent = "Tp_Fat=Interno"
        C97798 = (fn_ConvertValueCompiled(C97797, C97797DataType, AKB_DecimalPoint, False) = "Interno")
        C97798DataType = AKBTypeConst.cAKBTypeLogical
        If C97798 Then
            GoTo Comp_C96652
        Else
            GoTo Comp_C8824
        End If

    Comp_C111652:

        ' Sel_Flag
        sCurrComponent = "Sel_Flag"
        QueryC111652 = con.CreateCommand()
        QueryC111652.CommandText = QueryC111652.CommandText & " " & "Select WF_PEDIDO.Flag_Pos_Ped "
        QueryC111652.CommandText = QueryC111652.CommandText & " " & "from AKBUSER01.WF_PEDIDO "
        QueryC111652.CommandText = QueryC111652.CommandText & " " & "where WF_PEDIDO.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC111652.CommandText = QueryC111652.CommandText & " " & "and WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC111652.Transaction = txn
        RsC111652 = QueryC111652.ExecuteReader()
        Dim iC111652 As Short
        ReDim C111652DataType(RsC111652.FieldCount)
        For iC111652 = 0 to RsC111652.FieldCount - 1
            Select Case RsC111652.GetDataTypeName(iC111652).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C111652DataType(iC111652 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C111652DataType(iC111652 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C111652DataType(iC111652 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC111652
        RsC111652_EOF = Not RsC111652.Read()

        GoTo Comp_C139381

    Comp_C111653:

        ' Sel_Flag=4
        sCurrComponent = "Sel_Flag=4"
        C111653 = (fn_ConvertValueCompiled(RsC111652(0), C111652DataType(1), AKB_DecimalPoint, False) = 4)
        C111653DataType = AKBTypeConst.cAKBTypeLogical
        If C111653 Then
            GoTo Comp_C111654
        Else
            GoTo Comp_C8606
        End If

    Comp_C111654:

        ' Libera2
        sCurrComponent = "Libera2"
        QueryC111654 = con.CreateCommand()
        QueryC111654.CommandText = QueryC111654.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO SET WF_PEDIDO.Data_Manut_Credito = " & _ 
ConvertValue(C8597, C8597DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC111654.CommandText = QueryC111654.CommandText & " " & "                                                 WF_PEDIDO.Liberado_Credito = 'Liberado' , "
        QueryC111654.CommandText = QueryC111654.CommandText & " " & "                                                 WF_PEDIDO.Usuario_Manut_Credito = " & _ 
ConvertValue(C8618, C8618DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC111654.CommandText = QueryC111654.CommandText & " " & "WHERE WF_PEDIDO.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC111654.Transaction = txn
        C111654 = QueryC111654.ExecuteNonQuery()
        C111654DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C113272

    Comp_C113272:

        ' Ins_Hist
        sCurrComponent = "Ins_Hist"
        QueryC113272 = con.CreateCommand()
        QueryC113272.CommandText = QueryC113272.CommandText & " " & "Insert into AKBUSER01.WF_PED_HISTORICO  ( Seq_Historico, Dt_Ocorrencia, Usuario, Flag_Pos_Ped, Tp_Produto, Pedido ) "
        QueryC113272.CommandText = QueryC113272.CommandText & " " & ""
        QueryC113272.CommandText = QueryC113272.CommandText & " " & "VALUES ( AKBUSER01.SEQ_WF_PED_HISTORICO.nextval, SysDate, User, 2, " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC113272.Transaction = txn
        C113272 = QueryC113272.ExecuteNonQuery()
        C113272DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C8607

    Comp_C125373:

        ' ValMaxLib
        sCurrComponent = "ValMaxLib"
        QueryC125373 = con.CreateCommand()
        QueryC125373.CommandText = QueryC125373.CommandText & " " & "SELECT WF_TP_PRODUTO.Valor_Max_Liberacao_Automatica FROM AKBUSER01.WF_TP_PRODUTO WHERE WF_TP_PRODUTO.Tp_Produto = " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC125373.Transaction = txn
        RsC125373 = QueryC125373.ExecuteReader()
        Dim iC125373 As Short
        ReDim C125373DataType(RsC125373.FieldCount)
        For iC125373 = 0 to RsC125373.FieldCount - 1
            Select Case RsC125373.GetDataTypeName(iC125373).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C125373DataType(iC125373 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C125373DataType(iC125373 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C125373DataType(iC125373 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC125373
        RsC125373_EOF = Not RsC125373.Read()

        GoTo Comp_C125377

    Comp_C125377:

        ' ValPed
        sCurrComponent = "ValPed"
        QueryC125377 = con.CreateCommand()
        QueryC125377.CommandText = QueryC125377.CommandText & " " & "SELECT WF_PEDIDO.Total_Itens FROM AKBUSER01.WF_PEDIDO WHERE WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC125377.Transaction = txn
        RsC125377 = QueryC125377.ExecuteReader()
        Dim iC125377 As Short
        ReDim C125377DataType(RsC125377.FieldCount)
        For iC125377 = 0 to RsC125377.FieldCount - 1
            Select Case RsC125377.GetDataTypeName(iC125377).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C125377DataType(iC125377 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C125377DataType(iC125377 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C125377DataType(iC125377 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC125377
        RsC125377_EOF = Not RsC125377.Read()

        GoTo Comp_C126475

    Comp_C125379:

        ' ValPed>MaxLib?
        sCurrComponent = "ValPed>MaxLib?"
        C125379 = (fn_ConvertValueCompiled(RsC125377(0), C125377DataType(1), AKB_DecimalPoint, False) > fn_ConvertValueCompiled(RsC125373(0), C125373DataType(1), AKB_DecimalPoint, False))
        C125379DataType = AKBTypeConst.cAKBTypeLogical
        If C125379 Then
            GoTo Comp_C53541
        Else
            GoTo Comp_C96652
        End If

    Comp_C126475:

        ' Coligada
        sCurrComponent = "Coligada"
        QueryC126475 = con.CreateCommand()
        QueryC126475.CommandText = QueryC126475.CommandText & " " & "SELECT COUNT(WF_ESTAB_JURID_COLIGADAS.Cod_Pessoa_Estab_Juridico) FROM AKBUSER01.WF_ESTAB_JURID_COLIGADAS WHERE WF_ESTAB_JURID_COLIGADAS.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P2305, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_ESTAB_JURID_COLIGADAS.Cod_Pessoa = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC126475.Transaction = txn
        RsC126475 = QueryC126475.ExecuteReader()
        Dim iC126475 As Short
        ReDim C126475DataType(RsC126475.FieldCount)
        For iC126475 = 0 to RsC126475.FieldCount - 1
            Select Case RsC126475.GetDataTypeName(iC126475).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C126475DataType(iC126475 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C126475DataType(iC126475 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C126475DataType(iC126475 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC126475
        RsC126475_EOF = Not RsC126475.Read()

        GoTo Comp_C331780

    Comp_C126476:

        ' Coligada>0?
        sCurrComponent = "Coligada>0?"
        C126476 = (fn_ConvertValueCompiled(RsC126475(0), C126475DataType(1), AKB_DecimalPoint, False) > 0)
        C126476DataType = AKBTypeConst.cAKBTypeLogical
        If C126476 Then
            GoTo Comp_C571768
        Else
            GoTo Comp_C317473
        End If

    Comp_C128605:

        ' Sel_Pedido
        sCurrComponent = "Sel_Pedido"
        QueryC128605 = con.CreateCommand()
        QueryC128605.CommandText = QueryC128605.CommandText & " " & "Select WF_PEDIDO.Pagto_Antecipado , WF_PEDIDO.Pg_Antec_Prod , WF_PEDIDO.Bloquear_Producao "
        QueryC128605.CommandText = QueryC128605.CommandText & " " & "from AKBUSER01.WF_PEDIDO "
        QueryC128605.CommandText = QueryC128605.CommandText & " " & "where WF_PEDIDO.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128605.CommandText = QueryC128605.CommandText & " " & "and  WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128605.Transaction = txn
        RsC128605 = QueryC128605.ExecuteReader()
        Dim iC128605 As Short
        ReDim C128605DataType(RsC128605.FieldCount)
        For iC128605 = 0 to RsC128605.FieldCount - 1
            Select Case RsC128605.GetDataTypeName(iC128605).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C128605DataType(iC128605 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C128605DataType(iC128605 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C128605DataType(iC128605 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC128605
        RsC128605_EOF = Not RsC128605.Read()

        GoTo Comp_C128608

    Comp_C128606:

        ' Bloq_Prod
        sCurrComponent = "Bloq_Prod"
        C128606DataType = 0
        C128606DataType = C128605Datatype(3)
        If C128606DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC128605(2)) Then
          C128606 = Strings.RTrim(RsC128605(2))
        Else 
          C128606 = RsC128605(2)
        End If 

        GoTo Comp_C128609

    Comp_C128607:

        ' Pg_Prod
        sCurrComponent = "Pg_Prod"
        C128607DataType = 0
        C128607DataType = C128605Datatype(2)
        If C128607DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC128605(1)) Then
          C128607 = Strings.RTrim(RsC128605(1))
        Else 
          C128607 = RsC128605(1)
        End If 

        GoTo Comp_C128606

    Comp_C128608:

        ' Pagto_Antecipado
        sCurrComponent = "Pagto_Antecipado"
        C128608DataType = 0
        C128608 = RsC128605(0)
        C128608DataType = C128605Datatype(1)
        If C128608DataType = AKBTypeConst.cAKBTypeString Then
          C128608 = IIF(IsDBNull(C128608), C128608, Strings.RTrim(C128608))
        End If 

        GoTo Comp_C128607

    Comp_C128609:

        ' Pagto_Antec=1
        sCurrComponent = "Pagto_Antec=1"
        C128609 = (fn_ConvertValueCompiled(C128608, C128608DataType, AKB_DecimalPoint, False) = 1)
        C128609DataType = AKBTypeConst.cAKBTypeLogical
        If C128609 Then
            GoTo Comp_C128610
        Else
            GoTo Comp_C8618
        End If

    Comp_C128610:

        ' Bloq_Prod=1
        sCurrComponent = "Bloq_Prod=1"
        C128610 = (fn_ConvertValueCompiled(C128606, C128606DataType, AKB_DecimalPoint, False) = 1)
        C128610DataType = AKBTypeConst.cAKBTypeLogical
        If C128610 Then
            GoTo Comp_C128611
        Else
            GoTo Comp_C8618
        End If

    Comp_C128611:

        ' Pg_Prod=1
        sCurrComponent = "Pg_Prod=1"
        C128611 = (fn_ConvertValueCompiled(C128607, C128607DataType, AKB_DecimalPoint, False) = 1)
        C128611DataType = AKBTypeConst.cAKBTypeLogical
        If C128611 Then
            GoTo Comp_C163125
        Else
            GoTo Comp_C128612
        End If

    Comp_C128612:

        ' MSG16
        sCurrComponent = "MSG16"
        Dim row_C128612 As DataRow = msg.NewRow()
        row_C128612(0) = "Não é possível liberar o crédito do pedido " & _ 
P2280 & ", pois ele é pedido " & Chr(13) & "" & Chr(10) & "com Pagamento Antecipado e ainda não foi pago o" & _ 
C163127 & " % para começar a " & Chr(13) & "" & Chr(10) & "Produzir." & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "Favor entrar em contato com o pessoal do comercial." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C128612)
        C128612DataType = AKBTypeConst.cAKBTypeNumeric
        C128612 = 1

        GoTo Comp_C8584

    Comp_C139381:

        ' Verif_CartaCred
        sCurrComponent = "Verif_CartaCred"
        'C139381 = clsRuleDynamicallyCompiled_R8842.R8842(con, txn, IIf(Not IsDBNull(P2280), P2280, System.DBNull.Value), IIf(Not IsDBNull(P2289), P2289, System.DBNull.Value))
        C139381CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C139381) Then
          iColumns = C139381.Columns.Count
        End If
        ReDim C139381DataType(iColumns)
        For iCol = 1 To iColumns
            'C139381DataType(iCol) = clsRuleDynamicallyCompiled_R8842.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C139383

    Comp_C139383:

        ' Valida=Ok?
        sCurrComponent = "Valida=Ok?"
        C139383 = (fn_ConvertValueCompiled(C139381.rows(C139381CurrentRow - 1)(0), C139381DataType(1), AKB_DecimalPoint, False) = 1)
        C139383DataType = AKBTypeConst.cAKBTypeLogical
        If C139383 Then
            GoTo Comp_C572052
        Else
            GoTo Comp_C96653
        End If

    Comp_C161492:

        ' Sel_CliBolq
        sCurrComponent = "Sel_CliBolq"
        QueryC161492 = con.CreateCommand()
        QueryC161492.CommandText = QueryC161492.CommandText & " " & "SELECT COUNT(WF_CLIENTES.Cod_Cliente) "
        QueryC161492.CommandText = QueryC161492.CommandText & " " & "FROM AKBUSER01.WF_CLIENTES WHERE WF_CLIENTES.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC161492.CommandText = QueryC161492.CommandText & " " & "AND  WF_CLIENTES.Bloqueado = 1 "
        QueryC161492.Transaction = txn
        RsC161492 = QueryC161492.ExecuteReader()
        Dim iC161492 As Short
        ReDim C161492DataType(RsC161492.FieldCount)
        For iC161492 = 0 to RsC161492.FieldCount - 1
            Select Case RsC161492.GetDataTypeName(iC161492).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C161492DataType(iC161492 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C161492DataType(iC161492 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C161492DataType(iC161492 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC161492
        RsC161492_EOF = Not RsC161492.Read()

        GoTo Comp_C161494

    Comp_C161494:

        ' Cli Bloq?
        sCurrComponent = "Cli Bloq?"
        C161494 = (fn_ConvertValueCompiled(RsC161492(0), C161492DataType(1), AKB_DecimalPoint, False) > 0)
        C161494DataType = AKBTypeConst.cAKBTypeLogical
        If C161494 Then
            GoTo Comp_C161495
        Else
            GoTo Comp_C8592
        End If

    Comp_C161495:

        ' Msg_ErrCliBloq
        sCurrComponent = "Msg_ErrCliBloq"
        Dim row_C161495 As DataRow = msg.NewRow()
        row_C161495(0) = "O pedido " & _ 
P2280 & " não pode ser liberado devido ao cliente se encontrar bloqueado." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C161495)
        C161495DataType = AKBTypeConst.cAKBTypeNumeric
        C161495 = 1

        GoTo Comp_C8584

    Comp_C163125:

        ' Sel_VrAberto
        sCurrComponent = "Sel_VrAberto"
        QueryC163125 = con.CreateCommand()
        QueryC163125.CommandText = QueryC163125.CommandText & " " & "SELECT NVL(WF_DUPLIC_RECEBER.Valor_Aberto,0)"
        QueryC163125.CommandText = QueryC163125.CommandText & " " & "FROM AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC163125.CommandText = QueryC163125.CommandText & " " & "WHERE WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC163125.CommandText = QueryC163125.CommandText & " " & "AND  WF_DUPLIC_RECEBER.Tp_Produto = " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC163125.CommandText = QueryC163125.CommandText & " " & "AND  WF_DUPLIC_RECEBER.Ped_Antecipado_Prod = 1 "
        QueryC163125.CommandText = QueryC163125.CommandText & " " & "AND  WF_DUPLIC_RECEBER.Duplic_Cancelada = 0 "
        QueryC163125.Transaction = txn
        RsC163125 = QueryC163125.ExecuteReader()
        Dim iC163125 As Short
        ReDim C163125DataType(RsC163125.FieldCount)
        For iC163125 = 0 to RsC163125.FieldCount - 1
            Select Case RsC163125.GetDataTypeName(iC163125).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C163125DataType(iC163125 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C163125DataType(iC163125 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C163125DataType(iC163125 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC163125
        RsC163125_EOF = Not RsC163125.Read()

        GoTo Comp_C163126

    Comp_C163126:

        ' VrAberto = 0?
        sCurrComponent = "VrAberto = 0?"
        C163126 = (fn_ConvertValueCompiled(RsC163125(0), C163125DataType(1), AKB_DecimalPoint, False) = 0)
        C163126DataType = AKBTypeConst.cAKBTypeLogical
        If C163126 Then
            GoTo Comp_C8618
        Else
            GoTo Comp_C163128
        End If

    Comp_C163127:

        ' v_Msg
        sCurrComponent = "v_Msg"
        C163127 = System.DBNull.Value
        C163127DataType = 3
        GoTo Comp_C236085

    Comp_C163128:

        ' Att_vMgs
        sCurrComponent = "Att_vMgs"
        C163128DataType = 4
        C163127 = fn_ConvertValueCompiled(" total do", 3, AKB_DecimalPoint)
        C163128 = True
        GoTo Comp_C128612

    Comp_C165638:

        ' Evento
        sCurrComponent = "Evento"
        'C165638 = clsRuleDynamicallyCompiled_R9674.R9674(con, txn, IIf(Not IsDBNull(P2289), P2289, System.DBNull.Value), IIf(Not IsDBNull(P2280), P2280, System.DBNull.Value))
        C165638CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C165638) Then
          iColumns = C165638.Columns.Count
        End If
        ReDim C165638DataType(iColumns)
        For iCol = 1 To iColumns
            'C165638DataType(iCol) = clsRuleDynamicallyCompiled_R9674.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C165639

    Comp_C165639:

        ' Evento OK?
        sCurrComponent = "Evento OK?"
        C165639 = (fn_ConvertValueCompiled(C165638.rows(C165638CurrentRow - 1)(0), C165638DataType(1), AKB_DecimalPoint, False) = 1)
        C165639DataType = AKBTypeConst.cAKBTypeLogical
        If C165639 Then
            GoTo Comp_C96653
        Else
            GoTo Comp_C165641
        End If

    Comp_C165641:

        ' Transaction = 3
        sCurrComponent = "Transaction = 3"
        C165641 = (fn_ConvertValueCompiled(P20449, 4, AKB_DecimalPoint, False) = 1)
        C165641DataType = AKBTypeConst.cAKBTypeLogical
        If C165641 Then
            GoTo Comp_C8584
        Else
            GoTo Comp_C165642
        End If

    Comp_C165642:

        ' Rollback
        sCurrComponent = "Rollback"
        txn.Rollback()
        C165642 = True
        C165642DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C8584

    Comp_C236085:

        ' Moeda
        sCurrComponent = "Moeda"
        QueryC236085 = con.CreateCommand()
        QueryC236085.CommandText = QueryC236085.CommandText & " " & "SELECT WF_ESTAB_JURIDICO.Codigo_Moeda "
        QueryC236085.CommandText = QueryC236085.CommandText & " " & "FROM AKBUSER01.WF_ESTAB_JURIDICO "
        QueryC236085.CommandText = QueryC236085.CommandText & " " & "WHERE WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P2305, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC236085.Transaction = txn
        RsC236085 = QueryC236085.ExecuteReader()
        Dim iC236085 As Short
        ReDim C236085DataType(RsC236085.FieldCount)
        For iC236085 = 0 to RsC236085.FieldCount - 1
            Select Case RsC236085.GetDataTypeName(iC236085).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C236085DataType(iC236085 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C236085DataType(iC236085 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C236085DataType(iC236085 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC236085
        RsC236085_EOF = Not RsC236085.Read()

        GoTo Comp_C128605

    Comp_C317472:

        ' SelGeraDuplic?
        sCurrComponent = "SelGeraDuplic?"
        QueryC317472 = con.CreateCommand()
        QueryC317472.CommandText = QueryC317472.CommandText & " " & "SELECT WF_CONDICAO_PGTO.Nao_Gera_Duplicata "
        QueryC317472.CommandText = QueryC317472.CommandText & " " & "FROM AKBUSER01.WF_CONDICAO_PGTO ,AKBUSER01.WF_PEDIDO "
        QueryC317472.CommandText = QueryC317472.CommandText & " " & "WHERE WF_CONDICAO_PGTO.Cod_Pagto  = WF_PEDIDO.Cod_Pagto "
        QueryC317472.CommandText = QueryC317472.CommandText & " " & "AND WF_PEDIDO.Pedido  = " & _ 
ConvertValue(P2280, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC317472.CommandText = QueryC317472.CommandText & " " & "AND WF_PEDIDO.Tp_Produto  = " & _ 
ConvertValue(P2289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC317472.Transaction = txn
        RsC317472 = QueryC317472.ExecuteReader()
        Dim iC317472 As Short
        ReDim C317472DataType(RsC317472.FieldCount)
        For iC317472 = 0 to RsC317472.FieldCount - 1
            Select Case RsC317472.GetDataTypeName(iC317472).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C317472DataType(iC317472 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C317472DataType(iC317472 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C317472DataType(iC317472 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC317472
        RsC317472_EOF = Not RsC317472.Read()

        GoTo Comp_C53544

    Comp_C317473:

        ' GeraDuplic?
        sCurrComponent = "GeraDuplic?"
        C317473 = (fn_ConvertValueCompiled(RsC317472(0), C317472DataType(1), AKB_DecimalPoint, False) = 1)
        C317473DataType = AKBTypeConst.cAKBTypeLogical
        If C317473 Then
            GoTo Comp_C96652
        Else
            GoTo Comp_C97798
        End If

    Comp_C320568:

        ' Desvio forçado
        sCurrComponent = "Desvio forçado"
        C320568 = (1 = 1)
        C320568DataType = AKBTypeConst.cAKBTypeLogical
        If C320568 Then
            GoTo Comp_C96653
        Else
            GoTo Comp_C165638
        End If

    Comp_C331780:

        ' DESVIO1
        sCurrComponent = "DESVIO1"
        C331780 = (fn_ConvertValueCompiled(RsC125377(0), C125377DataType(1), AKB_DecimalPoint, False)  = 0)
        C331780DataType = AKBTypeConst.cAKBTypeLogical
        If C331780 Then
            GoTo Comp_C53541
        Else
            GoTo Comp_C125379
        End If

    Comp_C332854:

        ' ReservaPedProntaEntrega
        sCurrComponent = "ReservaPedProntaEntrega"
        'C332854 = clsRuleDynamicallyCompiled_R15662.R15662(con, txn, IIf(Not IsDBNull(P2280), P2280, System.DBNull.Value), IIf(Not IsDBNull(P2289), P2289, System.DBNull.Value), fn_ConvertValueCompiled( "0", 4, AKB_DecimalPoint), System.DBNull.Value, System.DBNull.Value, System.DBNull.Value)
        C332854CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C332854) Then
          iColumns = C332854.Columns.Count
        End If
        ReDim C332854DataType(iColumns)
        For iCol = 1 To iColumns
            'C332854DataType(iCol) = clsRuleDynamicallyCompiled_R15662.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C502251

    Comp_C502251:

        ' CalcularPrazoPedidoEtiqueta
        sCurrComponent = "CalcularPrazoPedidoEtiqueta"
        'C502251 = clsRuleDynamicallyCompiled_R21229.R21229(con, txn, IIf(Not IsDBNull(P2289), P2289, System.DBNull.Value), IIf(Not IsDBNull(P2280), P2280, System.DBNull.Value))
        C502251CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C502251) Then
          iColumns = C502251.Columns.Count
        End If
        ReDim C502251DataType(iColumns)
        For iCol = 1 To iColumns
            'C502251DataType(iCol) = clsRuleDynamicallyCompiled_R21229.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C320568

    Comp_C571768:

        ' DESVIO2
        sCurrComponent = "DESVIO2"
        C571768 = (fn_ConvertValueCompiled(P2305, 1, AKB_DecimalPoint, False) = 5 AND fn_ConvertValueCompiled(P2286, 1, AKB_DecimalPoint, False) = 207461)
        C571768DataType = AKBTypeConst.cAKBTypeLogical
        If C571768 Then
            GoTo Comp_C317473
        Else
            GoTo Comp_C96652
        End If

    Comp_C572052:

        ' CNPJ_Outros
        sCurrComponent = "CNPJ_Outros"
        QueryC572052 = con.CreateCommand()
        QueryC572052.CommandText = QueryC572052.CommandText & " " & "SELECT DISTINCT PS.Cod_Pessoa"
        QueryC572052.CommandText = QueryC572052.CommandText & " " & "FROM WF_PESSOAS_SOCIOS PS"
        QueryC572052.CommandText = QueryC572052.CommandText & " " & " WHERE EXISTS ("
        QueryC572052.CommandText = QueryC572052.CommandText & " " & " SELECT * FROM  WF_PESSOAS_SOCIOS PS1"
        QueryC572052.CommandText = QueryC572052.CommandText & " " & "     WHERE PS1.CPF_CNPJ = PS.CPF_CNPJ"
        QueryC572052.CommandText = QueryC572052.CommandText & " " & "     AND PS1.Cod_Pessoa = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   )"
        QueryC572052.CommandText = QueryC572052.CommandText & " " & ""
        QueryC572052.Transaction = txn
        RsC572052 = QueryC572052.ExecuteReader()
        Dim iC572052 As Short
        ReDim C572052DataType(RsC572052.FieldCount)
        For iC572052 = 0 to RsC572052.FieldCount - 1
            Select Case RsC572052.GetDataTypeName(iC572052).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C572052DataType(iC572052 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C572052DataType(iC572052 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C572052DataType(iC572052 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC572052
        RsC572052_EOF = Not RsC572052.Read()

        GoTo Comp_C572053

    Comp_C572053:

        ' FimCNPJ_Outros
        sCurrComponent = "FimCNPJ_Outros"
        C572053DataType = 4
        C572053 = RsC572052_EOF
        GoTo Comp_C572054

    Comp_C572054:

        ' FimCNPJ_Outros?
        sCurrComponent = "FimCNPJ_Outros?"
        C572054 = (fn_ConvertValueCompiled(C572053, C572053DataType, AKB_DecimalPoint, False) = 1)
        C572054DataType = AKBTypeConst.cAKBTypeLogical
        If C572054 Then
            GoTo Comp_C111653
        Else
            GoTo Comp_C572056
        End If

    Comp_C572056:

        ' vClienteOutro
        sCurrComponent = "vClienteOutro"
        C572056DataType = 0
        C572056 = RsC572052(0)
        C572056DataType = C572052Datatype(1)
        If C572056DataType = AKBTypeConst.cAKBTypeString Then
          C572056 = IIF(IsDBNull(C572056), C572056, Strings.RTrim(C572056))
        End If 

        GoTo Comp_C572057

    Comp_C572057:

        ' ClienteOutroBloqueado
        sCurrComponent = "ClienteOutroBloqueado"
        QueryC572057 = con.CreateCommand()
        QueryC572057.CommandText = QueryC572057.CommandText & " " & "SELECT NVL(COUNT(WF_CLIENTES.Cod_Cliente),0) "
        QueryC572057.CommandText = QueryC572057.CommandText & " " & "FROM AKBUSER01.WF_CLIENTES WHERE WF_CLIENTES.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC572057.CommandText = QueryC572057.CommandText & " " & "AND  WF_CLIENTES.Bloqueado = 1 "
        QueryC572057.CommandText = QueryC572057.CommandText & " " & ""
        QueryC572057.Transaction = txn
        RsC572057 = QueryC572057.ExecuteReader()
        Dim iC572057 As Short
        ReDim C572057DataType(RsC572057.FieldCount)
        For iC572057 = 0 to RsC572057.FieldCount - 1
            Select Case RsC572057.GetDataTypeName(iC572057).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C572057DataType(iC572057 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C572057DataType(iC572057 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C572057DataType(iC572057 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC572057
        RsC572057_EOF = Not RsC572057.Read()

        GoTo Comp_C572058

    Comp_C572058:

        ' ClienteOutroBloqueado>0?
        sCurrComponent = "ClienteOutroBloqueado>0?"
        C572058 = (fn_ConvertValueCompiled(RsC572057(0), C572057DataType(1), AKB_DecimalPoint, False) > 0)
        C572058DataType = AKBTypeConst.cAKBTypeLogical
        If C572058 Then
            GoTo Comp_C572435
        Else
            GoTo Comp_C572060
        End If

    Comp_C572059:

        ' MSG17
        sCurrComponent = "MSG17"
        Dim row_C572059 As DataRow = msg.NewRow()
        row_C572059(0) = "Cliente " & _ 
C572056 & " que tem partcipação no cliente " & _ 
P2286 & " está bloqueado." & Chr(13) & "" & Chr(10) & "Pedido " & _ 
P2280 & " não liberado." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C572059)
        C572059DataType = AKBTypeConst.cAKBTypeNumeric
        C572059 = 1

        GoTo Comp_C8584

    Comp_C572060:

        ' DuplClienteOutro
        sCurrComponent = "DuplClienteOutro"
        QueryC572060 = con.CreateCommand()
        QueryC572060.CommandText = QueryC572060.CommandText & " " & "SELECT  NVL(COUNT(*),0)"
        QueryC572060.CommandText = QueryC572060.CommandText & " " & "FROM  AKBUSER01.WF_DUPLIC_RECEBER DR , AKBUSER01.WF_ESTAB_JURIDICO  E"
        QueryC572060.CommandText = QueryC572060.CommandText & " " & "WHERE   DR.Valor_Aberto > 0 "
        QueryC572060.CommandText = QueryC572060.CommandText & " " & "AND  DR.Cod_Cliente = " & _ 
ConvertValue(C572056, C572056DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC572060.CommandText = QueryC572060.CommandText & " " & "AND  DR.Duplic_Cancelada = 0"
        QueryC572060.CommandText = QueryC572060.CommandText & " " & "AND  DR.Cod_Pessoa_Estab_Juridico = E.Cod_Pessoa_Estab_Juridico "
        QueryC572060.CommandText = QueryC572060.CommandText & " " & "AND  DR.Data_Vencimento + NVL(E.DIA_BLOQUEIO_DUPLICATA , 0)  < SYSDATE"
        QueryC572060.CommandText = QueryC572060.CommandText & " " & ""
        QueryC572060.Transaction = txn
        RsC572060 = QueryC572060.ExecuteReader()
        Dim iC572060 As Short
        ReDim C572060DataType(RsC572060.FieldCount)
        For iC572060 = 0 to RsC572060.FieldCount - 1
            Select Case RsC572060.GetDataTypeName(iC572060).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C572060DataType(iC572060 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C572060DataType(iC572060 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C572060DataType(iC572060 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC572060
        RsC572060_EOF = Not RsC572060.Read()

        GoTo Comp_C572061

    Comp_C572061:

        ' DuplClienteOutro>0?
        sCurrComponent = "DuplClienteOutro>0?"
        C572061 = (fn_ConvertValueCompiled(RsC572060(0), C572060DataType(1), AKB_DecimalPoint, False) > 0)
        C572061DataType = AKBTypeConst.cAKBTypeLogical
        If C572061 Then
            GoTo Comp_C572438
        Else
            GoTo Comp_C572063
        End If

    Comp_C572062:

        ' MSG18
        sCurrComponent = "MSG18"
        Dim row_C572062 As DataRow = msg.NewRow()
        row_C572062(0) = "Cliente " & _ 
C572056 & " que tem partcipação no cliente " & _ 
P2286 & " tem duplicatas atrasadas." & Chr(13) & "" & Chr(10) & "Pedido " & _ 
P2280 & " não liberado." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C572062)
        C572062DataType = AKBTypeConst.cAKBTypeNumeric
        C572062 = 1

        GoTo Comp_C8584

    Comp_C572063:

        ' ClienteOutroRestricao
        sCurrComponent = "ClienteOutroRestricao"
        QueryC572063 = con.CreateCommand()
        QueryC572063.CommandText = QueryC572063.CommandText & " " & "SELECT NVL(WF_CLIENTE_RESTICAO.Tem_Restricao,0) "
        QueryC572063.CommandText = QueryC572063.CommandText & " " & "FROM AKBUSER01.WF_CLIENTE_RESTICAO "
        QueryC572063.CommandText = QueryC572063.CommandText & " " & "WHERE WF_CLIENTE_RESTICAO.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC572063.CommandText = QueryC572063.CommandText & " " & "AND WF_CLIENTE_RESTICAO.Sequencia  = (SELECT MAX(WF_CLIENTE_RESTICAO.Sequencia ) FROM AKBUSER01.WF_CLIENTE_RESTICAO "
        QueryC572063.CommandText = QueryC572063.CommandText & " " & "WHERE WF_CLIENTE_RESTICAO.Cod_Cliente = " & _ 
ConvertValue(P2286, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  )"
        QueryC572063.CommandText = QueryC572063.CommandText & " " & "AND  WF_CLIENTE_RESTICAO.Tem_Restricao = 1"
        QueryC572063.CommandText = QueryC572063.CommandText & " " & ""
        QueryC572063.Transaction = txn
        RsC572063 = QueryC572063.ExecuteReader()
        Dim iC572063 As Short
        ReDim C572063DataType(RsC572063.FieldCount)
        For iC572063 = 0 to RsC572063.FieldCount - 1
            Select Case RsC572063.GetDataTypeName(iC572063).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C572063DataType(iC572063 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C572063DataType(iC572063 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C572063DataType(iC572063 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC572063
        RsC572063_EOF = Not RsC572063.Read()

        GoTo Comp_C572064

    Comp_C572064:

        ' ClienteOutroRestricao=0?
        sCurrComponent = "ClienteOutroRestricao=0?"
        C572064 = (fn_ConvertValueCompiled(RsC572063(0), C572063DataType(1), AKB_DecimalPoint, False) = 0)
        C572064DataType = AKBTypeConst.cAKBTypeLogical
        If C572064 Then
            GoTo Comp_C111653
        Else
            GoTo Comp_C572441
        End If

    Comp_C572065:

        ' MSG19
        sCurrComponent = "MSG19"
        Dim row_C572065 As DataRow = msg.NewRow()
        row_C572065(0) = "Cliente " & _ 
C572056 & " que tem partcipação no cliente " & _ 
P2286 & " tem restrição no cadastro." & Chr(13) & "" & Chr(10) & "Pedido " & _ 
P2280 & " não liberado." & Chr(13) & Chr(10) & "Ok" & Chr(13) & Chr(10)
        msg.Rows.Add(row_C572065)
        C572065DataType = AKBTypeConst.cAKBTypeNumeric
        C572065 = 1

        GoTo Comp_C8584

    Comp_C572425:

        ' NaoLiberaCredAut=0?x
        sCurrComponent = "NaoLiberaCredAut=0?x"
        C572425 = (fn_ConvertValueCompiled(P90926, 4, AKB_DecimalPoint, False) = 0)
        C572425DataType = AKBTypeConst.cAKBTypeLogical
        If C572425 Then
            GoTo Comp_C8891
        Else
            GoTo Comp_C572427
        End If

    Comp_C572427:

        ' DESVIO3
        sCurrComponent = "DESVIO3"
        C572427 = (1 = 1)
        C572427DataType = AKBTypeConst.cAKBTypeLogical
        If C572427 Then
            GoTo Comp_C8584
        Else
            GoTo Comp_C8708
        End If

    Comp_C572428:

        ' NaoLiberaCredAut=0?b
        sCurrComponent = "NaoLiberaCredAut=0?b"
        C572428 = (fn_ConvertValueCompiled(P90926, 4, AKB_DecimalPoint, False) = 0)
        C572428DataType = AKBTypeConst.cAKBTypeLogical
        If C572428 Then
            GoTo Comp_C8892
        Else
            GoTo Comp_C572430
        End If

    Comp_C572430:

        ' DESVIO4
        sCurrComponent = "DESVIO4"
        C572430 = (1 = 1)
        C572430DataType = AKBTypeConst.cAKBTypeLogical
        If C572430 Then
            GoTo Comp_C8584
        Else
            GoTo Comp_C8687
        End If

    Comp_C572431:

        ' NaoLiberaCredAut=0?c
        sCurrComponent = "NaoLiberaCredAut=0?c"
        C572431 = (fn_ConvertValueCompiled(P90926, 4, AKB_DecimalPoint, False)  = 0)
        C572431DataType = AKBTypeConst.cAKBTypeLogical
        If C572431 Then
            GoTo Comp_C8586
        Else
            GoTo Comp_C96652
        End If

    Comp_C572435:

        ' NaoLiberaCredAut=0?a1
        sCurrComponent = "NaoLiberaCredAut=0?a1"
        C572435 = (fn_ConvertValueCompiled(P90926, 4, AKB_DecimalPoint, False) = 0)
        C572435DataType = AKBTypeConst.cAKBTypeLogical
        If C572435 Then
            GoTo Comp_C572059
        Else
            GoTo Comp_C572437
        End If

    Comp_C572437:

        ' DESVIO5
        sCurrComponent = "DESVIO5"
        C572437 = (1 = 0)
        C572437DataType = AKBTypeConst.cAKBTypeLogical
        If C572437 Then
            GoTo Comp_C572060
        Else
            GoTo Comp_C8584
        End If

    Comp_C572438:

        ' NaoLiberaCredAut=0?a2
        sCurrComponent = "NaoLiberaCredAut=0?a2"
        C572438 = (fn_ConvertValueCompiled(P90926, 4, AKB_DecimalPoint, False) = 0)
        C572438DataType = AKBTypeConst.cAKBTypeLogical
        If C572438 Then
            GoTo Comp_C572062
        Else
            GoTo Comp_C572440
        End If

    Comp_C572440:

        ' DESVIO6
        sCurrComponent = "DESVIO6"
        C572440 = (1 = 0)
        C572440DataType = AKBTypeConst.cAKBTypeLogical
        If C572440 Then
            GoTo Comp_C572063
        Else
            GoTo Comp_C8584
        End If

    Comp_C572441:

        ' NaoLiberaCredAut=0?a3
        sCurrComponent = "NaoLiberaCredAut=0?a3"
        C572441 = (fn_ConvertValueCompiled(P90926, 4, AKB_DecimalPoint, False) = 0)
        C572441DataType = AKBTypeConst.cAKBTypeLogical
        If C572441 Then
            GoTo Comp_C572065
        Else
            GoTo Comp_C572443
        End If

    Comp_C572443:

        ' DESVIO7
        sCurrComponent = "DESVIO7"
        C572443 = (1 = 0)
        C572443DataType = AKBTypeConst.cAKBTypeLogical
        If C572443 Then
            GoTo Comp_C111653
        Else
            GoTo Comp_C8584
        End If

    Exit_R1480:

        Exit Function

    Err_R1480:
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
