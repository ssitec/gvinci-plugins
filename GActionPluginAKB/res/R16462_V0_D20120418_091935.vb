Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R16462

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

    Public Shared Function R16462(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P58959 As Object, ByVal P58960 As Object, ByVal P58961 As Object, ByVal P58962 As Object, ByVal P58963 As Object, ByVal P58964 As Object, ByVal P58968 As Object, ByVal P58969 As Object, ByVal P58970 As Object, ByVal P58971 As Object, ByVal P58972 As Object, ByVal P58973 As Object, ByVal P58975 As Object, ByVal P58976 As Object) As DataTable
    ' 
    ' 16462 - Insere Pedido Itens x Impostos - Versão: 0
    ' 
        'On Error GoTo Err_R16462
        Dim sCurrComponent as String

        Dim C359385 As Boolean
        Dim C359385DataType As Short
        Dim C359388 As Double
        Dim C359388DataType As Short
        Dim QueryC359389 As New Object
        Dim C359389 As Integer
        Dim C359389DataType As Short
        Dim QueryC359390 As New Object
        Dim C359390 As Integer
        Dim C359390DataType As Short
        Dim QueryC359391 As New Object
        Dim C359391 As Integer
        Dim C359391DataType As Short
        Dim QueryC359392 As New Object
        Dim C359392 As Integer
        Dim C359392DataType As Short
        Dim QueryC359393 As New Object
        Dim RsC359393 As New Object
        Dim C359393DataType() As Short
        Dim RsC359393_EOF As Boolean
        Dim C359394 As Double
        Dim C359394DataType As Short
        Dim QueryC359395 As New Object
        Dim RsC359395 As New Object
        Dim C359395DataType() As Short
        Dim RsC359395_EOF As Boolean
        Dim C359396 As Double
        Dim C359396DataType As Short
        Dim QueryC359397 As New Object
        Dim RsC359397 As New Object
        Dim C359397DataType() As Short
        Dim RsC359397_EOF As Boolean
        Dim C359398 As Double
        Dim C359398DataType As Short
        Dim C359401 As Double
        Dim C359401DataType As Short
        Dim C359403 As Double
        Dim C359403DataType As Short
        Dim QueryC359404 As New Object
        Dim C359404 As Integer
        Dim C359404DataType As Short
        Dim C359405DataType() As Short
        Dim QueryC360571 As New Object
        Dim RsC360571 As New Object
        Dim C360571DataType() As Short
        Dim RsC360571_EOF As Boolean
        Dim C360573 As Boolean
        Dim C360573DataType As Short
        Dim QueryC360574 As New Object
        Dim C360574 As Integer
        Dim C360574DataType As Short
        Dim QueryC360585 As New Object
        Dim C360585 As Integer
        Dim C360585DataType As Short
        Dim C360594 As Boolean
        Dim C360594DataType As Short
        Dim C360597 As Boolean
        Dim C360597DataType As Short
        Dim QueryC360600 As New Object
        Dim C360600 As Integer
        Dim C360600DataType As Short
        Dim QueryC360603 As New Object
        Dim RsC360603 As New Object
        Dim C360603DataType() As Short
        Dim RsC360603_EOF As Boolean
        Dim QueryC360663 As New Object
        Dim RsC360663 As New Object
        Dim C360663DataType() As Short
        Dim RsC360663_EOF As Boolean
        Dim C360664 As Boolean
        Dim C360664DataType As Short
        Dim QueryC360665 As New Object
        Dim C360665 As Integer
        Dim C360665DataType As Short
        Dim QueryC360666 As New Object
        Dim RsC360666 As New Object
        Dim C360666DataType() As Short
        Dim RsC360666_EOF As Boolean
        Dim C360667 As Boolean
        Dim C360667DataType As Short
        Dim QueryC360668 As New Object
        Dim C360668 As Integer
        Dim C360668DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C360571

    Comp_C359385:

        ' TemRedução?
        sCurrComponent = "TemRedução?"
        C359385 = (fn_ConvertValueCompiled(P58963, 1, AKB_DecimalPoint, False)  = 0)
        C359385DataType = AKBTypeConst.cAKBTypeLogical
        If C359385 Then
            GoTo Comp_C359401
        Else
            GoTo Comp_C359388
        End If

    Comp_C359388:

        ' ValorImpostoICMS
        sCurrComponent = "ValorImpostoICMS"
        C359388 = (fn_ConvertValueCompiled(P58973, 1, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(P58968, 1, AKB_DecimalPoint, True) ) / 100
        C359388DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360594

    Comp_C359389:

        ' InsertIemImpostoICMS
        sCurrComponent = "InsertIemImpostoICMS"
        QueryC359389 = con.CreateCommand()
        QueryC359389.CommandText = QueryC359389.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS ( WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto , WF_PEDIDO_ITENS_IMPOSTOS.Pedido , WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item , WF_PEDIDO_ITENS_IMPOSTOS.Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Aliquota , WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto ) VALUES( " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58960, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58968, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , (Select Round(" & _ 
ConvertValue(C359388, C359388DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",2) From Dual) , " & _ 
ConvertValue(P58973, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC359389.Transaction = txn
        C359389 = QueryC359389.ExecuteNonQuery()
        C359389DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360603

    Comp_C359390:

        ' InsertIemImpostoPIS
        sCurrComponent = "InsertIemImpostoPIS"
        QueryC359390 = con.CreateCommand()
        QueryC359390.CommandText = QueryC359390.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS ( WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto , WF_PEDIDO_ITENS_IMPOSTOS.Pedido , WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item , WF_PEDIDO_ITENS_IMPOSTOS.Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Aliquota , WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto ) VALUES( " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ," & _ 
ConvertValue(P58959, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(RsC359393(0), C359393DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  ,(Select Round( " & _ 
ConvertValue(C359394, C359394DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",2) From Dual) , " & _ 
ConvertValue(P58973, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC359390.Transaction = txn
        C359390 = QueryC359390.ExecuteNonQuery()
        C359390DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360663

    Comp_C359391:

        ' InsertIemImpostoIPI
        sCurrComponent = "InsertIemImpostoIPI"
        QueryC359391 = con.CreateCommand()
        QueryC359391.CommandText = QueryC359391.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS ( WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto , WF_PEDIDO_ITENS_IMPOSTOS.Pedido , WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item , WF_PEDIDO_ITENS_IMPOSTOS.Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Aliquota , WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto ) VALUES( " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ," & _ 
ConvertValue(P58962, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , " & _ 
ConvertValue(RsC359395(0), C359395DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , (Select Round( " & _ 
ConvertValue(C359396, C359396DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,2) From Dual) , " & _ 
ConvertValue(P58973, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC359391.Transaction = txn
        C359391 = QueryC359391.ExecuteNonQuery()
        C359391DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360666

    Comp_C359392:

        ' InsertIemImpostoCOFINS
        sCurrComponent = "InsertIemImpostoCOFINS"
        QueryC359392 = con.CreateCommand()
        QueryC359392.CommandText = QueryC359392.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS ( WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto , WF_PEDIDO_ITENS_IMPOSTOS.Pedido , WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item , WF_PEDIDO_ITENS_IMPOSTOS.Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Aliquota , WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto ) VALUES( " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ," & _ 
ConvertValue(P58961, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(RsC359397(0), C359397DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , (Select Round( " & _ 
ConvertValue(C359398, C359398DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,2) From Dual) , " & _ 
ConvertValue(P58973, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC359392.Transaction = txn
        C359392 = QueryC359392.ExecuteNonQuery()
        C359392DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359405

    Comp_C359393:

        ' AliquotaPIS
        sCurrComponent = "AliquotaPIS"
        QueryC359393 = con.CreateCommand()
        QueryC359393.CommandText = QueryC359393.CommandText & " " & "SELECT NVL(WF_IMPOSTOS_VALIDADE.ALIQUOTA,0)"
        QueryC359393.CommandText = QueryC359393.CommandText & " " & ""
        QueryC359393.CommandText = QueryC359393.CommandText & " " & "FROM WF_IMPOSTOS_VALIDADE "
        QueryC359393.CommandText = QueryC359393.CommandText & " " & ""
        QueryC359393.CommandText = QueryC359393.CommandText & " " & "WHERE WF_IMPOSTOS_VALIDADE.IMPOSTO = " & _ 
ConvertValue(P58959, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359393.CommandText = QueryC359393.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.COD_PESSOA_ESTAB_JURIDICO = " & _ 
ConvertValue(P58975, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359393.CommandText = QueryC359393.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.VALIDADE_INICIAL <= " & _ 
ConvertValue(P58976, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359393.CommandText = QueryC359393.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.VALIDADE_FINAL >=  " & _ 
ConvertValue(P58976, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359393.Transaction = txn
        RsC359393 = QueryC359393.ExecuteReader()
        Dim iC359393 As Short
        ReDim C359393DataType(RsC359393.FieldCount)
        For iC359393 = 0 to RsC359393.FieldCount - 1
            Select Case RsC359393.GetDataTypeName(iC359393).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359393DataType(iC359393 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359393DataType(iC359393 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359393DataType(iC359393 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359393
        RsC359393_EOF = Not RsC359393.Read()

        GoTo Comp_C359394

    Comp_C359394:

        ' ValorImpostoPIS
        sCurrComponent = "ValorImpostoPIS"
        C359394 = (fn_ConvertValueCompiled(P58973, 1, AKB_DecimalPoint, True) *fn_ConvertValueCompiled(RsC359393(0), C359393DataType(1), AKB_DecimalPoint, True) )/100
        C359394DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360597

    Comp_C359395:

        ' AliquotaIPI
        sCurrComponent = "AliquotaIPI"
        QueryC359395 = con.CreateCommand()
        QueryC359395.CommandText = QueryC359395.CommandText & " " & "SELECT NVL(WF_IMPOSTOS_VALIDADE.ALIQUOTA,0)"
        QueryC359395.CommandText = QueryC359395.CommandText & " " & ""
        QueryC359395.CommandText = QueryC359395.CommandText & " " & "FROM WF_IMPOSTOS_VALIDADE "
        QueryC359395.CommandText = QueryC359395.CommandText & " " & ""
        QueryC359395.CommandText = QueryC359395.CommandText & " " & "WHERE WF_IMPOSTOS_VALIDADE.IMPOSTO = " & _ 
ConvertValue(P58962, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359395.CommandText = QueryC359395.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.COD_PESSOA_ESTAB_JURIDICO = " & _ 
ConvertValue(P58975, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359395.CommandText = QueryC359395.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.VALIDADE_INICIAL <= " & _ 
ConvertValue(P58976, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359395.CommandText = QueryC359395.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.VALIDADE_FINAL >=  " & _ 
ConvertValue(P58976, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359395.Transaction = txn
        RsC359395 = QueryC359395.ExecuteReader()
        Dim iC359395 As Short
        ReDim C359395DataType(RsC359395.FieldCount)
        For iC359395 = 0 to RsC359395.FieldCount - 1
            Select Case RsC359395.GetDataTypeName(iC359395).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359395DataType(iC359395 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359395DataType(iC359395 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359395DataType(iC359395 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359395
        RsC359395_EOF = Not RsC359395.Read()

        GoTo Comp_C359396

    Comp_C359396:

        ' ValorImpostoIPI
        sCurrComponent = "ValorImpostoIPI"
        C359396 = (fn_ConvertValueCompiled(P58973, 1, AKB_DecimalPoint, True) *fn_ConvertValueCompiled(RsC359395(0), C359395DataType(1), AKB_DecimalPoint, True) )/100
        C359396DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360664

    Comp_C359397:

        ' AliquotaCOFINS
        sCurrComponent = "AliquotaCOFINS"
        QueryC359397 = con.CreateCommand()
        QueryC359397.CommandText = QueryC359397.CommandText & " " & "SELECT NVL(WF_IMPOSTOS_VALIDADE.ALIQUOTA,0)"
        QueryC359397.CommandText = QueryC359397.CommandText & " " & ""
        QueryC359397.CommandText = QueryC359397.CommandText & " " & "FROM WF_IMPOSTOS_VALIDADE "
        QueryC359397.CommandText = QueryC359397.CommandText & " " & ""
        QueryC359397.CommandText = QueryC359397.CommandText & " " & "WHERE WF_IMPOSTOS_VALIDADE.IMPOSTO = " & _ 
ConvertValue(P58961, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359397.CommandText = QueryC359397.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.COD_PESSOA_ESTAB_JURIDICO = " & _ 
ConvertValue(P58975, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359397.CommandText = QueryC359397.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.VALIDADE_INICIAL <= " & _ 
ConvertValue(P58976, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359397.CommandText = QueryC359397.CommandText & " " & "AND WF_IMPOSTOS_VALIDADE.VALIDADE_FINAL >=  " & _ 
ConvertValue(P58976, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359397.Transaction = txn
        RsC359397 = QueryC359397.ExecuteReader()
        Dim iC359397 As Short
        ReDim C359397DataType(RsC359397.FieldCount)
        For iC359397 = 0 to RsC359397.FieldCount - 1
            Select Case RsC359397.GetDataTypeName(iC359397).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359397DataType(iC359397 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359397DataType(iC359397 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359397DataType(iC359397 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359397
        RsC359397_EOF = Not RsC359397.Read()

        GoTo Comp_C359398

    Comp_C359398:

        ' ValorImpostoCOFINS
        sCurrComponent = "ValorImpostoCOFINS"
        C359398 = (fn_ConvertValueCompiled(P58973, 1, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(RsC359397(0), C359397DataType(1), AKB_DecimalPoint, True)  )/100
        C359398DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360667

    Comp_C359401:

        ' BaseImposto
        sCurrComponent = "BaseImposto"
        C359401 = fn_ConvertValueCompiled(P58973, 1, AKB_DecimalPoint, True) - ((fn_ConvertValueCompiled(P58969, 1, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(P58973, 1, AKB_DecimalPoint, True))/100)
        C359401DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359403

    Comp_C359403:

        ' ValorImpostoICMS2
        sCurrComponent = "ValorImpostoICMS2"
        C359403 = (fn_ConvertValueCompiled(C359401, C359401DataType, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(P58968, 1, AKB_DecimalPoint, True) ) / 100
        C359403DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360573

    Comp_C359404:

        ' InsertImpostoICMS2
        sCurrComponent = "InsertImpostoICMS2"
        QueryC359404 = con.CreateCommand()
        QueryC359404.CommandText = QueryC359404.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS ( WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto , WF_PEDIDO_ITENS_IMPOSTOS.Pedido , WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item , WF_PEDIDO_ITENS_IMPOSTOS.Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Aliquota , WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto , WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto ) VALUES( " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58960, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P58968, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,( Select Round(" & _ 
ConvertValue(C359403, C359403DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",2) From Dual)  ," & _ 
ConvertValue(C359401, C359401DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  )"
        QueryC359404.Transaction = txn
        C359404 = QueryC359404.ExecuteNonQuery()
        C359404DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360603

    Comp_C359405:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C359405 As DataTable = New DataTable()
        tb_C359405.Columns.Add("TemRedução?" & "")
        Dim row_C359405 As DataRow = tb_C359405.NewRow()
        row_C359405(0) = C359385
        tb_C359405.Rows.Add(row_C359405)
        R16462 = tb_C359405
        ReDim C359405DataType(1)
        C359405DataType(1) = C359385DataType
        ReturnDataType = C359405DataType

        GoTo Exit_R16462

    Comp_C360571:

        ' SelVerificaExiste
        sCurrComponent = "SelVerificaExiste"
        QueryC360571 = con.CreateCommand()
        QueryC360571.CommandText = QueryC360571.CommandText & " " & "SELECT  COUNT(*)"
        QueryC360571.CommandText = QueryC360571.CommandText & " " & "FROM WF_PEDIDO_ITENS_IMPOSTOS"
        QueryC360571.CommandText = QueryC360571.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.PEDIDO   = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360571.CommandText = QueryC360571.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.TP_PRODUTO = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360571.CommandText = QueryC360571.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.SEQ_ITEM   = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360571.CommandText = QueryC360571.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.IMPOSTO = " & _ 
ConvertValue(P58960, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360571.Transaction = txn
        RsC360571 = QueryC360571.ExecuteReader()
        Dim iC360571 As Short
        ReDim C360571DataType(RsC360571.FieldCount)
        For iC360571 = 0 to RsC360571.FieldCount - 1
            Select Case RsC360571.GetDataTypeName(iC360571).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C360571DataType(iC360571 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C360571DataType(iC360571 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C360571DataType(iC360571 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC360571
        RsC360571_EOF = Not RsC360571.Read()

        GoTo Comp_C359385

    Comp_C360573:

        ' Atualizar2?
        sCurrComponent = "Atualizar2?"
        C360573 = (fn_ConvertValueCompiled(RsC360571(0), C360571DataType(1), AKB_DecimalPoint, False) =1)
        C360573DataType = AKBTypeConst.cAKBTypeLogical
        If C360573 Then
            GoTo Comp_C360574
        Else
            GoTo Comp_C359404
        End If

    Comp_C360574:

        ' AtualICMS2
        sCurrComponent = "AtualICMS2"
        QueryC360574 = con.CreateCommand()
        QueryC360574.CommandText = QueryC360574.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS "
        QueryC360574.CommandText = QueryC360574.CommandText & " " & "SET WF_PEDIDO_ITENS_IMPOSTOS.Aliquota = " & _ 
ConvertValue(P58968, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC360574.CommandText = QueryC360574.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto = ( Select Round(" & _ 
ConvertValue(C359403, C359403DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,2) From Dual) , "
        QueryC360574.CommandText = QueryC360574.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto = " & _ 
ConvertValue(C359401, C359401DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360574.CommandText = QueryC360574.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360574.CommandText = QueryC360574.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Pedido = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360574.CommandText = QueryC360574.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC360574.CommandText = QueryC360574.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Imposto = " & _ 
ConvertValue(P58960, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360574.Transaction = txn
        C360574 = QueryC360574.ExecuteNonQuery()
        C360574DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360603

    Comp_C360585:

        ' AtualICMS
        sCurrComponent = "AtualICMS"
        QueryC360585 = con.CreateCommand()
        QueryC360585.CommandText = QueryC360585.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS "
        QueryC360585.CommandText = QueryC360585.CommandText & " " & "SET WF_PEDIDO_ITENS_IMPOSTOS.Aliquota = " & _ 
ConvertValue(P58968, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC360585.CommandText = QueryC360585.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto = ( Select Round(" & _ 
ConvertValue(C359388, C359388DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",2) From Dual) , "
        QueryC360585.CommandText = QueryC360585.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto = " & _ 
ConvertValue(P58973, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360585.CommandText = QueryC360585.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360585.CommandText = QueryC360585.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Pedido = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360585.CommandText = QueryC360585.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC360585.CommandText = QueryC360585.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Imposto = " & _ 
ConvertValue(P58960, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360585.Transaction = txn
        C360585 = QueryC360585.ExecuteNonQuery()
        C360585DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360603

    Comp_C360594:

        ' Atualizar1?
        sCurrComponent = "Atualizar1?"
        C360594 = (fn_ConvertValueCompiled(RsC360571(0), C360571DataType(1), AKB_DecimalPoint, False) =1)
        C360594DataType = AKBTypeConst.cAKBTypeLogical
        If C360594 Then
            GoTo Comp_C360585
        Else
            GoTo Comp_C359389
        End If

    Comp_C360597:

        ' Atualizar3?
        sCurrComponent = "Atualizar3?"
        C360597 = (fn_ConvertValueCompiled(RsC360603(0), C360603DataType(1), AKB_DecimalPoint, False) =1)
        C360597DataType = AKBTypeConst.cAKBTypeLogical
        If C360597 Then
            GoTo Comp_C360600
        Else
            GoTo Comp_C359390
        End If

    Comp_C360600:

        ' AtualPIS
        sCurrComponent = "AtualPIS"
        QueryC360600 = con.CreateCommand()
        QueryC360600.CommandText = QueryC360600.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS "
        QueryC360600.CommandText = QueryC360600.CommandText & " " & "SET WF_PEDIDO_ITENS_IMPOSTOS.Aliquota = " & _ 
ConvertValue(RsC359393(0), C359393DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , "
        QueryC360600.CommandText = QueryC360600.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto = ( Select Round(" & _ 
ConvertValue(C359394, C359394DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,2) From Dual) , "
        QueryC360600.CommandText = QueryC360600.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto = " & _ 
ConvertValue(P58973, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360600.CommandText = QueryC360600.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360600.CommandText = QueryC360600.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Pedido = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360600.CommandText = QueryC360600.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC360600.CommandText = QueryC360600.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Imposto = " & _ 
ConvertValue(P58959, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360600.Transaction = txn
        C360600 = QueryC360600.ExecuteNonQuery()
        C360600DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360663

    Comp_C360603:

        ' SelVerificaExiste2
        sCurrComponent = "SelVerificaExiste2"
        QueryC360603 = con.CreateCommand()
        QueryC360603.CommandText = QueryC360603.CommandText & " " & "SELECT  COUNT(*)"
        QueryC360603.CommandText = QueryC360603.CommandText & " " & "FROM WF_PEDIDO_ITENS_IMPOSTOS"
        QueryC360603.CommandText = QueryC360603.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.PEDIDO   = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360603.CommandText = QueryC360603.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.TP_PRODUTO = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360603.CommandText = QueryC360603.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.SEQ_ITEM   = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360603.CommandText = QueryC360603.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.IMPOSTO = " & _ 
ConvertValue(P58959, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360603.Transaction = txn
        RsC360603 = QueryC360603.ExecuteReader()
        Dim iC360603 As Short
        ReDim C360603DataType(RsC360603.FieldCount)
        For iC360603 = 0 to RsC360603.FieldCount - 1
            Select Case RsC360603.GetDataTypeName(iC360603).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C360603DataType(iC360603 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C360603DataType(iC360603 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C360603DataType(iC360603 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC360603
        RsC360603_EOF = Not RsC360603.Read()

        GoTo Comp_C359393

    Comp_C360663:

        ' SelVerificaExiste3
        sCurrComponent = "SelVerificaExiste3"
        QueryC360663 = con.CreateCommand()
        QueryC360663.CommandText = QueryC360663.CommandText & " " & "SELECT  COUNT(*)"
        QueryC360663.CommandText = QueryC360663.CommandText & " " & "FROM WF_PEDIDO_ITENS_IMPOSTOS"
        QueryC360663.CommandText = QueryC360663.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.PEDIDO   = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360663.CommandText = QueryC360663.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.TP_PRODUTO = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360663.CommandText = QueryC360663.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.SEQ_ITEM   = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360663.CommandText = QueryC360663.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.IMPOSTO =" & _ 
ConvertValue(P58962, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360663.Transaction = txn
        RsC360663 = QueryC360663.ExecuteReader()
        Dim iC360663 As Short
        ReDim C360663DataType(RsC360663.FieldCount)
        For iC360663 = 0 to RsC360663.FieldCount - 1
            Select Case RsC360663.GetDataTypeName(iC360663).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C360663DataType(iC360663 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C360663DataType(iC360663 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C360663DataType(iC360663 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC360663
        RsC360663_EOF = Not RsC360663.Read()

        GoTo Comp_C359395

    Comp_C360664:

        ' Atualizar4?
        sCurrComponent = "Atualizar4?"
        C360664 = (fn_ConvertValueCompiled(RsC360663(0), C360663DataType(1), AKB_DecimalPoint, False) =1)
        C360664DataType = AKBTypeConst.cAKBTypeLogical
        If C360664 Then
            GoTo Comp_C360665
        Else
            GoTo Comp_C359391
        End If

    Comp_C360665:

        ' AtualIPI
        sCurrComponent = "AtualIPI"
        QueryC360665 = con.CreateCommand()
        QueryC360665.CommandText = QueryC360665.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS "
        QueryC360665.CommandText = QueryC360665.CommandText & " " & "SET WF_PEDIDO_ITENS_IMPOSTOS.Aliquota = " & _ 
ConvertValue(RsC359395(0), C359395DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC360665.CommandText = QueryC360665.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto = ( Select Round(" & _ 
ConvertValue(C359396, C359396DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  ,2) From Dual) , "
        QueryC360665.CommandText = QueryC360665.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto = " & _ 
ConvertValue(P58973, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360665.CommandText = QueryC360665.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360665.CommandText = QueryC360665.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Pedido = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360665.CommandText = QueryC360665.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC360665.CommandText = QueryC360665.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Imposto = " & _ 
ConvertValue(P58962, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360665.Transaction = txn
        C360665 = QueryC360665.ExecuteNonQuery()
        C360665DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C360666

    Comp_C360666:

        ' SelVerificaExiste4
        sCurrComponent = "SelVerificaExiste4"
        QueryC360666 = con.CreateCommand()
        QueryC360666.CommandText = QueryC360666.CommandText & " " & "SELECT  COUNT(*)"
        QueryC360666.CommandText = QueryC360666.CommandText & " " & "FROM WF_PEDIDO_ITENS_IMPOSTOS"
        QueryC360666.CommandText = QueryC360666.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.PEDIDO   = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360666.CommandText = QueryC360666.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.TP_PRODUTO = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360666.CommandText = QueryC360666.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.SEQ_ITEM   = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360666.CommandText = QueryC360666.CommandText & " " & "AND WF_PEDIDO_ITENS_IMPOSTOS.IMPOSTO = " & _ 
ConvertValue(P58961, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360666.Transaction = txn
        RsC360666 = QueryC360666.ExecuteReader()
        Dim iC360666 As Short
        ReDim C360666DataType(RsC360666.FieldCount)
        For iC360666 = 0 to RsC360666.FieldCount - 1
            Select Case RsC360666.GetDataTypeName(iC360666).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C360666DataType(iC360666 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C360666DataType(iC360666 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C360666DataType(iC360666 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC360666
        RsC360666_EOF = Not RsC360666.Read()

        GoTo Comp_C359397

    Comp_C360667:

        ' Atualiza5?
        sCurrComponent = "Atualiza5?"
        C360667 = (fn_ConvertValueCompiled(RsC360666(0), C360666DataType(1), AKB_DecimalPoint, False) = 1)
        C360667DataType = AKBTypeConst.cAKBTypeLogical
        If C360667 Then
            GoTo Comp_C360668
        Else
            GoTo Comp_C359392
        End If

    Comp_C360668:

        ' AtualCOFINS
        sCurrComponent = "AtualCOFINS"
        QueryC360668 = con.CreateCommand()
        QueryC360668.CommandText = QueryC360668.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS_IMPOSTOS "
        QueryC360668.CommandText = QueryC360668.CommandText & " " & "SET WF_PEDIDO_ITENS_IMPOSTOS.Aliquota = " & _ 
ConvertValue(RsC359397(0), C359397DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , "
        QueryC360668.CommandText = QueryC360668.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Valor_Imposto = ( Select Round(" & _ 
ConvertValue(C359398, C359398DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   ,2) From Dual) , "
        QueryC360668.CommandText = QueryC360668.CommandText & " " & "WF_PEDIDO_ITENS_IMPOSTOS.Base_Imposto = " & _ 
ConvertValue(P58973, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360668.CommandText = QueryC360668.CommandText & " " & "WHERE WF_PEDIDO_ITENS_IMPOSTOS.Tp_Produto = " & _ 
ConvertValue(P58970, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360668.CommandText = QueryC360668.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Pedido = " & _ 
ConvertValue(P58971, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360668.CommandText = QueryC360668.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Seq_Item = " & _ 
ConvertValue(P58972, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC360668.CommandText = QueryC360668.CommandText & " " & "AND  WF_PEDIDO_ITENS_IMPOSTOS.Imposto = " & _ 
ConvertValue(P58961, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC360668.Transaction = txn
        C360668 = QueryC360668.ExecuteNonQuery()
        C360668DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359405

    Exit_R16462:

        Exit Function

    Err_R16462:
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
