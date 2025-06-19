Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R7671

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

    Public Shared Function R7671(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P23377 As Double, ByVal P23378 As Object, ByVal P23379 As Double, ByVal P23380 As Double, ByVal P23435 As Double, ByVal P23438 As Double, ByVal P23439 As Double, ByVal P23440 As Double, ByVal P23457 As Object, ByVal P24281 As Object, ByVal P24282 As Object, ByVal P24283 As Object, ByVal P27973 As Object, ByVal P28516 As Object, ByVal P28521 As Object, ByVal P36494 As Object, ByVal P36936 As Object, ByVal P36937 As Object, ByVal P53300 As Object, ByVal P67697 As Object) As DataTable
    ' 
    ' 7671 - Pedido antecipado - Versão: 0
    ' 
        'On Error GoTo Err_R7671
        Dim sCurrComponent as String

        Dim C114809 As Object
        Dim C114809DataType As Short
        Dim QueryC114810 As New Object
        Dim RsC114810 As New Object
        Dim C114810DataType() As Short
        Dim RsC114810_EOF As Boolean
        Dim C114811 As Boolean
        Dim C114811DataType As Short
        Dim C114812DataType() As Short
        Dim C114813 As Boolean
        Dim C114813DataType As Short
        Dim C114814 As Short
        Dim C114814DataType As Short
        Dim C114814ReturnDataType() As Short

        Dim C114815 As Object
        Dim C114815DataType As Short
        Dim QueryC114816 As New Object
        Dim RsC114816 As New Object
        Dim C114816DataType() As Short
        Dim RsC114816_EOF As Boolean
        Dim C114817 As Double
        Dim C114817DataType As Short
        Dim QueryC114823 As New Object
        Dim C114823 As Integer
        Dim C114823DataType As Short
        Dim QueryC114825 As New Object
        Dim C114825 As Integer
        Dim C114825DataType As Short
        Dim QueryC115244 As New Object
        Dim C115244 As Integer
        Dim C115244DataType As Short
        Dim C115256 As Object
        Dim C115256DataType As Short
        Dim C115274 As Double
        Dim C115274DataType As Short
        Dim C115275 As Object
        Dim C115275DataType As Short
        Dim C115334 As Short
        Dim C115334DataType As Short
        Dim C115334ReturnDataType() As Short

        Dim C115413DataType() As Short
        Dim C120414 As Boolean
        Dim C120414DataType As Short
        Dim C120418 As Boolean
        Dim C120418DataType As Short
        Dim C120421 As Short
        Dim C120421DataType As Short
        Dim C120421ReturnDataType() As Short

        Dim C128446 As Boolean
        Dim C128446DataType As Short
        Dim QueryC128489 As New Object
        Dim RsC128489 As New Object
        Dim C128489DataType() As Short
        Dim RsC128489_EOF As Boolean
        Dim C128492 As Object
        Dim C128492DataType As Short
        Dim C128493 As Boolean
        Dim C128493DataType As Short
        Dim QueryC128494 As New Object
        Dim RsC128494 As New Object
        Dim C128494DataType() As Short
        Dim RsC128494_EOF As Boolean
        Dim C128496 As Boolean
        Dim C128496DataType As Short
        Dim QueryC128498 As New Object
        Dim C128498 As Integer
        Dim C128498DataType As Short
        Dim C128499 As Short
        Dim C128499DataType As Short
        Dim C128499ReturnDataType() As Short

        Dim C128761 As Object
        Dim C128761DataType As Short
        Dim C128763 As Double
        Dim C128763DataType As Short
        Dim C128767 As Object
        Dim C128767DataType As Short
        Dim C128903 As DataTable
        Dim C128903CurrentRow As Integer
        Dim C128903DataType() As Short
        Dim C128906 As Boolean
        Dim C128906DataType As Short
        Dim C128907DataType() As Short
        Dim C128908 As Object
        Dim C128908DataType As Short
        Dim C142341 As Boolean
        Dim C142341DataType As Short
        Dim C142342 As Double
        Dim C142342DataType As Short
        Dim QueryC142343 As New Object
        Dim C142343 As Integer
        Dim C142343DataType As Short
        Dim QueryC142509 As New Object
        Dim RsC142509 As New Object
        Dim C142509DataType() As Short
        Dim RsC142509_EOF As Boolean
        Dim C142513 As Object
        Dim C142513DataType As Short
        Dim C148784 As Boolean
        Dim C148784DataType As Short
        Dim QueryC148785 As New Object
        Dim RsC148785 As New Object
        Dim C148785DataType() As Short
        Dim RsC148785_EOF As Boolean
        Dim C148786 As Object
        Dim C148786DataType As Short
        Dim C148792 As Object
        Dim C148792DataType As Short
        Dim C148793 As Object
        Dim C148793DataType As Short
        Dim C148795 As Object
        Dim C148795DataType As Short
        Dim C148797 As Object
        Dim C148797DataType As Short
        Dim C148798 As Object
        Dim C148798DataType As Short
        Dim C149084 As Object
        Dim C149084DataType As Short
        Dim C149085 As Object
        Dim C149085DataType As Short
        Dim C149090 As Double
        Dim C149090DataType As Short
        Dim C149117 As Object
        Dim C149117DataType As Short
        Dim C149119 As Object
        Dim C149119DataType As Short
        Dim C203465 As Boolean
        Dim C203465DataType As Short
        Dim C260391 As Double
        Dim C260391DataType As Short
        Dim QueryC260392 As New Object
        Dim RsC260392 As New Object
        Dim C260392DataType() As Short
        Dim RsC260392_EOF As Boolean
        Dim QueryC260393 As New Object
        Dim C260393 As Integer
        Dim C260393DataType As Short
        Dim QueryC265412 As New Object
        Dim RsC265412 As New Object
        Dim C265412DataType() As Short
        Dim RsC265412_EOF As Boolean
        Dim C265413 As Boolean
        Dim C265413DataType As Short
        Dim C266212 As Boolean
        Dim C266212DataType As Short
        Dim C320161 As Object
        Dim C320161DataType As Short
        Dim C321283 As Boolean
        Dim C321283DataType As Short
        Dim C321284 As Object
        Dim C321284DataType As Short
        Dim QueryC321300 As New Object
        Dim C321300 As Integer
        Dim C321300DataType As Short
        Dim C321306 As Boolean
        Dim C321306DataType As Short
        Dim C321307 As Object
        Dim C321307DataType As Short
        Dim C321308 As Object
        Dim C321308DataType As Short
        Dim C321363 As Boolean
        Dim C321363DataType As Short
        Dim C321364 As Object
        Dim C321364DataType As Short
        Dim C321365 As Boolean
        Dim C321365DataType As Short
        Dim QueryC414510 As New Object
        Dim RsC414510 As New Object
        Dim C414510DataType() As Short
        Dim RsC414510_EOF As Boolean
        Dim C414511 As Object
        Dim C414511DataType As Short
        Dim C414512 As Boolean
        Dim C414512DataType As Short
        Dim QueryC414513 As New Object
        Dim C414513 As Integer
        Dim C414513DataType As Short
        P23378 = IIf(IsDBNull(P23378), 0, P23378)

        P24281 = IIf(IsDBNull(P24281), 0, P24281)

        P24282 = IIf(IsDBNull(P24282), 0, P24282)

        P24283 = IIf(IsDBNull(P24283), 0, P24283)

        P28516 = IIf(IsDBNull(P28516), 0, P28516)

        P36494 = IIf(IsDBNull(P36494), 0, P36494)

        P36936 = IIf(IsDBNull(P36936), 0, P36936)

        P36937 = IIf(IsDBNull(P36937), 0, P36937)

        ReDim ReturnDatatype(0)

        GoTo Comp_C114809

    Comp_C114809:

        ' vRetorno
        sCurrComponent = "vRetorno"
        C114809 = 1
        C114809DataType = 4
        GoTo Comp_C321307

    Comp_C114810:

        ' Sel_Cond_Pagto
        sCurrComponent = "Sel_Cond_Pagto"
        QueryC114810 = con.CreateCommand()
        QueryC114810.CommandText = QueryC114810.CommandText & " " & "Select  NVL( WF_PEDIDO.Pagto_Antecipado, 0 )"
        QueryC114810.CommandText = QueryC114810.CommandText & " " & "from AKBUSER01.WF_PEDIDO "
        QueryC114810.CommandText = QueryC114810.CommandText & " " & "where WF_PEDIDO.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC114810.CommandText = QueryC114810.CommandText & " " & "and WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC114810.CommandText = QueryC114810.CommandText & " " & ""
        QueryC114810.Transaction = txn
        RsC114810 = QueryC114810.ExecuteReader()
        Dim iC114810 As Short
        ReDim C114810DataType(RsC114810.FieldCount)
        For iC114810 = 0 to RsC114810.FieldCount - 1
            Select Case RsC114810.GetDataTypeName(iC114810).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C114810DataType(iC114810 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C114810DataType(iC114810 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C114810DataType(iC114810 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC114810
        RsC114810_EOF = Not RsC114810.Read()

        GoTo Comp_C114811

    Comp_C114811:

        ' Sel_Cond_Pagto=0
        sCurrComponent = "Sel_Cond_Pagto=0"
        C114811 = (fn_ConvertValueCompiled(RsC114810(0), C114810DataType(1), AKB_DecimalPoint, False) = 0)
        C114811DataType = AKBTypeConst.cAKBTypeLogical
        If C114811 Then
            GoTo Comp_C128489
        Else
            GoTo Comp_C142509
        End If

    Comp_C114812:

        ' Fim
        sCurrComponent = "Fim"
        Dim tb_C114812 As DataTable = New DataTable()
        tb_C114812.Columns.Add("vRetorno" & "")
        Dim row_C114812 As DataRow = tb_C114812.NewRow()
        row_C114812(0) = C114809
        tb_C114812.Rows.Add(row_C114812)
        R7671 = tb_C114812
        ReDim C114812DataType(1)
        C114812DataType(1) = C114809DataType
        ReturnDataType = C114812DataType

        GoTo Exit_R7671

    Comp_C114813:

        ' Antec_Prod=0
        sCurrComponent = "Antec_Prod=0"
        C114813 = (1=2)
        C114813DataType = AKBTypeConst.cAKBTypeLogical
        If C114813 Then
            GoTo Comp_C114814
        Else
            GoTo Comp_C142341
        End If

    Comp_C114814:

        ' Msg_%Antec_Prod
        sCurrComponent = "Msg_%Antec_Prod"
        C114814 = 1
        C114814DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C114815

    Comp_C114815:

        ' AttvRet
        sCurrComponent = "AttvRet"
        C114815DataType = 4
        C114809 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C114815 = True
        GoTo Comp_C114812

    Comp_C114816:

        ' Sel_Total_Itens
        sCurrComponent = "Sel_Total_Itens"
        QueryC114816 = con.CreateCommand()
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "Select  "
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "Round(SUM( WF_PEDIDO_ITENS.Qtde_Pedido * (WF_PEDIDO_ITENS.Preco_Unit * WF_PEDIDO.COTACAO_MOEDA )"
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "                           ) -"
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "                          NVL ( WF_PEDIDO.BONIFICACAO * WF_PEDIDO.COTACAO_MOEDA,0 )"
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "             , 3),"
        QueryC114816.CommandText = QueryC114816.CommandText & " " & ""
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "Round(SUM( WF_PEDIDO_ITENS.Qtde_Pedido * WF_PEDIDO_ITENS.Preco_Unit ) -"
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "                                                                                                    NVL ( WF_PEDIDO.BONIFICACAO, 0 ) , 3)"
        QueryC114816.CommandText = QueryC114816.CommandText & " " & ""
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "FROM AKBUSER01.WF_PEDIDO_ITENS , AKBUSER01.WF_PEDIDO_SEQ, AKBUSER01.WF_PEDIDO "
        QueryC114816.CommandText = QueryC114816.CommandText & " " & ""
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "AND  WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "AND WF_PEDIDO_ITENS.Flag_Pos_Ped NOT IN (8 , 10 , 12 , 299)"
        QueryC114816.CommandText = QueryC114816.CommandText & " " & ""
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido =WF_PEDIDO_SEQ.Pedido "
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "AND  WF_PEDIDO_ITENS.Tp_Produto = WF_PEDIDO_SEQ.Tp_Produto "
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "AND WF_PEDIDO_SEQ.Seq = 1"
        QueryC114816.CommandText = QueryC114816.CommandText & " " & ""
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "AND WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "AND WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC114816.CommandText = QueryC114816.CommandText & " " & ""
        QueryC114816.CommandText = QueryC114816.CommandText & " " & "GROUP BY WF_PEDIDO.BONIFICACAO, WF_PEDIDO.COTACAO_MOEDA"
        QueryC114816.Transaction = txn
        RsC114816 = QueryC114816.ExecuteReader()
        Dim iC114816 As Short
        ReDim C114816DataType(RsC114816.FieldCount)
        For iC114816 = 0 to RsC114816.FieldCount - 1
            Select Case RsC114816.GetDataTypeName(iC114816).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C114816DataType(iC114816 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C114816DataType(iC114816 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C114816DataType(iC114816 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC114816
        RsC114816_EOF = Not RsC114816.Read()

        GoTo Comp_C148784

    Comp_C114817:

        ' Calc_Vlr_Antec_Prod
        sCurrComponent = "Calc_Vlr_Antec_Prod"
        C114817 = ( fn_ConvertValueCompiled(C149084, C149084DataType, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(P23378, 1, AKB_DecimalPoint, True) ) / 100
        C114817DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C148798

    Comp_C114823:

        ' Up_Vlr_Antecipado
        sCurrComponent = "Up_Vlr_Antecipado"
        QueryC114823 = con.CreateCommand()
        QueryC114823.CommandText = QueryC114823.CommandText & " " & "Update  AKBUSER01.WF_PEDIDO  "
        QueryC114823.CommandText = QueryC114823.CommandText & " " & "set WF_PEDIDO.Valor_Antec_Prod = Round( " & _ 
ConvertValue(C114817, C114817DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 2 ) ,"
        QueryC114823.CommandText = QueryC114823.CommandText & " " & "WF_PEDIDO.Porc_Pg_Antec_Prod = Round(" & _ 
ConvertValue(C149117, C149117DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 3 )"
        QueryC114823.CommandText = QueryC114823.CommandText & " " & "where WF_PEDIDO.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC114823.CommandText = QueryC114823.CommandText & " " & "and WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC114823.CommandText = QueryC114823.CommandText & " " & ""
        QueryC114823.Transaction = txn
        C114823 = QueryC114823.ExecuteNonQuery()
        C114823DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C115274

    Comp_C114825:

        ' Up_Vlr_Antec=0
        sCurrComponent = "Up_Vlr_Antec=0"
        QueryC114825 = con.CreateCommand()
        QueryC114825.CommandText = QueryC114825.CommandText & " " & "Update  AKBUSER01.WF_PEDIDO  set WF_PEDIDO.Valor_Antec_Prod = 0,"
        QueryC114825.CommandText = QueryC114825.CommandText & " " & "                                            WF_PEDIDO.Valor_Antec_Fat = 0,"
        QueryC114825.CommandText = QueryC114825.CommandText & " " & "                                            WF_PEDIDO.Porc_Pg_Antec_Prod = 0,"
        QueryC114825.CommandText = QueryC114825.CommandText & " " & "                                            WF_PEDIDO.Porc_Pg_Antec_Fat = 0,"
        QueryC114825.CommandText = QueryC114825.CommandText & " " & "                                            WF_PEDIDO.Bloquear_Faturamento = 0,"
        QueryC114825.CommandText = QueryC114825.CommandText & " " & "                                            WF_PEDIDO.Bloquear_Producao = 0"
        QueryC114825.CommandText = QueryC114825.CommandText & " " & "where WF_PEDIDO.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC114825.CommandText = QueryC114825.CommandText & " " & "and WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC114825.CommandText = QueryC114825.CommandText & " " & ""
        QueryC114825.Transaction = txn
        C114825 = QueryC114825.ExecuteNonQuery()
        C114825DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C114812

    Comp_C115244:

        ' Ins_Duplicata_Prod
        sCurrComponent = "Ins_Duplicata_Prod"
        QueryC115244 = con.CreateCommand()
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "INSERT INTO AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "( WF_DUPLIC_RECEBER.Ident_Duplicata , WF_DUPLIC_RECEBER.Cod_Pessoa_Estab_Juridico , "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Tp_Produto , WF_DUPLIC_RECEBER.Pedido ,"
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Data_Vencimento , WF_DUPLIC_RECEBER.Valor_Duplicata , "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Porc_Comissao , WF_DUPLIC_RECEBER.Valor_Comissao , "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Tipo_Carteira , WF_DUPLIC_RECEBER.Data_Geracao , "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Usuario_Geracao , WF_DUPLIC_RECEBER.Duplic_Cancelada , "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Valor_Aberto, WF_DUPLIC_RECEBER.Num_Conta_Bancaria, "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Cod_Cliente, WF_DUPLIC_RECEBER.Cod_Repres, "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Data_Emissao, WF_DUPLIC_RECEBER.Seq_Duplicata ,    "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   COMISSAO_PAGA_FAT ,  WF_DUPLIC_RECEBER.VALOR_PREMIO ,"
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Ped_Antecipado_Prod,  "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Vlr_Duplic_Moeda_Cliente,"
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "   WF_DUPLIC_RECEBER.Codigo_Moeda ) "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & ""
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "VALUES (  " & _ 
ConvertValue(C115256, C115256DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P23435, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C148786, C148786DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "                      Round( " & _ 
ConvertValue(C148797, C148797DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,2 ), " & _ 
ConvertValue(P23438, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , Round( " & _ 
ConvertValue(C115274, C115274DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", 2 ), 1,"
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "                      SysDate, " & _ 
ConvertValue(C115275, C115275DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 0, Round(" & _ 
ConvertValue(C148797, C148797DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,2 ), " & _ 
ConvertValue(C142513, C142513DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(P23439, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "                      " & _ 
ConvertValue(P23440, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P23457, 2, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C128761, C128761DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 0, 0, 1, "
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "                      Round(" & _ 
ConvertValue(C149090, C149090DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", 2),"
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "                      Decode (" & _ 
ConvertValue(P36937, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",0, " & _ 
ConvertValue(P36936, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(P36494, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC115244.CommandText = QueryC115244.CommandText & " " & "                    )"
        QueryC115244.Transaction = txn
        C115244 = QueryC115244.ExecuteNonQuery()
        C115244DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C128763

    Comp_C115256:

        ' Id_Duplicata_Prod
        sCurrComponent = "Id_Duplicata_Prod"
        C115256DataType = 1
        C115256 = ObjTable_NewIdentity (con, "WF_DUPLIC_RECEBER")
        GoTo Comp_C115244

    Comp_C115274:

        ' Calc_comissao
        sCurrComponent = "Calc_comissao"
        C115274 = (fn_ConvertValueCompiled(C148797, C148797DataType, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(P23438, 1, AKB_DecimalPoint, True) ) / 100
        C115274DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C414510

    Comp_C115275:

        ' Usuario
        sCurrComponent = "Usuario"
        C115275DataType = 3
        C115275 = Mid(con.ConnectionString, InStr(con.ConnectionString, "User Id=")+8, Instr(InStr(con.ConnectionString, "User Id="), con.ConnectionString, ";")-InStr(con.ConnectionString, "User Id=")-8)
        GoTo Comp_C114810

    Comp_C115334:

        ' Mgs_Duplicata
        sCurrComponent = "Mgs_Duplicata"
        C115334 = 1
        C115334DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C115413

    Comp_C115413:

        ' Fim2
        sCurrComponent = "Fim2"
        Dim tb_C115413 As DataTable = New DataTable()
        tb_C115413.Columns.Add("vRetorno" & "")
        Dim row_C115413 As DataRow = tb_C115413.NewRow()
        row_C115413(0) = C114809
        tb_C115413.Rows.Add(row_C115413)
        R7671 = tb_C115413
        ReDim C115413DataType(1)
        C115413DataType(1) = C114809DataType
        ReturnDataType = C115413DataType

        GoTo Exit_R7671

    Comp_C120414:

        ' Bloq_Prod
        sCurrComponent = "Bloq_Prod"
        C120414 = (fn_ConvertValueCompiled(P24281, 4, AKB_DecimalPoint, False) = 1)
        C120414DataType = AKBTypeConst.cAKBTypeLogical
        If C120414 Then
            GoTo Comp_C114813
        Else
            GoTo Comp_C120418
        End If

    Comp_C120418:

        ' Bloq_Fat
        sCurrComponent = "Bloq_Fat"
        C120418 = (fn_ConvertValueCompiled(P24282, 4, AKB_DecimalPoint, False) = 1)
        C120418DataType = AKBTypeConst.cAKBTypeLogical
        If C120418 Then
            GoTo Comp_C266212
        Else
            GoTo Comp_C128446
        End If

    Comp_C120421:

        ' Msg_Sem_bloq
        sCurrComponent = "Msg_Sem_bloq"
        C120421 = 1
        C120421DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C114815

    Comp_C128446:

        ' Sem_Evento
        sCurrComponent = "Sem_Evento"
        C128446 = (( fn_ConvertValueCompiled(P24282, 4, AKB_DecimalPoint, False) = 0 ) and ( fn_ConvertValueCompiled(P24281, 4, AKB_DecimalPoint, False) = 0 ))
        C128446DataType = AKBTypeConst.cAKBTypeLogical
        If C128446 Then
            GoTo Comp_C120421
        Else
            GoTo Comp_C115334
        End If

    Comp_C128489:

        ' Sel_Duplicatas
        sCurrComponent = "Sel_Duplicatas"
        QueryC128489 = con.CreateCommand()
        QueryC128489.CommandText = QueryC128489.CommandText & " " & "Select WF_DUPLIC_RECEBER.Ident_Duplicata "
        QueryC128489.CommandText = QueryC128489.CommandText & " " & "from AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC128489.CommandText = QueryC128489.CommandText & " " & "where WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128489.CommandText = QueryC128489.CommandText & " " & "and WF_DUPLIC_RECEBER.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128489.CommandText = QueryC128489.CommandText & " " & "and WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC128489.Transaction = txn
        RsC128489 = QueryC128489.ExecuteReader()
        Dim iC128489 As Short
        ReDim C128489DataType(RsC128489.FieldCount)
        For iC128489 = 0 to RsC128489.FieldCount - 1
            Select Case RsC128489.GetDataTypeName(iC128489).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C128489DataType(iC128489 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C128489DataType(iC128489 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C128489DataType(iC128489 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC128489
        RsC128489_EOF = Not RsC128489.Read()

        GoTo Comp_C320161

    Comp_C128492:

        ' ML_Duplic
        sCurrComponent = "ML_Duplic"
        C128492DataType = 5
        C128492 = ""
        Do Until RsC128489_EOF
            If RTrim(C128492) <> "" Then
                C128492 = C128492 & ", "
            End If
            C128492 = C128492 & ConvertValue(RsC128489(0), 0, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss")
            RsC128489_EOF = Not RsC128489.Read()
        Loop

        GoTo Comp_C128493

    Comp_C128493:

        ' ML_Duplic=0
        sCurrComponent = "ML_Duplic=0"
        C128493 = (fn_ConvertValueCompiled(C320161, C320161DataType, AKB_DecimalPoint, False) = 1)
        C128493DataType = AKBTypeConst.cAKBTypeLogical
        If C128493 Then
            GoTo Comp_C114825
        Else
            GoTo Comp_C128494
        End If

    Comp_C128494:

        ' Sel_Baixas
        sCurrComponent = "Sel_Baixas"
        QueryC128494 = con.CreateCommand()
        QueryC128494.CommandText = QueryC128494.CommandText & " " & "Select  count(*)"
        QueryC128494.CommandText = QueryC128494.CommandText & " " & "from AKBUSER01.WF_DUPL_RECEBER_BAIXAS "
        QueryC128494.CommandText = QueryC128494.CommandText & " " & "where WF_DUPL_RECEBER_BAIXAS.Ident_Duplicata in ( " & _ 
ConvertValue(C128492, C128492DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC128494.CommandText = QueryC128494.CommandText & " " & "and WF_DUPL_RECEBER_BAIXAS.Estornada = 0"
        QueryC128494.Transaction = txn
        RsC128494 = QueryC128494.ExecuteReader()
        Dim iC128494 As Short
        ReDim C128494DataType(RsC128494.FieldCount)
        For iC128494 = 0 to RsC128494.FieldCount - 1
            Select Case RsC128494.GetDataTypeName(iC128494).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C128494DataType(iC128494 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C128494DataType(iC128494 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C128494DataType(iC128494 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC128494
        RsC128494_EOF = Not RsC128494.Read()

        GoTo Comp_C128496

    Comp_C128496:

        ' Sel_Baixas=0
        sCurrComponent = "Sel_Baixas=0"
        C128496 = (fn_ConvertValueCompiled(RsC128494(0), C128494DataType(1), AKB_DecimalPoint, False) = 0)
        C128496DataType = AKBTypeConst.cAKBTypeLogical
        If C128496 Then
            GoTo Comp_C128498
        Else
            GoTo Comp_C265412
        End If

    Comp_C128498:

        ' Up_Duplicatas
        sCurrComponent = "Up_Duplicatas"
        QueryC128498 = con.CreateCommand()
        QueryC128498.CommandText = QueryC128498.CommandText & " " & "Update  AKBUSER01.WF_DUPLIC_RECEBER  set  WF_DUPLIC_RECEBER.Duplic_Cancelada = 1"
        QueryC128498.CommandText = QueryC128498.CommandText & " " & "where WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128498.CommandText = QueryC128498.CommandText & " " & "and WF_DUPLIC_RECEBER.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC128498.CommandText = QueryC128498.CommandText & " " & "and WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC128498.Transaction = txn
        C128498 = QueryC128498.ExecuteNonQuery()
        C128498DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C114825

    Comp_C128499:

        ' MSG1
        sCurrComponent = "MSG1"
        C128499 = 1
        C128499DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C114815

    Comp_C128761:

        ' Seq_Duplic
        sCurrComponent = "Seq_Duplic"
        C128761 = 1
        C128761DataType = 1
        GoTo Comp_C128903

    Comp_C128763:

        ' Seq+1
        sCurrComponent = "Seq+1"
        C128763 = fn_ConvertValueCompiled(C128761, C128761DataType, AKB_DecimalPoint, True) + 1
        C128763DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C128767

    Comp_C128767:

        ' Att_Seq_Dup
        sCurrComponent = "Att_Seq_Dup"
        C128767DataType = 4
        C128761 = fn_ConvertValueCompiled(C128763 , 1, AKB_DecimalPoint)
        C128767 = True
        GoTo Comp_C120418

    Comp_C128903:

        ' Verifica_Ped
        sCurrComponent = "Verifica_Ped"
        C128903 = clsRuleDynamicallyCompiled_R7681.R7681(con, txn, IIf(Not IsDBNull(P23380), P23380, System.DBNull.Value), IIf(Not IsDBNull(P23379), P23379, System.DBNull.Value), IIf(Not IsDBNull(P28521), P28521, System.DBNull.Value), System.DBNull.Value)
        C128903CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C128903) Then
          iColumns = C128903.Columns.Count
        End If
        ReDim C128903DataType(iColumns)
        For iCol = 1 To iColumns
            C128903DataType(iCol) = clsRuleDynamicallyCompiled_R7681.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C128906

    Comp_C128906:

        ' Verifica_Ped=1
        sCurrComponent = "Verifica_Ped=1"
        C128906 = (fn_ConvertValueCompiled(C128903.rows(C128903CurrentRow - 1)(0), C128903DataType(1), AKB_DecimalPoint, False) = 1)
        C128906DataType = AKBTypeConst.cAKBTypeLogical
        If C128906 Then
            GoTo Comp_C115275
        Else
            GoTo Comp_C203465
        End If

    Comp_C128907:

        ' Ret
        sCurrComponent = "Ret"
        Dim tb_C128907 As DataTable = New DataTable()
        tb_C128907.Columns.Add("vRetorno" & "")
        Dim row_C128907 As DataRow = tb_C128907.NewRow()
        row_C128907(0) = C114809
        tb_C128907.Rows.Add(row_C128907)
        R7671 = tb_C128907
        ReDim C128907DataType(1)
        C128907DataType(1) = C114809DataType
        ReturnDataType = C128907DataType

        GoTo Exit_R7671

    Comp_C128908:

        ' Att_vRet=0
        sCurrComponent = "Att_vRet=0"
        C128908DataType = 4
        C114809 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C128908 = True
        GoTo Comp_C128907

    Comp_C142341:

        ' %Pagto_Prod = 0?
        sCurrComponent = "%Pagto_Prod = 0?"
        C142341 = ((fn_ConvertValueCompiled(P23378, 1, AKB_DecimalPoint, False) = 0))
        C142341DataType = AKBTypeConst.cAKBTypeLogical
        If C142341 Then
            GoTo Comp_C142342
        Else
            GoTo Comp_C321365
        End If

    Comp_C142342:

        ' Calc_%_Ant_Prod
        sCurrComponent = "Calc_%_Ant_Prod"
        C142342 = ((fn_ConvertValueCompiled(P27973, 1, AKB_DecimalPoint, True) * 100) / fn_ConvertValueCompiled(C149084, C149084DataType, AKB_DecimalPoint, True) )
        C142342DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C321283

    Comp_C142343:

        ' Up_%_Ant_Prod
        sCurrComponent = "Up_%_Ant_Prod"
        QueryC142343 = con.CreateCommand()
        QueryC142343.CommandText = QueryC142343.CommandText & " " & "Update  AKBUSER01.WF_PEDIDO  "
        QueryC142343.CommandText = QueryC142343.CommandText & " " & "set WF_PEDIDO.Porc_Pg_Antec_Prod = Round(" & _ 
ConvertValue(C149117, C149117DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 3 )"
        QueryC142343.CommandText = QueryC142343.CommandText & " " & "where WF_PEDIDO.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC142343.CommandText = QueryC142343.CommandText & " " & "and WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC142343.CommandText = QueryC142343.CommandText & " " & ""
        QueryC142343.Transaction = txn
        C142343 = QueryC142343.ExecuteNonQuery()
        C142343DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C115274

    Comp_C142509:

        ' Sel_NumConta
        sCurrComponent = "Sel_NumConta"
        QueryC142509 = con.CreateCommand()
        QueryC142509.CommandText = QueryC142509.CommandText & " " & "SELECT WF_CONTAS_BANCARIAS.Num_Conta_Bancaria FROM AKBUSER01.WF_CONTAS_BANCARIAS WHERE WF_CONTAS_BANCARIAS.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P23435, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_CONTAS_BANCARIAS.GERACAO_DUPLIC_PEDIDO_ANTEC = 1 "
        QueryC142509.Transaction = txn
        RsC142509 = QueryC142509.ExecuteReader()
        Dim iC142509 As Short
        ReDim C142509DataType(RsC142509.FieldCount)
        For iC142509 = 0 to RsC142509.FieldCount - 1
            Select Case RsC142509.GetDataTypeName(iC142509).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C142509DataType(iC142509 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C142509DataType(iC142509 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C142509DataType(iC142509 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC142509
        RsC142509_EOF = Not RsC142509.Read()

        GoTo Comp_C142513

    Comp_C142513:

        ' NumConta
        sCurrComponent = "NumConta"
        C142513DataType = 0
        C142513 = RsC142509(0)
        C142513DataType = C142509Datatype(1)
        If C142513DataType = AKBTypeConst.cAKBTypeString Then
          C142513 = IIF(IsDBNull(C142513), C142513, Strings.RTrim(C142513))
        End If 

        GoTo Comp_C114816

    Comp_C148784:

        ' Dias_Entregar=0?
        sCurrComponent = "Dias_Entregar=0?"
        C148784 = (fn_ConvertValueCompiled(P28516, 1, AKB_DecimalPoint, False) = 0)
        C148784DataType = AKBTypeConst.cAKBTypeLogical
        If C148784 Then
            GoTo Comp_C148785
        Else
            GoTo Comp_C148795
        End If

    Comp_C148785:

        ' Sel_Dias
        sCurrComponent = "Sel_Dias"
        QueryC148785 = con.CreateCommand()
        QueryC148785.CommandText = QueryC148785.CommandText & " " & "SELECT NVL(WF_ESTAB_JURIDICO_PARAM.Dias_Venc_Duplic_Ped_Antec,0)"
        QueryC148785.CommandText = QueryC148785.CommandText & " " & "FROM AKBUSER01.WF_ESTAB_JURIDICO_PARAM "
        QueryC148785.CommandText = QueryC148785.CommandText & " " & "WHERE WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P23435, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC148785.Transaction = txn
        RsC148785 = QueryC148785.ExecuteReader()
        Dim iC148785 As Short
        ReDim C148785DataType(RsC148785.FieldCount)
        For iC148785 = 0 to RsC148785.FieldCount - 1
            Select Case RsC148785.GetDataTypeName(iC148785).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C148785DataType(iC148785 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C148785DataType(iC148785 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C148785DataType(iC148785 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC148785
        RsC148785_EOF = Not RsC148785.Read()

        GoTo Comp_C148793

    Comp_C148786:

        ' Dt_Vencim
        sCurrComponent = "Dt_Vencim"
        C148786DataType = 2
        C148786 = DateAdd("d", C148792, C148795)
        GoTo Comp_C149084

    Comp_C148792:

        ' vDias
        sCurrComponent = "vDias"
        C148792 = P28516 
        C148792DataType = 1
        GoTo Comp_C128761

    Comp_C148793:

        ' AttQtdeDias
        sCurrComponent = "AttQtdeDias"
        C148793DataType = 4
        C148792 = fn_ConvertValueCompiled(RsC148785(0) , 1, AKB_DecimalPoint)
        C148793 = True
        GoTo Comp_C148795

    Comp_C148795:

        ' SysDate
        sCurrComponent = "SysDate"
        C148795DataType = 2
        C148795 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C148786

    Comp_C148797:

        ' vVr_Antec_Prod
        sCurrComponent = "vVr_Antec_Prod"
        C148797 = P27973 
        C148797DataType = 1
        GoTo Comp_C148792

    Comp_C148798:

        ' Att_Vl_Antec_Prod
        sCurrComponent = "Att_Vl_Antec_Prod"
        C148798DataType = 4
        C148797 = fn_ConvertValueCompiled(C114817 , 1, AKB_DecimalPoint)
        C148798 = True
        GoTo Comp_C321363

    Comp_C149084:

        ' Total_Itens
        sCurrComponent = "Total_Itens"
        C149084DataType = 0
        C149084 = RsC114816(0)
        C149084DataType = C114816Datatype(1)
        If C149084DataType = AKBTypeConst.cAKBTypeString Then
          C149084 = IIF(IsDBNull(C149084), C149084, Strings.RTrim(C149084))
        End If 

        GoTo Comp_C149085

    Comp_C149085:

        ' Total_Itens_Moeda
        sCurrComponent = "Total_Itens_Moeda"
        C149085DataType = 0
        C149085DataType = C114816Datatype(2)
        If C149085DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC114816(1)) Then
          C149085 = Strings.RTrim(RsC114816(1))
        Else 
          C149085 = RsC114816(1)
        End If 

        GoTo Comp_C120414

    Comp_C149090:

        ' Calc_Vlr_Antec_Prod_Moeda
        sCurrComponent = "Calc_Vlr_Antec_Prod_Moeda"
        C149090 = (fn_ConvertValueCompiled(C149085, C149085DataType, AKB_DecimalPoint, True) * fn_ConvertValueCompiled(C149117, C149117DataType, AKB_DecimalPoint, True)) / 100
        C149090DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C115256

    Comp_C149117:

        ' v%Antec_Prod
        sCurrComponent = "v%Antec_Prod"
        C149117 = P23378 
        C149117DataType = 1
        GoTo Comp_C148797

    Comp_C149119:

        ' Att%AntecProd
        sCurrComponent = "Att%AntecProd"
        C149119DataType = 4
        C149117 = fn_ConvertValueCompiled(C142342 , 1, AKB_DecimalPoint)
        C149119 = True
        GoTo Comp_C142343

    Comp_C203465:

        ' Verifica_Ped = -1?
        sCurrComponent = "Verifica_Ped = -1?"
        C203465 = (fn_ConvertValueCompiled(C128903.rows(C128903CurrentRow - 1)(0), C128903DataType(1), AKB_DecimalPoint, False) = -1)
        C203465DataType = AKBTypeConst.cAKBTypeLogical
        If C203465 Then
            GoTo Comp_C128907
        Else
            GoTo Comp_C128908
        End If

    Comp_C260391:

        ' %Fat
        sCurrComponent = "%Fat"
        C260391 = 100 - fn_ConvertValueCompiled(RsC260392(0), C260392DataType(1), AKB_DecimalPoint, True)
        C260391DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C260393

    Comp_C260392:

        ' %Prod
        sCurrComponent = "%Prod"
        QueryC260392 = con.CreateCommand()
        QueryC260392.CommandText = QueryC260392.CommandText & " " & "SELECT NVL( WF_PEDIDO.Porc_Pg_Antec_Prod, 0)"
        QueryC260392.CommandText = QueryC260392.CommandText & " " & ""
        QueryC260392.CommandText = QueryC260392.CommandText & " " & "FROM  AKBUSER01.WF_PEDIDO  "
        QueryC260392.CommandText = QueryC260392.CommandText & " " & ""
        QueryC260392.CommandText = QueryC260392.CommandText & " " & "WHERE WF_PEDIDO.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC260392.CommandText = QueryC260392.CommandText & " " & "AND WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC260392.Transaction = txn
        RsC260392 = QueryC260392.ExecuteReader()
        Dim iC260392 As Short
        ReDim C260392DataType(RsC260392.FieldCount)
        For iC260392 = 0 to RsC260392.FieldCount - 1
            Select Case RsC260392.GetDataTypeName(iC260392).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C260392DataType(iC260392 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C260392DataType(iC260392 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C260392DataType(iC260392 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC260392
        RsC260392_EOF = Not RsC260392.Read()

        GoTo Comp_C260391

    Comp_C260393:

        ' up_%_Fat
        sCurrComponent = "up_%_Fat"
        QueryC260393 = con.CreateCommand()
        QueryC260393.CommandText = QueryC260393.CommandText & " " & "UPDATE  AKBUSER01.WF_PEDIDO  "
        QueryC260393.CommandText = QueryC260393.CommandText & " " & "SET WF_PEDIDO.Porc_Pg_Antec_Fat = Round(" & _ 
ConvertValue(C260391, C260391DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 3 )"
        QueryC260393.CommandText = QueryC260393.CommandText & " " & "WHERE WF_PEDIDO.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC260393.CommandText = QueryC260393.CommandText & " " & "AND WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC260393.Transaction = txn
        C260393 = QueryC260393.ExecuteNonQuery()
        C260393DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C115413

    Comp_C265412:

        ' selDuplicProd
        sCurrComponent = "selDuplicProd"
        QueryC265412 = con.CreateCommand()
        QueryC265412.CommandText = QueryC265412.CommandText & " " & "Select count( *)"
        QueryC265412.CommandText = QueryC265412.CommandText & " " & "from AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC265412.CommandText = QueryC265412.CommandText & " " & "where WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC265412.CommandText = QueryC265412.CommandText & " " & "and WF_DUPLIC_RECEBER.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC265412.CommandText = QueryC265412.CommandText & " " & "and WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC265412.CommandText = QueryC265412.CommandText & " " & "and WF_DUPLIC_RECEBER.Ped_Antecipado_Prod = 1"
        QueryC265412.Transaction = txn
        RsC265412 = QueryC265412.ExecuteReader()
        Dim iC265412 As Short
        ReDim C265412DataType(RsC265412.FieldCount)
        For iC265412 = 0 to RsC265412.FieldCount - 1
            Select Case RsC265412.GetDataTypeName(iC265412).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C265412DataType(iC265412 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C265412DataType(iC265412 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C265412DataType(iC265412 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC265412
        RsC265412_EOF = Not RsC265412.Read()

        GoTo Comp_C265413

    Comp_C265413:

        ' selDuplicProd > 0
        sCurrComponent = "selDuplicProd > 0"
        C265413 = (fn_ConvertValueCompiled(RsC265412(0), C265412DataType(1), AKB_DecimalPoint, False) > 0)
        C265413DataType = AKBTypeConst.cAKBTypeLogical
        If C265413 Then
            GoTo Comp_C128499
        Else
            GoTo Comp_C114812
        End If

    Comp_C266212:

        ' Bloq_Prod = 1
        sCurrComponent = "Bloq_Prod = 1"
        C266212 = (fn_ConvertValueCompiled(P24281, 4, AKB_DecimalPoint, False) = 1)
        C266212DataType = AKBTypeConst.cAKBTypeLogical
        If C266212 Then
            GoTo Comp_C260392
        Else
            GoTo Comp_C321306
        End If

    Comp_C320161:

        ' EOF Duplicatas
        sCurrComponent = "EOF Duplicatas"
        C320161DataType = 4
        C320161 = RsC128489_EOF
        GoTo Comp_C128492

    Comp_C321283:

        ' Calc_%_Ant_Prod > 100
        sCurrComponent = "Calc_%_Ant_Prod > 100"
        C321283 = (fn_ConvertValueCompiled(C142342, C142342DataType, AKB_DecimalPoint, False) > 100)
        C321283DataType = AKBTypeConst.cAKBTypeLogical
        If C321283 Then
            GoTo Comp_C321284
        Else
            GoTo Comp_C149119
        End If

    Comp_C321284:

        ' Atr -> v%Antec_Prod
        sCurrComponent = "Atr -> v%Antec_Prod"
        C321284DataType = 4
        C149117 = fn_ConvertValueCompiled(100, 1, AKB_DecimalPoint)
        C321284 = True
        GoTo Comp_C142343

    Comp_C321300:

        ' up_%_Fat1
        sCurrComponent = "up_%_Fat1"
        QueryC321300 = con.CreateCommand()
        QueryC321300.CommandText = QueryC321300.CommandText & " " & "UPDATE  AKBUSER01.WF_PEDIDO  "
        QueryC321300.CommandText = QueryC321300.CommandText & " " & "SET WF_PEDIDO.Porc_Pg_Antec_Fat = " & _ 
ConvertValue(C321307, C321307DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC321300.CommandText = QueryC321300.CommandText & " " & "WHERE WF_PEDIDO.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC321300.CommandText = QueryC321300.CommandText & " " & "AND WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(P23379, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC321300.Transaction = txn
        C321300 = QueryC321300.ExecuteNonQuery()
        C321300DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C115413

    Comp_C321306:

        ' %Pagto_Fat = 0
        sCurrComponent = "%Pagto_Fat = 0"
        C321306 = (fn_ConvertValueCompiled(P24283, 1, AKB_DecimalPoint, False) = 0)
        C321306DataType = AKBTypeConst.cAKBTypeLogical
        If C321306 Then
            GoTo Comp_C321308
        Else
            GoTo Comp_C321300
        End If

    Comp_C321307:

        ' v%Antec_Fat
        sCurrComponent = "v%Antec_Fat"
        C321307 = P24283 & " "
        C321307DataType = 1
        GoTo Comp_C149117

    Comp_C321308:

        ' Atr -> v%Antec_Fat
        sCurrComponent = "Atr -> v%Antec_Fat"
        C321308DataType = 4
        C321307 = fn_ConvertValueCompiled(100, 1, AKB_DecimalPoint)
        C321308 = True
        GoTo Comp_C321300

    Comp_C321363:

        ' %Pagto_Prod > 100
        sCurrComponent = "%Pagto_Prod > 100"
        C321363 = (fn_ConvertValueCompiled(P23378, 1, AKB_DecimalPoint, False) > 100)
        C321363DataType = AKBTypeConst.cAKBTypeLogical
        If C321363 Then
            GoTo Comp_C321364
        Else
            GoTo Comp_C114823
        End If

    Comp_C321364:

        ' Atr -> v%Antec_Prod1
        sCurrComponent = "Atr -> v%Antec_Prod1"
        C321364DataType = 4
        C149117 = fn_ConvertValueCompiled(100, 1, AKB_DecimalPoint)
        C321364 = True
        GoTo Comp_C114823

    Comp_C321365:

        ' calcula valor ?
        sCurrComponent = "calcula valor ?"
        C321365 = (( fn_ConvertValueCompiled(P27973, 1, AKB_DecimalPoint, False) > 0) AND (fn_ConvertValueCompiled(P23378, 1, AKB_DecimalPoint, False) = 100) AND   (fn_ConvertValueCompiled(P27973, 1, AKB_DecimalPoint, False) = fn_ConvertValueCompiled(P53300, 1, AKB_DecimalPoint, False) ))
        C321365DataType = AKBTypeConst.cAKBTypeLogical
        If C321365 Then
            GoTo Comp_C115274
        Else
            GoTo Comp_C114817
        End If

    Comp_C414510:

        ' SelDuplicataPrePedido
        sCurrComponent = "SelDuplicataPrePedido"
        QueryC414510 = con.CreateCommand()
        QueryC414510.CommandText = QueryC414510.CommandText & " " & "Select WF_DUPLIC_RECEBER.Ident_Duplicata "
        QueryC414510.CommandText = QueryC414510.CommandText & " " & "from AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC414510.CommandText = QueryC414510.CommandText & " " & "where WF_DUPLIC_RECEBER.Id_PrePedido = " & _ 
ConvertValue(P67697, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC414510.CommandText = QueryC414510.CommandText & " " & "and WF_DUPLIC_RECEBER.Pedido is Null"
        QueryC414510.CommandText = QueryC414510.CommandText & " " & "and WF_DUPLIC_RECEBER.Duplic_Cancelada = 0"
        QueryC414510.Transaction = txn
        RsC414510 = QueryC414510.ExecuteReader()
        Dim iC414510 As Short
        ReDim C414510DataType(RsC414510.FieldCount)
        For iC414510 = 0 to RsC414510.FieldCount - 1
            Select Case RsC414510.GetDataTypeName(iC414510).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C414510DataType(iC414510 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C414510DataType(iC414510 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C414510DataType(iC414510 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC414510
        RsC414510_EOF = Not RsC414510.Read()

        GoTo Comp_C414511

    Comp_C414511:

        ' SemDuplicPrePedido
        sCurrComponent = "SemDuplicPrePedido"
        C414511DataType = 4
        C414511 = RsC414510_EOF
        GoTo Comp_C414512

    Comp_C414512:

        ' Pré Pedido sem Duplicata?
        sCurrComponent = "Pré Pedido sem Duplicata?"
        C414512 = (fn_ConvertValueCompiled(C414511, C414511DataType, AKB_DecimalPoint, False) = 1)
        C414512DataType = AKBTypeConst.cAKBTypeLogical
        If C414512 Then
            GoTo Comp_C149090
        Else
            GoTo Comp_C414513
        End If

    Comp_C414513:

        ' AtualizaPedidoDuplicataAntec
        sCurrComponent = "AtualizaPedidoDuplicataAntec"
        QueryC414513 = con.CreateCommand()
        QueryC414513.CommandText = QueryC414513.CommandText & " " & "UPDATE AKBUSER01.WF_DUPLIC_RECEBER "
        QueryC414513.CommandText = QueryC414513.CommandText & " " & "SET WF_DUPLIC_RECEBER.Pedido = " & _ 
ConvertValue(P23380, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC414513.CommandText = QueryC414513.CommandText & " " & "WHERE WF_DUPLIC_RECEBER.Ident_Duplicata = " & _ 
ConvertValue(RsC414510(0), C414510DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC414513.Transaction = txn
        C414513 = QueryC414513.ExecuteNonQuery()
        C414513DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C128763

    Exit_R7671:

        Exit Function

    Err_R7671:
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
