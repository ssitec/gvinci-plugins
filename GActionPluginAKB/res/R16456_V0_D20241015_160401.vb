Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R16456

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

    Public Shared Function R16456(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P58928 As Double, ByVal P58946 As Double, ByVal P58947 As Double) As DataTable
    ' 
    ' 16456 - Calcula Custo Presumido e PUD do Item do Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R16456
        Dim sCurrComponent as String

        Dim C359034 As Object
        Dim C359034DataType As Short
        Dim QueryC359060 As New Object
        Dim RsC359060 As New Object
        Dim C359060DataType() As Short
        Dim RsC359060_EOF As Boolean
        Dim QueryC359061 As New Object
        Dim RsC359061 As New Object
        Dim C359061DataType() As Short
        Dim RsC359061_EOF As Boolean
        Dim QueryC359062 As New Object
        Dim RsC359062 As New Object
        Dim C359062DataType() As Short
        Dim RsC359062_EOF As Boolean
        Dim QueryC359063 As New Object
        Dim RsC359063 As New Object
        Dim C359063DataType() As Short
        Dim RsC359063_EOF As Boolean
        Dim C359064 As Object
        Dim C359064DataType As Short
        Dim C359065 As Object
        Dim C359065DataType As Short
        Dim C359071 As Object
        Dim C359071DataType As Short
        Dim C359072 As Boolean
        Dim C359072DataType As Short
        Dim C359073 As Object
        Dim C359073DataType As Short
        Dim C359074 As Object
        Dim C359074DataType As Short
        Dim QueryC359087 As New Object
        Dim RsC359087 As New Object
        Dim C359087DataType() As Short
        Dim RsC359087_EOF As Boolean
        Dim C359088 As Object
        Dim C359088DataType As Short
        Dim C359089 As Double
        Dim C359089DataType As Short
        Dim C359090 As Double
        Dim C359090DataType As Short
        Dim C359108 As Object
        Dim C359108DataType As Short
        Dim C359124 As Object
        Dim C359124DataType As Short
        Dim QueryC359125 As New Object
        Dim RsC359125 As New Object
        Dim C359125DataType() As Short
        Dim RsC359125_EOF As Boolean
        Dim C359126 As Boolean
        Dim C359126DataType As Short
        Dim C359127 As Object
        Dim C359127DataType As Short
        Dim C359128 As Object
        Dim C359128DataType As Short
        Dim C359129 As Object
        Dim C359129DataType As Short
        Dim C359130 As Object
        Dim C359130DataType As Short
        Dim C359131 As Object
        Dim C359131DataType As Short
        Dim C359132 As Boolean
        Dim C359132DataType As Short
        Dim QueryC359133 As New Object
        Dim RsC359133 As New Object
        Dim C359133DataType() As Short
        Dim RsC359133_EOF As Boolean
        Dim C359134 As Object
        Dim C359134DataType As Short
        Dim C359135 As Boolean
        Dim C359135DataType As Short
        Dim C359136 As Object
        Dim C359136DataType As Short
        Dim C359137 As Object
        Dim C359137DataType As Short
        Dim QueryC359138 As New Object
        Dim RsC359138 As New Object
        Dim C359138DataType() As Short
        Dim RsC359138_EOF As Boolean
        Dim C359139 As Object
        Dim C359139DataType As Short
        Dim C359140 As Boolean
        Dim C359140DataType As Short
        Dim C359141 As Object
        Dim C359141DataType As Short
        Dim C359142 As Object
        Dim C359142DataType As Short
        Dim C359143 As Object
        Dim C359143DataType As Short
        Dim QueryC359144 As New Object
        Dim RsC359144 As New Object
        Dim C359144DataType() As Short
        Dim RsC359144_EOF As Boolean
        Dim QueryC359145 As New Object
        Dim RsC359145 As New Object
        Dim C359145DataType() As Short
        Dim RsC359145_EOF As Boolean
        Dim C359146 As Object
        Dim C359146DataType As Short
        Dim C359147 As Object
        Dim C359147DataType As Short
        Dim C359148 As Object
        Dim C359148DataType As Short
        Dim C359149 As Object
        Dim C359149DataType As Short
        Dim C359150 As Object
        Dim C359150DataType As Short
        Dim C359151 As Boolean
        Dim C359151DataType As Short
        Dim C359152 As Object
        Dim C359152DataType As Short
        Dim C359153 As Object
        Dim C359153DataType As Short
        Dim C359154 As Object
        Dim C359154DataType As Short
        Dim C359155 As Object
        Dim C359155DataType As Short
        Dim C359156 As Object
        Dim C359156DataType As Short
        Dim C359157 As Object
        Dim C359157DataType As Short
        Dim C359158 As Object
        Dim C359158DataType As Short
        Dim C359159 As Object
        Dim C359159DataType As Short
        Dim C359160 As Boolean
        Dim C359160DataType As Short
        Dim C359161 As Boolean
        Dim C359161DataType As Short
        Dim C359162 As Object
        Dim C359162DataType As Short
        Dim QueryC359178 As New Object
        Dim RsC359178 As New Object
        Dim C359178DataType() As Short
        Dim RsC359178_EOF As Boolean
        Dim C359179 As Object
        Dim C359179DataType As Short
        Dim C359180 As Object
        Dim C359180DataType As Short
        Dim C359205 As Object
        Dim C359205DataType As Short
        Dim C359230 As Object
        Dim C359230DataType As Short
        Dim C359231 As Object
        Dim C359231DataType As Short
        Dim C359232 As Boolean
        Dim C359232DataType As Short
        Dim C359233 As Object
        Dim C359233DataType As Short
        Dim C359234 As Object
        Dim C359234DataType As Short
        Dim C359235 As Object
        Dim C359235DataType As Short
        Dim C359236 As Boolean
        Dim C359236DataType As Short
        Dim C359238 As Object
        Dim C359238DataType As Short
        Dim C359239 As Boolean
        Dim C359239DataType As Short
        Dim C359240 As Object
        Dim C359240DataType As Short
        Dim C359246 As Object
        Dim C359246DataType As Short
        Dim C359247 As Double
        Dim C359247DataType As Short
        Dim C359248 As Object
        Dim C359248DataType As Short
        Dim C359249 As Object
        Dim C359249DataType As Short
        Dim C359250 As Double
        Dim C359250DataType As Short
        Dim C359251 As Double
        Dim C359251DataType As Short
        Dim QueryC359253 As New Object
        Dim RsC359253 As New Object
        Dim C359253DataType() As Short
        Dim RsC359253_EOF As Boolean
        Dim C359254 As Boolean
        Dim C359254DataType As Short
        Dim C359255 As Object
        Dim C359255DataType As Short
        Dim C359256 As Object
        Dim C359256DataType As Short
        Dim C359257 As Object
        Dim C359257DataType As Short
        Dim C359258 As Object
        Dim C359258DataType As Short
        Dim C359259 As Object
        Dim C359259DataType As Short
        Dim C359260 As Object
        Dim C359260DataType As Short
        Dim C359261 As Object
        Dim C359261DataType As Short
        Dim C359262 As Object
        Dim C359262DataType As Short
        Dim C359263 As Object
        Dim C359263DataType As Short
        Dim C359264 As Object
        Dim C359264DataType As Short
        Dim C359265 As Object
        Dim C359265DataType As Short
        Dim C359266 As Object
        Dim C359266DataType As Short
        Dim C359267 As Object
        Dim C359267DataType As Short
        Dim C359268 As Object
        Dim C359268DataType As Short
        Dim QueryC359291 As New Object
        Dim RsC359291 As New Object
        Dim C359291DataType() As Short
        Dim RsC359291_EOF As Boolean
        Dim C359292 As Object
        Dim C359292DataType As Short
        Dim C359293 As Boolean
        Dim C359293DataType As Short
        Dim C359299 As Object
        Dim C359299DataType As Short
        Dim C359300 As Object
        Dim C359300DataType As Short
        Dim C359301 As Object
        Dim C359301DataType As Short
        Dim C359302 As Object
        Dim C359302DataType As Short
        Dim C359303 As Object
        Dim C359303DataType As Short
        Dim C359304 As Object
        Dim C359304DataType As Short
        Dim C359305 As Object
        Dim C359305DataType As Short
        Dim C359308 As Object
        Dim C359308DataType As Short
        Dim C359310 As Object
        Dim C359310DataType As Short
        Dim C359311 As Object
        Dim C359311DataType As Short
        Dim C359312 As Object
        Dim C359312DataType As Short
        Dim C359313 As Object
        Dim C359313DataType As Short
        Dim QueryC359315 As New Object
        Dim RsC359315 As New Object
        Dim C359315DataType() As Short
        Dim RsC359315_EOF As Boolean
        Dim C359330DataType() As Short
        Dim QueryC359348 As New Object
        Dim RsC359348 As New Object
        Dim C359348DataType() As Short
        Dim RsC359348_EOF As Boolean
        Dim QueryC359350 As New Object
        Dim RsC359350 As New Object
        Dim C359350DataType() As Short
        Dim RsC359350_EOF As Boolean
        Dim C359351 As Object
        Dim C359351DataType As Short
        Dim C359352 As Object
        Dim C359352DataType As Short
        Dim QueryC359358 As New Object
        Dim C359358 As Integer
        Dim C359358DataType As Short
        Dim C359359 As Object
        Dim C359359DataType As Short
        Dim C359360 As Object
        Dim C359360DataType As Short
        Dim C359406 As DataTable
        Dim C359406CurrentRow As Integer
        Dim C359406DataType() As Short
        Dim C359407 As Double
        Dim C359407DataType As Short
        Dim C368769 As Boolean
        Dim C368769DataType As Short
        Dim QueryC368770 As New Object
        Dim C368770 As Integer
        Dim C368770DataType As Short
        Dim C369726 As Object
        Dim C369726DataType As Short
        Dim C369727 As Object
        Dim C369727DataType As Short
        Dim QueryC533250 As New Object
        Dim RsC533250 As New Object
        Dim C533250DataType() As Short
        Dim RsC533250_EOF As Boolean
        Dim C533251 As Boolean
        Dim C533251DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C359205

    Comp_C359034:

        ' Next
        sCurrComponent = "Next"
        C359034DataType = 4
        RsC359291_EOF = Not RsC359291.Read()
        C359034 = True

        GoTo Comp_C359292

    Comp_C359060:

        ' IdCusto
        sCurrComponent = "IdCusto"
        QueryC359060 = con.CreateCommand()
        QueryC359060.CommandText = QueryC359060.CommandText & " " & "SELECT NVL(MAX(WF_CUSTO_PRODUTO.Id_Custo),0)"
        QueryC359060.CommandText = QueryC359060.CommandText & " " & ""
        QueryC359060.CommandText = QueryC359060.CommandText & " " & " FROM AKBUSER01.WF_CUSTO_PRODUTO , AKBUSER01.WF_CUSTO_STANDARD "
        QueryC359060.CommandText = QueryC359060.CommandText & " " & ""
        QueryC359060.CommandText = QueryC359060.CommandText & " " & " WHERE WF_CUSTO_PRODUTO.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC359060.CommandText = QueryC359060.CommandText & " " & " AND  WF_CUSTO_PRODUTO.Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359060.CommandText = QueryC359060.CommandText & " " & " AND WF_CUSTO_PRODUTO.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   "
        QueryC359060.CommandText = QueryC359060.CommandText & " " & " AND WF_CUSTO_STANDARD.Id_Custo = WF_CUSTO_PRODUTO.Id_Custo "
        QueryC359060.CommandText = QueryC359060.CommandText & " " & " AND WF_CUSTO_STANDARD.Data_Geracao >= " & _ 
ConvertValue(C359088, C359088DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359060.CommandText = QueryC359060.CommandText & " " & " AND WF_CUSTO_STANDARD.Data_Geracao <= " & _ 
ConvertValue(C359231, C359231DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359060.CommandText = QueryC359060.CommandText & " " & " AND WF_CUSTO_PRODUTO.Custo_e_Simulacao = 0"
        QueryC359060.CommandText = QueryC359060.CommandText & " " & "AND WF_CUSTO_PRODUTO.Custo_Total > 0"
        QueryC359060.Transaction = txn
        RsC359060 = QueryC359060.ExecuteReader()
        Dim iC359060 As Short
        ReDim C359060DataType(RsC359060.FieldCount)
        For iC359060 = 0 to RsC359060.FieldCount - 1
            Select Case RsC359060.GetDataTypeName(iC359060).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359060DataType(iC359060 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359060DataType(iC359060 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359060DataType(iC359060 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359060
        RsC359060_EOF = Not RsC359060.Read()

        GoTo Comp_C359137

    Comp_C359061:

        ' SeqCusto
        sCurrComponent = "SeqCusto"
        QueryC359061 = con.CreateCommand()
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "SELECT NVL(MAX (WF_CUSTO_PRODUTO.Seq_Custo) ,0)"
        QueryC359061.CommandText = QueryC359061.CommandText & " " & ""
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "FROM AKBUSER01.WF_CUSTO_PRODUTO , "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AKBUSER01.WF_CUSTO_STANDARD "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & ""
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "WHERE WF_CUSTO_PRODUTO.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Id_Custo =  " & _ 
ConvertValue(C359137, C359137DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_STANDARD.Id_Custo = WF_CUSTO_PRODUTO.Id_Custo "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Custo_Total > 0"
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Custo_Total ="
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "(SELECT  " & _ 
ConvertValue(C359238, C359238DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "(WF_CUSTO_PRODUTO.Custo_Total) "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "FROM AKBUSER01.WF_CUSTO_PRODUTO , "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AKBUSER01.WF_CUSTO_STANDARD "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "WHERE WF_CUSTO_PRODUTO.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Id_Custo = " & _ 
ConvertValue(C359137, C359137DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_STANDARD.Id_Custo = WF_CUSTO_PRODUTO.Id_Custo "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Custo_Total > 0"
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Parametro_Qtde ="
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "(SELECT MAX (WF_CUSTO_PRODUTO.Parametro_Qtde) "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "FROM AKBUSER01.WF_CUSTO_PRODUTO ,"
        QueryC359061.CommandText = QueryC359061.CommandText & " " & " AKBUSER01.WF_CUSTO_STANDARD "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "WHERE WF_CUSTO_PRODUTO.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Id_Custo = " & _ 
ConvertValue(C359137, C359137DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_STANDARD.Id_Custo = WF_CUSTO_PRODUTO.Id_Custo "
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "AND WF_CUSTO_PRODUTO.Custo_Total > 0"
        QueryC359061.CommandText = QueryC359061.CommandText & " " & "" & _ 
ConvertValue(C359124, C359124DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  )"
        QueryC359061.CommandText = QueryC359061.CommandText & " " & ")"
        QueryC359061.CommandText = QueryC359061.CommandText & " " & ""
        QueryC359061.Transaction = txn
        RsC359061 = QueryC359061.ExecuteReader()
        Dim iC359061 As Short
        ReDim C359061DataType(RsC359061.FieldCount)
        For iC359061 = 0 to RsC359061.FieldCount - 1
            Select Case RsC359061.GetDataTypeName(iC359061).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359061DataType(iC359061 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359061DataType(iC359061 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359061DataType(iC359061 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359061
        RsC359061_EOF = Not RsC359061.Read()

        GoTo Comp_C359062

    Comp_C359062:

        ' CustoProdSP
        sCurrComponent = "CustoProdSP"
        QueryC359062 = con.CreateCommand()
        QueryC359062.CommandText = QueryC359062.CommandText & " " & "SELECT NVL(MAX(WF_CUSTO_PRODUTO.Custo_Insumos_Acumulado),0) ,"
        QueryC359062.CommandText = QueryC359062.CommandText & " " & " NVL(MAX(WF_CUSTO_PRODUTO.Custo_Direto_Acumulado),0) ,"
        QueryC359062.CommandText = QueryC359062.CommandText & " " & "NVL(MAX(WF_CUSTO_PRODUTO.Custo_Indireto_Acumulado),0),"
        QueryC359062.CommandText = QueryC359062.CommandText & " " & "NVL(MAX(WF_CUSTO_PRODUTO.Custo_Maquina_Acum),0),"
        QueryC359062.CommandText = QueryC359062.CommandText & " " & "NVL(MAX(WF_CUSTO_PRODUTO.Custo_itens_sem_desemb_cx_ac),0)"
        QueryC359062.CommandText = QueryC359062.CommandText & " " & ""
        QueryC359062.CommandText = QueryC359062.CommandText & " " & " FROM AKBUSER01.WF_CUSTO_PRODUTO , AKBUSER01.WF_CUSTO_STANDARD "
        QueryC359062.CommandText = QueryC359062.CommandText & " " & " WHERE WF_CUSTO_PRODUTO.Id_Custo = " & _ 
ConvertValue(C359137, C359137DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359062.CommandText = QueryC359062.CommandText & " " & " AND  WF_CUSTO_PRODUTO.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "    "
        QueryC359062.CommandText = QueryC359062.CommandText & " " & " AND  WF_CUSTO_PRODUTO.Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359062.CommandText = QueryC359062.CommandText & " " & " AND WF_CUSTO_PRODUTO.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   "
        QueryC359062.CommandText = QueryC359062.CommandText & " " & " AND  WF_CUSTO_PRODUTO.Seq_Custo = " & _ 
ConvertValue(RsC359061(0), C359061DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC359062.CommandText = QueryC359062.CommandText & " " & " AND WF_CUSTO_STANDARD.Id_Custo = WF_CUSTO_PRODUTO.Id_Custo "
        QueryC359062.Transaction = txn
        RsC359062 = QueryC359062.ExecuteReader()
        Dim iC359062 As Short
        ReDim C359062DataType(RsC359062.FieldCount)
        For iC359062 = 0 to RsC359062.FieldCount - 1
            Select Case RsC359062.GetDataTypeName(iC359062).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359062DataType(iC359062 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359062DataType(iC359062 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359062DataType(iC359062 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359062
        RsC359062_EOF = Not RsC359062.Read()

        GoTo Comp_C359230

    Comp_C359063:

        ' FatConv
        sCurrComponent = "FatConv"
        QueryC359063 = con.CreateCommand()
        QueryC359063.CommandText = QueryC359063.CommandText & " " & "SELECT WF_EMB_COMP_VDA_PROD.Fator_Conv_Unid_1 FROM AKBUSER01.WF_EMB_COMP_VDA_PROD"
        QueryC359063.CommandText = QueryC359063.CommandText & " " & " WHERE WF_EMB_COMP_VDA_PROD.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "    "
        QueryC359063.CommandText = QueryC359063.CommandText & " " & " AND  WF_EMB_COMP_VDA_PROD.Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359063.CommandText = QueryC359063.CommandText & " " & " AND  WF_EMB_COMP_VDA_PROD.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359063.Transaction = txn
        RsC359063 = QueryC359063.ExecuteReader()
        Dim iC359063 As Short
        ReDim C359063DataType(RsC359063.FieldCount)
        For iC359063 = 0 to RsC359063.FieldCount - 1
            Select Case RsC359063.GetDataTypeName(iC359063).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359063DataType(iC359063 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359063DataType(iC359063 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359063DataType(iC359063 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359063
        RsC359063_EOF = Not RsC359063.Read()

        GoTo Comp_C359071

    Comp_C359064:

        ' CustoDirSP
        sCurrComponent = "CustoDirSP"
        C359064DataType = 0
        C359064DataType = C359062Datatype(2)
        If C359064DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359062(1)) Then
          C359064 = Strings.RTrim(RsC359062(1))
        Else 
          C359064 = RsC359062(1)
        End If 

        GoTo Comp_C359065

    Comp_C359065:

        ' CustoIndSP
        sCurrComponent = "CustoIndSP"
        C359065DataType = 0
        C359065DataType = C359062Datatype(3)
        If C359065DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359062(2)) Then
          C359065 = Strings.RTrim(RsC359062(2))
        Else 
          C359065 = RsC359062(2)
        End If 

        GoTo Comp_C359248

    Comp_C359071:

        ' vFatConv
        sCurrComponent = "vFatConv"
        C359071 = 1
        C359071DataType = 1
        GoTo Comp_C359072

    Comp_C359072:

        ' FatConv = Nulo
        sCurrComponent = "FatConv = Nulo"
        C359072 = (fn_ConvertValueCompiled(RsC359063(0), C359063DataType(1), AKB_DecimalPoint, False) Is System.DBNull.Value OR fn_ConvertValueCompiled(RsC359063(0), C359063DataType(1), AKB_DecimalPoint, False)=0)
        C359072DataType = AKBTypeConst.cAKBTypeLogical
        If C359072 Then
            GoTo Comp_C359246
        Else
            GoTo Comp_C359073
        End If

    Comp_C359073:

        ' AtrValorvFatConv
        sCurrComponent = "AtrValorvFatConv"
        C359073DataType = 4
        C359071 = fn_ConvertValueCompiled(RsC359063(0) , 1, AKB_DecimalPoint)
        C359073 = True
        GoTo Comp_C359246

    Comp_C359074:

        ' DataSistema
        sCurrComponent = "DataSistema"
        C359074DataType = 2
        C359074 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy HH:mm:ss"))
        GoTo Comp_C359291

    Comp_C359087:

        ' MesesCusto
        sCurrComponent = "MesesCusto"
        QueryC359087 = con.CreateCommand()
        QueryC359087.CommandText = QueryC359087.CommandText & " " & "SELECT NVL(WF_SIGLA_PROD_AGRUP.Meses_Max_Validade_Custo, 0) "
        QueryC359087.CommandText = QueryC359087.CommandText & " " & ""
        QueryC359087.CommandText = QueryC359087.CommandText & " " & "FROM AKBUSER01.WF_SIGLA_PROD_AGRUP "
        QueryC359087.CommandText = QueryC359087.CommandText & " " & ""
        QueryC359087.CommandText = QueryC359087.CommandText & " " & "WHERE WF_SIGLA_PROD_AGRUP.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359087.Transaction = txn
        RsC359087 = QueryC359087.ExecuteReader()
        Dim iC359087 As Short
        ReDim C359087DataType(RsC359087.FieldCount)
        For iC359087 = 0 to RsC359087.FieldCount - 1
            Select Case RsC359087.GetDataTypeName(iC359087).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359087DataType(iC359087 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359087DataType(iC359087 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359087DataType(iC359087 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359087
        RsC359087_EOF = Not RsC359087.Read()

        GoTo Comp_C359088

    Comp_C359088:

        ' DtValCT
        sCurrComponent = "DtValCT"
        C359088DataType = 2
        C359088 = DateAdd("m", -RsC359087(0), C359074)
        GoTo Comp_C359231

    Comp_C359089:

        ' UnitCDSP
        sCurrComponent = "UnitCDSP"
        C359089 = (fn_ConvertValueCompiled(C359064, C359064DataType, AKB_DecimalPoint, True)/fn_ConvertValueCompiled(C359071, C359071DataType, AKB_DecimalPoint, True) ) *  fn_ConvertValueCompiled(C359304, C359304DataType, AKB_DecimalPoint, True)
        C359089DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359090

    Comp_C359090:

        ' UnitCISP
        sCurrComponent = "UnitCISP"
        C359090 = (fn_ConvertValueCompiled(C359065, C359065DataType, AKB_DecimalPoint, True)/fn_ConvertValueCompiled(C359071, C359071DataType, AKB_DecimalPoint, True) ) *  fn_ConvertValueCompiled(C359304, C359304DataType, AKB_DecimalPoint, True)
        C359090DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359247

    Comp_C359108:

        ' VlImpVar
        sCurrComponent = "VlImpVar"
        C359108 = 0
        C359108DataType = 1
        GoTo Comp_C359145

    Comp_C359124:

        ' query
        sCurrComponent = "query"
        C359124 = "AND 1 = 1"
        C359124DataType = 5
        GoTo Comp_C359238

    Comp_C359125:

        ' Count
        sCurrComponent = "Count"
        QueryC359125 = con.CreateCommand()
        QueryC359125.CommandText = QueryC359125.CommandText & " " & "SELECT COUNT (*) "
        QueryC359125.CommandText = QueryC359125.CommandText & " " & "FROM WF_PARAM_CUSTO_PROD"
        QueryC359125.CommandText = QueryC359125.CommandText & " " & "WHERE Sigla_Prod  = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359125.CommandText = QueryC359125.CommandText & " " & "       AND Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC359125.CommandText = QueryC359125.CommandText & " " & "       AND Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC359125.Transaction = txn
        RsC359125 = QueryC359125.ExecuteReader()
        Dim iC359125 As Short
        ReDim C359125DataType(RsC359125.FieldCount)
        For iC359125 = 0 to RsC359125.FieldCount - 1
            Select Case RsC359125.GetDataTypeName(iC359125).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359125DataType(iC359125 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359125DataType(iC359125 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359125DataType(iC359125 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359125
        RsC359125_EOF = Not RsC359125.Read()

        GoTo Comp_C359126

    Comp_C359126:

        ' Count<=1
        sCurrComponent = "Count<=1"
        C359126 = (fn_ConvertValueCompiled(RsC359125(0), C359125DataType(1), AKB_DecimalPoint, False) <=1)
        C359126DataType = AKBTypeConst.cAKBTypeLogical
        If C359126 Then
            GoTo Comp_C359061
        Else
            GoTo Comp_C359128
        End If

    Comp_C359127:

        ' AtrValorQuery
        sCurrComponent = "AtrValorQuery"
        C359127DataType = 4
        C359124 = fn_ConvertValueCompiled("AND WF_CUSTO_PRODUTO.Parametro_Qtde <= " & C359128 , 5, AKB_DecimalPoint)
        C359127 = True
        GoTo Comp_C359061

    Comp_C359128:

        ' Troca
        sCurrComponent = "Troca"
        C359128DataType = 3
        C359128 = Replace(C359304 , ",", ".")
        GoTo Comp_C359127

    Comp_C359129:

        ' vSigla
        sCurrComponent = "vSigla"
        C359129 = C359301 
        C359129DataType = 3
        GoTo Comp_C359130

    Comp_C359130:

        ' vAcesso
        sCurrComponent = "vAcesso"
        C359130 = C359302 & " "
        C359130DataType = 1
        GoTo Comp_C359131

    Comp_C359131:

        ' vCodEmb
        sCurrComponent = "vCodEmb"
        C359131 = C359303 & " "
        C359131DataType = 1
        GoTo Comp_C359234

    Comp_C359132:

        ' IdCusto>0?
        sCurrComponent = "IdCusto>0?"
        C359132 = (1 = 1)
        C359132DataType = AKBTypeConst.cAKBTypeLogical
        If C359132 Then
            GoTo Comp_C359124
        Else
            GoTo Comp_C359144
        End If

    Comp_C359133:

        ' SelSemelh
        sCurrComponent = "SelSemelh"
        QueryC359133 = con.CreateCommand()
        QueryC359133.CommandText = QueryC359133.CommandText & " " & "SELECT DISTINCT WF_VISAO_COMERCIAL_DADOS.Sigla_Prod , WF_VISAO_COMERCIAL_DADOS.Acesso"
        QueryC359133.CommandText = QueryC359133.CommandText & " " & "FROM WF_EMB_COMP_VDA_PROD , WF_VISAO_COMERCIAL_DADOS"
        QueryC359133.CommandText = QueryC359133.CommandText & " " & "WHERE WF_VISAO_COMERCIAL_DADOS.Sigla_Prod = WF_EMB_COMP_VDA_PROD.Sigla_Prod"
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " AND  WF_VISAO_COMERCIAL_DADOS.Acesso = WF_EMB_COMP_VDA_PROD.Acesso"
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " AND  WF_EMB_COMP_VDA_PROD.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " AND  WF_VISAO_COMERCIAL_DADOS.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " AND  WF_VISAO_COMERCIAL_DADOS.Acesso <> " & _ 
ConvertValue(C359302, C359302DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " " & _ 
ConvertValue(C359152, C359152DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " " & _ 
ConvertValue(C359153, C359153DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " " & _ 
ConvertValue(C359154, C359154DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " " & _ 
ConvertValue(C359155, C359155DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359133.CommandText = QueryC359133.CommandText & " " & ""
        QueryC359133.CommandText = QueryC359133.CommandText & " " & " ORDER BY WF_VISAO_COMERCIAL_DADOS.Acesso Desc  "
        QueryC359133.Transaction = txn
        RsC359133 = QueryC359133.ExecuteReader()
        Dim iC359133 As Short
        ReDim C359133DataType(RsC359133.FieldCount)
        For iC359133 = 0 to RsC359133.FieldCount - 1
            Select Case RsC359133.GetDataTypeName(iC359133).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359133DataType(iC359133 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359133DataType(iC359133 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359133DataType(iC359133 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359133
        RsC359133_EOF = Not RsC359133.Read()

        GoTo Comp_C359134

    Comp_C359134:

        ' FimSelSem
        sCurrComponent = "FimSelSem"
        C359134DataType = 4
        C359134 = RsC359133_EOF
        GoTo Comp_C359135

    Comp_C359135:

        ' FimSelSem?
        sCurrComponent = "FimSelSem?"
        C359135 = (fn_ConvertValueCompiled(C359134, C359134DataType, AKB_DecimalPoint, False) = 1  )
        C359135DataType = AKBTypeConst.cAKBTypeLogical
        If C359135 Then
            GoTo Comp_C359162
        Else
            GoTo Comp_C359142
        End If

    Comp_C359136:

        ' AtrValorvAcesso
        sCurrComponent = "AtrValorvAcesso"
        C359136DataType = 4
        C359130 = fn_ConvertValueCompiled(C359143 , 1, AKB_DecimalPoint)
        C359136 = True
        GoTo Comp_C359138

    Comp_C359137:

        ' vIDCusto
        sCurrComponent = "vIDCusto"
        C359137 = RsC359060(0) & " "
        C359137DataType = 1
        GoTo Comp_C359132

    Comp_C359138:

        ' IdCusto1
        sCurrComponent = "IdCusto1"
        QueryC359138 = con.CreateCommand()
        QueryC359138.CommandText = QueryC359138.CommandText & " " & "SELECT NVL(MAX(WF_CUSTO_PRODUTO.Id_Custo),0) FROM AKBUSER01.WF_CUSTO_PRODUTO , AKBUSER01.WF_CUSTO_STANDARD "
        QueryC359138.CommandText = QueryC359138.CommandText & " " & " WHERE WF_CUSTO_PRODUTO.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC359138.CommandText = QueryC359138.CommandText & " " & " AND  WF_CUSTO_PRODUTO.Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359138.CommandText = QueryC359138.CommandText & " " & " AND WF_CUSTO_PRODUTO.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   "
        QueryC359138.CommandText = QueryC359138.CommandText & " " & " AND WF_CUSTO_STANDARD.Id_Custo = WF_CUSTO_PRODUTO.Id_Custo "
        QueryC359138.CommandText = QueryC359138.CommandText & " " & " AND WF_CUSTO_PRODUTO.Custo_e_Simulacao = 0"
        QueryC359138.CommandText = QueryC359138.CommandText & " " & " AND WF_CUSTO_PRODUTO.Custo_Total > 0"
        QueryC359138.CommandText = QueryC359138.CommandText & " " & ""
        QueryC359138.Transaction = txn
        RsC359138 = QueryC359138.ExecuteReader()
        Dim iC359138 As Short
        ReDim C359138DataType(RsC359138.FieldCount)
        For iC359138 = 0 to RsC359138.FieldCount - 1
            Select Case RsC359138.GetDataTypeName(iC359138).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359138DataType(iC359138 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359138DataType(iC359138 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359138DataType(iC359138 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359138
        RsC359138_EOF = Not RsC359138.Read()

        GoTo Comp_C359140

    Comp_C359139:

        ' NextSelSem
        sCurrComponent = "NextSelSem"
        C359139DataType = 4
        RsC359133_EOF = Not RsC359133.Read()
        C359139 = True

        GoTo Comp_C359134

    Comp_C359140:

        ' IdCusto1>0?
        sCurrComponent = "IdCusto1>0?"
        C359140 = (fn_ConvertValueCompiled(RsC359138(0), C359138DataType(1), AKB_DecimalPoint, False) > 0)
        C359140DataType = AKBTypeConst.cAKBTypeLogical
        If C359140 Then
            GoTo Comp_C359141
        Else
            GoTo Comp_C359139
        End If

    Comp_C359141:

        ' AtrValorvIdCusto
        sCurrComponent = "AtrValorvIdCusto"
        C359141DataType = 4
        C359137 = fn_ConvertValueCompiled(RsC359138(0) , 1, AKB_DecimalPoint)
        C359141 = True
        GoTo Comp_C359124

    Comp_C359142:

        ' SiglaSem
        sCurrComponent = "SiglaSem"
        C359142DataType = 0
        C359142 = RsC359133(0)
        C359142DataType = C359133Datatype(1)
        If C359142DataType = AKBTypeConst.cAKBTypeString Then
          C359142 = IIF(IsDBNull(C359142), C359142, Strings.RTrim(C359142))
        End If 

        GoTo Comp_C359143

    Comp_C359143:

        ' AcessSem
        sCurrComponent = "AcessSem"
        C359143DataType = 0
        C359143DataType = C359133Datatype(2)
        If C359143DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359133(1)) Then
          C359143 = Strings.RTrim(RsC359133(1))
        Else 
          C359143 = RsC359133(1)
        End If 

        GoTo Comp_C359136

    Comp_C359144:

        ' SeqVisao
        sCurrComponent = "SeqVisao"
        QueryC359144 = con.CreateCommand()
        QueryC359144.CommandText = QueryC359144.CommandText & " " & "SELECT MAX(Sequencia)-1"
        QueryC359144.CommandText = QueryC359144.CommandText & " " & " FROM WF_VISAO_COMERCIAL"
        QueryC359144.CommandText = QueryC359144.CommandText & " " & " WHERE Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359144.CommandText = QueryC359144.CommandText & " " & ""
        QueryC359144.Transaction = txn
        RsC359144 = QueryC359144.ExecuteReader()
        Dim iC359144 As Short
        ReDim C359144DataType(RsC359144.FieldCount)
        For iC359144 = 0 to RsC359144.FieldCount - 1
            Select Case RsC359144.GetDataTypeName(iC359144).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359144DataType(iC359144 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359144DataType(iC359144 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359144DataType(iC359144 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359144
        RsC359144_EOF = Not RsC359144.Read()

        GoTo Comp_C359152

    Comp_C359145:

        ' SelVisao
        sCurrComponent = "SelVisao"
        QueryC359145 = con.CreateCommand()
        QueryC359145.CommandText = QueryC359145.CommandText & " " & "SELECT NVL(COL1,NULL),NVL(COL2,NULL),NVL(COL3,NULL),NVL(COL4,NULL),NVL(COL5,NULL)"
        QueryC359145.CommandText = QueryC359145.CommandText & " " & " FROM WF_VISAO_COMERCIAL_DADOS"
        QueryC359145.CommandText = QueryC359145.CommandText & " " & "  WHERE Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359145.CommandText = QueryC359145.CommandText & " " & "         AND Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359145.Transaction = txn
        RsC359145 = QueryC359145.ExecuteReader()
        Dim iC359145 As Short
        ReDim C359145DataType(RsC359145.FieldCount)
        For iC359145 = 0 to RsC359145.FieldCount - 1
            Select Case RsC359145.GetDataTypeName(iC359145).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359145DataType(iC359145 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359145DataType(iC359145 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359145DataType(iC359145 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359145
        RsC359145_EOF = Not RsC359145.Read()

        GoTo Comp_C359146

    Comp_C359146:

        ' Visao1
        sCurrComponent = "Visao1"
        C359146DataType = 0
        C359146 = RsC359145(0)
        C359146DataType = C359145Datatype(1)
        If C359146DataType = AKBTypeConst.cAKBTypeString Then
          C359146 = IIF(IsDBNull(C359146), C359146, Strings.RTrim(C359146))
        End If 

        GoTo Comp_C359147

    Comp_C359147:

        ' Visao2
        sCurrComponent = "Visao2"
        C359147DataType = 0
        C359147DataType = C359145Datatype(2)
        If C359147DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359145(1)) Then
          C359147 = Strings.RTrim(RsC359145(1))
        Else 
          C359147 = RsC359145(1)
        End If 

        GoTo Comp_C359148

    Comp_C359148:

        ' Visao3
        sCurrComponent = "Visao3"
        C359148DataType = 0
        C359148DataType = C359145Datatype(3)
        If C359148DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359145(2)) Then
          C359148 = Strings.RTrim(RsC359145(2))
        Else 
          C359148 = RsC359145(2)
        End If 

        GoTo Comp_C359149

    Comp_C359149:

        ' Visao4
        sCurrComponent = "Visao4"
        C359149DataType = 0
        C359149DataType = C359145Datatype(4)
        If C359149DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359145(3)) Then
          C359149 = Strings.RTrim(RsC359145(3))
        Else 
          C359149 = RsC359145(3)
        End If 

        GoTo Comp_C359150

    Comp_C359150:

        ' Visao5
        sCurrComponent = "Visao5"
        C359150DataType = 0
        C359150DataType = C359145Datatype(5)
        If C359150DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359145(4)) Then
          C359150 = Strings.RTrim(RsC359145(4))
        Else 
          C359150 = RsC359145(4)
        End If 

        GoTo Comp_C359151

    Comp_C359151:

        ' SeqVC=4?
        sCurrComponent = "SeqVC=4?"
        C359151 = (fn_ConvertValueCompiled(RsC359144(0), C359144DataType(1), AKB_DecimalPoint, False) = 4)
        C359151DataType = AKBTypeConst.cAKBTypeLogical
        If C359151 Then
            GoTo Comp_C359156
        Else
            GoTo Comp_C359160
        End If

    Comp_C359152:

        ' vVC1
        sCurrComponent = "vVC1"
        C359152 = "AND 1 = 1"
        C359152DataType = 5
        GoTo Comp_C359153

    Comp_C359153:

        ' vVC2
        sCurrComponent = "vVC2"
        C359153 = "AND 1 = 1"
        C359153DataType = 5
        GoTo Comp_C359154

    Comp_C359154:

        ' vVC3
        sCurrComponent = "vVC3"
        C359154 = "AND 1 = 1"
        C359154DataType = 5
        GoTo Comp_C359155

    Comp_C359155:

        ' vVC4
        sCurrComponent = "vVC4"
        C359155 = "AND 1 = 1"
        C359155DataType = 5
        GoTo Comp_C359108

    Comp_C359156:

        ' att4
        sCurrComponent = "att4"
        C359156DataType = 4
        C359155 = fn_ConvertValueCompiled("AND WF_VISAO_COMERCIAL_DADOS.COL4 = '" & C359149 & "'", 5, AKB_DecimalPoint)
        C359156 = True
        GoTo Comp_C359157

    Comp_C359157:

        ' att3
        sCurrComponent = "att3"
        C359157DataType = 4
        C359154 = fn_ConvertValueCompiled("AND WF_VISAO_COMERCIAL_DADOS.COL3 = '" & C359148 & "'", 5, AKB_DecimalPoint)
        C359157 = True
        GoTo Comp_C359158

    Comp_C359158:

        ' att2
        sCurrComponent = "att2"
        C359158DataType = 4
        C359153 = fn_ConvertValueCompiled("AND WF_VISAO_COMERCIAL_DADOS.COL2 = '" & C359147 & "'", 5, AKB_DecimalPoint)
        C359158 = True
        GoTo Comp_C359159

    Comp_C359159:

        ' att1
        sCurrComponent = "att1"
        C359159DataType = 4
        C359152 = fn_ConvertValueCompiled("AND WF_VISAO_COMERCIAL_DADOS.COL1 = '" & C359146 & "'", 5, AKB_DecimalPoint)
        C359159 = True
        GoTo Comp_C359133

    Comp_C359160:

        ' SeqVC=3?
        sCurrComponent = "SeqVC=3?"
        C359160 = (fn_ConvertValueCompiled(RsC359144(0), C359144DataType(1), AKB_DecimalPoint, False) = 3)
        C359160DataType = AKBTypeConst.cAKBTypeLogical
        If C359160 Then
            GoTo Comp_C359157
        Else
            GoTo Comp_C359161
        End If

    Comp_C359161:

        ' SeqVC=2?
        sCurrComponent = "SeqVC=2?"
        C359161 = (fn_ConvertValueCompiled(RsC359144(0), C359144DataType(1), AKB_DecimalPoint, False) = 2)
        C359161DataType = AKBTypeConst.cAKBTypeLogical
        If C359161 Then
            GoTo Comp_C359158
        Else
            GoTo Comp_C359159
        End If

    Comp_C359162:

        ' AtrValor0vAcesso
        sCurrComponent = "AtrValor0vAcesso"
        C359162DataType = 4
        C359130 = fn_ConvertValueCompiled(0, 1, AKB_DecimalPoint)
        C359162 = True
        GoTo Comp_C359124

    Comp_C359178:

        ' SelectImpostos
        sCurrComponent = "SelectImpostos"
        QueryC359178 = con.CreateCommand()
        QueryC359178.CommandText = QueryC359178.CommandText & " " & "SELECT NVL(WF_ESTAB_JURIDICO_PARAM.Cod_Imposto_PIS,0) PIS,"
        QueryC359178.CommandText = QueryC359178.CommandText & " " & "  NVL(WF_ESTAB_JURIDICO_PARAM.Cod_Imposto_COFINS,0) COFINS,"
        QueryC359178.CommandText = QueryC359178.CommandText & " " & "  NVL(WF_ESTAB_JURIDICO_PARAM.COD_IMPOSTO_ICM,0) ICM,"
        QueryC359178.CommandText = QueryC359178.CommandText & " " & "  NVL(WF_ESTAB_JURIDICO_PARAM.COD_IMPOSTO_IPI,0) IPI"
        QueryC359178.CommandText = QueryC359178.CommandText & " " & ""
        QueryC359178.CommandText = QueryC359178.CommandText & " " & "FROM WF_ESTAB_JURIDICO_PARAM"
        QueryC359178.CommandText = QueryC359178.CommandText & " " & "WHERE WF_ESTAB_JURIDICO_PARAM.COD_PESSOA_ESTAB_JURIDICO = " & _ 
ConvertValue(P58928, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359178.Transaction = txn
        RsC359178 = QueryC359178.ExecuteReader()
        Dim iC359178 As Short
        ReDim C359178DataType(RsC359178.FieldCount)
        For iC359178 = 0 to RsC359178.FieldCount - 1
            Select Case RsC359178.GetDataTypeName(iC359178).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359178DataType(iC359178 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359178DataType(iC359178 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359178DataType(iC359178 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359178
        RsC359178_EOF = Not RsC359178.Read()

        GoTo Comp_C359179

    Comp_C359179:

        ' CodPIS
        sCurrComponent = "CodPIS"
        C359179DataType = 0
        C359179 = RsC359178(0)
        C359179DataType = C359178Datatype(1)
        If C359179DataType = AKBTypeConst.cAKBTypeString Then
          C359179 = IIF(IsDBNull(C359179), C359179, Strings.RTrim(C359179))
        End If 

        GoTo Comp_C359180

    Comp_C359180:

        ' CodCOFINS
        sCurrComponent = "CodCOFINS"
        C359180DataType = 0
        C359180DataType = C359178Datatype(2)
        If C359180DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359178(1)) Then
          C359180 = Strings.RTrim(RsC359178(1))
        Else 
          C359180 = RsC359178(1)
        End If 

        GoTo Comp_C359359

    Comp_C359205:

        ' Ret
        sCurrComponent = "Ret"
        C359205 = 1
        C359205DataType = 4
        GoTo Comp_C359178

    Comp_C359230:

        ' ValCSPDir
        sCurrComponent = "ValCSPDir"
        C359230DataType = 0
        C359230DataType = C359062Datatype(2)
        If C359230DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359062(1)) Then
          C359230 = Strings.RTrim(RsC359062(1))
        Else 
          C359230 = RsC359062(1)
        End If 

        GoTo Comp_C359236

    Comp_C359231:

        ' vDT
        sCurrComponent = "vDT"
        C359231 = C359308 & " "
        C359231DataType = 2
        GoTo Comp_C359060

    Comp_C359232:

        ' ValCSPDir>0?
        sCurrComponent = "ValCSPDir>0?"
        C359232 = (fn_ConvertValueCompiled(C359230, C359230DataType, AKB_DecimalPoint, False) > 0)
        C359232DataType = AKBTypeConst.cAKBTypeLogical
        If C359232 Then
            GoTo Comp_C359063
        Else
            GoTo Comp_C359233
        End If

    Comp_C359233:

        ' AtrValorvDT
        sCurrComponent = "AtrValorvDT"
        C359233DataType = 4
        C359231 = fn_ConvertValueCompiled(CDate("01/01/2100"), 2, AKB_DecimalPoint)
        C359233 = True
        GoTo Comp_C359235

    Comp_C359234:

        ' Controle
        sCurrComponent = "Controle"
        C359234 = 1
        C359234DataType = 1
        GoTo Comp_C359087

    Comp_C359235:

        ' AtrValorControle
        sCurrComponent = "AtrValorControle"
        C359235DataType = 4
        C359234 = fn_ConvertValueCompiled(2, 1, AKB_DecimalPoint)
        C359235 = True
        GoTo Comp_C359060

    Comp_C359236:

        ' Controle=2?
        sCurrComponent = "Controle=2?"
        C359236 = (fn_ConvertValueCompiled(C359234, C359234DataType, AKB_DecimalPoint, False) = 2)
        C359236DataType = AKBTypeConst.cAKBTypeLogical
        If C359236 Then
            GoTo Comp_C359063
        Else
            GoTo Comp_C359232
        End If

    Comp_C359238:

        ' MaxMin
        sCurrComponent = "MaxMin"
        C359238 = "MAX"
        C359238DataType = 5
        GoTo Comp_C359315

    Comp_C359239:

        ' UsaCtMax?
        sCurrComponent = "UsaCtMax?"
        C359239 = (fn_ConvertValueCompiled(RsC359315(0), C359315DataType(1), AKB_DecimalPoint, False) = 0)
        C359239DataType = AKBTypeConst.cAKBTypeLogical
        If C359239 Then
            GoTo Comp_C359125
        Else
            GoTo Comp_C359240
        End If

    Comp_C359240:

        ' AtrValorMaxMin
        sCurrComponent = "AtrValorMaxMin"
        C359240DataType = 4
        C359238 = fn_ConvertValueCompiled("MIN", 5, AKB_DecimalPoint)
        C359240 = True
        GoTo Comp_C359125

    Comp_C359246:

        ' CustoInsSP
        sCurrComponent = "CustoInsSP"
        C359246DataType = 0
        C359246 = RsC359062(0)
        C359246DataType = C359062Datatype(1)
        If C359246DataType = AKBTypeConst.cAKBTypeString Then
          C359246 = IIF(IsDBNull(C359246), C359246, Strings.RTrim(C359246))
        End If 

        GoTo Comp_C359064

    Comp_C359247:

        ' UnitInsSP
        sCurrComponent = "UnitInsSP"
        C359247 = (fn_ConvertValueCompiled(C359246, C359246DataType, AKB_DecimalPoint, True)/fn_ConvertValueCompiled(C359071, C359071DataType, AKB_DecimalPoint, True) ) *  fn_ConvertValueCompiled(C359304, C359304DataType, AKB_DecimalPoint, True)
        C359247DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359250

    Comp_C359248:

        ' CustoMaqSP
        sCurrComponent = "CustoMaqSP"
        C359248DataType = 0
        C359248DataType = C359062Datatype(4)
        If C359248DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359062(3)) Then
          C359248 = Strings.RTrim(RsC359062(3))
        Else 
          C359248 = RsC359062(3)
        End If 

        GoTo Comp_C359249

    Comp_C359249:

        ' CustoInsSCxSP
        sCurrComponent = "CustoInsSCxSP"
        C359249DataType = 0
        C359249DataType = C359062Datatype(5)
        If C359249DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359062(4)) Then
          C359249 = Strings.RTrim(RsC359062(4))
        Else 
          C359249 = RsC359062(4)
        End If 

        GoTo Comp_C359089

    Comp_C359250:

        ' UnitCMSP
        sCurrComponent = "UnitCMSP"
        C359250 = (fn_ConvertValueCompiled(C359248, C359248DataType, AKB_DecimalPoint, True)/fn_ConvertValueCompiled(C359071, C359071DataType, AKB_DecimalPoint, True) ) *  fn_ConvertValueCompiled(C359304, C359304DataType, AKB_DecimalPoint, True)
        C359250DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359251

    Comp_C359251:

        ' UnitCIsCXSP
        sCurrComponent = "UnitCIsCXSP"
        C359251 = (fn_ConvertValueCompiled(C359249, C359249DataType, AKB_DecimalPoint, True)/fn_ConvertValueCompiled(C359071, C359071DataType, AKB_DecimalPoint, True) ) *  fn_ConvertValueCompiled(C359304, C359304DataType, AKB_DecimalPoint, True)
        C359251DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359255

    Comp_C359253:

        ' SelPUtab
        sCurrComponent = "SelPUtab"
        QueryC359253 = con.CreateCommand()
        QueryC359253.CommandText = QueryC359253.CommandText & " " & "SELECT  NVL(MAX (C.Tabela_Preco_Custo) ,0),"
        QueryC359253.CommandText = QueryC359253.CommandText & " " & " NVL(MAX (CP.Preco),0)"
        QueryC359253.CommandText = QueryC359253.CommandText & " " & "FROM WF_TAB_CT_COMPR_REC_CT T, "
        QueryC359253.CommandText = QueryC359253.CommandText & " " & "WF_TAB_PRECO_CUSTO C, "
        QueryC359253.CommandText = QueryC359253.CommandText & " " & "WF_TAB_PRC_CUSTO_ACESSO CP"
        QueryC359253.CommandText = QueryC359253.CommandText & " " & "WHERE T.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P58928, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359253.CommandText = QueryC359253.CommandText & " " & " AND T.Tabela_Preco_Custo = C.Tabela_Preco_Custo"
        QueryC359253.CommandText = QueryC359253.CommandText & " " & " AND C.Data_Validade_Inicial <=" & _ 
ConvertValue(C359308, C359308DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359253.CommandText = QueryC359253.CommandText & " " & " AND (C.Data_Validade_Final >= " & _ 
ConvertValue(C359308, C359308DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  OR C.Data_Validade_Final IS NULL )"
        QueryC359253.CommandText = QueryC359253.CommandText & " " & " AND C.Tabela_Preco_Custo = CP.Tabela_Preco_Custo"
        QueryC359253.CommandText = QueryC359253.CommandText & " " & " AND CP.Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359253.CommandText = QueryC359253.CommandText & " " & " AND CP.Acesso = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359253.CommandText = QueryC359253.CommandText & " " & " AND CP.Cod_Embalagem = " & _ 
ConvertValue(C359131, C359131DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359253.CommandText = QueryC359253.CommandText & " " & ""
        QueryC359253.Transaction = txn
        RsC359253 = QueryC359253.ExecuteReader()
        Dim iC359253 As Short
        ReDim C359253DataType(RsC359253.FieldCount)
        For iC359253 = 0 to RsC359253.FieldCount - 1
            Select Case RsC359253.GetDataTypeName(iC359253).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359253DataType(iC359253 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359253DataType(iC359253 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359253DataType(iC359253 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359253
        RsC359253_EOF = Not RsC359253.Read()

        GoTo Comp_C359265

    Comp_C359254:

        ' PUtab=0?
        sCurrComponent = "PUtab=0?"
        C359254 = (fn_ConvertValueCompiled(C359266, C359266DataType, AKB_DecimalPoint, False) = 0)
        C359254DataType = AKBTypeConst.cAKBTypeLogical
        If C359254 Then
            GoTo Comp_C359348
        Else
            GoTo Comp_C359268
        End If

    Comp_C359255:

        ' vUnitCDSP
        sCurrComponent = "vUnitCDSP"
        C359255 = C359089 & " "
        C359255DataType = 1
        GoTo Comp_C359256

    Comp_C359256:

        ' vUnitCISP
        sCurrComponent = "vUnitCISP"
        C359256 = C359090 & " "
        C359256DataType = 1
        GoTo Comp_C359257

    Comp_C359257:

        ' vUnitInsSP
        sCurrComponent = "vUnitInsSP"
        C359257 = C359247 & " "
        C359257DataType = 1
        GoTo Comp_C359258

    Comp_C359258:

        ' vUnitCMSP
        sCurrComponent = "vUnitCMSP"
        C359258 = C359250 & " "
        C359258DataType = 1
        GoTo Comp_C359259

    Comp_C359259:

        ' vUnitCICXSP
        sCurrComponent = "vUnitCICXSP"
        C359259 = C359251 & " "
        C359259DataType = 1
        GoTo Comp_C359253

    Comp_C359260:

        ' AtrValorvUnitCDSP
        sCurrComponent = "AtrValorvUnitCDSP"
        C359260DataType = 4
        C359255 = fn_ConvertValueCompiled(C359266 , 1, AKB_DecimalPoint)
        C359260 = True
        GoTo Comp_C359261

    Comp_C359261:

        ' AtrValorvUnitCISP
        sCurrComponent = "AtrValorvUnitCISP"
        C359261DataType = 4
        C359256 = fn_ConvertValueCompiled(0, 1, AKB_DecimalPoint)
        C359261 = True
        GoTo Comp_C359262

    Comp_C359262:

        ' AtrValorvUnitInsSP
        sCurrComponent = "AtrValorvUnitInsSP"
        C359262DataType = 4
        C359257 = fn_ConvertValueCompiled(0, 1, AKB_DecimalPoint)
        C359262 = True
        GoTo Comp_C359263

    Comp_C359263:

        ' AtrValorvUnitCMSP
        sCurrComponent = "AtrValorvUnitCMSP"
        C359263DataType = 4
        C359258 = fn_ConvertValueCompiled(0, 1, AKB_DecimalPoint)
        C359263 = True
        GoTo Comp_C359264

    Comp_C359264:

        ' AtrValorvUnitCICXSP
        sCurrComponent = "AtrValorvUnitCICXSP"
        C359264DataType = 4
        C359259 = fn_ConvertValueCompiled(0, 1, AKB_DecimalPoint)
        C359264 = True
        GoTo Comp_C359348

    Comp_C359265:

        ' TabCT
        sCurrComponent = "TabCT"
        C359265DataType = 0
        C359265 = RsC359253(0)
        C359265DataType = C359253Datatype(1)
        If C359265DataType = AKBTypeConst.cAKBTypeString Then
          C359265 = IIF(IsDBNull(C359265), C359265, Strings.RTrim(C359265))
        End If 

        GoTo Comp_C359266

    Comp_C359266:

        ' PUtab
        sCurrComponent = "PUtab"
        C359266DataType = 0
        C359266DataType = C359253Datatype(2)
        If C359266DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359253(1)) Then
          C359266 = Strings.RTrim(RsC359253(1))
        Else 
          C359266 = RsC359253(1)
        End If 

        GoTo Comp_C359267

    Comp_C359267:

        ' FlagPUct
        sCurrComponent = "FlagPUct"
        C359267 = 0
        C359267DataType = 4
        GoTo Comp_C359254

    Comp_C359268:

        ' AtrValorFlagPUct
        sCurrComponent = "AtrValorFlagPUct"
        C359268DataType = 4
        C359267 = fn_ConvertValueCompiled(1, 4, AKB_DecimalPoint)
        C359268 = True
        GoTo Comp_C359260

    Comp_C359291:

        ' SelectPedidoItens
        sCurrComponent = "SelectPedidoItens"
        QueryC359291 = con.CreateCommand()
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "SELECT WF_PEDIDO.COD_CLIENTE,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO_ITENS.SEQ_ITEM,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO_ITENS.SIGLA_PROD,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO_ITENS.ACESSO,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO_ITENS.COD_EMBALAGEM,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO_ITENS.QTDE_PEDIDO,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO_ITENS.PRECO_UNIT,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO_ITENS.DT_GERACAO,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_ESTADOS.CODIGO_PAIS,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO.NAO_POSSUI_RED_ICMS,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_ESTADOS.ALIQUOTA_ICMS,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_ESTADOS.PORC_REDUCAO_ICMS"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "FROM WF_PEDIDO_ITENS,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PEDIDO,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_PESSOAS,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_CIDADES,"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "  WF_ESTADOS "
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "WHERE WF_PEDIDO_ITENS.PEDIDO = " & _ 
ConvertValue(P58946, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "AND WF_PEDIDO_ITENS.TP_PRODUTO =" & _ 
ConvertValue(P58947, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "AND WF_PESSOAS.COD_PESSOA = WF_PEDIDO.COD_CLIENTE"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "AND WF_PESSOAS.CODIGO_CIDADE = WF_CIDADES.CODIGO_CIDADE"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "AND WF_CIDADES.SIGLA_ESTADO = WF_ESTADOS.SIGLA_ESTADO"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "AND WF_PEDIDO.PEDIDO = WF_PEDIDO_ITENS.PEDIDO"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & "AND WF_PEDIDO.TP_PRODUTO = WF_PEDIDO_ITENS.TP_PRODUTO"
        QueryC359291.CommandText = QueryC359291.CommandText & " " & ""
        QueryC359291.CommandText = QueryC359291.CommandText & " " & ""
        QueryC359291.Transaction = txn
        RsC359291 = QueryC359291.ExecuteReader()
        Dim iC359291 As Short
        ReDim C359291DataType(RsC359291.FieldCount)
        For iC359291 = 0 to RsC359291.FieldCount - 1
            Select Case RsC359291.GetDataTypeName(iC359291).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359291DataType(iC359291 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359291DataType(iC359291 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359291DataType(iC359291 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359291
        RsC359291_EOF = Not RsC359291.Read()

        GoTo Comp_C359292

    Comp_C359292:

        ' EOF_SelectPedidoItens
        sCurrComponent = "EOF_SelectPedidoItens"
        C359292DataType = 4
        C359292 = RsC359291_EOF
        GoTo Comp_C359293

    Comp_C359293:

        ' EOF_SelectPedidoItens?
        sCurrComponent = "EOF_SelectPedidoItens?"
        C359293 = (fn_ConvertValueCompiled(C359292, C359292DataType, AKB_DecimalPoint, False)  = 1)
        C359293DataType = AKBTypeConst.cAKBTypeLogical
        If C359293 Then
            GoTo Comp_C359330
        Else
            GoTo Comp_C533250
        End If

    Comp_C359299:

        ' Cod_Cliente
        sCurrComponent = "Cod_Cliente"
        C359299DataType = 0
        C359299 = RsC359291(0)
        C359299DataType = C359291Datatype(1)
        If C359299DataType = AKBTypeConst.cAKBTypeString Then
          C359299 = IIF(IsDBNull(C359299), C359299, Strings.RTrim(C359299))
        End If 

        GoTo Comp_C359300

    Comp_C359300:

        ' Seq_Item
        sCurrComponent = "Seq_Item"
        C359300DataType = 0
        C359300DataType = C359291Datatype(2)
        If C359300DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(1)) Then
          C359300 = Strings.RTrim(RsC359291(1))
        Else 
          C359300 = RsC359291(1)
        End If 

        GoTo Comp_C359301

    Comp_C359301:

        ' Sigla_Prod
        sCurrComponent = "Sigla_Prod"
        C359301DataType = 0
        C359301DataType = C359291Datatype(3)
        If C359301DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(2)) Then
          C359301 = Strings.RTrim(RsC359291(2))
        Else 
          C359301 = RsC359291(2)
        End If 

        GoTo Comp_C359302

    Comp_C359302:

        ' Acesso
        sCurrComponent = "Acesso"
        C359302DataType = 0
        C359302DataType = C359291Datatype(4)
        If C359302DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(3)) Then
          C359302 = Strings.RTrim(RsC359291(3))
        Else 
          C359302 = RsC359291(3)
        End If 

        GoTo Comp_C359303

    Comp_C359303:

        ' Cod_Emb
        sCurrComponent = "Cod_Emb"
        C359303DataType = 0
        C359303DataType = C359291Datatype(5)
        If C359303DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(4)) Then
          C359303 = Strings.RTrim(RsC359291(4))
        Else 
          C359303 = RsC359291(4)
        End If 

        GoTo Comp_C359304

    Comp_C359304:

        ' Qtde_Ped
        sCurrComponent = "Qtde_Ped"
        C359304DataType = 0
        C359304DataType = C359291Datatype(6)
        If C359304DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(5)) Then
          C359304 = Strings.RTrim(RsC359291(5))
        Else 
          C359304 = RsC359291(5)
        End If 

        GoTo Comp_C359305

    Comp_C359305:

        ' Preço_Unit
        sCurrComponent = "Preço_Unit"
        C359305DataType = 0
        C359305DataType = C359291Datatype(7)
        If C359305DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(6)) Then
          C359305 = Strings.RTrim(RsC359291(6))
        Else 
          C359305 = RsC359291(6)
        End If 

        GoTo Comp_C359308

    Comp_C359308:

        ' Dt_Geração
        sCurrComponent = "Dt_Geração"
        C359308DataType = 0
        C359308DataType = C359291Datatype(8)
        If C359308DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(7)) Then
          C359308 = Strings.RTrim(RsC359291(7))
        Else 
          C359308 = RsC359291(7)
        End If 

        GoTo Comp_C359310

    Comp_C359310:

        ' Cod_Pais
        sCurrComponent = "Cod_Pais"
        C359310DataType = 0
        C359310DataType = C359291Datatype(9)
        If C359310DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(8)) Then
          C359310 = Strings.RTrim(RsC359291(8))
        Else 
          C359310 = RsC359291(8)
        End If 

        GoTo Comp_C359311

    Comp_C359311:

        ' Nao_Possui_Red_Icms
        sCurrComponent = "Nao_Possui_Red_Icms"
        C359311DataType = 0
        C359311DataType = C359291Datatype(10)
        If C359311DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(9)) Then
          C359311 = Strings.RTrim(RsC359291(9))
        Else 
          C359311 = RsC359291(9)
        End If 

        GoTo Comp_C359312

    Comp_C359312:

        ' Aliquota_Icms
        sCurrComponent = "Aliquota_Icms"
        C359312DataType = 0
        C359312DataType = C359291Datatype(11)
        If C359312DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(10)) Then
          C359312 = Strings.RTrim(RsC359291(10))
        Else 
          C359312 = RsC359291(10)
        End If 

        GoTo Comp_C359313

    Comp_C359313:

        ' Porc_Redução_Icms
        sCurrComponent = "Porc_Redução_Icms"
        C359313DataType = 0
        C359313DataType = C359291Datatype(12)
        If C359313DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359291(11)) Then
          C359313 = Strings.RTrim(RsC359291(11))
        Else 
          C359313 = RsC359291(11)
        End If 

        GoTo Comp_C359129

    Comp_C359315:

        ' CustoMin
        sCurrComponent = "CustoMin"
        QueryC359315 = con.CreateCommand()
        QueryC359315.CommandText = QueryC359315.CommandText & " " & "SELECT NVL(Utilizar_Custo_Min_Comp_Custo,0)"
        QueryC359315.CommandText = QueryC359315.CommandText & " " & " FROM WF_SIGLA_PROD_AGRUP"
        QueryC359315.CommandText = QueryC359315.CommandText & " " & " WHERE Sigla_Prod = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359315.Transaction = txn
        RsC359315 = QueryC359315.ExecuteReader()
        Dim iC359315 As Short
        ReDim C359315DataType(RsC359315.FieldCount)
        For iC359315 = 0 to RsC359315.FieldCount - 1
            Select Case RsC359315.GetDataTypeName(iC359315).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359315DataType(iC359315 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359315DataType(iC359315 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359315DataType(iC359315 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359315
        RsC359315_EOF = Not RsC359315.Read()

        GoTo Comp_C359239

    Comp_C359330:

        ' Retorno
        sCurrComponent = "Retorno"
        Dim tb_C359330 As DataTable = New DataTable()
        tb_C359330.Columns.Add("Ret" & "")
        Dim row_C359330 As DataRow = tb_C359330.NewRow()
        row_C359330(0) = C359205
        tb_C359330.Rows.Add(row_C359330)
        R16456 = tb_C359330
        ReDim C359330DataType(1)
        C359330DataType(1) = C359205DataType
        ReturnDataType = C359330DataType

        GoTo Exit_R16456

    Comp_C359348:

        ' SelTipoDesconto
        sCurrComponent = "SelTipoDesconto"
        QueryC359348 = con.CreateCommand()
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "SELECT NVL(MAX(WF_PED_SEQ_DESCONTO.TIPO_DESCONTO) ,0) ,"
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "  NVL(MAX(WF_PED_SEQ_DESCONTO.PORC_DESCONTO),0)"
        QueryC359348.CommandText = QueryC359348.CommandText & " " & ""
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "FROM WF_PED_SEQ_DESCONTO"
        QueryC359348.CommandText = QueryC359348.CommandText & " " & ""
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.PEDIDO   = " & _ 
ConvertValue(P58946, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.TP_PRODUTO = " & _ 
ConvertValue(P58947, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.SEQ   ="
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "  (SELECT MAX(WF_PED_SEQ_DESCONTO.SEQ )"
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "  FROM WF_PED_SEQ_DESCONTO"
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "  WHERE WF_PED_SEQ_DESCONTO.PEDIDO   = " & _ 
ConvertValue(P58946, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "  AND WF_PED_SEQ_DESCONTO.TP_PRODUTO = " & _ 
ConvertValue(P58947, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "  )"
        QueryC359348.CommandText = QueryC359348.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.TIPO_DESCONTO = 280 "
        QueryC359348.CommandText = QueryC359348.CommandText & " " & ""
        QueryC359348.Transaction = txn
        RsC359348 = QueryC359348.ExecuteReader()
        Dim iC359348 As Short
        ReDim C359348DataType(RsC359348.FieldCount)
        For iC359348 = 0 to RsC359348.FieldCount - 1
            Select Case RsC359348.GetDataTypeName(iC359348).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359348DataType(iC359348 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359348DataType(iC359348 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359348DataType(iC359348 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359348
        RsC359348_EOF = Not RsC359348.Read()

        GoTo Comp_C359351

    Comp_C359350:

        ' SelValorBrutoPUD
        sCurrComponent = "SelValorBrutoPUD"
        QueryC359350 = con.CreateCommand()
        QueryC359350.CommandText = QueryC359350.CommandText & " " & "SELECT  (WF_PEDIDO_ITENS.QTDE_PEDIDO * WF_PEDIDO_ITENS.PRECO_UNIT) VALOR_BRUTO ,"
        QueryC359350.CommandText = QueryC359350.CommandText & " " & "(WF_PEDIDO_ITENS.PRECO_SEM_DESCONTO * (1-(NVL(WF_PEDIDO_ITENS.DPUD,0)/100))) PUD"
        QueryC359350.CommandText = QueryC359350.CommandText & " " & "                       "
        QueryC359350.CommandText = QueryC359350.CommandText & " " & "FROM WF_PEDIDO_ITENS"
        QueryC359350.CommandText = QueryC359350.CommandText & " " & "WHERE WF_PEDIDO_ITENS.PEDIDO     = " & _ 
ConvertValue(P58946, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359350.CommandText = QueryC359350.CommandText & " " & "AND WF_PEDIDO_ITENS.TP_PRODUTO   = " & _ 
ConvertValue(P58947, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359350.CommandText = QueryC359350.CommandText & " " & "AND WF_PEDIDO_ITENS.SEQ_ITEM     = " & _ 
ConvertValue(C359300, C359300DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359350.Transaction = txn
        RsC359350 = QueryC359350.ExecuteReader()
        Dim iC359350 As Short
        ReDim C359350DataType(RsC359350.FieldCount)
        For iC359350 = 0 to RsC359350.FieldCount - 1
            Select Case RsC359350.GetDataTypeName(iC359350).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C359350DataType(iC359350 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C359350DataType(iC359350 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C359350DataType(iC359350 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC359350
        RsC359350_EOF = Not RsC359350.Read()

        GoTo Comp_C369726

    Comp_C359351:

        ' Tipo_Desconto
        sCurrComponent = "Tipo_Desconto"
        C359351DataType = 0
        C359351 = RsC359348(0)
        C359351DataType = C359348Datatype(1)
        If C359351DataType = AKBTypeConst.cAKBTypeString Then
          C359351 = IIF(IsDBNull(C359351), C359351, Strings.RTrim(C359351))
        End If 

        GoTo Comp_C359352

    Comp_C359352:

        ' Porc_Desconto
        sCurrComponent = "Porc_Desconto"
        C359352DataType = 0
        C359352DataType = C359348Datatype(2)
        If C359352DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359348(1)) Then
          C359352 = Strings.RTrim(RsC359348(1))
        Else 
          C359352 = RsC359348(1)
        End If 

        GoTo Comp_C359350

    Comp_C359358:

        ' AtualCustos
        sCurrComponent = "AtualCustos"
        QueryC359358 = con.CreateCommand()
        QueryC359358.CommandText = QueryC359358.CommandText & " " & "UPDATE WF_PEDIDO_ITENS "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ""
        QueryC359358.CommandText = QueryC359358.CommandText & " " & "SET WF_PEDIDO_ITENS.CUSTO_DIRETO_STD_PROD = " & _ 
ConvertValue(C359255, C359255DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.CUSTO_INDIRETO_STD_PROD =" & _ 
ConvertValue(C359256, C359256DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.CUSTO_INSUMO_STD_PROD = " & _ 
ConvertValue(C359257, C359257DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.CUSTO_MAQUINA_STD_PROD =" & _ 
ConvertValue(C359258, C359258DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.CUSTO_ITENS_SEM_DESEMB_CX = " & _ 
ConvertValue(C359259, C359259DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.ID_CUSTO = " & _ 
ConvertValue(C359137, C359137DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.SIGLA_PROD3 = " & _ 
ConvertValue(C359129, C359129DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.ACESSO3 = " & _ 
ConvertValue(C359130, C359130DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.SEQ_CUSTO = " & _ 
ConvertValue(RsC359061(0), C359061DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ""
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ",WF_PEDIDO_ITENS.VALOR_BRUTO = " & _ 
ConvertValue(C369726, C369726DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ", WF_PEDIDO_ITENS.PUD =" & _ 
ConvertValue(C369727, C369727DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & ""
        QueryC359358.CommandText = QueryC359358.CommandText & " " & "WHERE WF_PEDIDO_ITENS.PEDIDO =  " & _ 
ConvertValue(P58946, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & "AND WF_PEDIDO_ITENS.TP_PRODUTO = " & _ 
ConvertValue(P58947, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.CommandText = QueryC359358.CommandText & " " & "AND WF_PEDIDO_ITENS.SEQ_ITEM = " & _ 
ConvertValue(C359300, C359300DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359358.Transaction = txn
        C359358 = QueryC359358.ExecuteNonQuery()
        C359358DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359407

    Comp_C359359:

        ' CodICM
        sCurrComponent = "CodICM"
        C359359DataType = 0
        C359359DataType = C359178Datatype(3)
        If C359359DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359178(2)) Then
          C359359 = Strings.RTrim(RsC359178(2))
        Else 
          C359359 = RsC359178(2)
        End If 

        GoTo Comp_C359360

    Comp_C359360:

        ' CodIPI
        sCurrComponent = "CodIPI"
        C359360DataType = 0
        C359360DataType = C359178Datatype(4)
        If C359360DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359178(3)) Then
          C359360 = Strings.RTrim(RsC359178(3))
        Else 
          C359360 = RsC359178(3)
        End If 

        GoTo Comp_C359074

    Comp_C359406:

        ' Insere Pedido Itens x Impostos
        sCurrComponent = "Insere Pedido Itens x Impostos"
        C359406 = clsRuleDynamicallyCompiled_R16462.R16462(con, txn, IIf(Not IsDBNull(C359179), C359179, System.DBNull.Value), IIf(Not IsDBNull(C359359), C359359, System.DBNull.Value), IIf(Not IsDBNull(C359180), C359180, System.DBNull.Value), IIf(Not IsDBNull(C359360), C359360, System.DBNull.Value), IIf(Not IsDBNull(C359311), C359311, System.DBNull.Value), IIf(Not IsDBNull(C359310), C359310, System.DBNull.Value), IIf(Not IsDBNull(C359312), C359312, System.DBNull.Value), IIf(Not IsDBNull(C359313), C359313, System.DBNull.Value), IIf(Not IsDBNull(P58947), P58947, System.DBNull.Value), IIf(Not IsDBNull(P58946), P58946, System.DBNull.Value), IIf(Not IsDBNull(C359300), C359300, System.DBNull.Value), IIf(Not IsDBNull(C369726), C369726, System.DBNull.Value), IIf(Not IsDBNull(P58928), P58928, System.DBNull.Value), IIf(Not IsDBNull(C359308), C359308, System.DBNull.Value))
        C359406CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C359406) Then
          iColumns = C359406.Columns.Count
        End If
        ReDim C359406DataType(iColumns)
        For iCol = 1 To iColumns
            C359406DataType(iCol) = clsRuleDynamicallyCompiled_R16462.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C359034

    Comp_C359407:

        ' ValorItem
        sCurrComponent = "ValorItem"
        C359407 = fn_ConvertValueCompiled(C359305, C359305DataType, AKB_DecimalPoint, True) *fn_ConvertValueCompiled(C359304, C359304DataType, AKB_DecimalPoint, True)
        C359407DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359406

    Comp_C368769:

        ' NaoTemCusto?
        sCurrComponent = "NaoTemCusto?"
        C368769 = (fn_ConvertValueCompiled(C359137, C359137DataType, AKB_DecimalPoint, False)  = 0 OR fn_ConvertValueCompiled(RsC359061(0), C359061DataType(1), AKB_DecimalPoint, False)  = 0)
        C368769DataType = AKBTypeConst.cAKBTypeLogical
        If C368769 Then
            GoTo Comp_C368770
        Else
            GoTo Comp_C359358
        End If

    Comp_C368770:

        ' AtualPUDPreçoBruto
        sCurrComponent = "AtualPUDPreçoBruto"
        QueryC368770 = con.CreateCommand()
        QueryC368770.CommandText = QueryC368770.CommandText & " " & "UPDATE WF_PEDIDO_ITENS "
        QueryC368770.CommandText = QueryC368770.CommandText & " " & ""
        QueryC368770.CommandText = QueryC368770.CommandText & " " & "SET"
        QueryC368770.CommandText = QueryC368770.CommandText & " " & ""
        QueryC368770.CommandText = QueryC368770.CommandText & " " & "WF_PEDIDO_ITENS.VALOR_BRUTO = " & _ 
ConvertValue(C369726, C369726DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368770.CommandText = QueryC368770.CommandText & " " & ", WF_PEDIDO_ITENS.PUD =" & _ 
ConvertValue(C369727, C369727DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC368770.CommandText = QueryC368770.CommandText & " " & ""
        QueryC368770.CommandText = QueryC368770.CommandText & " " & "WHERE WF_PEDIDO_ITENS.PEDIDO =  " & _ 
ConvertValue(P58946, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368770.CommandText = QueryC368770.CommandText & " " & "AND WF_PEDIDO_ITENS.TP_PRODUTO = " & _ 
ConvertValue(P58947, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368770.CommandText = QueryC368770.CommandText & " " & "AND WF_PEDIDO_ITENS.SEQ_ITEM = " & _ 
ConvertValue(C359300, C359300DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368770.Transaction = txn
        C368770 = QueryC368770.ExecuteNonQuery()
        C368770DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C359407

    Comp_C369726:

        ' ValorBruto
        sCurrComponent = "ValorBruto"
        C369726DataType = 0
        C369726 = RsC359350(0)
        C369726DataType = C359350Datatype(1)
        If C369726DataType = AKBTypeConst.cAKBTypeString Then
          C369726 = IIF(IsDBNull(C369726), C369726, Strings.RTrim(C369726))
        End If 

        GoTo Comp_C369727

    Comp_C369727:

        ' PUD
        sCurrComponent = "PUD"
        C369727DataType = 0
        C369727DataType = C359350Datatype(2)
        If C369727DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC359350(1)) Then
          C369727 = Strings.RTrim(RsC359350(1))
        Else 
          C369727 = RsC359350(1)
        End If 

        GoTo Comp_C368769

    Comp_C533250:

        ' Banco
        sCurrComponent = "Banco"
        QueryC533250 = con.CreateCommand()
        QueryC533250.CommandText = QueryC533250.CommandText & " " & "select ora_database_name from dual"
        QueryC533250.Transaction = txn
        RsC533250 = QueryC533250.ExecuteReader()
        Dim iC533250 As Short
        ReDim C533250DataType(RsC533250.FieldCount)
        For iC533250 = 0 to RsC533250.FieldCount - 1
            Select Case RsC533250.GetDataTypeName(iC533250).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C533250DataType(iC533250 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C533250DataType(iC533250 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C533250DataType(iC533250 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC533250
        RsC533250_EOF = Not RsC533250.Read()

        GoTo Comp_C533251

    Comp_C533251:

        ' Dif_PROG?
        sCurrComponent = "Dif_PROG?"
        C533251 = (fn_ConvertValueCompiled(RsC533250(0), C533250DataType(1), AKB_DecimalPoint, False) <> "PROG")
        C533251DataType = AKBTypeConst.cAKBTypeLogical
        If C533251 Then
            GoTo Comp_C359299
        Else
            GoTo Comp_C359330
        End If

    Exit_R16456:

        Exit Function

    Err_R16456:
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
