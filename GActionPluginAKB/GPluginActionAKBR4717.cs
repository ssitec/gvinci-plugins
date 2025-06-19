using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR4717 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0008-C827AF67A9C2";

        public override string Name => "R4717 - Grava Alterações no Pré-Pedido";

        public override string Description => "Grava Alterações no Pré-Pedido";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR4717(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 13143, Name = "IdPréPedido", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 13144, Name = "Cliente", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 13419, Name = "Pagto", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 13420, Name = "Depto", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 33305, Name = "Desc OBS", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 31029, Name = "Depto_OBS", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

                    };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {

            // parâmetro IdPréPedido. Fernando 30/10/24
            string P13143 = "";
            IGPluginControl controlP13143 = this.Parameters[0].Value as IGPluginControl;

            if (controlP13143?.ControlType == "GCOMBOBOX")
            {
                P13143 += controlP13143.Name + ".Value";
            }
            else if (controlP13143?.ControlType == "GTEXTBOX")
            {
                P13143 += "Convert.ToDouble(" + controlP13143.Name + ".Text)";
            }
            else
            {
                P13143 = "var result";
            }

            // parâmetro Cliente. Fernando 30/10/24
            string P13144 = "";
            IGPluginControl controlP13144 = this.Parameters[1].Value as IGPluginControl;

            if (controlP13144?.ControlType == "GCOMBOBOX")
            {
                P13144 += controlP13144.Name + ".Value";
            }
            else if (controlP13144?.ControlType == "GTEXTBOX")
            {
                P13144 += controlP13144.Name + "Field";
            }
            else
            {
                P13144 = "var result";
            }

            // parâmetro Pagto. Fernando 30/10/24
            string P13419 = "";
            IGPluginControl controlP13419 = this.Parameters[2].Value as IGPluginControl;

            if (controlP13419?.ControlType == "GCOMBOBOX")
            {
                P13419 += controlP13419.Name + ".Value";
            }
            else if (controlP13419?.ControlType == "GTEXTBOX")
            {
                P13419 += controlP13419.Name + ".Text != \"\" ? " + controlP13419.Name + "Field : Convert.DBNull";
            }
            else
            {
                P13419 = "var result";
            }

            // parâmetro Depto. Fernando 30/10/24
            string P13420 = "";
            IGPluginControl controlP13420 = this.Parameters[3].Value as IGPluginControl;

            if (controlP13420?.ControlType == "GCOMBOBOX")
            {
                P13420 += controlP13420.Name + ".Value";
            }
            else if (controlP13420?.ControlType == "GTEXTBOX")
            {
                P13420 += controlP13420.Name + "Field != \"\" ? " + controlP13420.Name + "Field : Convert.DBNull";
            }
            else
            {
                P13420 = "var result";
            }

            // parâmetro Desc OBS. Fernando 30/10/24
            string P33305 = "";
            IGPluginControl controlP33305 = this.Parameters[4].Value as IGPluginControl;

            if (controlP33305?.ControlType == "GCOMBOBOX")
            {
                P33305 += controlP33305.Name + ".Value";
            }
            else if (controlP33305?.ControlType == "GTEXTBOX")
            {
                P33305 += controlP33305.Name + "Field != \"\" ? " + controlP33305.Name + "Field : Convert.DBNull";
            }
            else
            {
                P33305 = "var result";
            }

            // parâmetro Depto_OBS. Fernando 30/10/24
            string P31029 = "";
            IGPluginControl controlP31029 = this.Parameters[5].Value as IGPluginControl;

            if (controlP31029?.ControlType == "GCOMBOBOX")
            {
                P31029 += controlP31029.Name + ".Value";
            }
            else if (controlP31029?.ControlType == "GTEXTBOX")
            {
                P31029 += controlP31029.Name + "Field != \"\" ? " + controlP31029.Name + "Field : Convert.DBNull";
            }
            else
            {
                P31029 = "var result";
            }

            string idenStr = new string('\t', Identation);

            Builder.AppendLine(idenStr + "string connString = \"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=179.125.172.19)(PORT=1516))(CONNECT_DATA=(SERVICE_NAME=progtst)))\";");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleConnection conn = new System.Data.OracleClient.OracleConnection(\"Data Source=\" + connString + \";User Id=akbuser01;Password=Betania1200\");");
            Builder.AppendLine(idenStr + "conn.Open();");

            // Não abre transação nova. Fernando 03/06/2023
            //Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = conn.BeginTransaction();");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = null;");

            // Fernando 22/10/24
            Builder.AppendLine(idenStr + "DataTable msg = new DataTable();");
            Builder.AppendLine(idenStr + "msg.Columns.Add(\"Mensagem\");");

            // Public Shared Function R4717(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P13143 As Double, ByVal P13144 As Double, ByVal P13419 As Object, ByVal P13420 As Object, ByVal P31029 As Object, ByVal P33305 As Object) As DataTable
            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R4717.R4717(ref conn, ref trans, ref msg, " + P13143 + ", " + P13144 + ", " + P13419 + ", " + P13420 + ", " + P31029 + ", " + P33305 + ");");

            //Builder.AppendLine(idenStr + "trans.Commit();");
            Builder.AppendLine(idenStr + "conn.Close();");

            // Fernando 22/10/24
            Builder.AppendLine(idenStr + "if (msg.Rows.Count > 0)");
            Builder.AppendLine(idenStr + "{");

            Builder.AppendLine(idenStr + "  for (int i = 0; i < msg.Rows.Count; i++)");
            Builder.AppendLine(idenStr + "  {");
            Builder.AppendLine(idenStr + "        PageErrors.Add(\"Mensagem\", Convert.ToString(msg.Rows[i][0]));");
            Builder.AppendLine(idenStr + "  }");

            Builder.AppendLine(idenStr + "    ShowErrors();");
            Builder.AppendLine(idenStr + "}");

            Builder.AppendLine(idenStr + "if (Convert.ToDouble(ret.Rows[0][0]) == 1)");
            Builder.AppendLine(idenStr + "{");
            Builder.AppendLine(idenStr + "    grdPrePedidos.Rebind();");
            Builder.AppendLine(idenStr + "}");
        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                // Fernando 27/09/24
                new GPluginDependency() { FileName = "R4717_V0_D20241030_173400.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }


    }

}