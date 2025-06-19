using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR4554 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0004-C827AF67A9C2";

        public override string Name => "R4554";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR4554(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 12594, Name = "PréPedido", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 20448, Name = "Trans", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 24036, Name = "Dt_Pedido", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 24035, Name = "Cliente", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 24037, Name = "Representante", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 92047, Name = "PMotivo", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

                    };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {

            // parâmetro PréPedido. Fernando 24/10/24
            string P12594 = "";
            IGPluginControl controlP12594 = this.Parameters[0].Value as IGPluginControl;

            if (controlP12594?.ControlType == "GCOMBOBOX")
            {
                P12594 += controlP12594.Name + ".Value";
            }
            else if (controlP12594?.ControlType == "GTEXTBOX")
            {
                P12594 += "Convert.ToDouble(" + controlP12594.Name + ".Text)";
            }
            else
            {
                P12594 = "var result";
            }

            // parâmetro Trans. Fernando 24/10/24
            string P20448 = "";
            IGPluginControl controlP20448 = this.Parameters[1].Value as IGPluginControl;

            if (controlP20448?.ControlType == "GCOMBOBOX")
            {
                P20448 += controlP20448.Name + ".Value";
            }
            else if (controlP20448?.ControlType == "GTEXTBOX")
            {
                P20448 += controlP20448.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP20448.Name + ".Text) : Convert.DBNull";
            }
            else
            {
                P20448 = "var result";
            }

            // parâmetro Dt_Pedido. Fernando 24/10/24
            string P24036 = "";
            IGPluginControl controlP24036 = this.Parameters[2].Value as IGPluginControl;

            if (controlP24036?.ControlType == "GCOMBOBOX")
            {
                P24036 += controlP24036.Name + ".Value";
            }
            else if (controlP24036?.ControlType == "GTEXTBOX")
            {
                P24036 += "Convert.ToDateTime(" + controlP24036.Name + ".Text)";
            }
            
            else
            {
                P24036 = "var result";
            }

            // parâmetro Cliente. Fernando 24/10/24
            string P24035 = "";
            IGPluginControl controlP24035 = this.Parameters[3].Value as IGPluginControl;

            if (controlP24035?.ControlType == "GCOMBOBOX")
            {
                P24035 += controlP24035.Name + ".Value";
            }
            else if (controlP24035?.ControlType == "GTEXTBOX")
            {
                P24035 += controlP24035.Name + "Field";
            }
            else
            {
                P24035 = "var result";
            }

            // parâmetro Representante. Fernando 24/10/24
            string P24037 = "";
            IGPluginControl controlP24037 = this.Parameters[4].Value as IGPluginControl;

            if (controlP24037?.ControlType == "GCOMBOBOX")
            {
                P24037 += controlP24037.Name + ".Value";
            }
            else if (controlP24037?.ControlType == "GTEXTBOX")
            {
                P24037 += "Convert.ToDouble(" + controlP24037.Name + ".Text)";
            }
            else
            {
                P24037 = "var result";
            }

            // parâmetro PMotivo. Fernando 24/10/24
            string P92047 = "";
            IGPluginControl controlP92047 = this.Parameters[5].Value as IGPluginControl;

            if (controlP92047?.ControlType == "GCOMBOBOX")
            {
                P92047 += controlP92047.Name + ".Value";
            }
            else if (controlP92047?.ControlType == "GTEXTBOX")
            {
                P92047 += controlP92047.Name + ".Text != \"\" ? " + "" + controlP92047.Name + ".Text : Convert.DBNull";
            }
            else
            {
                P92047 = "var result";
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

            // ByVal P12594 As Double, ByVal P20448 As Object, ByVal P24035 As Double, ByVal P24036 As Object, ByVal P24037 As Double, ByVal P92047 As Object
            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R4554.R4554(ref conn, ref trans, ref msg, " + P12594 + ", " + P20448 + ", " + P24035 + ", " + P24036 + ", " + P24037 + ", " + P92047 + ");");
            
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
                new GPluginDependency() { FileName = "R4554_V0_D20241024_114751.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }


    }

}