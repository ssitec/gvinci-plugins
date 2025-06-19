using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR21393 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0006-C827AF67A9C2";

        public override string Name => "R21393 - Valida Dias e % CF Risco Sac - 1 pre-pedido";

        public override string Description => "Valida Dias e % CF Risco Sac - 1 pre-pedido";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR21393(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 82034, Name = "PrePedido", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 82032, Name = "CodPag", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 82033, Name = "CodCliente", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

                    };
        }
        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {

            // parâmetro PréPedido. Fernando 25/10/24
            string P82034 = "";
            IGPluginControl controlP82034 = this.Parameters[0].Value as IGPluginControl;

            if (controlP82034?.ControlType == "GCOMBOBOX")
            {
                P82034 += controlP82034.Name + ".Value";
            }
            else if (controlP82034?.ControlType == "GTEXTBOX")
            {
                P82034 += "Convert.ToDouble(" + controlP82034.Name + ".Text)";
            }
            else
            {
                P82034 = "var result";
            }

            // parâmetro CodPag. Fernando 25/10/24
            string P82032 = "";
            IGPluginControl controlP82032 = this.Parameters[1].Value as IGPluginControl;

            if (controlP82032?.ControlType == "GCOMBOBOX")
            {
                P82032 += controlP82032.Name + ".Value";
            }
            else if (controlP82032?.ControlType == "GTEXTBOX")
            {
                P82032 += controlP82032.Name + "Field";
            }
            else
            {
                P82032 = "var result";
            }

            // parâmetro CodPag. Fernando 25/10/24
            string P82033 = "";
            IGPluginControl controlP82033 = this.Parameters[2].Value as IGPluginControl;

            if (controlP82033?.ControlType == "GCOMBOBOX")
            {
                P82033 += controlP82033.Name + ".Value";
            }
            else if (controlP82033.ControlType == "GTEXTBOX")
            {
                P82033 += controlP82033.Name + "Field";
            }
            else
            {
                P82033 = "var result";
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

            // Public Shared Function R21393(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P82032 As Double, ByVal P82033 As Double, ByVal P82034 As Double) As DataTable
            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R21393.R21393(ref conn, ref trans, ref msg, " + P82032 + ", " + P82033 + ", " + P82034 + ");");

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

            Builder.AppendLine(idenStr + "if (Convert.ToDouble(ret.Rows[0][0]) == 0)");
            Builder.AppendLine(idenStr + "{");
            Builder.AppendLine(idenStr + "    ActionSucceeded_1 = false;");
            Builder.AppendLine(idenStr + "}");
        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                // Fernando 25/10/24
                new GPluginDependency() { FileName = "R21393_V0_D20241025_200524.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }
    }

}
