using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR6662 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0007-C827AF67A9C2";

        public override string Name => "R6662 - Retirar liberação do Pré-Pedido";

        public override string Description => "Retirar liberação do Pré-Pedido";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR6662(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 20170, Name = "ID", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

                    };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {

            // parâmetro ID. Fernando 30/10/24
            string P20170 = "";
            IGPluginControl controlP20170 = this.Parameters[0].Value as IGPluginControl;

            if (controlP20170?.ControlType == "GCOMBOBOX")
            {
                P20170 += controlP20170.Name + ".Value";
            }
            else if (controlP20170?.ControlType == "GTEXTBOX")
            {
                P20170 += "Convert.ToDouble(" + controlP20170.Name + ".Text)";
            }
            else
            {
                P20170 = "var result";
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

            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R6662.R6662(ref conn, ref trans, ref msg, " + P20170 + ");");

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
                new GPluginDependency() { FileName = "R6662_V0_D20241030_151351.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }


    }

}