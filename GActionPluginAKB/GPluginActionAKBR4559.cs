using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR4559 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0005-C827AF67A9C2";

        public override string Name => "R4559 - Apaga Item do Pré Pedido";

        public override string Description => "Apaga Item do Pré Pedido";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR4559(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 12598, Name = "PréPedido", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 12599, Name = "SeqItem", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 12949, Name = "Transaction", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

                    };
        }
        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {

            // parâmetro PréPedido. Fernando 25/10/24
            string P12598 = "";
            IGPluginControl controlP12598 = this.Parameters[0].Value as IGPluginControl;

            if (controlP12598?.ControlType == "GCOMBOBOX")
            {
                P12598 += controlP12598.Name + ".Value";
            }
            else if (controlP12598?.ControlType == "GTEXTBOX")
            {
                P12598 += "Convert.ToDouble(" + controlP12598.Name + ".Text)";
            }
            else
            {
                P12598 = "var result";
            }

            // parâmetro SeqItem. Fernando 25/10/24
            string P12599 = "";
            IGPluginControl controlP12599 = this.Parameters[1].Value as IGPluginControl;

            if (controlP12599?.ControlType == "GCOMBOBOX")
            {
                P12599 += controlP12599.Name + ".Value";
            }
            else if (controlP12599?.ControlType == "GTEXTBOX")
            {
                P12599 += controlP12599.Name + "Field";
            }
            else
            {
                P12599 = "var result";
            }

            // parâmetro Transaction. Fernando 25/10/24
            string P12949 = "";
            IGPluginControl controlP12949 = this.Parameters[2].Value as IGPluginControl;

            if (controlP12949?.ControlType == "GCOMBOBOX")
            {
                P12949 += controlP12949.Name + ".Value";
            }
            else if (controlP12949?.ControlType == "GTEXTBOX")
            {
                P12949 += controlP12949.Name + "Field != \"\" ? " + controlP12949.Name + "Field : Convert.DBNull";
            }

            else
            {
                P12949 = "var result";
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

            // Public Shared Function R4559(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P12598 As Double, ByVal P12599 As Double, ByVal P12949 As Object) As DataTable
            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R4559.R4559(ref conn, ref trans, ref msg, " + P12598 + ", " + P12599 + ", " + P12949 + ");");

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
            Builder.AppendLine(idenStr + "    Itens.Rebind();");
            Builder.AppendLine(idenStr + "}");

        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                // Fernando 27/09/24
                new GPluginDependency() { FileName = "R4559_V0_D20241024_100008.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }
    }
}
