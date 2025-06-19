using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR241 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0010-C827AF67A9C2";

        public override string Name => "R241 - Retorna Código Estabelecimento Default";

        public override string Description => "Retorna Código Estabelecimento Default";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR241(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {

                    };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {


            string idenStr = new string('\t', Identation);

            Builder.AppendLine(idenStr + "string connString = \"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=179.125.172.19)(PORT=1516))(CONNECT_DATA=(SERVICE_NAME=progtst)))\";");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleConnection conn = new System.Data.OracleClient.OracleConnection(\"Data Source=\" + connString + \";User Id=S2D-FERNANDO;Password=Gvinci2024\");");
            Builder.AppendLine(idenStr + "conn.Open();");

            // Não abre transação nova. Fernando 03/06/2023
            //Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = conn.BeginTransaction();");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = null;");

            // Fernando 22/10/24
            Builder.AppendLine(idenStr + "DataTable msg = new DataTable();");
            Builder.AppendLine(idenStr + "msg.Columns.Add(\"Mensagem\");");

            // Public Shared Function R241(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable) As DataTable
            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R241.R241(ref conn, ref trans, ref msg);");

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

            Builder.AppendLine(idenStr + "if (ret.Rows.Count > 0)");
            Builder.AppendLine(idenStr + "{");
            Builder.AppendLine(idenStr + "    txtCodEstab.Text = Convert.ToString(ret.Rows[0][0]);");
            Builder.AppendLine(idenStr + "    txtEstabNome.Text = Convert.ToString(ret.Rows[0][1]);");
            Builder.AppendLine(idenStr + "}");

        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                // Fernando 27/09/24
                new GPluginDependency() { FileName = "R241_V0_D20191202_171620.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }


    }

}
