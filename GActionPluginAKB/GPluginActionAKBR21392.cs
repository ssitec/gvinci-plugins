using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR21392 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0011-C827AF67A9C2";

        public override string Name => "R21392 - Valida Dias e % CF Risco Sac - varios pre-ped";

        public override string Description => "Valida Dias e % CF Risco Sac - varios pre-ped";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR21392(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 82023, Name = "Pre-Pedido", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 82022, Name = "Gerar", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 82024, Name = "CodPag", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 82027, Name = "CodCliente", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

                    };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {

            bool parametrosMultiplos = false;

            string idenStr = new string('\t', Identation);

            string[] redimStr = new string[0];
            string[] atribStr = new string[0];

            Builder.AppendLine(idenStr + "string connString = \"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=179.125.172.19)(PORT=1516))(CONNECT_DATA=(SERVICE_NAME=progtst)))\";");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleConnection conn = new System.Data.OracleClient.OracleConnection(\"Data Source=\" + connString + \";User Id=S2D-FERNANDO;Password=Gvinci2024\");");
            Builder.AppendLine(idenStr + "conn.Open();");

            // Não abre transação nova. Fernando 03/06/2023
            //Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = conn.BeginTransaction();");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = null;");

            // Fernando 22/10/24
            Builder.AppendLine(idenStr + "DataTable msg = new DataTable();");
            Builder.AppendLine(idenStr + "msg.Columns.Add(\"Mensagem\");");

            // parâmetro Pre-Pedido. Fernando 04/11/24
            string P82023 = "";
            IGPluginControl controlP82023 = this.Parameters[0].Value as IGPluginControl;

            if (controlP82023?.ControlType == "GCOMBOBOX")
            {
                P82023 += controlP82023.Name + ".Value";
            }
            else if (controlP82023?.ControlType == "GTEXTBOX")
            {
                P82023 += "Convert.ToDouble(" + controlP82023.Name + ".Text)";
            }
            else if (controlP82023?.ControlType == "GGRIDCOLUMN")
            {
                P82023 += "ref P82023";
                Builder.AppendLine(idenStr + "object[] P82023 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P82023, P82023.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P82023[P82023.Length - 1] = Convert.ToDouble(item.Cells[5].Text);";

                parametrosMultiplos = true;
            }
            else
            {
                P82023 = "var result";
            }

            // parâmetro Gerar. Fernando 04/11/24
            string P82022 = "";
            IGPluginControl controlP82022 = this.Parameters[1].Value as IGPluginControl;

            if (controlP82022?.ControlType == "GCOMBOBOX")
            {
                P82022 += controlP82022.Name + ".Value";
            }
            else if (controlP82022?.ControlType == "GTEXTBOX")
            {
                P82022 += "Convert.ToDouble(" + controlP82022.Name + ".Text)";
            }
            else if (controlP82022?.ControlType == "GGRIDCOLUMN")
            {
                P82022 += "ref P82022";
                Builder.AppendLine(idenStr + "object[] P82022 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P82022, P82022.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P82022[P82022.Length - 1] = ((CheckBox)item.Cells[3].Controls[1]).Checked ? 1 : 0;";

                parametrosMultiplos = true;
            }
            else
            {
                P82022 = "var result";
            }

            // parâmetro CodPag. Fernando 01/11/24
            string P82024 = "";
            IGPluginControl controlP82024 = this.Parameters[2].Value as IGPluginControl;

            if (controlP82024?.ControlType == "GCOMBOBOX")
            {
                P82024 += controlP82024.Name + ".Value";
            }
            else if (controlP82024?.ControlType == "GTEXTBOX")
            {
                P82024 += "Convert.ToDouble(" + controlP82024.Name + ".Text)";
            }
            else if (controlP82024?.ControlType == "GGRIDCOLUMN")
            {
                P82024 += "ref P82024";
                Builder.AppendLine(idenStr + "object[] P82024 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P82024, P82024.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 04/11/24
                atribStr[atribStr.Length - 1] = "P82024[P82024.Length - 1] = Convert.ToDouble(item.Cells[29].Text);";

                parametrosMultiplos = true;
            }
            else
            {
                P82024 = "var result";
            }

            // parâmetro CodCliente. Fernando 04/11/24
            string P82027 = "";
            IGPluginControl controlP82027 = this.Parameters[3].Value as IGPluginControl;

            if (controlP82027?.ControlType == "GCOMBOBOX")
            {
                P82027 += controlP82027.Name + ".Value";
            }
            else if (controlP82027?.ControlType == "GTEXTBOX")
            {
                P82027 += "Convert.ToDouble(" + controlP82027.Name + ".Text)";
            }
            else if (controlP82027?.ControlType == "GGRIDCOLUMN")
            {
                P82027 += "ref P82027";
                Builder.AppendLine(idenStr + "object[] P82027 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P82027, P82027.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P82027[P82027.Length - 1] = item.Cells[16].Text;";

                parametrosMultiplos = true;
            }
            else
            {
                P82027 = "var result";
            }

            if (parametrosMultiplos)
            {
                Builder.AppendLine(idenStr + "foreach (GridDataItem item in grdPrePedidos.Items)");
                Builder.AppendLine(idenStr + "{");

                foreach (string linha in redimStr)
                {
                    Builder.AppendLine(idenStr + "  " + linha);
                }

                foreach (string linha in atribStr)
                {
                    Builder.AppendLine(idenStr + "  " + linha);
                }

                Builder.AppendLine(idenStr + "}");
            }

            // Public Shared Function R21392(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByRef P82022() As Object, ByRef P82023() As Object, ByRef P82024() As Object, ByRef P82027() As Object) As DataTable
            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R21392.R21392(ref conn, ref trans, ref msg, " + P82022 + ", " + P82023 + ", " + P82024 + ", " + P82027 + ");");

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
                // Fernando 27/09/24
                new GPluginDependency() { FileName = "R21392_V0_D20241104_170946.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }


    }

}
