using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR4555 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0009-C827AF67A9C2";

        public override string Name => "R4555 - Gerar Pedido do Pré Pedido - Múltiplos";

        public override string Description => "Gerar Pedido do Pré Pedido - Múltiplos";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR4555(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 12595, Name = "PréPedidos", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 28838, Name = "Gerar", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 28839, Name = "Proforma", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 68165, Name = "Depto", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 76805, Name = "CodEstab", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 89541, Name = "SintegraCliente", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 89540, Name = "SintegraRedesp", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

                    };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {

            bool parametrosMultiplos = false;

            string idenStr = new string('\t', Identation);

            string[] redimStr = new string[0];
            string[] atribStr = new string[0];

            Builder.AppendLine(idenStr + "string connString = \"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=179.125.172.19)(PORT=1516))(CONNECT_DATA=(SERVICE_NAME=progtst)))\";");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleConnection conn = new System.Data.OracleClient.OracleConnection(\"Data Source=\" + connString + \";User Id=akbuser01;Password=Betania1200\");");
            Builder.AppendLine(idenStr + "conn.Open();");

            // Não abre transação nova. Fernando 03/06/2023
            //Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = conn.BeginTransaction();");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = null;");

            // Fernando 22/10/24
            Builder.AppendLine(idenStr + "DataTable msg = new DataTable();");
            Builder.AppendLine(idenStr + "msg.Columns.Add(\"Mensagem\");");

            // parâmetro PréPedidos. Fernando 01/11/24
            string P12595 = "";
            IGPluginControl controlP12595 = this.Parameters[0].Value as IGPluginControl;

            if (controlP12595?.ControlType == "GCOMBOBOX")
            {
                P12595 += controlP12595.Name + ".Value";
            }
            else if (controlP12595?.ControlType == "GTEXTBOX")
            {
                P12595 += "Convert.ToDouble(" + controlP12595.Name + ".Text)";
            }
            else if (controlP12595?.ControlType == "GGRIDCOLUMN")
            {
                P12595 += "ref P12595";
                Builder.AppendLine(idenStr + "object[] P12595 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P12595, P12595.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P12595[P12595.Length - 1] = Convert.ToDouble(item.Cells[5].Text);";

                parametrosMultiplos = true;
            }
            else
            {
                P12595 = "var result";
            }

            // parâmetro Gerar. Fernando 01/11/24
            string P28838 = "";
            IGPluginControl controlP28838 = this.Parameters[1].Value as IGPluginControl;

            if (controlP28838?.ControlType == "GCOMBOBOX")
            {
                P28838 += controlP28838.Name + ".Value";
            }
            else if (controlP28838?.ControlType == "GTEXTBOX")
            {
                P28838 += "Convert.ToDouble(" + controlP28838.Name + ".Text)";
            }
            else if (controlP28838?.ControlType == "GGRIDCOLUMN")
            {
                P28838 += "ref P28838";
                Builder.AppendLine(idenStr + "object[] P28838 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P28838, P28838.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P28838[P28838.Length - 1] = ((CheckBox)item.Cells[3].Controls[1]).Checked ? 1 : 0;";

                parametrosMultiplos = true;
            }
            else
            {
                P28838 = "var result";
            }

            // parâmetro Proforma. Fernando 01/11/24
            string P28839 = "";
            IGPluginControl controlP28839 = this.Parameters[2].Value as IGPluginControl;

            if (controlP28839?.ControlType == "GCOMBOBOX")
            {
                P28839 += controlP28839.Name + ".Value";
            }
            else if (controlP28839?.ControlType == "GTEXTBOX")
            {
                P28839 += "Convert.ToDouble(" + controlP28839.Name + ".Text)";
            }
            else if (controlP28839?.ControlType == "GGRIDCOLUMN")
            {
                P28839 += "ref P28839";
                Builder.AppendLine(idenStr + "object[] P28839 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P28839, P28839.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P28839[P28839.Length - 1] = ((CheckBox)item.Cells[13].Controls[0]).Checked ? 1 : 0;";

                parametrosMultiplos = true;
            }
            else
            {
                P28839 = "var result";
            }

            // parâmetro Depto. Fernando 01/11/24
            string P68165 = "";
            IGPluginControl controlP68165 = this.Parameters[3].Value as IGPluginControl;

            if (controlP68165?.ControlType == "GCOMBOBOX")
            {
                P68165 += controlP68165.Name + ".Value";
            }
            else if (controlP68165?.ControlType == "GTEXTBOX")
            {
                P68165 += "Convert.ToDouble(" + controlP68165.Name + ".Text)";
            }
            else if (controlP68165?.ControlType == "GGRIDCOLUMN")
            {
                P68165 += "ref P68165";
                Builder.AppendLine(idenStr + "object[] P68165 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P68165, P68165.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P68165[P68165.Length - 1] = item.Cells[33].Text;";

                parametrosMultiplos = true;
            }
            else
            {
                P68165 = "var result";
            }

            // parâmetro CodEstab. Fernando 03/10/24
            string P76805 = "";
            IGPluginControl controlP76805 = this.Parameters[4].Value as IGPluginControl;

            if (controlP76805?.ControlType == "GCOMBOBOX")
            {
                P76805 += controlP76805.Name + ".Value";
            }
            else if (controlP76805?.ControlType == "GTEXTBOX")
            {
                P76805 += "Convert.ToDouble(" + controlP76805.Name + ".Text)";
            }
            else
            {
                P76805 = "var result";
            }

            // parâmetro SintegraCliente. Fernando 01/11/24
            string P89541 = "";
            IGPluginControl controlP89541 = this.Parameters[3].Value as IGPluginControl;

            if (controlP89541?.ControlType == "GCOMBOBOX")
            {
                P89541 += controlP89541.Name + ".Value";
            }
            else if (controlP89541?.ControlType == "GTEXTBOX")
            {
                P89541 += "Convert.ToDouble(" + controlP89541.Name + ".Text)";
            }
            else if (controlP89541?.ControlType == "GGRIDCOLUMN")
            {
                P89541 += "ref P89541";
                Builder.AppendLine(idenStr + "object[] P89541 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P89541, P89541.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P89541[P89541.Length - 1] = ((CheckBox)item.Cells[6].Controls[0]).Checked ? 1 : 0;";

                parametrosMultiplos = true;
            }
            else
            {
                P89541 = "var result";
            }

            // parâmetro SintegraRedesp. Fernando 01/11/24
            string P89540 = "";
            IGPluginControl controlP89540 = this.Parameters[3].Value as IGPluginControl;

            if (controlP89540?.ControlType == "GCOMBOBOX")
            {
                P89540 += controlP89540.Name + ".Value";
            }
            else if (controlP89540?.ControlType == "GTEXTBOX")
            {
                P89540 += "Convert.ToDouble(" + controlP89540.Name + ".Text)";
            }
            else if (controlP89540?.ControlType == "GGRIDCOLUMN")
            {
                P89540 += "ref P89540";
                Builder.AppendLine(idenStr + "object[] P89540 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P89540, P89540.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P89540[P89540.Length - 1] = ((CheckBox)item.Cells[7].Controls[0]).Checked ? 1 : 0;";

                parametrosMultiplos = true;
            }
            else
            {
                P89540 = "var result";
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

            // Public Shared Function R4555(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByRef P12595() As Object, ByRef P28838() As Object, ByRef P28839() As Object, ByRef P68165() As Object, ByVal P76805 As Double, ByRef P89540() As Object, ByRef P89541() As Object) As DataTable
            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R4555.R4555(ref conn, ref trans, ref msg, " + P12595 + ", " + P28838 + ", " + P28839 + ", " + P68165 + ", " + P76805 + ", " + P89540 + ", " + P89541 + ");");

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

            Builder.AppendLine(idenStr + "if (Convert.ToBoolean(ret.Rows[0][0]))");
            Builder.AppendLine(idenStr + "{");
            Builder.AppendLine(idenStr + "    grdPrePedidos.Rebind();");
            Builder.AppendLine(idenStr + "}");
        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                // Fernando 27/09/24
                new GPluginDependency() { FileName = "R4555_V0_D20221228_105132.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }


    }

}
