using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR16268 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0012-C827AF67A9C2";

        public override string Name => "R16268 - Gera Pedido e Carrega Grid Automático";

        public override string Description => "Gera Pedido e Carrega Grid Automático";

        private List<GPluginActionParameter> _Parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR16268(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 70895, Name = "Dpto", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 70897, Name = "MarcarDpto", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 58159, Name = "Estab", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 70896, Name = "Celula", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

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

            // parâmetro Depto. Fernando 05/11/24
            string P70895 = "";
            IGPluginControl controlP70895 = this.Parameters[0].Value as IGPluginControl;

            if (controlP70895?.ControlType == "GCOMBOBOX")
            {
                P70895 += controlP70895.Name + ".Value";
            }
            else if (controlP70895?.ControlType == "GTEXTBOX")
            {
                P70895 += "Convert.ToDouble(" + controlP70895.Name + ".Text)";
            }
            else if (controlP70895?.ControlType == "GGRIDCOLUMN")
            {
                P70895 += "ref P70895";
                Builder.AppendLine(idenStr + "object[] P70895 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P70895, P70895.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P70895[P70895.Length - 1] = item.Cells[3].Text;";

                parametrosMultiplos = true;
            }
            else
            {
                P70895 = "var result";
            }

            // parâmetro MarcarDpto. Fernando 05/11/24
            string P70897 = "";
            IGPluginControl controlP70897 = this.Parameters[1].Value as IGPluginControl;

            if (controlP70897?.ControlType == "GCOMBOBOX")
            {
                P70897 += controlP70897.Name + ".Value";
            }
            else if (controlP70897?.ControlType == "GTEXTBOX")
            {
                P70897 += "Convert.ToDouble(" + controlP70897.Name + ".Text)";
            }
            else if (controlP70897?.ControlType == "GGRIDCOLUMN")
            {
                P70897 += "ref P70897";
                Builder.AppendLine(idenStr + "object[] P70897 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P70897, P70897.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P70897[P70897.Length - 1] = ((CheckBox)item.Cells[2].Controls[1]).Checked ? 1 : 0;";

                parametrosMultiplos = true;
            }
            else
            {
                P70897 = "var result";
            }

            // parâmetro Estab. Fernando 05/11/24
            string P58159 = "";
            IGPluginControl controlP58159 = this.Parameters[2].Value as IGPluginControl;

            if (controlP58159?.ControlType == "GCOMBOBOX")
            {
                P58159 += controlP58159.Name + ".Value";
            }
            else if (controlP58159?.ControlType == "GTEXTBOX")
            {
                P58159 += controlP58159.Name + ".Text";
            }
            else if (controlP58159?.ControlType == "GGRIDCOLUMN")
            {
                P58159 += "ref P58159";
                Builder.AppendLine(idenStr + "object[] P58159 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P58159, P58159.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 04/11/24
                atribStr[atribStr.Length - 1] = "P58159[P58159.Length - 1] = Convert.ToDouble(item.Cells[29].Text);";

                parametrosMultiplos = true;
            }
            else
            {
                P58159 = "var result";
            }

            // parâmetro Celula. Fernando 04/11/24
            string P70896 = "";
            IGPluginControl controlP70896 = this.Parameters[3].Value as IGPluginControl;

            if (controlP70896?.ControlType == "GCOMBOBOX")
            {
                P70896 += controlP70896.Name + ".Value";
            }
            else if (controlP70896?.ControlType == "GTEXTBOX")
            {
                P70896 += controlP70896.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP70896.Name + ".Text) : Convert.DBNull";
            }
            else if (controlP70896?.ControlType == "GGRIDCOLUMN")
            {
                P70896 += "ref P70896";
                Builder.AppendLine(idenStr + "object[] P70896 = new object[1];");

                Array.Resize(ref redimStr, redimStr.Length + 1);
                redimStr[redimStr.Length - 1] = "Array.Resize(ref P70896, P70896.Length + 1);";

                Array.Resize(ref atribStr, atribStr.Length + 1);
                // TO_DO: trocar número fixo da coluna. Fernando 01/11/24
                atribStr[atribStr.Length - 1] = "P70896[P70896.Length - 1] = item.Cells[16].Text;";

                parametrosMultiplos = true;
            }
            else
            {
                P70896 = "var result";
            }

            if (parametrosMultiplos)
            {
                Builder.AppendLine(idenStr + "foreach (GridDataItem item in GridDeptos.Items)");
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

            // Public Shared Function R16268(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P58159 As String, ByRef P70895() As Object, ByVal P70896 As Object, ByRef P70897() As Object) As DataTable
            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R16268.R16268(ref conn, ref trans, ref msg, " + P58159 + ", " + P70895 + ", " + P70896 + ", " + P70897 + ");");

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
        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                // Fernando 27/09/24
                new GPluginDependency() { FileName = "R16268_V0_D20241106_152912.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }


    }

}

