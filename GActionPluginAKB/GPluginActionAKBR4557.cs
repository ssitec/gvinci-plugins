using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR4557 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0003-C827AF67A9C2";

        public override string Name => "R4557";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR4557(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 12596, Name = "PréPedido", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 20438, Name = "NaoAbre_Trans", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 20439, Name = "NaoFecha_trans", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 22403, Name = "Mensagem?", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 26981, Name = "Proforma", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 56293, Name = "Iniciar_Transação?", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 68164, Name = "Depto", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 76804, Name = "CodEstab", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 78985, Name = "CópiadePedido?", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 88682, Name = "SintegraCliente", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 89539, Name = "SintegraRedesp", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

                    };
        }

        /*
        public override string GetActionCall()
        {
            string code = "";
            IGPluginControl controlReturn = this.Parameters[3].Value as IGPluginControl;

            if (controlReturn?.ControlType == "GCOMBOBOX")
            {
                code += controlReturn.Name + ".Value";
            }
            else if (controlReturn?.ControlType == "GTEXTBOX")
            {
                code += controlReturn.Name + ".Text";
            }
            else
            {
                code = "var result";
            }
            //code += $" = GvinciAsaas.Assas_CustomerCreateOrUpdate({this.Parameters[0].Value}, {this.Parameters[1].Value}, {this.Parameters[2].Value}, {this.Parameters[3].Value}, {this.Parameters[4].Value}, {this.Parameters[5].Value}, {this.Parameters[6].Value}, {this.Parameters[7].Value}, {this.Parameters[8].Value}, {this.Parameters[9].Value});";
            code += $" = clsRuleDynamicallyCompiled_R11301.R11301((ref conn, ref trans, {this.Parameters[0].Value}, {this.Parameters[1].Value}, {this.Parameters[2].Value});";
            return code;
        }
        */
        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {

            // parâmetro PréPedido. Fernando 06/08/2024
            string P12596 = "";
            IGPluginControl controlP12596 = this.Parameters[0].Value as IGPluginControl;

            if (controlP12596?.ControlType == "GCOMBOBOX")
            {
                P12596 += controlP12596.Name + ".Value";
            }
            else if (controlP12596?.ControlType == "GTEXTBOX")
            {
                P12596 += "Convert.ToDouble(" + controlP12596.Name + ".Text)";
            }
            else
            {
                P12596 = "var result";
            }

            // parâmetro NaoAbre_Trans. Fernando 06/08/2024
            string P20438 = "";
            IGPluginControl controlP20438 = this.Parameters[1].Value as IGPluginControl;

            if (controlP20438?.ControlType == "GCOMBOBOX")
            {
                P20438 += controlP20438.Name + ".Value";
            }
            else if (controlP20438?.ControlType == "GTEXTBOX")
            {
                P20438 += controlP20438.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP20438.Name + ".Text) : Convert.DBNull";
            }
            else
            {
                P20438 = "var result";
            }

            // parâmetro NaoFecha_trans. Fernando 03/10/24
            string P20439 = "";
            IGPluginControl controlP20439 = this.Parameters[2].Value as IGPluginControl;

            if (controlP20439?.ControlType == "GCOMBOBOX")
            {
                P20439 += controlP20439.Name + ".Value";
            }
            else if (controlP20439?.ControlType == "GTEXTBOX")
            {
                P20439 += controlP20439.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP20439.Name + ".Text) : Convert.DBNull";
            }
            else
            {
                P20439 = "var result";
            }

            // parâmetro Mensagem?. Fernando 03/10/24
            string P22403 = "";
            IGPluginControl controlP22403 = this.Parameters[3].Value as IGPluginControl;

            if (controlP22403?.ControlType == "GCOMBOBOX")
            {
                P22403 += controlP22403.Name + ".Value";
            }
            else if (controlP22403?.ControlType == "GTEXTBOX")
            {
                P22403 += controlP22403.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP22403.Name + ".Text) : Convert.DBNull";
            }
            else
            {
                P22403 = "var result";
            }

            // parâmetro Proforma. Fernando 03/10/24
            string P26981 = "";
            IGPluginControl controlP26981 = this.Parameters[4].Value as IGPluginControl;

            if (controlP26981?.ControlType == "GCOMBOBOX")
            {
                P26981 += controlP26981.Name + ".Value";
            }
            else if (controlP26981?.ControlType == "GTEXTBOX")
            {
                P26981 += controlP26981.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP26981.Name + ".Text) : Convert.DBNull";
            }
            else
            {
                P26981 = "var result";
            }

            // parâmetro Iniciar_Transação?. Fernando 03/10/24
            string P56293 = "";
            IGPluginControl controlP56293 = this.Parameters[5].Value as IGPluginControl;

            if (controlP56293?.ControlType == "GCOMBOBOX")
            {
                P56293 += controlP56293.Name + ".Value";
            }
            else if (controlP56293?.ControlType == "GTEXTBOX")
            {
                P56293 += controlP56293.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP56293.Name + ".Text) : Convert.DBNull";
            }
            else
            {
                P56293 = "var result";
            }

            // parâmetro Depto. Fernando 03/10/2024
            string P68164 = "";
            IGPluginControl controlP68164 = this.Parameters[6].Value as IGPluginControl;

            if (controlP68164?.ControlType == "GCOMBOBOX")
            {
                P68164 += controlP68164.Name + ".Value";
            }
            else if (controlP68164?.ControlType == "GTEXTBOX")
            {
                P68164 += controlP68164.Name + ".Text != \"\" ? " + "" + controlP68164.Name + ".Text : Convert.DBNull";
            }
            else
            {
                P68164 = "var result";
            }

            // parâmetro CodEstab. Fernando 03/10/24
            string P76804 = "";
            IGPluginControl controlP76804 = this.Parameters[7].Value as IGPluginControl;

            if (controlP76804?.ControlType == "GCOMBOBOX")
            {
                P76804 += controlP76804.Name + ".Value";
            }
            else if (controlP76804?.ControlType == "GTEXTBOX")
            {
                P76804 += "Convert.ToDouble(" + controlP76804.Name + ".Text)";
            }
            else
            {
                P76804 = "var result";
            }

            // parâmetro CópiadePedido?. Fernando 03/10/24
            string P78985 = "";
            IGPluginControl controlP78985 = this.Parameters[8].Value as IGPluginControl;

            if (controlP78985?.ControlType == "GCOMBOBOX")
            {
                P78985 += controlP78985.Name + ".Value";
            }
            else if (controlP78985?.ControlType == "GTEXTBOX")
            {
                P78985 += controlP78985.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP78985.Name + ".Text) : Convert.DBNull";
            }
            else
            {
                P78985 = "var result";
            }

            // parâmetro SintegraCliente. Fernando 03/10/24
            string P88682 = "";
            IGPluginControl controlP88682 = this.Parameters[9].Value as IGPluginControl;

            if (controlP88682?.ControlType == "GCOMBOBOX")
            {
                P88682 += controlP88682.Name + ".Value";
            }
            else if (controlP88682?.ControlType == "GTEXTBOX")
            {
                P88682 += controlP88682.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP88682.Name + ".Text) : Convert.DBNull";
            }
            else
            {
                P88682 = "var result";
            }

            // parâmetro SintegraRedesp. Fernando 03/10/2024
            string P89539 = "";
            IGPluginControl controlP89539 = this.Parameters[10].Value as IGPluginControl;

            if (controlP89539?.ControlType == "GCOMBOBOX")
            {
                P89539 += controlP89539.Name + ".Value";
            }
            else if (controlP89539?.ControlType == "GTEXTBOX")
            {
                P89539 += controlP89539.Name + ".Text != \"\" ? " + "Convert.ToDouble(" + controlP89539.Name + ".Text) : Convert.DBNull";

            }
            else
            {
                P89539 = "var result";
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

            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R4557.R4557(ref conn, ref trans, ref msg, " + P12596 + ", " + P20438 + ", " + P20439 + ", " + P22403 + ", " + P26981 + ", " + P56293 + ", " + P68164 + ", " + P76804 + ", " + P78985 + ", " + P88682 + ", " + P89539 + ");");

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
                new GPluginDependency() { FileName = "R4557_V0_D20241009_160451.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }

    }

}
