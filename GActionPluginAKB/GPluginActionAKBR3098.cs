using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR3098 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0002-C827AF67A9C2";

        public override string Name => "R3098";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR3098(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 7202, Name = "CidadeR", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },
                        new GPluginActionParameter() { ID = 7203, Name = "CidadeS", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" }  },

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

            // parâmetro text box. Fernando 25/09/2023
            string CidadeR = "";
            IGPluginControl controlCidadeR = this.Parameters[0].Value as IGPluginControl;

            if (controlCidadeR?.ControlType == "GCOMBOBOX")
            {
                CidadeR += controlCidadeR.Name + ".Value";
            }
            else if (controlCidadeR?.ControlType == "GTEXTBOX")
            {
                CidadeR += "Convert.ToDouble(" + controlCidadeR.Name + ".Text)";
            }
            else
            {
                CidadeR = "var result";
            }


            // parâmetro text box. Fernando 25/09/2023
            string CidadeS = "";
            IGPluginControl controlCidadeS = this.Parameters[1].Value as IGPluginControl;

            if (controlCidadeS?.ControlType == "GCOMBOBOX")
            {
                CidadeS += controlCidadeS.Name + ".Value";
            }
            else if (controlCidadeS?.ControlType == "GTEXTBOX")
            {
                CidadeS += "Convert.ToDouble(" + controlCidadeS.Name + ".Text)";
            }
            else
            {
                CidadeS = "var result";
            }

            string idenStr = new string('\t', Identation);

            Builder.AppendLine(idenStr + "string connString = \"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=191.13.138.110)(PORT=2226))(CONNECT_DATA=(SERVICE_NAME=dassistest)))\";");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleConnection conn = new System.Data.OracleClient.OracleConnection(\"Data Source=\" + connString + \";User ID=akbit;Password=Fama1200\");");
            Builder.AppendLine(idenStr + "conn.Open();");

            // Não abre transação nova. Fernando 03/06/2023
            //Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = conn.BeginTransaction();");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = null;");

            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R3098.R3098(ref conn, ref trans, " + CidadeR + ", " + CidadeS + ");");

            //Builder.AppendLine(idenStr + "trans.Commit();");
            Builder.AppendLine(idenStr + "conn.Close();");



        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                new GPluginDependency() { FileName = "R3098_V3.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }

    }
}
