using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAKBR11301 : GPluginAction
    {
        public override string ID => "6CACD40E-4F99-4D96-0001-C827AF67A9C2";

        public override string Name => "R11301";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAKBR11301(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 37254, Name = "UnidadeMedida", Type = PluginActionParameterTypeEnum.DECIMAL },
                        new GPluginActionParameter() { ID = 37255, Name = "SiglaProd", Type = PluginActionParameterTypeEnum.STRING },

                        // Novo parâmetro. Fernando 03/06/2023
                        new GPluginActionParameter() { ID = 82669, Name = "SiglaDestino", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" } },

                        // Retorno. Fernando 19/06/2023
                        new GPluginActionParameter() { ID = 3, Name = "ReturnControl", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX", "GGRID", "GGRIDCOLUMN" } },
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
            string unidade = this.Parameters[0].Value.ToString();
            string sigla = this.Parameters[1].Value.ToString();

            // Novo parâmetro. Fernando 03/06/2023
            //string SiglaDestino = this.Parameters[2].Value.ToString();
            string SiglaDestino = "";

            // parâmetro text box. Fernando 28/06/2023
            IGPluginControl controlSiglaDestino = this.Parameters[2].Value as IGPluginControl;

            if (controlSiglaDestino?.ControlType == "GCOMBOBOX")
            {
                SiglaDestino += controlSiglaDestino.Name + ".Value";
            }
            else if (controlSiglaDestino?.ControlType == "GTEXTBOX")
            {
                SiglaDestino += controlSiglaDestino.Name + ".Text";
            }
            else
            {
                SiglaDestino = "var result";
            }

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

            string idenStr = new string('\t', Identation);

            Builder.AppendLine(idenStr + "string connString = ((DatabaseInfo)HttpContext.Current.Session[\"DBINFO_AKBUSER01\"]).StringConnection;");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleConnection conn = new System.Data.OracleClient.OracleConnection(connString);");
            Builder.AppendLine(idenStr + "conn.Open();");

            // Não abre transação nova. Fernando 03/06/2023
            //Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = conn.BeginTransaction();");
            Builder.AppendLine(idenStr + "System.Data.OracleClient.OracleTransaction trans = null;");

            Builder.AppendLine(idenStr + "DataTable ret = clsRuleDynamicallyCompiled_R11301.R11301(ref conn, ref trans, " + unidade + ", " + sigla + ", " + SiglaDestino + ");");

            Builder.AppendLine(idenStr + code + " = ret.Rows[0][0].ToString();");

            //Builder.AppendLine(idenStr + "trans.Commit();");
            Builder.AppendLine(idenStr + "conn.Close();");



        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                new GPluginDependency() { FileName = "R11301.vb", DestinationRelativePath = "App_Code/vb", AllowReplace = true }
            };
        }

    }


}