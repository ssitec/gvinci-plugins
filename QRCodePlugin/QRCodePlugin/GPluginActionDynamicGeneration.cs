using Gvinci.Plugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QRCodePlugin
{
    internal class GPluginActionDynamicGeneration : GPluginAction
    {
        public override string ID => "FCCB254B-4920-472A-0002-BEB5226B3DF2";

        public override string Name => "Gerar QrCode Dinâmico";

        public override string Description => "";

        private List<GPluginActionParameter> _parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _parameters;
            }
        }

        public GPluginActionDynamicGeneration(IGPlugin Plugin) : base(Plugin)
        {
            _parameters = new List<GPluginActionParameter>()
            {
                new GPluginActionParameter() {ID = 1, Name= "Url",                      Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" }},
                new GPluginActionParameter() {ID = 2, Name= "Response (Link Or Image)", Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GIMAGE", "GLINK" }},
            };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {
            IGPluginControl url = this.Parameters[0].Value as IGPluginControl;
            IGPluginControl response = this.Parameters[1].Value as IGPluginControl;


            string identStr = new string('\t', Identation);

            Builder.AppendLine(identStr + $"var image = Gvinci.Plugin.QRCodePlugin.QrCodeHelper.GenerateByteArray({url.Name}.Text);");

            if (response.ControlType == "GIMAGE")
            {
                Builder.AppendLine(identStr + $"{response.Name}.ImageUrl = \"data:image/png;base64,\" + Convert.ToBase64String(image);");
            }
            else
            {
                Builder.AppendLine(identStr + $"{response.Name}.Attributes.Add(\"href\", \"data:image/png;base64,\" + Convert.ToBase64String(image));");
                Builder.AppendLine(identStr + $"{response.Name}.Attributes.Add(\"download\", \"QrCode\"); ");
            }

        }
    }
}
