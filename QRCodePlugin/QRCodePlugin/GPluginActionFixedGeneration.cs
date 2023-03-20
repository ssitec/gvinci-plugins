using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    internal class GPluginActionFixedGeneration : GPluginAction
    {
        public override string ID => "FCCB254B-4920-472A-0001-BEB5226B3DF2";

        public override string Name => "Gerar QrCode Fixo";

        public override string Description => "";

        private List<GPluginActionParameter> _parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _parameters;
            }
        }

        public GPluginActionFixedGeneration(IGPlugin Plugin) : base(Plugin)
        {
            _parameters = new List<GPluginActionParameter>()
            {
                new GPluginActionParameter() {ID = 1, Name= "Url",                      Description = "", Required = true, Type= PluginActionParameterTypeEnum.STRING },
                new GPluginActionParameter() {ID = 2, Name= "Response (Link Or Image)", Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GIMAGE", "GLINK" }},
            };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {
            var url = this.Parameters[0].Value.ToString();
            IGPluginControl response = this.Parameters[1].Value as IGPluginControl;


            string identStr = new string('\t', Identation);

            Builder.AppendLine(identStr + $"var image = Gvinci.Plugin.QRCodePlugin.QrCodeHelper.GenerateByteArray({url});");

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
