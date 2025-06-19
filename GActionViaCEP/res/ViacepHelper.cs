using Gvinci.Plugin.Viacep;
using RestSharp;
using RestSharp.Serialization.Json;


namespace Gvinci.Plugin.Viacep
{
    public static class ViacepHelper
    {
        public static ViacepModel BuscarCep(string cep)
        {

            RestClient restClient = new RestClient(string.Format("http://viacep.com.br/ws/{0}/json/", cep.Replace("-", "").Replace(".", "")));
            RestRequest restRequest = new RestRequest(Method.GET);

            IRestResponse restResponse = restClient.Execute(restRequest);

            return new JsonDeserializer().Deserialize<ViacepModel>(restResponse);
        }
    }
}
