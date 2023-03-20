using Newtonsoft.Json;
using System;
using System.Net;
using RestSharp;
using RestSharp.Authenticators;

namespace Gvinci.Plugin.OpenAI
{
    public class OpenAIHelper
    {
        public static string CallOpenAIChatGPT(string apiToken, int tokens, string input, string engine, decimal temperature, int topP, decimal frequencyPenalty, decimal presencePenalty)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
            var openAiKey = apiToken;

            var apiCall = "https://api.openai.com/v1/engines/" + engine + "/completions";

            var restClient = new RestClient(apiCall);
            restClient.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(openAiKey, "Bearer");

            RestRequest restRequest = new RestRequest();
            restRequest.Method = Method.POST;
            restRequest.RequestFormat = DataFormat.Json;
            restRequest.AddHeader("Content-Type", "application/json");

            var body = new
            {
                prompt = input,
                temperature = temperature,
                max_tokens = tokens,
                top_p = topP,
                frequency_penalty = frequencyPenalty,
                presence_penalty = presencePenalty
            };

            restRequest.AddJsonBody(body);

            var response = restClient.Execute(restRequest);

            if (!response.IsSuccessful)
            {
                throw new Exception(response.Content);
            }

            OpenAIResponse openAIResponse = JsonConvert.DeserializeObject<OpenAIResponse>(response.Content);

            return openAIResponse.choices[0].text;
        }

        public static ResponseModel CallOpenAiDallE(string apiToken, string input, int number, string size)
        {

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
            var openAiKey = apiToken;

            var apiCall = "https://api.openai.com/v1/images/generations.";

            var restClient = new RestClient(apiCall);
            restClient.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(openAiKey, "Bearer");

            RestRequest restRequest = new RestRequest();
            restRequest.Method = Method.POST;
            restRequest.RequestFormat = DataFormat.Json;
            restRequest.AddHeader("Content-Type", "application/json");

            var body = new
            {
                prompt = input,
                n = number,
                size = size
            };

            restRequest.AddJsonBody(body);

            var response = restClient.Execute(restRequest);

            if (!response.IsSuccessful)
            {
                throw new Exception(response.Content);
            }

            ResponseModel openAIResponse = JsonConvert.DeserializeObject<ResponseModel>(response.Content);

            return openAIResponse;
        }
    }
}
