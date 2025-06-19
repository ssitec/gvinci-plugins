using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;


public static class GvinciAsaas
{
    public static string Assas_CustomerCreateOrUpdate(string Token, string Name, string CpfCnpj, string Email, string Phone, string Mobilephone, string PostalCode, string AddressNumber, string ExternalReference, string Environment)
    {
        string CustomerId = "";
        try
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
            string linkAsaas = "https://sandbox.asaas.com/api/v3/";
            if (Environment == "P") linkAsaas = "https://www.asaas.com/api/v3/";
            string url = linkAsaas + "/customers?email=" + (Email);
            var client = new RestClient(url);
            var request = new RestRequest();
            AsaasModel.CustomerResponse customerResponse = new AsaasModel.CustomerResponse();
            request.AddHeader("access_token", Token);
            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                AsaasModel.CustomerResponseList listCustomers = JsonConvert.DeserializeObject<AsaasModel.CustomerResponseList>(response.Content);
                if (listCustomers == null)
                {
                    throw new Exception("Erro - " + response.ErrorMessage); ;
                }
                if (listCustomers.totalCount == 0)
                {
                    request = new RestRequest(Method.POST);
                    url = linkAsaas + "/customers";
                }
                else
                {
                    request = new RestRequest(Method.PUT);
                }

                request.AddHeader("access_token", Token);
                client = new RestClient(url);
                var customerRequest = new AsaasModel.CustomerRequest();
                if (listCustomers.totalCount > 0) customerRequest.id = listCustomers.data[0].id;
                customerRequest.name = Name;
                customerRequest.cpfCnpj = CpfCnpj;
                customerRequest.email = Email;
                customerRequest.phone = Phone;
                customerRequest.mobilePhone = Mobilephone;
                customerRequest.addressNumber = AddressNumber;
                customerRequest.postalCode = PostalCode;
                customerRequest.externalReference = ExternalReference;
                request.AddJsonBody(customerRequest);

                response = client.Execute(request);
                if (!response.IsSuccessful)
                {
                    throw new Exception(response.Content);
                }
                customerResponse = JsonConvert.DeserializeObject<AsaasModel.CustomerResponse>(response.Content);
                return customerResponse.id;
            }
            else
            {
                throw new Exception(response.Content);
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public static string Assas_CreateBoleto(string Token, string Name, string CpfCnpj, string Email, string Phone, string Mobilephone, string PostalCode, string AddressNumber, string ExternalReference, string LinkAsaas, DateTime Vencimento, decimal Valor, string Descricao, string DocRef)
    {
        var CustomerID = Assas_CustomerCreateOrUpdate(Token, Name, CpfCnpj, Email, Phone, Mobilephone, PostalCode, AddressNumber, ExternalReference, LinkAsaas);
        return Assas_CreateBoleto(CustomerID, LinkAsaas, Token, Vencimento, Valor, Descricao, DocRef);
    }

    public static string Assas_CreateBoleto(string CustomerID, string Environment, string Token, DateTime DueDate, decimal Value, string Description, string DocRef)
    {
        try
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
            string LinkAsaas = "https://sandbox.asaas.com/api/v3/";
            if (Environment == "P") LinkAsaas = "https://www.asaas.com/api/v3/";
            var client = new RestClient(LinkAsaas + "/payments");
            var request = new RestRequest();
            request.AddHeader("access_token", Token);
            var cobrancaRequest = new AsaasModel.CobrancaBoletoPixRequest();

            cobrancaRequest.customer = CustomerID; //Customer ID

            DateTime venc = DueDate;
            cobrancaRequest.dueDate = venc.ToString("yyyy-MM-dd");  //"15/12/2016";
            cobrancaRequest.value = (Decimal)Value;

            cobrancaRequest.description = Description;

            cobrancaRequest.externalReference = DocRef;

            cobrancaRequest.billingType = "BOLETO";
            cobrancaRequest.interest = new AsaasModel.Interest { value = 2 }; // cobranca.PercentualJurosMes;
            cobrancaRequest.fine = new AsaasModel.Fine { value = 1 };//cobranca.PercentualMulta;
            request.AddJsonBody(cobrancaRequest);
            AsaasModel.CobrancaResponse cobrancaResponse = new AsaasModel.CobrancaResponse();
            var response = client.Post(request);
            if (!response.IsSuccessful)
            {
                throw new Exception(response.Content);
            }
            cobrancaResponse = JsonConvert.DeserializeObject<AsaasModel.CobrancaResponse>(response.Content);
            return cobrancaResponse.id;
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }
}