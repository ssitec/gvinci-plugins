using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for AsaasModel
/// </summary>
public class AsaasModel
{

    public class CustomerRequest
    {
        public string id { get; set; }
        public string name { get; set; }
        public string email { get; set; }
        public string phone { get; set; }
        public string mobilePhone { get; set; }
        public string cpfCnpj { get; set; }
        public string postalCode { get; set; }
        public string address { get; set; }
        public string addressNumber { get; set; }
        public string complement { get; set; }
        public string province { get; set; }
        public string externalReference { get; set; }
        public bool notificationDisabled { get; set; }
        public string additionalEmails { get; set; }
        public string municipalInscription { get; set; }
        public string stateInscription { get; set; }
        public string observations { get; set; }
    }

    public class CustomerResponse
    {
        public string @object { get; set; }
        public string id { get; set; }
        public string dateCreated { get; set; }
        public string name { get; set; }
        public string email { get; set; }
        public object company { get; set; } 
        public string phone { get; set; }
        public string mobilePhone { get; set; }
        public string address { get; set; }
        public string addressNumber { get; set; }
        public string complement { get; set; }
        public string province { get; set; }
        public string postalCode { get; set; }
        public string cpfCnpj { get; set; }
        public string personType { get; set; }
        public bool deleted { get; set; }
        public string additionalEmails { get; set; }
        public string externalReference { get; set; }
        public bool notificationDisabled { get; set; }
        public string observations { get; set; }
        public string municipalInscription { get; set; }
        public string stateInscription { get; set; }
        public bool canDelete { get; set; }
        public object cannotBeDeletedReason { get; set; }
        public bool canEdit { get; set; }
        public object cannotEditReason { get; set; }
        public bool foreignCustomer { get; set; }
        public int? city { get; set; }
        public string state { get; set; }
        public string country { get; set; }

        //Erros
        public List<Error> errors { get; set; }
    }

    public class CobrancaRequest
    {
        public string customer { get; set; }
        public string billingType { get; set; }
        public string dueDate { get; set; }
        public decimal? value { get; set; }
        public string description { get; set; }
        public string externalReference { get; set; }
    }

    public class CobrancaBoletoPixRequest : CobrancaRequest
    {
        public Discount discount { get; set; }
        public Fine fine { get; set; }
        public Interest interest { get; set; }
        public bool postalService { get; set; }
        public string creditCarDtoken { get; set; }
    }

    public class CobrancaCartaoRequest : CobrancaRequest
    {
        public CreditCard creditCard { get; set; }
        public CreditCardHolderInfo creditCardHolderInfo { get; set; }
        public string creditCarDtoken { get; set; }
    }

    public class CreditCard
    {
        public string holderName { get; set; }
        public string number { get; set; }
        public string expiryMonth { get; set; }
        public string expiryYear { get; set; }
        public string ccv { get; set; }
    }

    public class CreditCardHolderInfo
    {
        public string name { get; set; }
        public string email { get; set; }
        public string cpfCnpj { get; set; }
        public string postalCode { get; set; }
        public string addressNumber { get; set; }
        public object addressComplement { get; set; }
        public string phone { get; set; }
        public string mobilePhone { get; set; }
    }

    public class CobrancaResponse
    {
        public string @object { get; set; }
        public string id { get; set; }
        public string dateCreated { get; set; }
        public string customer { get; set; }
        public string paymentLink { get; set; }
        public string dueDate { get; set; }
        public decimal? value { get; set; }
        public decimal? netValue { get; set; }
        public string billingType { get; set; }
        public object pixTransaction { get; set; }
        public string status { get; set; }
        public string description { get; set; }
        public string externalReference { get; set; }
        public string confirmedDate { get; set; } // cartao
        public object originalValue { get; set; }
        public object interestValue { get; set; }
        public string originalDueDate { get; set; }
        public object paymentDate { get; set; }
        public object clientPaymentDate { get; set; }
        public string transactionReceiptUrl { get; set; }
        public object nossoNumero { get; set; }
        public string invoiceUrl { get; set; }
        public object bankSlipUrl { get; set; }
        public string invoiceNumber { get; set; }
        public bool deleted { get; set; }
        public bool postalService { get; set; }
        public bool anticipated { get; set; }
        public CreditCardResponse creditCard { get; set; }

        //Erros
        public List<Error> errors { get; set; }


        //BOLETO
        public Discount discount { get; set; }
        public Fine fine { get; set; }
        public Interest interest { get; set; }
        public object refunds { get; set; }

    }

    public class CreditCardResponse
    {
        public string creditCardNumber { get; set; }
        public string creditCardBrand { get; set; }
        public string creditCarDtoken { get; set; }
    }

    //WEBHOOK
    public class PaymentNotification
    {
        public string @event { get; set; }
        public Payment payment { get; set; }
    }

    public class Chargeback
    {
        public string status { get; set; }
        public string reason { get; set; }
    }

    public class Discount
    {
        public decimal? value { get; set; }
        public int dueDateLimitDays { get; set; }
        public string type { get; set; }
    }

    public class Fine
    {
        public decimal? value { get; set; }
        public string type { get; set; }
    }

    public class Interest
    {
        public decimal value { get; set; }
        public string type { get; set; }
    }

    public class Payment
    {
        public string @object { get; set; }
        public string id { get; set; }
        public string dateCreated { get; set; }
        public string customer { get; set; }
        public string subscription { get; set; }
        public string installment { get; set; }
        public string paymentLink { get; set; }
        public string dueDate { get; set; }
        public string originalDueDate { get; set; }
        public decimal value { get; set; }
        public double netValue { get; set; }
        public object originalValue { get; set; }
        public object interestValue { get; set; }
        public string description { get; set; }
        public string externalReference { get; set; }
        public string billingType { get; set; }
        public string status { get; set; }
        public string confirmedDate { get; set; }
        public string paymentDate { get; set; }
        public string clientPaymentDate { get; set; }
        public string creditDate { get; set; }
        public string estimatedCreditDate { get; set; }
        public string invoiceUrl { get; set; }
        public object bankSlipUrl { get; set; }
        public string transactionReceiptUrl { get; set; }
        public string invoiceNumber { get; set; }
        public bool deleted { get; set; }
        public bool anticipated { get; set; }
        public string lastInvoiceViewedDate { get; set; }
        public object lastBankSlipViewedDate { get; set; }
        public bool postalService { get; set; }
        public CreditCardResponse creditCard { get; set; }
        public Discount discount { get; set; }
        public Fine fine { get; set; }
        public Interest interest { get; set; }
        public List<Split> split { get; set; }
        public Chargeback chargeback { get; set; }
        public object refunds { get; set; }
    }

    public class Split
    {
        public string walletId { get; set; }
        public int fixedValue { get; set; }
        public int? percentualValue { get; set; }
    }

    public class Cancel
    {
        public bool deleted { get; set; }
        public string id { get; set; }
    }

    //ERROS
    public class Error
    {
        public string code { get; set; }
        public string description { get; set; }
    }

    public class CustomerResponseList
    {
        public string _object { get; set; }
        public bool hasMore { get; set; }
        public int totalCount { get; set; }
        public int limit { get; set; }
        public int offset { get; set; }
        public CustomerResponse[] data { get; set; }
    }
}