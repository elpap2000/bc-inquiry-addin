window.ADDIN_CONFIG = {
  tenantId: "PASTE-YOUR-ENTRA-TENANT-ID",
  clientId: "PASTE-YOUR-APP-CLIENT-ID",

  // Business Central Online delegated scope
  // Keep this unless your IT confirms a different BC resource URI in your tenant.
  bcScope: "https://api.businesscentral.dynamics.com/user_impersonation",

  baseUrl: "https://businesscentral.dynamics.com/Tototheo/",
  company: "Tototheo Maritime Ltd",
  inquiryListPage: 70355879,

  odataBase: "https://api.businesscentral.dynamics.com/v2.0/e5296d4f-699b-43ea-a129-066a3f7010e3/Tototheo/ODataV4",
  odataEntity: "InquiryCard",

  customerRefField: "Customer_Ref_No",
  bcFilterField: "Customer Ref. No"
};
