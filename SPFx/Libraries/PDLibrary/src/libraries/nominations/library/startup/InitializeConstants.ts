import { IConstants } from '../../models/IConstants';

interface IEnvironmentConstants {
  dev: IConstants;
  uat: IConstants;
  prod: IConstants;
}

export class ConstantsConfig {

  public static GetConstants(): IConstants {
    return this.environmentConstants[this.getCurrentTenant()];

  }
  private static tenantEnvironment: any = {
    "dev": "spo365dev",
    "uat": "spo365dev",
    "prod": "spo365dev"
  };

  private static getCurrentTenant(): string {
    var currentWebUrl = document.URL.toLowerCase();
    if (currentWebUrl.indexOf("https://" + this.tenantEnvironment.dev + ".sharepoint.com") === 0)
      return "dev";
    else if (currentWebUrl.indexOf("https://" + this.tenantEnvironment.uat + ".sharepoint.com") === 0)
      return "uat";
    else if (currentWebUrl.indexOf("https://" + this.tenantEnvironment.prod + ".sharepoint.com") === 0)
      return "prod";
    else
      return "dev";
  }

  private static devConstants: IConstants = {
    "ENV": "dev",
    "LOG_SOURCE": "PD Nomination",
    "SP_LIST_NAMES": {
      "NominationDocumentLibraryName": "Nomination Attachments",
      "LANominationList": "Nomination Details By Local Admin",
      "MasterNominationList": "PD Nominations",
      "QCNominationList": "Nomination Details By QC",
      "NominationReviewersList": "Nomination Reviewers",
      "PtpacNominationList": "Nomination Details By PTPAC",
      "LegalNominationList": "Nomination Details By GCS Legal",
      "NotificationList": "Nomination Notifications",
      "ReferencesList":"Track References"
    },
    "ENDPOINTS": {
      "BaseAzureApiUrl": "https://millimandevbts-data.azurewebsites.net/api",
      "AadClientResourceIdentifier": "1e40af73-1485-48f7-b46d-4ec45939a254",
      "NomineeDetailsFetch": "IntegrationBD/GetNomineeDetailByEmail",
      "DisciplinesFetch": "IntegrationBD/GetDisciplines",
      "ProfDesignationsFetch": "IntegrationBD/GetProfessionalDesignations",
      "ProficientLanguageFetch": "IntegrationBD/GetProficientLanguages",
      "PDSubcategoriesFetch": "IntegrationBD/GetPDSubCategories",
      "EmployeeDetailsUpdate": "IntegrationBD/UpdateEmployeeDirectoryForNominee",
      "EmployeeProfessionalDesignationsByFinanceUserId": "IntegrationBD/GetEmployeeProfessionalDesignationsByFinanceUserId",
      "EmployeeInformationByFinanceUserId": "IntegrationBD/GetEmployeeProfessionalDesignationsByFinanceUserId",
    },

    "SP_TEST_USER": "donna.boyle@milliman.com",

    "SEND_EMAIL": true,
    "BREAK_ATTACHMENT_PERMISSION":true,


    "SP_Content_Type": {
      "DocSetContentType": "0x0120D520"
    }
    ,
    "SP_Group": {
      "LegalGroup": "GCS Legal"
    }
  };
  private static uatConstants: IConstants = {
    "ENV": "test",
    "LOG_SOURCE": "PD Nomination",
    "SP_LIST_NAMES": {
      "NominationDocumentLibraryName": "Nomination Attachments",
      "LANominationList": "Nomination Details By Local Admin",
      "MasterNominationList": "PD Nominations",
      "QCNominationList": "Nomination Details By QC",
      "NominationReviewersList": "Nomination Reviewers",
      "PtpacNominationList": "Nomination Details By PTPAC",
      "LegalNominationList": "Nomination Details By GCS Legal",
      "NotificationList": "Nomination Notifications",
      "ReferencesList":"Track References"

    },
    "ENDPOINTS": {
      "BaseAzureApiUrl": "https://millimantestbts-data.azurewebsites.net/api",
      "AadClientResourceIdentifier": "67e1f85b-bd41-4d56-8925-a633ae8dbf72",
      "NomineeDetailsFetch": "IntegrationBD/GetNomineeDetailByEmail",
      "DisciplinesFetch": "IntegrationBD/GetDisciplines",
      "ProfDesignationsFetch": "IntegrationBD/GetProfessionalDesignations",
      "ProficientLanguageFetch": "IntegrationBD/GetProficientLanguages",
      "PDSubcategoriesFetch": "IntegrationBD/GetPDSubCategories",
      "EmployeeDetailsUpdate": "IntegrationBD/UpdateEmployeeDirectoryForNominee",
      "EmployeeProfessionalDesignationsByFinanceUserId": "IntegrationBD/GetEmployeeProfessionalDesignationsByFinanceUserId",
      "EmployeeInformationByFinanceUserId": "IntegrationBD/GetEmployeeProfessionalDesignationsByFinanceUserId",

    },

    "SP_TEST_USER": "",
    "SEND_EMAIL": true,
    "BREAK_ATTACHMENT_PERMISSION":true,


    "SP_Content_Type": {
      "DocSetContentType": "0x0120D520"
    }
    ,
    "SP_Group": {
      "LegalGroup": "GCS Legal"
    }
  };
  private static prodConstants: IConstants = {
    "ENV": "production",
    "LOG_SOURCE": "PD Nomination",
    "SP_LIST_NAMES": {
      "NominationDocumentLibraryName": "Nomination Attachments",
      "LANominationList": "Nomination Details By Local Admin",
      "MasterNominationList": "PD Nominations",
      "QCNominationList": "Nomination Details By QC",
      "NominationReviewersList": "Nomination Reviewers",
      "PtpacNominationList": "Nomination Details By PTPAC",
      "LegalNominationList": "Nomination Details By GCS Legal",
      "NotificationList": "Nomination Notifications",
      "ReferencesList":"Track References"
    },
    "ENDPOINTS": {
      "BaseAzureApiUrl": "https://millimanbts-data.azurewebsites.net/api",
      "AadClientResourceIdentifier": "456cc820-8311-437f-ada7-d75d6e0493a6",
      "NomineeDetailsFetch": "IntegrationBD/GetNomineeDetailByEmail",
      "DisciplinesFetch": "IntegrationBD/GetDisciplines",
      "ProfDesignationsFetch": "IntegrationBD/GetProfessionalDesignations",
      "ProficientLanguageFetch": "IntegrationBD/GetProficientLanguages",
      "PDSubcategoriesFetch": "IntegrationBD/GetPDSubCategories",
      "EmployeeDetailsUpdate": "IntegrationBD/UpdateEmployeeDirectoryForNominee",
      "EmployeeProfessionalDesignationsByFinanceUserId": "IntegrationBD/GetEmployeeProfessionalDesignationsByFinanceUserId",
      "EmployeeInformationByFinanceUserId": "IntegrationBD/GetEmployeeProfessionalDesignationsByFinanceUserId",
    },

    "SP_TEST_USER": "",
    "SEND_EMAIL": true,
    "BREAK_ATTACHMENT_PERMISSION":true,


    "SP_Content_Type": {
      "DocSetContentType": "0x0120D520"
    },
    "SP_Group": {
      "LegalGroup": "GCS Legal"
    }

  };

  private static environmentConstants: IEnvironmentConstants = {
    dev: ConstantsConfig.devConstants,
    uat: ConstantsConfig.uatConstants,
    prod: ConstantsConfig.prodConstants
  };

  public static get(): IConstants {
    return this.environmentConstants[this.getCurrentTenant()];
  }

}
