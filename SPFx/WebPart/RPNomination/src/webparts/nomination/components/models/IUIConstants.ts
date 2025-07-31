import { IComboBoxOption } from "@fluentui/react/lib/ComboBox";

export const INITIAL_NOTIFY_OPTIONS: IComboBoxOption[] = [
    { key: 'Nominee', text: 'Nominee'},
    { key: 'Nominator', text: 'Nominator(s)' },
    { key: 'EPNominators', text: 'EP Nominator(s)' },         
];

export const INITIAL_CANDIDATE_NOMINATED: IComboBoxOption[] = [
  { key: 'Highest Credentialed Professional ', text: 'Highest Credentialed Professional '},
  { key: 'Experienced Professional', text: 'Experienced Professional' },
  { key: 'Product and Technology Professional', text: 'Product and Technology Professional' },         
];

export const INITIAL_REFERENCESPASSED_AND_QARPASSED_OPTIONS: IComboBoxOption[] = [
  { key: 'Yes', text: 'Yes'},
  { key: 'No', text: 'No' },
  { key: 'N/A', text: 'N/A' },         
];

export const INITIAL_TRACK_REFERENCES_OPTIONS: IComboBoxOption[] = [
  { key: 'Blank', text: 'Blank'},
  { key: 'Pending', text: 'Pending' },
  { key: 'Complete', text: 'Complete' },
  { key: 'Unqualified', text: 'Unqualified' },
  { key: 'Unavailable', text: 'Unavailable' },         
];


export interface IConstants {
  readonly PowerAutomateFlowUrl?: string;
  readonly PermissionPowerAutomateFlowUrl?: string;
  readonly SP_LIST_NAMES: {
    NominationDocumentLibraryName: string;
  };
}


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
      "dev": "millimandev",
      "uat": "millimantest",
      "prod": "milliman"
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
        "PowerAutomateFlowUrl":"https://prod-41.westus.logic.azure.com:443/workflows/1f5543c73d2c465cac88a6ef7468ae42/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=JjUxFNMpH9NJDXpJuqtedYcNQgmAlSWay9v2fVALk-g",
        "PermissionPowerAutomateFlowUrl":"https://prod-177.westus.logic.azure.com:443/workflows/d5e1ad42b32a4e08935c112115987c9a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Vp_Rd1jbY03oYvSMqkxx9e_ftEYr9Nt0nqIqwyIswts",
        "SP_LIST_NAMES": {
          "NominationDocumentLibraryName": "Nomination Attachments"
        }
      };
      private static uatConstants: IConstants = {
        "PowerAutomateFlowUrl":"https://prod-144.westus.logic.azure.com:443/workflows/659f37c80d0e486c9126ae3f03b7d500/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NoaJblGx73_Ozq_WNN7CFoVkWveblsbwQJzkPG_GpWI",
        "PermissionPowerAutomateFlowUrl":"https://prod-44.westus.logic.azure.com:443/workflows/d9eff74ac28f4890a01efff75d33a274/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=cdHo3fVmhZ68KTYBT8QPHVkUeGDioIjOa5bqcEWy_lY",
        "SP_LIST_NAMES": {
          "NominationDocumentLibraryName": "Nomination Attachments"
        }
      };
      private static prodConstants: IConstants = {
        "PowerAutomateFlowUrl":"https://prod-136.westus.logic.azure.com:443/workflows/3bc7288f417c4007afa58e9fbc896895/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=eZG9WMGEU_faEQi-UdOjgQ04uH9eBqxjDla5McAn_c4",                               
        "PermissionPowerAutomateFlowUrl":"https://prod-118.westus.logic.azure.com:443/workflows/503692776d5a485687044783455b5961/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=pk8-xiuviLfWbKKcSZ0W21DtaRPdWiR9C4hplqqYcx4",
        "SP_LIST_NAMES": {
          "NominationDocumentLibraryName": "Nomination Attachments"
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