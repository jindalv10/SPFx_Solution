import BaseService from "../BaseService";
import { IIntakeNominationLibrary } from "./IIntakeNominationLibrary";
import { IMasterDetails } from "../../models/IMasterDetails";
import { INomineeDetails } from "../../models/INomineeDetails";
import { NominationStatus } from "../../models/IConstants";
import { INomineeExist } from "../../models/IUserDetails";

export default class IntakeNominationLibrary extends BaseService implements IIntakeNominationLibrary {
    private currentUrl: string = null;
    constructor(context: any) {
        super(context);
        this.currentUrl = context.pageContext.web.absoluteUrl;
    }

    private getAttachmentTypeChoices(): Promise<string[]> {
        let queryEndpoint = this.currentUrl + "/_api/web/lists/GetByTitle('" + this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName + "')/fields?$filter=EntityPropertyName eq 'AttachmentType'";
        return this.spGet(
            queryEndpoint
        )
            .then((response) => {
                return response.json();
            })
            .then((data) => {
                if (data && data.value && data.value.length > 0) {
                    return data.value[0].Choices.map(choice => {
                        let _attachementTypeChoice: string = choice;
                        return _attachementTypeChoice;
                    });
                }
            })
            .catch(e => {
                console.log("Failed to fetch attachment type choices");
                return Promise.reject([]);
            });
    }

    public async checkIfValidNomineeWithDiscAndPDStatus(financeId: string, selectedNomineePDStatus: string, selectedNomineePDDiscipline: string): Promise<INomineeExist> {
      try {
          // <Or><Neq><FieldRef Name='NominationStatus'/><Value Type='Text'>" ${NominationStatus.ApproveCompleted}"</Value></Neq><Neq><FieldRef Name='NominationStatus'/><Value Type='Text'>" ${NominationStatus.WithdrawnCompleted}"</Value></Neq></Or></And>
          let camlQueryString = `<View><Query><Where><And><And><Neq><FieldRef Name='NominationStatus'/><Value Type='Text'>${NominationStatus.ApproveCompleted}</Value></Neq><Neq><FieldRef Name='NominationStatus'/><Value Type='Text'>${NominationStatus.WithdrawnCompleted}</Value></Neq></And><And><And><Eq><FieldRef Name='PDStatus'/><Value Type='Text'>${selectedNomineePDStatus}</Value></Eq><Eq><FieldRef Name='PDDiscipline'/><Value Type='Text'>${selectedNomineePDDiscipline}</Value></Eq></And><Eq><FieldRef Name='FinanceUserID' /><Value Type='Text'>${financeId}</Value></Eq></And></And></Where></Query></View>`;
          const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
          if (nominationsData && nominationsData.length > 0) {
              return {isNomineeExist: false, EPNominator: nominationsData[0].Author[0].title};
          }
          else
           return {isNomineeExist: true};
      }
      catch (e) {
          return Promise.reject(e);
      }
  }
    public async checkIfValidNominee(nomineeId: number): Promise<INomineeExist> {
        try {
            // <Or><Neq><FieldRef Name='NominationStatus'/><Value Type='Text'>" ${NominationStatus.ApproveCompleted}"</Value></Neq><Neq><FieldRef Name='NominationStatus'/><Value Type='Text'>" ${NominationStatus.WithdrawnCompleted}"</Value></Neq></Or></And>
            let camlQueryString = `<View><Query><Where><And><And><Neq><FieldRef Name='NominationStatus'/><Value Type='Text'>${NominationStatus.ApproveCompleted}</Value></Neq><Neq><FieldRef Name='NominationStatus'/><Value Type='Text'>${NominationStatus.WithdrawnCompleted}</Value></Neq></And><Eq><FieldRef Name='NomineeName' LookupId='True'/><Value Type='Lookup'>${nomineeId}</Value></Eq></And></Where></Query></View>`;
            const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
            if (nominationsData && nominationsData.length > 0) {
                return {isNomineeExist: false, EPNominator: nominationsData[0].Author[0].title};
            }

            else
             return {isNomineeExist: true};
        }
        catch (e) {
            return Promise.reject(e);
        }
    }
    public getNomineeDetailsFromEmpDB(email: string): Promise<INomineeDetails> {

        if (email) {
            if(this.Constants.SP_TEST_USER)
            {
              email = this.Constants.SP_TEST_USER;
            }
            return this.aadSecureGet(
                `${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.NomineeDetailsFetch}/${email}`
            )
                .then((response) => {
                    if (response.status)
                        return response.json();
                })
                .then((nominee) => {
                    if (nominee) {
                        let dateToCheck = nominee.grantedOn ? new Date(nominee.grantedOn) : null;

                        let _nomineeDetail: INomineeDetails = {
                            financeUserId: nominee.financeUserId,
                            email: nominee.email,
                            discipline: nominee.discipline,
                            practice: nominee.practice,
                            office: nominee.office,
                            grantedOn: nominee.grantedOn,
                            designation: nominee.designation,
                            hireDate: nominee.hireDate,
                            isStatusGrantedAfter2016: dateToCheck && dateToCheck .getFullYear() > 2016 ? true : false
                        };
                        return _nomineeDetail;

                    }
                })
                .catch(e => {
                    console.error("Failed to fetch Nominee Details");
                    return null;
                });
        }
    }

    private getProfessionalDesignations(): Promise<string[]> {
        return this.aadSecureGet(
            `${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.ProfDesignationsFetch}`
        )
            .then((response) => {
                return response.json();
            })
            .then((data) => {
                if (data && data.length > 0) {
                    return data.map(pd => {
                        let _professionalDesignationId: string = pd.id;
                        let _professionalDesignationCode: string = pd.code;
                        let _professionalDesignationTitle: string = pd.title;
                        return {"_professionalDesignationId":_professionalDesignationId, "_professionalDesignationCode": _professionalDesignationCode,"_professionalDesignationTitle":_professionalDesignationTitle};
                    });
                }
            })
            .catch(e => {
                console.error("Failed to fetch professional designations");
                return Promise.reject(e);
            });
    }
    private getPDSubcategory(): Promise<string[]> {
        return this.aadSecureGet(
            `${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.PDSubcategoriesFetch}`
        )
            .then((response) => {
                return response.json();
            })
            .then((data) => {
                if (data && data.length > 0) {
                    return data.map(pds => {
                        let _pdId: number = pds.id;
                        let _pdSub: string = pds.subCategory;
                        return {"_pdId":_pdId,"_pdSub":_pdSub};
                    });
                }
            })
            .catch(e => {
                console.error("Failed to fetch Professional designation subcateory");
                return Promise.reject(e);
            });
    }
    private getProficientLanguage(): Promise<string[]> {
        return this.aadSecureGet(
            `${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.ProficientLanguageFetch}`
        )
            .then((response) => {
                return response.json();
            })
            .then((data) => {
                if (data && data.length > 0) {
                    return data.map(lang => {
                        let _langId: number = lang.id;
                        let _langText: string = lang.language;
                        return {"_langId":_langId,"_langText":_langText};
                    });
                }
            })
            .catch(e => {
                console.error("Failed to fetch proficient language");
                return Promise.reject(e);
            });
    }
    private getDisciplines(): Promise<string[]> {
        return this.aadSecureGet(
            `${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.DisciplinesFetch}`
        )
            .then((response) => {
                return response.json();
            })
            .then((data) => {
                if (data && data.length > 0) {
                    return data.map(dis => {
                        let _disciplineId: string = dis.id;
                        let _disciplineName: string = dis.name;
                        let _disciplineFriendlyName: string = dis.friendlyName;
                        return {"_disciplineId": _disciplineId,"_disciplineName": _disciplineName,"_disciplineFriendlyName":_disciplineFriendlyName};
                    });
                }
            })
            .catch(e => {
                console.error("Failed to fetch Disciplines");
                return Promise.reject(e);
            });
    }


    public getEmployeeInformation(financeUserId:string): Promise<string[]> {
      return this.aadSecureGet(
          `${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.EmployeeInformationByFinanceUserId}/${financeUserId}`
      )
          .then((response) => {
              return response.json();
          })
          .then((data) => {
            if (data && data.length > 0) {
                  return data.map(pdData => {
                      let _pdSubCategory: string[] = pdData.subCategory;
                      let _pdDesignation: string[] = pdData.professionalDesignation;
                      let _pdDiscipline: string[] = pdData.disciplineId;
                      return {"pdDesignation": _pdDesignation, "_employeeSubCategory": _pdSubCategory, pdDiscipline: _pdDiscipline};
                  });
              }
          })
          .catch(e => {
              console.error("Failed to fetch Disciplines");
              return Promise.reject(e);
          });
    }

    public async getMasterDetails(): Promise<IMasterDetails> {
        let _masterDetails: IMasterDetails = null;
        try {
            const attachmentTypePromise: Promise<string[]> = this.getAttachmentTypeChoices();
            const disciplinePromise: Promise<string[]> = this.getDisciplines();
            const profDesignationPromise: Promise<string[]> = this.getProfessionalDesignations();
            const proficientLanguagePromise: Promise<string[]> = this.getProficientLanguage();
            const pdSubCategoryPromise: Promise<string[]> = this.getPDSubcategory();
            return Promise.all([attachmentTypePromise, disciplinePromise, profDesignationPromise, proficientLanguagePromise, pdSubCategoryPromise])
                .then(([attachmentTypeChoices, disciplines, professionalDesignation, proficientLanguages, pdSubCategory]) => {
                    _masterDetails = {
                        professionalDesignation: professionalDesignation,
                        discipline: disciplines,
                        attachmentType: attachmentTypeChoices,
                        pdSubCategory: pdSubCategory,
                        language: proficientLanguages

                    };
                    return Promise.resolve(_masterDetails);
                });



        }
        catch (e) {
            return Promise.reject(e);
        }
    }
}
