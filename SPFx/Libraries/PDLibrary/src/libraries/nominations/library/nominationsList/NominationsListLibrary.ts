import { ISiteGroupInfo, PromotedState } from "@pnp/sp/presets/all";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { AllRoles, NominationStatus, PdStatus, QCReviewStatus, ReviewStatus } from "../../models/IConstants";
import { IEmployee, IEmployeeUpdateProperties } from "../../models/IEmployee";
import { INominationListViewItem } from "../../models/INominationListViewItem";
import { INominationReviewer } from "../../models/INominationReviewer";
import { ISpUser } from "../../models/ISpUser";
import { IUserDetails } from "../../models/IUserDetails";
import SPService from "../SPService";
import { Mapper } from "../startup/Mapper";
import { INominationsListLibrary } from "./INominationsListLibrary";
import { IHttpClientOptions} from '@microsoft/sp-http';
import { IDiscipline } from "../../models/IDiscipline";
import { IProfessionalDesignationDetailed } from "../../models/IProfessionalDesignation";
import { Utility } from "../startup/Utility";

export default class NominationsListLibrary extends SPService implements INominationsListLibrary {

    private currentUrl: string = null;
    constructor(context: any) {
        super(context);
        this.currentUrl = context.pageContext.web.absoluteUrl;

    }
    private async GetCurrentUserPractice(): Promise<string> {
        let practice = null;
        let currentUserEmail = this._webpartContext.pageContext.user.email;
        if(this.Constants.SP_TEST_USER)
        {
          currentUserEmail = this.Constants.SP_TEST_USER;
        }
        if (currentUserEmail) {
            try {
                let employee = await this.aadSecureGet(`${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.NomineeDetailsFetch}/${currentUserEmail}`)
                    .then((response) => {
                        if (response.status)
                            return response.json();
                    });
                if (employee) {
                    practice = employee.practice;
                }
            }

            catch (e) {
                console.error("Failed to fetch User Practice");
                return Promise.reject(e);
            }
        }
        return Promise.resolve(practice);
    }
    private async CheckIfUserExistInLegalGroup(): Promise<boolean> {
        const userGroups: ISiteGroupInfo[] = await this.getCurrentUserGroup();
        let isUserInLegalGroup = false;
        if (userGroups) {
            let legalGroup = userGroups.filter((m) => {
                return m.Title == this.Constants.SP_Group.LegalGroup;
            });
            if (legalGroup) {
                isUserInLegalGroup = true;
            }
        }
        return isUserInLegalGroup;
    }
    private async CurrentUserPTPACDiscipline(user: ISiteUserInfo): Promise<any[]> {
        let queryEndpoint = this.currentUrl + "/_api/web/lists/GetByTitle('" + this.Constants.SP_LIST_NAMES.NominationReviewersList + "')/items?";
        queryEndpoint += "&$select=AuthorizedPTPACId,PDDiscipline";
        let discipline = null;
        try {
            let data = await this.spGet(queryEndpoint).then((response) => { return response.json(); });
            if (data && data.value && data.value.length > 0) {
                let ptpacDisciplineRow = data.value.filter((m) => {
                    return m.AuthorizedPTPACId && m.AuthorizedPTPACId.indexOf(user.Id) > -1;
                });

                discipline = ptpacDisciplineRow && ptpacDisciplineRow.length > 0 && ptpacDisciplineRow.map(disc => disc.PDDiscipline);
            }
        }
        catch (e) {
            console.error("Error in method CurrentUserPTPACDiscipline");
        }
        return discipline;
    }
    private async CurrentUserQCDiscipline(user: ISiteUserInfo): Promise<any[]> {
        let queryEndpoint = this.currentUrl + "/_api/web/lists/GetByTitle('" + this.Constants.SP_LIST_NAMES.NominationReviewersList + "')/items?";
        queryEndpoint += "&$select=AuthorizedQCId,PDDiscipline";
        let discipline = null;
        try {
            let data = await this.spGet(queryEndpoint).then((response) => { return response.json(); });
            if (data && data.value && data.value.length > 0) {
                let qcDisciplineRow = data.value.filter((m) => {
                    return m.AuthorizedQCId && m.AuthorizedQCId.indexOf(user.Id) > -1;
                });

                discipline = qcDisciplineRow && qcDisciplineRow.length > 0 && qcDisciplineRow.map(disc => disc.PDDiscipline);
            }
        }
        catch (e) {
            console.error("Error in method CurrentUserQCDiscipline");
        }
        return discipline;
    }

    public async getQCDisciplineUsers(pdDisciplineVal: string): Promise<INominationReviewer[]> {
    if(pdDisciplineVal){
      try {
      let camlQueryString = `<View><Query><Where>`;
            camlQueryString += pdDisciplineVal ? `<Eq><FieldRef Name='PDDiscipline' /><Value Type='Text'>${pdDisciplineVal}</Value></Eq>`: "";
            camlQueryString += `</Where><OrderBy><FieldRef Name='Id' Ascending='True'/></OrderBy></Query></View>`;
            const notificationData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.NominationReviewersList, camlQueryString);
            if (notificationData) {
                return notificationData.map((element) => {
                    return Mapper.mapQCReviewerDisciplineDetails(element);
                });

            }
      }
      catch (e) {
          console.error("Error in method CurrentUserQCDiscipline");
      }
    }
  }

  public async getNominationList(currentUser: IUserDetails): Promise<INominationListViewItem[]> {
      const isNominator = Utility.ciEquals(currentUser.role, AllRoles.NOMINATOR);
      const isLocalAdmin = Utility.ciEquals(currentUser.role, AllRoles.LA);
      const isQualityCoordinator = Utility.ciEquals(currentUser.role, AllRoles.QC);
      const isPTPACChair = Utility.ciEquals(currentUser.role, AllRoles.PTPAC_CHAIR);
      const isPTPACReviewer = Utility.ciEquals(currentUser.role, AllRoles.PTPAC_REVIEWER);
      const isLegal = Utility.ciEquals(currentUser.role, AllRoles.LEGAL);

      const user: ISiteUserInfo = await this.getCurrentSPUser();
      if (Utility.ciEquals(currentUser.role, AllRoles.EP_NOMINATOR)) {
          let userPractice = await this.GetCurrentUserPractice();
          let camlQueryString = `<View><Query><Where><Neq><FieldRef Name='NominationStatus' /><Value Type='Text'>${NominationStatus.DraftByNominator}</Value></Neq>`;
          camlQueryString += userPractice ? `<Or>` : "";
          camlQueryString += `<Contains><FieldRef Name='EPNominator' LookupId='True'/><Value Type='Lookup'>${user.Id}</Value></Contains>`;
          camlQueryString += userPractice ? `<Eq><FieldRef Name='NomineePractice' /><Value Type='Text'>${userPractice}</Value></Eq>` : "";
          camlQueryString += userPractice ? `</Or>` : "";
          camlQueryString += `</Where><OrderBy><FieldRef Name='SubmissionDate' Ascending='True'/><FieldRef Name='NominationStatus' Ascending='True' /></OrderBy></Query></View>`;
          const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
          if (nominationsData) {
              return nominationsData.map((element) => {
                  return Mapper.mapNominationListDetails(element);
              });

          }
      }
      if (isNominator && user) {
          //let camlQueryString = `<View><Query><Where><Neq><FieldRef Name='Author' LookupId='True'/><Value Type='Lookup'>` + user.Id + `</Value></Neq></Where></Query></View>`;
          let camlQueryString = `<View><Query><Where><And><Or><Neq><FieldRef Name='NominationStatus' /><Value Type='Text'>${NominationStatus.ApproveCompleted}</Value></Neq><Neq><FieldRef Name='NominationStatus' /><Value Type='Text'>${NominationStatus.WithdrawnCompleted}</Value></Neq></Or><Eq><FieldRef Name='Author' LookupId='True'/><Value Type='Lookup'>` + user.Id + `</Value></Eq></And></Where></Query></View>`;

          const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
          if (nominationsData) {
              return nominationsData.map((element) => {
                  return Mapper.mapNominationListDetails(element);
              });

          }
      }
      if (isLegal && user) {
          const isUserInLegalGroup = await this.CheckIfUserExistInLegalGroup();
          if (isUserInLegalGroup) {
              let camlQueryString = `<View><Query><Where><Or><Eq><FieldRef Name='NominationStatus'/><Value Type='Text'>" ${NominationStatus.PendingWithQCAndLegal}"</Value></Eq><Eq><FieldRef Name='NominationStatus'/><Value Type='Text'>" ${NominationStatus.SubmittedByQCPendingWithLegal}"</Value></Eq></Or></Where>`;
              camlQueryString += `<OrderBy><FieldRef Name='SubmissionDate' Ascending='True'/><FieldRef Name='NominationStatus' Ascending='True' /></OrderBy></Query></View>`;

              const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
              if (nominationsData) {
                  return nominationsData.map((element) => {
                      return Mapper.mapNominationListDetails(element);
                  });
              }
          }
      }
      if (isLocalAdmin && user) {
          let queryEndpoint = this.currentUrl + "/_api/web/lists/GetByTitle('" + this.Constants.SP_LIST_NAMES.LANominationList + "')/items?";
          queryEndpoint += "&$filter=(AssigneeId eq '" + user.Id + "') and ((Nomintaion/NominationStatus eq '" + NominationStatus.PendingWithLocalAdmin + "') or (Nomintaion/NominationStatus eq '" + NominationStatus.WithdrawRPStatus + "'))";
          queryEndpoint += "&$top=5000&$expand=Nomintaion";
          queryEndpoint += "&$select=NomintaionId,Nomintaion/NominationStatus";

          return this.spGet(queryEndpoint).then((response) => { return response.json(); })
              .then(async (data) => {
                  if (data && data.value && data.value.length > 0) {

                      let camlQueryString = `<View><Query><Where><In><FieldRef Name="ID"/><Values>`;
                      data.value.forEach(element => {
                          camlQueryString += `<Value Type="Number">${element.NomintaionId}</Value>`;
                      });
                      camlQueryString += `</Values></In></Where>`;
                      camlQueryString += `<OrderBy><FieldRef Name='SubmissionDate' Ascending='True'/><FieldRef Name='NominationStatus' Ascending='True' /></OrderBy></Query></View>`;

                      const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
                      if (nominationsData) {
                          return nominationsData.map((element) => {
                              return Mapper.mapNominationListDetails(element);
                          });

                      }
                  }
                  else {

                  }

              })
              .catch(e => {
                  console.error("Failed to Nomination List");
                  return Promise.reject([]);
              });
      }


      if (isQualityCoordinator && user) {
          let QCListDetails = null;
          let queryEndpoint = this.currentUrl + "/_api/web/lists/GetByTitle('" + this.Constants.SP_LIST_NAMES.QCNominationList + "')/items?";
          //queryEndpoint += "&$filter=(((ReviewerId eq '" + user.Id + "') or (AdditionalReviewerId eq '" + user.Id + "')))";

          queryEndpoint += "$top=1000&$expand=Nomintaion&$select=Id,NomintaionId,Nomintaion/NominationStatus,SentToScDate,NominationPasses,ReferencesPassed,QARPassed&$orderby=Id desc";

          let currentUserDiscipline: any[] = await this.CurrentUserQCDiscipline(user);
          if(currentUserDiscipline && currentUserDiscipline.length > 0 ){
            return this.spGet(queryEndpoint).then((response) => { return response.json(); })
                .then(async (data) => {
                    let camlQueryString = `<View><Query><Where>`;
                    if (data && data.value && data.value.length > 0) {
                        QCListDetails =  data && data.value && data.value.length  > 0 ? data.value.map((element) => {return Mapper.mapQcDetails(element);}) : null;
                         //camlQueryString += currentUserDiscipline ? `<And>` : "";
                        //camlQueryString += `<In><FieldRef Name="ID"/><Values>`;
                        //data.value.forEach(element => {
                        //    camlQueryString += `<Value Type="Number">${element.NomintaionId}</Value>`;
                        //});
                        //camlQueryString += `</Values></In>`;
                    }
                    camlQueryString += `<Or><Or><Or><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.PendingWithQC}</Value></Eq><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.ApproveCompleted}</Value></Eq></Or><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.PendingWithPTPACChair}</Value></Eq></Or><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.PendingWithPTPACReviewer}</Value></Eq></Or>`;

                    //camlQueryString += currentUserDiscipline ? `<And><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.PendingWithQC}</Value></Eq>` : "";
                    //camlQueryString += currentUserDiscipline && currentUserDiscipline.length > 1 ? `<Or>` : ""
                    //camlQueryString += currentUserDiscipline && currentUserDiscipline.length > 1 ? currentUserDiscipline.map(elem => `<Eq><FieldRef Name='PDDiscipline' /><Value Type="Text">${elem}</Value></Eq>`).join("")  : currentUserDiscipline.map(elem => `<Eq><FieldRef Name='PDDiscipline' /><Value Type="Text">${elem}</Value></Eq>`).join("");
                    //camlQueryString += currentUserDiscipline && currentUserDiscipline.length > 1 ? `</Or>` :""
                    //camlQueryString += currentUserDiscipline ? `</And>` : "";
                    //camlQueryString += currentUserDiscipline && data && data.value && data.value.length > 0 ? `</Or>` : "";
                    camlQueryString += `</Where><OrderBy><FieldRef Name='SubmissionDate' Ascending='True'/><FieldRef Name='NominationStatus' Ascending='True' /></OrderBy></Query></View>`;
                    const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
                    if (nominationsData) {
                      const filtered =  QCListDetails ? Utility.mergeByID("ID","nominationId", nominationsData.filter(element => {return currentUserDiscipline.indexOf(element.PDDiscipline) !== -1;}), QCListDetails): nominationsData.filter(element => {return currentUserDiscipline.indexOf(element.PDDiscipline) !== -1;});
                        //var filtered: any = nominationsData.filter(function(element) {return currentUserDiscipline.indexOf(element.PDDiscipline) !== -1});
                        return filtered.map((element) => {
                            return Mapper.mapNominationListDetails(element);
                        });
                    }

                })
                .catch(e => {
                    console.error("Failed to Nomination List");
                    return Promise.reject([]);
                });
            }
      }

      if (isPTPACChair && user) {
        let PTPACListDetails = null;
        let queryEndpoint = this.currentUrl + "/_api/web/lists/GetByTitle('" + this.Constants.SP_LIST_NAMES.PtpacNominationList + "')/items?";
          queryEndpoint += "&$expand=Nomintaion,Reviewer/Id&$select=NomintaionId,Nomintaion/NominationStatus,ReviewDueDate,internalReviewDueDate,Reviewer/Name,Reviewer/Title&$filter=Nomintaion/NominationStatus ne ''&$top=5000";


        let currentUserDiscipline: any[] = await this.CurrentUserPTPACDiscipline(user);
        if(currentUserDiscipline){
          return this.spGet(queryEndpoint).then((response) => { return response.json(); })
              .then(async (data) => {
                  let camlQueryString = `<View><Query><Where>`;
                  if (data && data.value && data.value.length > 0) {
                    PTPACListDetails =  data && data.value && data.value.length  > 0 ? data.value.map((element) => {return Mapper.mapPTPACDetails(element);}) : null;
                      //camlQueryString += currentUserDiscipline ? `<And>` : "";
                      camlQueryString += `<In><FieldRef Name="ID"/><Values>`;
                      data.value.forEach(element => {
                          camlQueryString += `<Value Type="Number">${element.NomintaionId}</Value>`;
                      });
                      camlQueryString += `</Values></In>`;

                  }
                  camlQueryString += `<Or><Or><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.ApproveCompleted}</Value></Eq><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.PendingWithPTPACChair}</Value></Eq></Or><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.PendingWithPTPACReviewer}</Value></Eq></Or>`;

                  //camlQueryString += currentUserDiscipline ? `<And><Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.PendingWithQC}</Value></Eq>` : "";
                  //camlQueryString += currentUserDiscipline && currentUserDiscipline.length > 1 ? `<Or>` : ""
                  //camlQueryString += currentUserDiscipline && currentUserDiscipline.length > 1 ? currentUserDiscipline.map(elem => `<Eq><FieldRef Name='PDDiscipline' /><Value Type="Text">${elem}</Value></Eq>`).join("")  : currentUserDiscipline.map(elem => `<Eq><FieldRef Name='PDDiscipline' /><Value Type="Text">${elem}</Value></Eq>`).join("");
                  //camlQueryString += currentUserDiscipline && currentUserDiscipline.length > 1 ? `</Or>` :""
                  //camlQueryString += currentUserDiscipline ? `</And>` : "";
                  //camlQueryString += currentUserDiscipline && data && data.value && data.value.length > 0 ? `</And>` : "";
                  camlQueryString += `</Where><OrderBy><FieldRef Name='SubmissionDate' Ascending='True'/><FieldRef Name='NominationStatus' Ascending='True' /></OrderBy></Query><RowLimit>4999</RowLimit></View>`;
                  const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
                  if (nominationsData) {
                    const filtered =  PTPACListDetails ? Utility.mergeByID("ID","nominationId", nominationsData.filter(element => {return currentUserDiscipline.indexOf(element.PDDiscipline) !== -1;}), PTPACListDetails): nominationsData.filter(element => {return currentUserDiscipline.indexOf(element.PDDiscipline) !== -1;});
                      //var filtered: any = nominationsData.filter(function(element) {return currentUserDiscipline.indexOf(element.PDDiscipline) !== -1});
                      return filtered.map((element) => {
                          return Mapper.mapNominationListDetails(element);
                      });
                  }
              })
              .catch(e => {
                  console.error("Failed to Nomination List");
                  return Promise.reject([]);
              });
          }
    }


      if (isPTPACReviewer && user) {
          let PTPACListDetails = null;
          let queryEndpoint = this.currentUrl + "/_api/web/lists/GetByTitle('" + this.Constants.SP_LIST_NAMES.PtpacNominationList + "')/items?";
          queryEndpoint += "&$filter=(ReviewerId eq '" + user.Id + "')";
          queryEndpoint += "&$expand=Nomintaion,Reviewer/Id&$select=NomintaionId,Nomintaion/NominationStatus,ReviewDueDate,internalReviewDueDate,Reviewer/Name,Reviewer/Title";
          return this.spGet(queryEndpoint).then((response) => { return response.json(); })
              .then(async (data) => {
                  if (data && data.value && data.value.length > 0) {
                      PTPACListDetails =  data && data.value && data.value.length  > 0 ? data.value.map((element) => {return Mapper.mapPTPACDetails(element);}) : null;

                      let camlQueryString = `<View><Query><Where><And>`;
                      camlQueryString += `<In><FieldRef Name="ID"/><Values>`;
                      data.value.forEach(element => {
                          camlQueryString += `<Value Type="Number">${element.NomintaionId}</Value>`;
                      });
                      camlQueryString += `</Values></In>`;
                      camlQueryString += `<Eq><FieldRef Name='NominationStatus' /><Value Type="Text">${NominationStatus.PendingWithPTPACReviewer}</Value></Eq>`;
                      camlQueryString += `</And></Where><OrderBy><FieldRef Name='SubmissionDate' Ascending='True'/><FieldRef Name='NominationStatus' Ascending='True' /></OrderBy></Query></View>`;
                      const nominationsData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.MasterNominationList, camlQueryString);
                      if (nominationsData) {
                        const filtered =  PTPACListDetails ? Utility.mergeByID("ID","nominationId", nominationsData, PTPACListDetails): nominationsData;
                        return filtered.map((element) => {
                          return Mapper.mapNominationListDetails(element);
                        });
                      }
                  }

              })
              .catch(e => {
                  console.error("Failed to Nomination List");
                  return Promise.reject([]);
              });
      }
      return null;
  }

  public async updateNomineeEmployeeDetails(employeeUpdateObject: IEmployeeUpdateProperties): Promise<IEmployee> {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    requestHeaders.append('Accept', 'application/json');

    const payload: string = JSON.stringify(employeeUpdateObject);

    const httpClientOptions: IHttpClientOptions = {
        method: "PUT",
        headers: requestHeaders,
        body: payload
    };
    return this.aadSecureFetch(`${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.EmployeeDetailsUpdate}`,httpClientOptions)
        .then((response) => {
            return response.json();
        })
        .then((data) => {
            //return this.mapToEmployeeFromServerFormat(data);
        })
        .catch(e => {
            return null;
        });
  }

  public getProfessionalDesignationsByFinanceUserId(financeUserID: number, disciplines?: IDiscipline[]): Promise<IProfessionalDesignationDetailed[]> {
    const _endpoint = `${this.Constants.ENDPOINTS.BaseAzureApiUrl}/${this.Constants.ENDPOINTS.EmployeeProfessionalDesignationsByFinanceUserId}/${financeUserID}`;

    return this.aadSecureGet(
        _endpoint
        // `${Constants.Endpoints.BaseAzureApiUrl}/${Constants.Endpoints.EmployeeProfessionalDesignationsByFinanceUserId}/${1638}` //test
    )
    .then((response) => {
          return response.json();
      })
      .then((data) => {
          if (!Utility.isObjectNullOrEmpty(data)) {
              if (data.length == 0) {
                  return [];
              }
              return data.map(pd => {

                  let _professionalDesignation: IProfessionalDesignationDetailed = {
                      id: pd.id,
                      financeUserId: pd.financeUserId,
                      designationId: pd.designationId,
                      pdSubategoryId: pd.pdSubategoryId,
                      discipline: pd.disciplineId,
                      grantedOn: pd.grantedOn ? new Date(pd.grantedOn) : null,
                      removedOn: pd.removedOn ? new Date(pd.removedOn) : null,
                      level: pd.level,
                      restrictions: pd.restrictions,
                      restrictionDate: pd.restrictionDate,
                      professionalDesignation: pd.professionalDesignation,
                      subCategory: pd.subCategory,
                      code: pd.code,
                      abbreviation: pd.abbreviation,
                      friendlyName: pd.friendlyName,
                      created: pd.created ? new Date(pd.created) : null,
                      modified: pd.modified ? new Date(pd.modified) : null,
                      createdBy: pd.createdBy,
                      modifiedBy: pd.modifiedBy,
                  };
                  return _professionalDesignation;
              });
          }
          else {
              return [];
          }
      })
      .catch(e => {
          console.log("Failed to fetch Professional Designations");
          return Promise.reject(null);
      });
}

}
