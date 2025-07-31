import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ExtensionContext } from '@microsoft/sp-extension-base';
import { QueryType } from '../models/IConstants';
import { PnpAttachmentsRequest, PnpBatchReuqest } from '../models/PnpBatchReuqest';
import { sp, ISiteGroupInfo, PermissionKind } from "@pnp/sp/presets/all";
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import BaseService from './BaseService';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { ISpUser } from '../models/ISpUser';

export default class SPService extends BaseService {

    private serverRelativeUrl = null;

    protected _webpartContext: WebPartContext | ExtensionContext;

    constructor(ctx: WebPartContext | ExtensionContext) {
        super(ctx);
        this.serverRelativeUrl = this._webpartContext.pageContext.site.serverRelativeUrl;
    }
    protected async isUserProfileActive(email: string): Promise<boolean> {
        let loginName = "i:0#.f|membership|" + email;
        let employmentStatus = await sp.profiles.getUserProfilePropertyFor(loginName, "EmploymentStatus");

        return employmentStatus && employmentStatus.toLowerCase().indexOf("current") > -1 ? Promise.resolve(true) : Promise.resolve(false);
    }
    protected async getCurrentUserGroup(): Promise<ISiteGroupInfo[]> {
        let groups = await sp.web.currentUser.groups();
        return groups;
    }
    protected async getCurrentSPUser(): Promise<ISiteUserInfo> {
        return await sp.web.currentUser.get();
    }
    protected async spBatchGet(batchArray: PnpBatchReuqest[]): Promise<any> {
        var batch = sp.web.createBatch();
        let allCallsResult = [];
        if (batchArray) {
            batchArray.forEach(async (element: PnpBatchReuqest) => {
                if (element.type == QueryType.GETITEM) {
                    if (element.id)
                        sp.web.lists.getByTitle(element.list).items.getById(element.id).expand(element.expand).select(element.select).inBatch(batch).get().then((resultSet) => {
                            allCallsResult.push(resultSet);
                        });
                    else
                        sp.web.lists.getByTitle(element.list).items.inBatch(batch).filter(element.filter).expand(element.expand).select(element.select).get().then((resultSet) => {
                            allCallsResult.push(resultSet);
                        });

                }
                if (element.type == QueryType.GETFILES) {

                    let folderUrl = this.serverRelativeUrl + "/" + element.list + "/" + element.docSet + "/" + element.folder;
                    sp.web
                        .getFolderByServerRelativeUrl(folderUrl)
                        .files
                        .expand('ListItemAllFields').inBatch(batch)
                        .get()
                        .then((files: any[]) => {
                            allCallsResult.push(files);
                        });
                }
            });

            await batch.execute().then(async () => {

            });

        }
        return Promise.resolve(allCallsResult);

    }
    protected async spBatchPostAll(batchArray: PnpBatchReuqest[], attachmentRequest?: PnpAttachmentsRequest): Promise<any> {

        var batch = sp.web.createBatch();
        let allCalls = [];
        let allResponse = [];
        if (batchArray) {
            batchArray.forEach((element: PnpBatchReuqest) => {
                if (element.type == QueryType.UPDATE) {
                    let list = sp.web.lists.getByTitle(element.list);
                    const updatedItem = list.items.getById(element.id).inBatch(batch).update(element.data);
                    allCalls.push(updatedItem);
                }
                if (element.type == QueryType.ADD) {
                    let list = sp.web.lists.getByTitle(element.list);
                    const addedItem = list.items.inBatch(batch).add(element.data);
                    allCalls.push(addedItem);
                }
                if (element.type == QueryType.DELETE) {
                    let list = sp.web.lists.getByTitle(element.list);
                    list.items.getById(element.id).inBatch(batch).delete();
                }

            });
        }
        if (attachmentRequest && attachmentRequest.list && attachmentRequest.docSet) {
            let docSetUrl = this.serverRelativeUrl + "/" + attachmentRequest.list + "/" + attachmentRequest.docSet;
            attachmentRequest.attachments.forEach((attachment) => {
                if (attachment.file) {
                    sp.web.getFolderByServerRelativeUrl(docSetUrl + "/" + attachment.attachmentBy).files.inBatch(batch).add(attachment.file.name, attachment.file, true).then(async (addedFile) => {
                        if (attachment.attachmentType) {
                            await addedFile.file.getItem().then(async (item) => {
                                console.info("added file :" + attachment.file.name);
                                await item.update({
                                    Title: attachment.attachmentName,
                                    AttachmentType: attachment.attachmentType
                                }).then((data) => {
                                    console.info("Updated file :" + attachment.file.name);
                                });
                            });
                        }
                    });
                }
                else if (!attachment.id) {
                    sp.web.getFolderByServerRelativeUrl(docSetUrl + "/" + attachment.attachmentBy).files.getByName(attachment.attachmentName).inBatch(batch).delete();
                }

            });
        }

        await batch.execute().then(async () => {
            await allCalls.forEach(async(element, index) => {
                await element.then(async (res: any) => {
                    allResponse.push(res.data);
                }).catch(e => {
                    allResponse.push(e);
                });
            });

        });
        return Promise.resolve(allResponse);
    }


    protected async createDocumentSet(libraryName: string, folderName: string, documentSetContentTypeId: string, folders: string[]): Promise<any> {
        try {
            var batch = sp.web.createBatch();
            const lib = sp.web.lists.getByTitle(libraryName).inBatch(batch);
            lib.rootFolder.inBatch(batch).addSubFolderUsingPath(folderName);
            if (folders) {
                folders.forEach(element => {
                    lib.rootFolder.folders.getByName(folderName).inBatch(batch).addSubFolderUsingPath(element);
                });
            }
            lib.rootFolder.folders.getByName(folderName).listItemAllFields.inBatch(batch).get().then(async item => {
                await lib.items.getById(item.Id).update({
                    ContentTypeId: documentSetContentTypeId
                });
            });
            await batch.execute().then(async (e) => {
              Promise.resolve(true);
            });
        }
        catch {
            Promise.reject(false);
        }

    }

    protected async assignItemLevelPermission(libraryName: string, usersForPermissions: ISpUser[], folderName: string, groupName: string[]): Promise<boolean> {
      try {
          const {Id: fullControlDefId} = await sp.web.roleDefinitions.getByName('Full Control').get();
          const folder =  sp.web
                      .getFolderByServerRelativePath(libraryName+"/"+folderName);
          const item = await folder
          .select('ListItemAllFields/Id','ListItemAllFields/HasUniqueRoleAssignments')
          .expand('ListItemAllFields','HasUniqueRoleAssignments')
          .get<{ ListItemAllFields: { Id: number, HasUniqueRoleAssignments: boolean }}>();
          let user = await sp.web.currentUser();

          const list = sp.web.lists.getByTitle(libraryName);
          await list.breakRoleInheritance(false);

            let permission = list.items.getById(item.ListItemAllFields.Id);

            await permission.breakRoleInheritance(false, true);

              if(groupName && groupName.length >0){
                /*
                if(!permission.userHasPermissions(usersForPermissions[i],PermissionKind.ViewListItems))
                {
                  sp.web.siteGroups.getByName(groupName.toString()).users
                  .add(usersForPermissions[i]).then(function(d){
                    console.info("Added user to Group");
                  });
                }
                */

                if(!item.ListItemAllFields.HasUniqueRoleAssignments){
                  permission.roleAssignments.add(user.Id ,fullControlDefId);
                }

                for(var i=0; i< groupName.length; i++) {
                  let assignGroupName = await sp.web.siteGroups.getByName(groupName[i].toString()).get();
                  await permission.roleAssignments.add(assignGroupName.Id ,fullControlDefId);
                }
              }
              else if(usersForPermissions && usersForPermissions.length > 0 ){
                for(var j=0; j< usersForPermissions.length; j++) {
                  await permission.roleAssignments.add(usersForPermissions[j].id ,fullControlDefId);
                  await sp.web.siteGroups.getByName("Professional Designation Nomination PTPAC Reviewer Group").users.add("i:0#.f|membership|" + usersForPermissions[j].email);
                }
              }
              return Promise.resolve(true);
      }
      catch {
          Promise.reject(false);
      }

    }

    protected async deleteDocumentSet(libraryName: string, folderName: string): Promise<boolean> {
        try {
            sp.web.lists.getByTitle(libraryName).rootFolder.folders.getByName(folderName).delete();
            return true;
        }
        catch {
            Promise.reject(false);
        }
    }

    protected async delFile(libraryName: string, folderName: string, subFolderName: string, fileName: string): Promise<boolean> {
      try {
          sp.web.lists.getByTitle(libraryName).rootFolder.folders.getByName(folderName).folders.getByName(subFolderName).files.getByName(fileName).delete();
          return true;
      }
      catch {
          Promise.reject(false);
      }

    }


    public async callPowerAutomate(body: string, postURL: string) {
      try {
          const requestHeaders: Headers = new Headers();
          requestHeaders.append('Content-type', 'application/json');

          const httpClientOptions: IHttpClientOptions = {
              body: body,
              headers: requestHeaders
          };

          const response = await this._webpartContext.httpClient.post(
              postURL,
              HttpClient.configurations.v1,
              httpClientOptions
          );

          if (response.status === 200) {
              console.info("Successfully Completed");
              return await response.json();
          } else {
              console.error(`Request Error. Status: ${response.status}`);
              throw new Error("Failed to complete request.");
          }
      } catch (error) {
          console.error("Error in Request:", error);
          throw error;
      }
    }


}
