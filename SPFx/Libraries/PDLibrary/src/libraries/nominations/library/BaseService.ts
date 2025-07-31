import { SPHttpClient, IHttpClientOptions, ISPHttpClientBatchCreationOptions, SPHttpClientResponse, SPHttpClientBatch, AadHttpClient, SPHttpClientConfiguration, AadHttpClientConfiguration } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ExtensionContext } from '@microsoft/sp-extension-base';
import { ConstantsConfig } from './startup/InitializeConstants';
import { IConstants, QueryType } from '../models/IConstants';
import { PnpBatchReuqest } from '../models/PnpBatchReuqest';
import { sp, IItemAddResult, ICamlQuery, ISiteGroupInfo } from "@pnp/sp/presets/all";
import { ISiteUserInfo } from '@pnp/sp/site-users/types';

export default class BaseService {

    protected Constants: IConstants = null;

    protected _webpartContext: WebPartContext | ExtensionContext;

    constructor(ctx: WebPartContext | ExtensionContext) {
        this.getAadHttpClient = this.getAadHttpClient.bind(this);
        this._webpartContext = ctx;
        this.Constants = ConstantsConfig.GetConstants();
        sp.setup({
            spfxContext: ctx
        });
    }


    protected spGet(queryEndpoint: string, HttpConfig?: SPHttpClientConfiguration, options?: IHttpClientOptions): Promise<any> {


        let _httpConfig = HttpConfig != null ? HttpConfig : SPHttpClient.configurations.v1;
        let spOptions: IHttpClientOptions = options ? options : {
            headers: {
                'Accept': 'application/json;odata=nometadata',
            }
        };
        return this._webpartContext.spHttpClient.get(
            queryEndpoint,
            _httpConfig,
            options
        )
            .then(
                (response: any) => {
                    if (response.status >= 200 && response.status < 300) {
                        return response;
                    } else {
                        return Promise.reject(JSON.stringify(response));
                    }
                });
    }

    protected async spGetByCamlQuery(list: string, camlQuery): Promise<any> {
        const caml: ICamlQuery = {
            ViewXml: camlQuery,
        };
        return await sp.web.lists.getByTitle(list).renderListDataAsStream(caml).then((data) => {
            return Promise.resolve(data && data.Row);
        }).catch((error) => {
            console.log(error);
        });

    }
    protected async spPost(request: PnpBatchReuqest): Promise<any> {

        if (request && request.type == QueryType.ADD) {
            const iar: IItemAddResult = await sp.web.lists.getByTitle(request.list).items.add(request.data);
            return Promise.resolve(iar.data.Id);
        }
        else {
            return Promise.reject(null);
        }
    }

    protected aadSecureGet(postEndpoint: string, aadHttpConfig?: AadHttpClientConfiguration): Promise<any> {

        let _aadHttpConfig = aadHttpConfig != null ? aadHttpConfig : AadHttpClient.configurations.v1;
        return this.getAadHttpClient().then((aadHttpClient: AadHttpClient): Promise<any> => {
            return aadHttpClient.get(
                postEndpoint,
                _aadHttpConfig
            )
                .then((response) => {
                    if (response.status >= 200 && response.status < 300) {
                        return response;
                    } else {
                        return Promise.reject(null);
                    }
                })
                .catch(e => {
                    console.log(e);
                });
        });

    }

    protected aadSecureFetch(postEndpoint: string, httpClientOptions: IHttpClientOptions, aadHttpConfig?: AadHttpClientConfiguration): Promise<any> {

        let _aadHttpConfig = aadHttpConfig != null ? aadHttpConfig : AadHttpClient.configurations.v1;

        return this.getAadHttpClient().then((aadHttpClient: AadHttpClient): Promise<any> => {
            return aadHttpClient.fetch(
                postEndpoint,
                _aadHttpConfig,
                httpClientOptions,
            )
                .then((response) => {
                    if (response.status >= 200 && response.status < 300) {
                        return response;
                    } else {
                        return Promise.reject(JSON.stringify(response));
                    }
                });
        });

    }




    private getAadHttpClient(): Promise<AadHttpClient> {

        return this._webpartContext.aadHttpClientFactory
            .getClient(this.Constants.ENDPOINTS.AadClientResourceIdentifier)
            .then((client: AadHttpClient): AadHttpClient => {
                return client;
            })
            .catch(e => {
                console.log(e);
                return Promise.reject(null);
            });
    }
    protected async createDocumentSet(libraryName: string, folderName: string, documentSetContentTypeId: string, folders: string[],nominationId : number): Promise<any> {
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


}
