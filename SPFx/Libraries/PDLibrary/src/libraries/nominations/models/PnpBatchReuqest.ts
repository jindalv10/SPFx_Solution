import { IAttachment } from "../models/IAttachment";

export type PnpBatchReuqest = {
    list: string;
    id?: number;
    filter?: string;
    docSet?: string;
    data?: object;
    type: string;
    folder?: string;
    expand?: string;
    select?:string;
};

export type PnpAttachmentsRequest = {
    list: string;
    docSet: string;
    attachments: IAttachment[];
};