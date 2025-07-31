
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICommonMethods } from "./ICommonMethods";

export default class CommonMethods  implements ICommonMethods {
 

    private static _canvas: HTMLCanvasElement | undefined = undefined;
    public parse_query_string(query: string): Promise<string> {
        let vars = query.split("&");
        let query_string;
        for (var i = 0; i < vars.length; i++) {
            
            var pair = vars[i].split("=");
            var key = decodeURIComponent(pair[0]);
            var value = decodeURIComponent(pair[1]);
            
            // If first entry with this name
            if (typeof query_string[key] === "undefined") {
                query_string[key] = decodeURIComponent(value);
                // If second entry with this name
            } 
            else if (typeof query_string[key] === "string") {
                var arr = [query_string[key], decodeURIComponent(value)];
                query_string[key] = arr;
            // If third or later entry with this name
            } 
            else {
                query_string[key].push(decodeURIComponent(value));
            }
        }
        return query_string;
    }

    public static getSPFormatDate = (date: Date) => {
        if (date) {
            //return date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate() + "T07:00:00Z";
            //}
            var current = new Date();
            return new Date(date.getFullYear() + "/" + (date.getMonth() + 1) + "/" + date.getDate() + " " + current.getHours() + ":" + current.getMinutes() + ":" + current.getSeconds()).toISOString();
        }
        else
            return null;
    }

    public static getTextWidth(text?: string, font?: string): number {
        if (text == null) { return 0; }
        // re-use canvas object for better performance
        const canvas = this._canvas || (this._canvas = document.createElement("canvas"));
        const context = canvas.getContext("2d");
        if (context == null) { return 0; }
        if (font != null) {
            context.font = font;
        }
        const metrics = context.measureText(text);
        return metrics.width;
    }

    public static getFieldWidth(data: any[], label: string, field: string, min: number) {
        const font = '12px "Segoe UI"';
        const labelWidth = this.getTextWidth(label, font);
        if (labelWidth > min) { min = labelWidth; }
        data.forEach(d => {
            let obj: any = d;
            const props = field.split('.');
            if (props.length > 0) {
                for (let i = 0; i < props.length; i++) {
                    obj = obj[props[i]];
                }
            } else { obj = obj[field]; }
            if (obj != null && obj.length > 0) {
                const l = this.getTextWidth(obj, font);
                if (l > min) { min = l; }
            }
        });
        return min;
    }

    public static setPermissionOnAttachment(context: WebPartContext, attachmentRootFolderNameValue:string,Actor:string, libraryDisplayNameVal: string, ptpacReviewer?: any)
    {
        if(attachmentRootFolderNameValue){
            const body: string = JSON.stringify({
                'currentUser': context.pageContext.user.email,
                'siteUrl': context.pageContext.site.absoluteUrl,
                'libraryDisplayName': libraryDisplayNameVal,
                'attachmentRootFolderName': attachmentRootFolderNameValue,
                'buttonAction':"Intake Submission",
                'actor': Actor,
                'ptpacReviewerEmail':ptpacReviewer
            });
            return body;
        }
    }

    
}