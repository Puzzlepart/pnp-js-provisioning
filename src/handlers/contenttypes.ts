import { HandlerBase } from "./handlerbase";
import { IContentType, IContentTypeFieldRef } from "../schema";
import { Web, Logger, LogLevel } from "sp-pnp-js";

export class ContentTypes extends HandlerBase {

    constructor() {
        super("ContentTypes");
    }

    public async ProvisionObjects(web: Web, contentTypes: IContentType[]): Promise<void> {
        super.scope_started();
        try {
            if (contentTypes) {
                await contentTypes.reduce((chain, ct) => chain.then(_ => this.processContentType(web, ct)), Promise.resolve());
            }
            super.scope_ended();
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }

    private async processContentType(web: Web, ct: IContentType): Promise<void> {
        const payload: any = {
            Description: ct.Description || "",
            Group: ct.Group || "Custom Content Types",
            Id: { StringValue: ct.StringId },
            Name: ct.Name,
            StringId: ct.StringId,
        };

        if (ct.ParentStringId) {
            payload.Parent = { StringValue: ct.ParentStringId };
        }

        try {
            await web.post("_api/web/contenttypes", {
                body: JSON.stringify(payload),
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-type": "application/json;odata=verbose",
                },
            });
            Logger.log({ level: LogLevel.Info, message: `Content type ${ct.Name} (${ct.StringId}) created.` });
        } catch (err) {
            Logger.log({ level: LogLevel.Warning, message: `Content type ${ct.Name} (${ct.StringId}) create failed, attempting update.` });
            try {
                await web.contentTypes.getById(ct.StringId).update({
                    Description: ct.Description,
                    Group: ct.Group,
                    Name: ct.Name,
                });
                Logger.log({ level: LogLevel.Info, message: `Content type ${ct.Name} (${ct.StringId}) updated.` });
            } catch (updateErr) {
                Logger.log({ level: LogLevel.Warning, message: `Failed to upsert content type ${ct.Name} (${ct.StringId}).` });
            }
        }

        if (ct.FieldRefs && ct.FieldRefs.length > 0) {
            await ct.FieldRefs.reduce((chain, fr) => chain.then(_ => this.processFieldRef(web, ct, fr)), Promise.resolve());
        }
    }

    private async processFieldRef(web: Web, ct: IContentType, fieldRef: IContentTypeFieldRef): Promise<void> {
        try {
            await web.contentTypes.getById(ct.StringId).fieldLinks.add(fieldRef.FieldInternalName);
            Logger.log({ level: LogLevel.Info, message: `Field ${fieldRef.FieldInternalName} linked to content type ${ct.StringId}.` });
        } catch (err) {
            Logger.log({ level: LogLevel.Warning, message: `Failed to link field ${fieldRef.FieldInternalName} to content type ${ct.StringId}.` });
        }
    }
}
