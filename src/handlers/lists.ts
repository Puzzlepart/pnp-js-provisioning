import * as xmljs from "xml-js";
import { HandlerBase } from "./handlerbase";
import { IContentTypeBinding, IList, IListInstanceFieldRef, IListView } from "../schema";
import { Web, List, Logger, LogLevel } from "sp-pnp-js";

/**
 * Describes the Lists Object Handler
 */
export class Lists extends HandlerBase {
    private lists: any[];
    private tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z ]*}/g;

    /**
     * Creates a new instance of the Lists class
     */
    constructor() {
        super("Lists");
        this.lists = [];
    }

    /**
     * Provisioning lists
     *
     * @param {Web} web The web
     * @param {Array<IList>} lists The lists to provision
     */
    public async ProvisionObjects(web: Web, lists: IList[]): Promise<void> {
        super.scope_started();
        try {
            await lists.reduce((chain, list) => chain.then(_ => this.processList(web, list)), Promise.resolve());
            await lists.reduce((chain, list) => chain.then(_ => this.processFields(web, list)), Promise.resolve());
            await lists.reduce((chain, list) => chain.then(_ => this.processFieldRefs(web, list)), Promise.resolve());
            await lists.reduce((chain, list) => chain.then(_ => this.processViews(web, list)), Promise.resolve());
            super.scope_ended();
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }

    /**
     * Processes a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list
     */
    private async processList(web: Web, lc: IList): Promise<void> {
        const { created, list, data } = await web.lists.ensure(lc.Title, lc.Description, lc.Template, lc.ContentTypesEnabled, lc.AdditionalSettings);
        this.lists.push(data);
        if (created) {
            Logger.log({ data: list, level: LogLevel.Info, message: `List ${lc.Title} created successfully.` });
        }
        await this.processContentTypeBindings(lc, list, lc.ContentTypeBindings, lc.RemoveExistingContentTypes);
    }

    /**
     * Processes content type bindings for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {Array<IContentTypeBinding>} contentTypeBindings Content type bindings
     * @param {boolean} removeExisting Remove existing content type bindings
     */
    private async processContentTypeBindings(lc: IList, list: List, contentTypeBindings: IContentTypeBinding[], removeExisting: boolean): Promise<any> {
        if (contentTypeBindings) {
            await contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(lc, list, ct.ContentTypeID)), Promise.resolve());
            if (removeExisting) {
                let promises = [];
                const contentTypes = await list.contentTypes.get();
                contentTypes.forEach(({ Id: { StringValue: ContentTypeId } }) => {
                    let shouldRemove = (contentTypeBindings.filter(ctb => ContentTypeId.indexOf(ctb.ContentTypeID) !== -1).length === 0)
                        && (ContentTypeId.indexOf("0x0120") === -1);
                    if (shouldRemove) {
                        Logger.write(`Removing content type ${ContentTypeId} from list ${lc.Title}`, LogLevel.Info);
                        promises.push(list.contentTypes.getById(ContentTypeId).delete());
                    }
                });
                await Promise.all(promises);
            }
        }
    }

    /**
     * Processes a content type binding for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {string} contentTypeID The Content Type ID
     */
    private async processContentTypeBinding(lc: IList, list: List, contentTypeID: string): Promise<any> {
        await list.contentTypes.addAvailableContentType(contentTypeID);
        Logger.log({ level: LogLevel.Info, message: `Content Type ${contentTypeID} added successfully to list ${lc.Title}.` });
    }


    /**
     * Processes fields for a list
     *
     * @param {Web} web The web
     * @param {IList} list The pnp list
     */
    private async processFields(web: Web, list: IList): Promise<any> {
        if (list.Fields) {
            await list.Fields.reduce((chain, field) => chain.then(_ => this.processField(web, list, field)), Promise.resolve());
        }
    }

    /**
     * Processes a field for a lit
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {string} fieldXml Field xml
     */
    private async processField(web: Web, lc: IList, fieldXml: string): Promise<any> {
        const list = web.lists.getByTitle(lc.Title);
        const fXmlJson = JSON.parse(xmljs.xml2json(fieldXml));
        const fAttr = fXmlJson.elements[0].attributes;
        const internalName = fAttr.InternalName;
        const displayName = fAttr.DisplayName;
        Logger.log({ message: `Processing field ${internalName} (${displayName}) for list ${lc.Title}.`, level: LogLevel.Info });
        fieldXml = xmljs.json2xml(fXmlJson);
        fXmlJson.elements[0].attributes.DisplayName = internalName;

        // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
        try {
            let field = await list.fields.getById(fAttr.ID);
            await field.delete();
        } catch (err) {
            Logger.log({ message: `Failed to remove field '${displayName}' from list ${lc.Title}.`, level: LogLevel.Warning });
        }

        let fieldAddResult = await list.fields.createFieldAsXml(this.replaceFieldXmlTokens(fieldXml));
        await fieldAddResult.field.update({ Title: displayName });
        Logger.log({ message: `Field '${displayName}' added successfully to list ${lc.Title}.`, level: LogLevel.Info });
    }

    /**
   * Processes field refs for a list
   *
   * @param {Web} web The web
   * @param {IList} list The pnp list
   */
    private async processFieldRefs(web: Web, list: IList): Promise<any> {
        if (list.FieldRefs) {
            await list.FieldRefs.reduce((chain, fieldRef) => chain.then(_ => this.processFieldRef(web, list, fieldRef)), Promise.resolve());
        }
    }

    /**
     *
     * Processes a field ref for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListInstanceFieldRef} fieldRef The list field ref
     */
    private async processFieldRef(web: Web, lc: IList, fieldRef: IListInstanceFieldRef): Promise<any> {
        const list = web.lists.getByTitle(lc.Title);
        await list.fields.getById(fieldRef.ID).update({ Hidden: fieldRef.Hidden, Required: fieldRef.Required, Title: fieldRef.DisplayName });
        Logger.log({ data: fieldRef, level: LogLevel.Info, message: `Field '${fieldRef.ID}' updated for list ${lc.Title}.` });
    }

    /**
     * Processes views for a list
     *
     * @param web The web
     * @param lc The list configuration
     */
    private async processViews(web: Web, lc: IList): Promise<any> {
        if (lc.Views) {
            await lc.Views.reduce((chain, view) => chain.then(_ => this.processView(web, lc, view)), Promise.resolve());
        }
    }

    /**
     * Processes a view for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListView} lvc The view configuration
     */
    private async processView(web: Web, lc: IList, lvc: IListView): Promise<void> {
        let view = web.lists.getByTitle(lc.Title).views.getByTitle(lvc.Title);
        try {
            await view.get();
            await view.update(lvc.AdditionalSettings);
            await this.processViewFields(view, lvc.ViewFields);
        } catch (err) {
            const result = await web.lists.getByTitle(lc.Title).views.add(lvc.Title, lvc.PersonalView, lvc.AdditionalSettings);
            Logger.log({ level: LogLevel.Info, message: `View ${lvc.Title} added successfully to list ${lc.Title}.` });
            await this.processViewFields(result.view, lvc.ViewFields);
        }
    }

    /**
     * Processes view fields for a view
     *
     * @param {any} view The pnp view
     * @param {Array<string>} viewFields Array of view fields
     */
    private async processViewFields(view, viewFields: string[]): Promise<void> {
        await view.fields.removeAll();
        await viewFields.reduce((chain, viewField) => chain.then(_ => view.fields.add(viewField)), Promise.resolve());
    }

    /**
     * Replaces tokens in field xml
     *
     * @param {string} fieldXml The field xml
     */
    private replaceFieldXmlTokens(fieldXml: string) {
        let m;
        while ((m = this.tokenRegex.exec(fieldXml)) !== null) {
            if (m.index === this.tokenRegex.lastIndex) {
                this.tokenRegex.lastIndex++;
            }
            m.forEach((match) => {
                let [Type, Value] = match.replace(/[\{\}]/g, "").split(":");
                switch (Type) {
                    case "listid": {
                        let list = this.lists.filter(l => l.Title === Value);
                        if (list.length === 1) {
                            fieldXml = fieldXml.replace(match, list[0].Id);
                        }
                    }
                }
            });
        }
        return fieldXml;
    }
}
