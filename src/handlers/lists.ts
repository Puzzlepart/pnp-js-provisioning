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
     * @param {IList} listConfig The list
     */
    private async processList(web: Web, listConfig: IList): Promise<void> {
        const { created, list, data } = await web.lists.ensure(listConfig.Title, listConfig.Description, listConfig.Template, listConfig.ContentTypesEnabled, listConfig.AdditionalSettings);
        this.lists.push(data);
        if (created) {
            Logger.log({ data: list, level: LogLevel.Info, message: `List ${listConfig.Title} created successfully.` });
        }
        await this.processContentTypeBindings(listConfig, list, listConfig.ContentTypeBindings, listConfig.RemoveExistingContentTypes);
    }

    /**
     * Processes content type bindings for a list
     *
     * @param {IList} listConfig The list configuration
     * @param {List} list The pnp list
     * @param {Array<IContentTypeBinding>} contentTypeBindings Content type bindings
     * @param {boolean} removeExisting Remove existing content type bindings
     */
    private async processContentTypeBindings(listConfig: IList, list: List, contentTypeBindings: IContentTypeBinding[], removeExisting: boolean): Promise<any> {
        if (contentTypeBindings) {
            await contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(listConfig, list, ct.ContentTypeID)), Promise.resolve());
            if (removeExisting) {
                let promises = [];
                const contentTypes = await list.contentTypes.get();
                contentTypes.forEach(({ Id: { StringValue: ContentTypeId } }) => {
                    let shouldRemove = (contentTypeBindings.filter(ctb => ContentTypeId.indexOf(ctb.ContentTypeID) !== -1).length === 0)
                        && (ContentTypeId.indexOf("0x0120") === -1);
                    if (shouldRemove) {
                        Logger.write(`Removing content type ${ContentTypeId} from list ${listConfig.Title}`, LogLevel.Info);
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
     * @param {IList} listConfig The list configuration
     * @param {List} list The pnp list
     * @param {string} contentTypeID The Content Type ID
     */
    private async processContentTypeBinding(listConfig: IList, list: List, contentTypeID: string): Promise<any> {
        await list.contentTypes.addAvailableContentType(contentTypeID);
        Logger.log({ level: LogLevel.Info, message: `Content Type ${contentTypeID} added successfully to list ${listConfig.Title}.` });
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
     * @param {IList} listConfig The list configuration
     * @param {string} fieldXml Field xml
     */
    private async processField(web: Web, listConfig: IList, fieldXml: string): Promise<any> {
        const list = web.lists.getByTitle(listConfig.Title);
        const fXmlJson = JSON.parse(xmljs.xml2json(fieldXml));
        const fAttr = fXmlJson.elements[0].attributes;
        const internalName = fAttr.InternalName;
        const displayName = fAttr.DisplayName;
        fieldXml = xmljs.json2xml(fXmlJson);
        fXmlJson.elements[0].attributes.DisplayName = internalName;
        try {
            // Looks like e.g. lookup fields can't be updated, so we'll need to reac
            let field = await list.fields.getById(fAttr.ID);
            await field.delete();
        } catch (err) { }
        let fieldAddResult = await list.fields.createFieldAsXml(this.replaceFieldXmlTokens(fieldXml));
        await fieldAddResult.field.update({ Title: displayName });
        Logger.log({ message: `Field '${displayName}' added successfully to list ${listConfig.Title}.`, level: LogLevel.Info });
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
     * @param {IList} listConfig The list configuration
     * @param {IListInstanceFieldRef} fieldRef The list field ref
     */
    private async processFieldRef(web: Web, listConfig: IList, fieldRef: IListInstanceFieldRef): Promise<any> {
        const list = web.lists.getByTitle(listConfig.Title);
        await list.fields.getById(fieldRef.ID).update({ Hidden: fieldRef.Hidden, Required: fieldRef.Required, Title: fieldRef.DisplayName })
        Logger.log({ data: fieldRef, level: LogLevel.Info, message: `Field '${fieldRef.ID}' updated for list ${listConfig.Title}.` });
    }

    /**
     * Processes views for a list
     *
     * @param web The web
     * @param listConfig The list configuration
     */
    private async processViews(web: Web, listConfig: IList): Promise<any> {
        if (listConfig.Views) {
            await listConfig.Views.reduce((chain, view) => chain.then(_ => this.processView(web, listConfig, view)), Promise.resolve());
        }
    }

    /**
     * Processes a view for a list
     *
     * @param {Web} web The web
     * @param {IList} listConfig The list configuration
     * @param {IListView} view The view configuration
     */
    private async processView(web: Web, listConfig: IList, view: IListView): Promise<void> {
        let _view = web.lists.getByTitle(listConfig.Title).views.getByTitle(view.Title);
        try {
            await _view.get();
            await _view.update(view.AdditionalSettings);
            await this.processViewFields(_view, view.ViewFields);
        } catch (err) {
            const result = await web.lists.getByTitle(listConfig.Title).views.add(view.Title, view.PersonalView, view.AdditionalSettings);
            Logger.log({ level: LogLevel.Info, message: `View ${view.Title} added successfully to list ${listConfig.Title}.` });
            await this.processViewFields(result.view, view.ViewFields)
        }
    }

    /**
     * Processes view fields for a view
     *
     * @param {any} view The pnp view
     * @param {Array<string>} viewFields Array of view fields
     */
    private async processViewFields(view: any, viewFields: string[]): Promise<void> {
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
