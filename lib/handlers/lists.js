"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const xmljs = require("xml-js");
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the Lists Object Handler
 */
class Lists extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the Lists class
     */
    constructor() {
        super("Lists");
        this.tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z ]*}/g;
        this.lists = [];
    }
    /**
     * Provisioning lists
     *
     * @param {Web} web The web
     * @param {Array<IList>} lists The lists to provision
     */
    ProvisionObjects(web, lists) {
        const _super = name => super[name];
        return __awaiter(this, void 0, void 0, function* () {
            _super("scope_started").call(this);
            try {
                yield lists.reduce((chain, list) => chain.then(_ => this.processList(web, list)), Promise.resolve());
                yield lists.reduce((chain, list) => chain.then(_ => this.processFields(web, list)), Promise.resolve());
                yield lists.reduce((chain, list) => chain.then(_ => this.processFieldRefs(web, list)), Promise.resolve());
                yield lists.reduce((chain, list) => chain.then(_ => this.processViews(web, list)), Promise.resolve());
                _super("scope_ended").call(this);
            }
            catch (err) {
                _super("scope_ended").call(this);
                throw err;
            }
        });
    }
    /**
     * Processes a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list
     */
    processList(web, lc) {
        return __awaiter(this, void 0, void 0, function* () {
            const { created, list, data } = yield web.lists.ensure(lc.Title, lc.Description, lc.Template, lc.ContentTypesEnabled, lc.AdditionalSettings);
            this.lists.push(data);
            if (created) {
                sp_pnp_js_1.Logger.log({ data: list, level: sp_pnp_js_1.LogLevel.Info, message: `List ${lc.Title} created successfully.` });
            }
            yield this.processContentTypeBindings(lc, list, lc.ContentTypeBindings, lc.RemoveExistingContentTypes);
        });
    }
    /**
     * Processes content type bindings for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {Array<IContentTypeBinding>} contentTypeBindings Content type bindings
     * @param {boolean} removeExisting Remove existing content type bindings
     */
    processContentTypeBindings(lc, list, contentTypeBindings, removeExisting) {
        return __awaiter(this, void 0, void 0, function* () {
            if (contentTypeBindings) {
                yield contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(lc, list, ct.ContentTypeID)), Promise.resolve());
                if (removeExisting) {
                    let promises = [];
                    const contentTypes = yield list.contentTypes.get();
                    contentTypes.forEach(({ Id: { StringValue: ContentTypeId } }) => {
                        let shouldRemove = (contentTypeBindings.filter(ctb => ContentTypeId.indexOf(ctb.ContentTypeID) !== -1).length === 0)
                            && (ContentTypeId.indexOf("0x0120") === -1);
                        if (shouldRemove) {
                            sp_pnp_js_1.Logger.write(`Removing content type ${ContentTypeId} from list ${lc.Title}`, sp_pnp_js_1.LogLevel.Info);
                            promises.push(list.contentTypes.getById(ContentTypeId).delete());
                        }
                    });
                    yield Promise.all(promises);
                }
            }
        });
    }
    /**
     * Processes a content type binding for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {string} contentTypeID The Content Type ID
     */
    processContentTypeBinding(lc, list, contentTypeID) {
        return __awaiter(this, void 0, void 0, function* () {
            yield list.contentTypes.addAvailableContentType(contentTypeID);
            sp_pnp_js_1.Logger.log({ level: sp_pnp_js_1.LogLevel.Info, message: `Content Type ${contentTypeID} added successfully to list ${lc.Title}.` });
        });
    }
    /**
     * Processes fields for a list
     *
     * @param {Web} web The web
     * @param {IList} list The pnp list
     */
    processFields(web, list) {
        return __awaiter(this, void 0, void 0, function* () {
            if (list.Fields) {
                yield list.Fields.reduce((chain, field) => chain.then(_ => this.processField(web, list, field)), Promise.resolve());
            }
        });
    }
    /**
     * Processes a field for a lit
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {string} fieldXml Field xml
     */
    processField(web, lc, fieldXml) {
        return __awaiter(this, void 0, void 0, function* () {
            const list = web.lists.getByTitle(lc.Title);
            const fXmlJson = JSON.parse(xmljs.xml2json(fieldXml));
            const fAttr = fXmlJson.elements[0].attributes;
            const internalName = fAttr.InternalName;
            const displayName = fAttr.DisplayName;
            fieldXml = xmljs.json2xml(fXmlJson);
            fXmlJson.elements[0].attributes.DisplayName = internalName;
            try {
                // Looks like e.g. lookup fields can't be updated, so we'll need to reac
                let field = yield list.fields.getById(fAttr.ID);
                yield field.delete();
            }
            catch (err) {
                sp_pnp_js_1.Logger.log({ message: `Failed to remove field '${displayName}' from list ${lc.Title}.`, level: sp_pnp_js_1.LogLevel.Warning });
                throw err;
            }
            let fieldAddResult = yield list.fields.createFieldAsXml(this.replaceFieldXmlTokens(fieldXml));
            yield fieldAddResult.field.update({ Title: displayName });
            sp_pnp_js_1.Logger.log({ message: `Field '${displayName}' added successfully to list ${lc.Title}.`, level: sp_pnp_js_1.LogLevel.Info });
        });
    }
    /**
   * Processes field refs for a list
   *
   * @param {Web} web The web
   * @param {IList} list The pnp list
   */
    processFieldRefs(web, list) {
        return __awaiter(this, void 0, void 0, function* () {
            if (list.FieldRefs) {
                yield list.FieldRefs.reduce((chain, fieldRef) => chain.then(_ => this.processFieldRef(web, list, fieldRef)), Promise.resolve());
            }
        });
    }
    /**
     *
     * Processes a field ref for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListInstanceFieldRef} fieldRef The list field ref
     */
    processFieldRef(web, lc, fieldRef) {
        return __awaiter(this, void 0, void 0, function* () {
            const list = web.lists.getByTitle(lc.Title);
            yield list.fields.getById(fieldRef.ID).update({ Hidden: fieldRef.Hidden, Required: fieldRef.Required, Title: fieldRef.DisplayName });
            sp_pnp_js_1.Logger.log({ data: fieldRef, level: sp_pnp_js_1.LogLevel.Info, message: `Field '${fieldRef.ID}' updated for list ${lc.Title}.` });
        });
    }
    /**
     * Processes views for a list
     *
     * @param web The web
     * @param lc The list configuration
     */
    processViews(web, lc) {
        return __awaiter(this, void 0, void 0, function* () {
            if (lc.Views) {
                yield lc.Views.reduce((chain, view) => chain.then(_ => this.processView(web, lc, view)), Promise.resolve());
            }
        });
    }
    /**
     * Processes a view for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListView} view The view configuration
     */
    processView(web, lc, view) {
        return __awaiter(this, void 0, void 0, function* () {
            let _view = web.lists.getByTitle(lc.Title).views.getByTitle(view.Title);
            try {
                yield _view.get();
                yield _view.update(view.AdditionalSettings);
                yield this.processViewFields(_view, view.ViewFields);
            }
            catch (err) {
                const result = yield web.lists.getByTitle(lc.Title).views.add(view.Title, view.PersonalView, view.AdditionalSettings);
                sp_pnp_js_1.Logger.log({ level: sp_pnp_js_1.LogLevel.Info, message: `View ${view.Title} added successfully to list ${lc.Title}.` });
                yield this.processViewFields(result.view, view.ViewFields);
            }
        });
    }
    /**
     * Processes view fields for a view
     *
     * @param {any} view The pnp view
     * @param {Array<string>} viewFields Array of view fields
     */
    processViewFields(view, viewFields) {
        return __awaiter(this, void 0, void 0, function* () {
            yield view.fields.removeAll();
            yield viewFields.reduce((chain, viewField) => chain.then(_ => view.fields.add(viewField)), Promise.resolve());
        });
    }
    /**
     * Replaces tokens in field xml
     *
     * @param {string} fieldXml The field xml
     */
    replaceFieldXmlTokens(fieldXml) {
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
exports.Lists = Lists;
