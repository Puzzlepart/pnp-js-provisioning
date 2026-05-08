import { TypedHash } from "sp-pnp-js";
import { HandlerBase } from "./handlerbase";
import { ComposedLook } from "./composedlook";
import { CustomActions } from "./customactions";
import { Features } from "./features";
import { WebSettings } from "./websettings";
import { Navigation } from "./navigation";
import { Lists } from "./lists";
import { Files } from "./files";
import { PropertyBagEntries } from "./propertybagentries";
import { ContentTypes } from "./contenttypes";

export const DefaultHandlerMap: TypedHash<HandlerBase> = {
    ComposedLook: new ComposedLook(),
    CustomActions: new CustomActions(),
    Features: new Features(),
    Files: new Files(),
    ContentTypes: new ContentTypes(),
    Lists: new Lists(),
    Navigation: new Navigation(),
    PropertyBagEntries: new PropertyBagEntries(),
    WebSettings: new WebSettings(),
};

export const DefaultHandlerSort: TypedHash<number> = {
    ComposedLook: 7,
    ContentTypes: 2,
    CustomActions: 6,
    Features: 3,
    Files: 5,
    Lists: 4,
    Navigation: 8,
    PropertyBagEntries: 9,
    WebSettings: 1,
};

