// Reused by courtesy of PnPJS
// original at: https://github.com/pnp/pnpjs/blob/b4336b370c9b10950a22d12c48ad69789d1382fc/packages/sp/src/clientsidepages.ts
export var CanvasSectionTemplate;
(function (CanvasSectionTemplate) {
    /// <summary>
    /// One column
    /// </summary>
    CanvasSectionTemplate[CanvasSectionTemplate["OneColumn"] = 0] = "OneColumn";
    /// <summary>
    /// One column, full browser width. This one only works for communication sites in combination with image or hero webparts
    /// </summary>
    CanvasSectionTemplate[CanvasSectionTemplate["OneColumnFullWidth"] = 1] = "OneColumnFullWidth";
    /// <summary>
    /// Two columns of the same size
    /// </summary>
    CanvasSectionTemplate[CanvasSectionTemplate["TwoColumn"] = 2] = "TwoColumn";
    /// <summary>
    /// Three columns of the same size
    /// </summary>
    CanvasSectionTemplate[CanvasSectionTemplate["ThreeColumn"] = 3] = "ThreeColumn";
    /// <summary>
    /// Two columns, left one is 2/3, right one 1/3
    /// </summary>
    CanvasSectionTemplate[CanvasSectionTemplate["TwoColumnLeft"] = 4] = "TwoColumnLeft";
    /// <summary>
    /// Two columns, left one is 1/3, right one 2/3
    /// </summary>
    CanvasSectionTemplate[CanvasSectionTemplate["TwoColumnRight"] = 5] = "TwoColumnRight";
    /// <summary>
    /// Vertical
    /// </summary>
    CanvasSectionTemplate[CanvasSectionTemplate["Vertical"] = 6] = "Vertical";
    /// <summary>
    /// Flexible
    /// </summary>
    CanvasSectionTemplate[CanvasSectionTemplate["Flexible"] = 7] = "Flexible";
})(CanvasSectionTemplate || (CanvasSectionTemplate = {}));
/**
 * Section background shading
 * 0 - None
 * 1 - Neutral
 * 2 - Soft
 * 3 - Strong
 */
export var ZoneEmphasis;
(function (ZoneEmphasis) {
    ZoneEmphasis[ZoneEmphasis["None"] = 0] = "None";
    ZoneEmphasis[ZoneEmphasis["Neutral"] = 1] = "Neutral";
    ZoneEmphasis[ZoneEmphasis["Soft"] = 2] = "Soft";
    ZoneEmphasis[ZoneEmphasis["Strong"] = 3] = "Strong";
})(ZoneEmphasis || (ZoneEmphasis = {}));
/**
 * Shorthand for Object.hasOwnProperty
 *
 * @param o Object to check for
 * @param p Name of the property
 */
function hOP(o, p) {
    return Object.hasOwnProperty.call(o, p);
}
/**
 * Provides functionality to extend the given object by doing a shallow copy
 *
 * @param target The object to which properties will be copied
 * @param source The source object from which properties will be copied
 * @param noOverwrite If true existing properties on the target are not overwritten from the source
 *
 */
function extend(target, source, noOverwrite = false) {
    if (!objectDefinedNotNull(source)) {
        return target;
    }
    // ensure we don't overwrite things we don't want overwritten
    const check = noOverwrite ? (o, i) => !(i in o) : () => true;
    return Object.getOwnPropertyNames(source)
        .filter((v) => check(target, v))
        .reduce((t, v) => {
        t[v] = source[v];
        return t;
    }, target);
}
/**
 * Determines if an object is both defined and not null
 * @param obj Object to test
 */
function objectDefinedNotNull(obj) {
    return typeof obj !== "undefined" && obj !== null;
}
/**
 * Gets a random GUID value
 *
 * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
 */
function getGUID() {
    let d = new Date().getTime();
    const guid = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
        const r = (d + Math.random() * 16) % 16 | 0;
        d = Math.floor(d / 16);
        return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
    });
    return guid;
}
/**
 * Gets the next order value 1 based for the provided collection
 *
 * @param collection Collection of orderable things
 */
function getNextOrder(collection) {
    if (collection.length < 1) {
        return 1;
    }
    return Math.max.apply(null, collection.map(i => i.order)) + 1;
}
/**
 * After https://stackoverflow.com/questions/273789/is-there-a-version-of-javascripts-string-indexof-that-allows-for-regular-expr/274094#274094
 *
 * @param this Types the called context this to a string in which the search will be conducted
 * @param regex A regex or string to match
 * @param startpos A starting position from which the search will begin
 */
function regexIndexOf(regex, startpos = 0) {
    const indexOf = this.substring(startpos).search(regex);
    return (indexOf >= 0) ? (indexOf + (startpos)) : indexOf;
}
/**
 * Gets an attribute value from an html string block
 *
 * @param html HTML to search
 * @param attrName The name of the attribute to find
 */
function getAttrValueFromString(html, attrName) {
    const reg = new RegExp(`${attrName}="([^"]*?)"`, "i");
    const match = reg.exec(html);
    return match && match.length > 0 ? match[1] : null;
}
/**
 * Finds bounded blocks of markup bounded by divs, ensuring to match the ending div even with nested divs in the interstitial markup
 *
 * @param html HTML to search
 * @param boundaryStartPattern The starting pattern to find, typically a div with attribute
 * @param collector A func to take the found block and provide a way to form it into a useful return that is added into the return array
 */
function getBoundedDivMarkup(html, boundaryStartPattern, collector) {
    const blocks = [];
    if (typeof html === "undefined" || html === null) {
        return blocks;
    }
    // remove some extra whitespace if present
    const cleanedHtml = html.replace(/[\t\r\n]/g, "");
    // find the first div
    let startIndex = regexIndexOf.call(cleanedHtml, boundaryStartPattern);
    if (startIndex < 0) {
        // we found no blocks in the supplied html
        return blocks;
    }
    // this loop finds each of the blocks
    while (startIndex > -1) {
        // we have one open div counting from the one found above using boundaryStartPattern so we need to ensure we find it's close
        let openCounter = 1;
        let searchIndex = startIndex + 1;
        let nextDivOpen;
        let nextCloseDiv;
        // this loop finds the </div> tag that matches the opening of the control
        while (true) {
            // find both the next opening and closing div tags from our current searching index
            nextDivOpen = regexIndexOf.call(cleanedHtml, /<div[^>]*>/i, searchIndex);
            nextCloseDiv = regexIndexOf.call(cleanedHtml, /<\/div>/i, searchIndex);
            if (nextDivOpen < 0) {
                // we have no more opening divs, just set this to simplify checks below
                nextDivOpen = cleanedHtml.length + 1;
            }
            // determine which we found first, then increment or decrement our counter
            // and set the location to begin searching again
            if (nextDivOpen < nextCloseDiv) {
                openCounter++;
                searchIndex = nextDivOpen + 1;
            }
            else if (nextCloseDiv < nextDivOpen) {
                openCounter--;
                searchIndex = nextCloseDiv + 1;
            }
            // once we have no open divs back to the level of the opening control div
            // meaning we have all of the markup we intended to find
            if (openCounter === 0) {
                // get the bounded markup, +6 is the size of the ending </div> tag
                const markup = cleanedHtml.substring(startIndex, nextCloseDiv + 6).trim();
                // save the control data we found to the array
                blocks.push(collector(markup));
                // get out of our while loop
                break;
            }
            if (openCounter > 1000 || openCounter < 0) {
                // this is an arbitrary cut-off but likely we will not have 1000 nested divs
                // something has gone wrong above and we are probably stuck in our while loop
                // let's get out of our while loop and not hang everything
                throw new Error("getBoundedDivMarkup exceeded depth parameters.");
            }
        }
        // get the start of the next control
        startIndex = regexIndexOf.call(cleanedHtml, boundaryStartPattern, nextCloseDiv);
    }
    return blocks;
}
/**
 * Normalizes the order value for all the sections, columns, and controls to be 1 based and stepped (1, 2, 3...)
 *
 * @param collection The collection to normalize
 */
function reindex(collection) {
    if (!collection) {
        return;
    }
    for (let i = 0; i < collection.length; i++) {
        collection[i].order = i + 1;
        if (hOP(collection[i], "columns")) {
            reindex(collection[i].columns);
        }
        else if (hOP(collection[i], "controls")) {
            reindex(collection[i].controls);
        }
    }
}
/**
 * Represents the data and methods associated with client side "modern" pages
 */
export class ClientSidePage {
    constructor() {
        this.sections = [];
    }
    /**
     * Converts a json object to an escaped string appropriate for use in attributes when storing client-side controls
     *
     * @param json The json object to encode into a string
     */
    static jsonToEscapedString(json) {
        return JSON.stringify(json)
            .replace(/"/g, "&quot;")
            .replace(/:/g, "&#58;")
            .replace(/{/g, "&#123;")
            .replace(/}/g, "&#125;")
            .replace(/\[/g, "[")
            .replace(/\]/g, "]")
            .replace(/\*/g, "*")
            .replace(/\$/g, "$")
            .replace(/\./g, ".");
    }
    /**
     * Converts an escaped string from a client-side control attribute to a json object
     *
     * @param escapedString
     */
    static escapedStringToJson(escapedString) {
        if (!escapedString) {
            return {};
        }
        const unespace = (escaped) => {
            const mapDict = [
                [/\\/g, "\\\\"], [/&quot;/g, "\""], [/&#58;/g, ":"], [/&#123;/g, "{"], [/&#125;/g, "}"],
                [/\\\\/g, "\\"], [/\\\?/g, "?"], [/\\\./g, "."], [/\\\[/g, "["], [/\\\]/g, "]"],
                [/\\\(/g, "("], [/\\\)/g, ")"], [/\\\|/g, "|"], [/\\\+/g, "+"], [/\\\*/g, "*"],
                [/\\\$/g, "$"]
            ];
            return mapDict.reduce((r, m) => r.replace(m[0], m[1]), escaped);
        };
        return JSON.parse(unespace(escapedString));
    }
    /**
     * Converts this page's content to html markup
     */
    toHtml() {
        // trigger reindex of the entire tree
        reindex(this.sections);
        const html = [];
        html.push("<div>");
        for (let i = 0; i < this.sections.length; i++) {
            html.push(this.sections[i].toHtml());
        }
        if (this.pageSettings) {
            html.push(this.pageSettings.toHtml());
        }
        if (this.backgroundSettings) {
            html.push(this.backgroundSettings.toHtml());
        }
        html.push("</div>");
        return html.join("");
    }
    /**
     * Loads this page instance's content from the supplied html
     *
     * @param html html string representing the page's content
     */
    static fromHtml(html) {
        const page = new ClientSidePage();
        // reset sections
        page.sections = [];
        // gather our controls from the supplied html
        getBoundedDivMarkup(html, /<div\b[^>]*data-sp-canvascontrol[^>]*?>/i, markup => {
            // get the control type
            const ct = /controlType&quot;&#58;(\d*?)(,|&)/i.exec(markup);
            // if no control type is present this is a column which we give type 0 to let us process it
            const controlType = ct === null || ct.length < 0 ? -1 : parseInt(ct[1], 10);
            let control;
            switch (controlType) {
                case -1:
                    // empty canvas column
                    control = new CanvasColumn(null, 0);
                    control.fromHtml(markup);
                    page.mergeColumnToTree(control);
                    break;
                case 0:
                    // page settings
                    control = new PageSettings();
                    control.fromHtml(markup);
                    page.pageSettings = control;
                    break;
                case 1:
                    // empty canvas column
                    control = new CanvasColumn(null, 0);
                    control.fromHtml(markup);
                    page.mergeColumnToTree(control);
                    break;
                case 3:
                    // client side webpart
                    control = new ClientSideWebpart("");
                    control.fromHtml(markup);
                    page.mergePartToTree(control);
                    break;
                case 4:
                    // client side text
                    control = new ClientSideText();
                    control.fromHtml(markup);
                    page.mergePartToTree(control);
                    break;
                case 14:
                    // backgroundSection
                    control = new BackgroundSettings();
                    control.fromHtml(markup);
                    page.backgroundSettings = control;
                    break;
            }
        });
        // refresh all the orders within the tree
        reindex(page.sections);
        return page;
    }
    /**
     * Finds a control within this page's control tree using the supplied predicate
     *
     * @param predicate Takes a control and returns true or false, if true that control is returned by findControl
     */
    findControl(predicate) {
        // check all sections
        for (let i = 0; i < this.sections.length; i++) {
            // check all columns
            for (let j = 0; j < this.sections[i].columns.length; j++) {
                // check all controls
                for (let k = 0; k < this.sections[i].columns[j].controls.length; k++) {
                    // check to see if the predicate likes this control
                    if (predicate(this.sections[i].columns[j].controls[k])) {
                        return this.sections[i].columns[j].controls[k];
                    }
                }
            }
        }
        // we found nothing so give nothing back
        return null;
    }
    /**
     * Merges the control into the tree of sections and columns for this page
     *
     * @param control The control to merge
     */
    mergePartToTree(control) {
        let section;
        let column;
        let sectionFactor = 12;
        let sectionIndex = 0;
        let zoneIndex = 0;
        if (control.controlData) {
            // handle case where we don't have position data
            if (hOP(control.controlData, "position")) {
                if (hOP(control.controlData.position, "zoneIndex")) {
                    zoneIndex = control.controlData.position.zoneIndex;
                }
                if (hOP(control.controlData.position, "sectionIndex")) {
                    sectionIndex = control.controlData.position.sectionIndex;
                }
                if (hOP(control.controlData.position, "sectionFactor")) {
                    sectionFactor = control.controlData.position.sectionFactor;
                }
            }
        }
        const sections = this.sections.filter(s => s.order === zoneIndex && s.layoutIndex === (control.controlData?.position.layoutIndex ?? 1));
        if (sections.length < 1) {
            section = new CanvasSection(this, zoneIndex, [], control?.controlData);
            this.sections.push(section);
        }
        else {
            section = sections[0];
        }
        const columns = section.columns.filter(c => c.order === sectionIndex);
        if (columns.length < 1) {
            column = new CanvasColumn(section, sectionIndex, sectionFactor);
            section.columns.push(column);
        }
        else {
            column = columns[0];
        }
        control.column = column;
        column.addControl(control);
    }
    /**
     * Merges the supplied column into the tree
     *
     * @param column Column to merge
     * @param position The position data for the column
     */
    mergeColumnToTree(column) {
        const order = column?.controlData?.position.zoneIndex || 0;
        let section;
        const sections = this.sections.filter(s => s.order === order && s.layoutIndex === (column?.controlData?.position.layoutIndex ?? 1));
        if (sections.length < 1) {
            section = new CanvasSection(this, order, [], column.controlData);
            this.sections.push(section);
        }
        else {
            section = sections[0];
        }
        column.section = section;
        section.columns.push(column);
    }
}
export class CanvasSection {
    constructor(page, order, columns = [], controlData) {
        this.page = page;
        this.order = order;
        this.columns = columns;
        this.controlData = controlData;
        this.zoneId = this.controlData?.position.zoneId || getGUID();
        this.zoneGroupMetadata = this.controlData?.zoneGroupMetadata;
        this.emphasis = this.controlData?.emphasis;
        this.layoutIndex = this.controlData?.position.layoutIndex ?? 1;
        this.isLayoutReflowOnTop = this.controlData?.position.isLayoutReflowOnTop;
    }
    /**
     * Default column (this.columns[0]) for this section
     */
    get defaultColumn() {
        if (this.columns.length < 1) {
            this.addColumn(12);
        }
        return this.columns[0];
    }
    /**
     * Adds a new column to this section
     */
    addColumn(factor) {
        const column = new CanvasColumn(this, getNextOrder(this.columns), factor);
        this.columns.push(column);
        return column;
    }
    toHtml() {
        const html = [];
        for (let i = 0; i < this.columns.length; i++) {
            html.push(this.columns[i].toHtml());
        }
        return html.join("");
    }
}
class CanvasControl {
    constructor(controlType, dataVersion, column, order = 1, id = getGUID(), controlData, dynamicDataPaths = null, dynamicDataValues = null) {
        this.controlType = controlType;
        this.dataVersion = dataVersion;
        this.column = column;
        this.order = order;
        this.id = id;
        this.controlData = controlData;
        this.dynamicDataPaths = dynamicDataPaths;
        this.dynamicDataValues = dynamicDataValues;
    }
    /**
     * Value of the control's "data-sp-controldata" attribute
     */
    get jsonData() {
        return ClientSidePage.jsonToEscapedString(this.getControlData());
    }
    fromHtml(html) {
        this.controlData = ClientSidePage.escapedStringToJson(getAttrValueFromString(html, "data-sp-controldata"));
        this.dataVersion = getAttrValueFromString(html, "data-sp-canvasdataversion");
        this.controlType = this.controlData.controlType;
        this.id = this.controlData.id;
    }
}
export class PageSettings extends CanvasControl {
    constructor() {
        super(0, "1.0");
    }
    getControlData() {
        return this.controlData;
    }
    toHtml() {
        return `<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this.dataVersion}" data-sp-controldata="${this.jsonData}"></div>`;
    }
    fromHtml(html) {
        super.fromHtml(html);
    }
}
export class CanvasColumn extends CanvasControl {
    constructor(section, order, factor = 12, controls = [], dataVersion = "1.0") {
        super(0, dataVersion);
        this.section = section;
        this.order = order;
        this.factor = factor;
        this.controls = controls;
    }
    addControl(control) {
        control.column = this;
        this.controls.push(control);
        return this;
    }
    insertControl(control, index) {
        if (typeof index === 'undefined' ||
            index < 0 ||
            index >= this.controls.length) {
            this.addControl(control);
        }
        else {
            control.column = this;
            control.order = index;
            this.controls.splice(index, 0, control);
        }
        return this;
    }
    toHtml() {
        const html = [];
        if (this.controls.length < 1) {
            html.push(`<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this.dataVersion}" data-sp-controldata="${this.jsonData}"></div>`);
        }
        else {
            for (let i = 0; i < this.controls.length; i++) {
                html.push(this.controls[i].toHtml(i + 1));
            }
        }
        return html.join("");
    }
    fromHtml(html) {
        super.fromHtml(html);
        this.controlData = ClientSidePage.escapedStringToJson(getAttrValueFromString(html, "data-sp-controldata"));
        if (hOP(this.controlData, "position")) {
            if (hOP(this.controlData.position, "sectionFactor")) {
                this.factor = this.controlData.position.sectionFactor;
            }
            if (hOP(this.controlData.position, "sectionIndex")) {
                this.order = this.controlData.position.sectionIndex;
            }
        }
    }
    getControlData() {
        const controlData = {
            position: {
                sectionFactor: this.factor,
                sectionIndex: this.order,
                zoneIndex: this.section?.order || 0,
                zoneId: this.section?.zoneId,
                layoutIndex: this.section?.layoutIndex
            },
            zoneGroupMetadata: this.section?.zoneGroupMetadata,
            emphasis: this.section?.emphasis
        };
        if (this.column?.section?.isLayoutReflowOnTop !== undefined) {
            controlData.position.isLayoutReflowOnTop = this.column.section.isLayoutReflowOnTop;
        }
        const isEmptyColumn = this.controls.length === 0;
        if (isEmptyColumn) {
            controlData.id = "emptySection";
            controlData.controlType = 1;
        }
        return controlData;
    }
    /**
     * Removes this column and all contained controls from the collection
     */
    remove() {
        if (this.section) {
            this.section.columns = this.section.columns.filter(column => column.id !== this.id);
            if (this.column) {
                reindex(this.column.controls);
            }
        }
    }
}
/**
 * Abstract class with shared functionality for parts
 */
export class ClientSidePart extends CanvasControl {
    /**
     * Removes this column and all contained controls from the collection
     */
    remove() {
        if (this.column) {
            this.column.controls = this.column.controls.filter(control => control.id !== this.id);
            reindex(this.column.controls);
        }
    }
}
export class BackgroundSettings extends ClientSidePart {
    constructor() {
        super(0, "1.0");
        this.propertieJson = {};
        this.serverProcessedContent = null;
    }
    getControlData() {
        return {
            controlType: this.controlType
        };
    }
    toHtml() {
        // will form the value of the data-sp-webpartdata attribute
        const data = {
            dataVersion: this.dataVersion,
            instanceId: this.id,
            properties: this.propertieJson,
            serverProcessedContent: this.serverProcessedContent
        };
        const html = [];
        html.push(`<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this.dataVersion}" data-sp-controldata="${this.jsonData}">`);
        html.push(`<div data-sp-webpart="" data-sp-webpartdataversion="${this.dataVersion}" data-sp-webpartdata="${ClientSidePage.jsonToEscapedString(data)}">`);
        html.push(`<div data-sp-componentid="">`);
        html.push("</div>");
        html.push(`<div data-sp-htmlproperties="">`);
        for (let imageSource in this.serverProcessedContent?.imageSources) {
            html.push(`<img data-sp-prop-name="${imageSource}" src="${this.serverProcessedContent?.imageSources[imageSource]}" />`);
        }
        html.push("</div>");
        html.push("</div>");
        html.push("</div>");
        return html.join("");
    }
    setProperties(properties) {
        this.propertieJson = extend(this.propertieJson, properties);
        return this;
    }
    fromHtml(html) {
        super.fromHtml(html);
        const webPartData = ClientSidePage.escapedStringToJson(getAttrValueFromString(html, "data-sp-webpartdata"));
        this.setProperties(webPartData.properties);
        if (typeof webPartData.serverProcessedContent !== "undefined") {
            this.serverProcessedContent = webPartData.serverProcessedContent;
        }
        if (typeof webPartData.dynamicDataPaths !== "undefined") {
            this.dynamicDataPaths = webPartData.dynamicDataPaths;
        }
        if (typeof webPartData.dynamicDataValues !== "undefined") {
            this.dynamicDataValues = webPartData.dynamicDataValues;
        }
    }
}
export class ClientSideText extends ClientSidePart {
    constructor(text = "") {
        super(4, "1.0");
        this._text = '';
        this.text = text;
    }
    /**
     * The text markup of this control
     */
    get text() {
        return this._text;
    }
    set text(text) {
        if (!text.startsWith("<p>")) {
            text = `<p>${text}</p>`;
        }
        this._text = text;
    }
    getControlData() {
        const controlData = {
            controlType: this.controlType,
            editorType: "CKEditor",
            id: this.id,
            position: {
                controlIndex: this.order,
                sectionFactor: this.column ? this.column.factor : 0,
                sectionIndex: this.column ? this.column.order : 0,
                zoneIndex: this.column && this.column.section ? this.column.section.order : 0,
                zoneId: this.column?.section?.zoneId,
                layoutIndex: this.column?.section?.layoutIndex
            },
            zoneGroupMetadata: this.column?.section?.zoneGroupMetadata,
            emphasis: this.column?.section?.emphasis
        };
        if (this.column?.section?.isLayoutReflowOnTop !== undefined) {
            controlData.position.isLayoutReflowOnTop = this.column.section.isLayoutReflowOnTop;
        }
        return controlData;
    }
    toHtml(index) {
        // set our order to the value passed in
        this.order = index;
        const html = [];
        html.push(`<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this.dataVersion}" data-sp-controldata="${this.jsonData}">`);
        html.push("<div data-sp-rte=\"\">");
        html.push(`${this.text}`);
        html.push("</div>");
        html.push("</div>");
        return html.join("");
    }
    fromHtml(html) {
        super.fromHtml(html);
        this.text = "";
        getBoundedDivMarkup(html, /<div[^>]*data-sp-rte[^>]*>/i, (s) => {
            // now we need to grab the inner text between the divs
            const match = /<div[^>]*data-sp-rte[^>]*>(.*?)<\/div>$/i.exec(s);
            this.text = match && match.length > 1 ? match[1] : "";
        });
    }
}
export class ClientSideWebpart extends ClientSidePart {
    constructor(title, description = "", propertieJson = {}, webPartId = "", htmlProperties = "", serverProcessedContent = null, canvasDataVersion = "1.0", dynamicDataPaths = "", dynamicDataValues = "") {
        super(3, "1.0");
        this.title = title;
        this.description = description;
        this.propertieJson = propertieJson;
        this.webPartId = webPartId;
        this.htmlProperties = htmlProperties;
        this.serverProcessedContent = serverProcessedContent;
        this.canvasDataVersion = canvasDataVersion;
        this.dynamicDataPaths = dynamicDataPaths;
        this.dynamicDataValues = dynamicDataValues;
    }
    import(component) {
        this.webPartId = component.Id.replace(/^\{|\}$/g, "").toLowerCase();
        const manifest = JSON.parse(component.Manifest);
        this.title = manifest.preconfiguredEntries[0].title.default;
        this.description = manifest.preconfiguredEntries[0].description.default;
        this.dataVersion = "1.0";
        this.propertieJson = this.parseJsonProperties(manifest.preconfiguredEntries[0].properties);
    }
    setProperties(properties) {
        this.propertieJson = extend(this.propertieJson, properties);
        return this;
    }
    toHtml(index) {
        // set our order to the value passed in
        this.order = index;
        // will form the value of the data-sp-webpartdata attribute
        const data = {
            dataVersion: this.dataVersion,
            description: this.description,
            id: this.webPartId,
            instanceId: this.id,
            properties: this.propertieJson,
            serverProcessedContent: this.serverProcessedContent,
            title: this.title
        };
        if (this.dynamicDataPaths) {
            data['dynamicDataPaths'] = this.dynamicDataPaths;
        }
        if (this.dynamicDataValues) {
            data['dynamicDataValues'] = this.dynamicDataValues;
        }
        const html = [];
        html.push(`<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this.canvasDataVersion}" data-sp-controldata="${this.jsonData}">`);
        html.push(`<div data-sp-webpart="" data-sp-webpartdataversion="${this.dataVersion}" data-sp-webpartdata="${ClientSidePage.jsonToEscapedString(data)}">`);
        html.push(`<div data-sp-componentid>`);
        html.push(this.webPartId);
        html.push("</div>");
        html.push(`<div data-sp-htmlproperties="">`);
        html.push(this.renderHtmlProperties());
        html.push("</div>");
        html.push("</div>");
        html.push("</div>");
        return html.join("");
    }
    fromHtml(html) {
        super.fromHtml(html);
        const webPartData = ClientSidePage.escapedStringToJson(getAttrValueFromString(html, "data-sp-webpartdata"));
        this.title = webPartData.title;
        this.description = webPartData.description;
        this.webPartId = webPartData.id;
        this.canvasDataVersion = (getAttrValueFromString(html, "data-sp-canvasdataversion") || '').replace(/\\\./, ".");
        this.dataVersion = (getAttrValueFromString(html, "data-sp-webpartdataversion") || '').replace(/\\\./, ".");
        this.setProperties(webPartData.properties);
        if (typeof webPartData.serverProcessedContent !== "undefined") {
            this.serverProcessedContent = webPartData.serverProcessedContent;
        }
        if (typeof webPartData.dynamicDataPaths !== "undefined") {
            this.dynamicDataPaths = webPartData.dynamicDataPaths;
        }
        if (typeof webPartData.dynamicDataValues !== "undefined") {
            this.dynamicDataValues = webPartData.dynamicDataValues;
        }
        // get our html properties
        const htmlProps = getBoundedDivMarkup(html, /<div\b[^>]*data-sp-htmlproperties[^>]*?>/i, markup => {
            return markup.replace(/^<div\b[^>]*data-sp-htmlproperties[^>]*?>/i, "").replace(/<\/div>$/i, "");
        });
        this.htmlProperties = htmlProps.length > 0 ? htmlProps[0] : "";
    }
    getControlData() {
        const controlData = {
            controlType: this.controlType,
            id: this.id,
            position: {
                controlIndex: this.order,
                sectionFactor: this.column ? this.column.factor : 0,
                sectionIndex: this.column ? this.column.order : 0,
                zoneIndex: this.column && this.column.section ? this.column.section.order : 0,
                zoneId: this.column?.section?.zoneId,
                layoutIndex: this.column?.section?.layoutIndex
            },
            webPartId: this.webPartId,
            zoneGroupMetadata: this.column?.section?.zoneGroupMetadata,
            emphasis: this.column?.section?.emphasis
        };
        if (this.column?.section?.isLayoutReflowOnTop !== undefined) {
            controlData.position.isLayoutReflowOnTop = this.column.section.isLayoutReflowOnTop;
        }
        return controlData;
    }
    renderHtmlProperties() {
        const html = [];
        if (typeof this.serverProcessedContent === "undefined" || this.serverProcessedContent === null) {
            html.push(this.htmlProperties);
        }
        else if (typeof this.serverProcessedContent !== "undefined") {
            if (typeof this.serverProcessedContent.searchablePlainTexts !== "undefined") {
                const keys = Object.keys(this.serverProcessedContent.searchablePlainTexts);
                for (let i = 0; i < keys.length; i++) {
                    html.push(`<div data-sp-prop-name="${keys[i]}" data-sp-searchableplaintext="true">`);
                    html.push(this.serverProcessedContent.searchablePlainTexts[keys[i]]);
                    html.push("</div>");
                }
            }
            if (typeof this.serverProcessedContent.imageSources !== "undefined") {
                const keys = Object.keys(this.serverProcessedContent.imageSources);
                for (let i = 0; i < keys.length; i++) {
                    html.push(`<img data-sp-prop-name="${keys[i]}" src="${this.serverProcessedContent.imageSources[keys[i]]}" />`);
                }
            }
            if (typeof this.serverProcessedContent.links !== "undefined") {
                const keys = Object.keys(this.serverProcessedContent.links);
                for (let i = 0; i < keys.length; i++) {
                    html.push(`<a data-sp-prop-name="${keys[i]}" href="${this.serverProcessedContent.links[keys[i]]}"></a>`);
                }
            }
        }
        return html.join("");
    }
    parseJsonProperties(props) {
        // If the web part has the serverProcessedContent property then keep this one as it might be needed as input to render the web part HTML later on
        if (typeof props.webPartData !== "undefined" && typeof props.webPartData.serverProcessedContent !== "undefined") {
            this.serverProcessedContent = props.webPartData.serverProcessedContent;
        }
        else if (typeof props.serverProcessedContent !== "undefined") {
            this.serverProcessedContent = props.serverProcessedContent;
        }
        else {
            this.serverProcessedContent = null;
        }
        if (typeof props.webPartData !== "undefined" && typeof props.webPartData.dynamicDataPaths !== "undefined") {
            this.dynamicDataPaths = props.webPartData.dynamicDataPaths;
        }
        else if (typeof props.dynamicDataPaths !== "undefined") {
            this.dynamicDataPaths = props.dynamicDataPaths;
        }
        else {
            this.dynamicDataPaths = null;
        }
        if (typeof props.webPartData !== "undefined" && typeof props.webPartData.dynamicDataValues !== "undefined") {
            this.dynamicDataValues = props.webPartData.dynamicDataValues;
        }
        else if (typeof props.dynamicDataValues !== "undefined") {
            this.dynamicDataValues = props.dynamicDataValues;
        }
        else {
            this.dynamicDataValues = null;
        }
        if (typeof props.webPartData !== "undefined" && typeof props.webPartData.properties !== "undefined") {
            return props.webPartData.properties;
        }
        else if (typeof props.properties !== "undefined") {
            return props.properties;
        }
        else {
            return props;
        }
    }
}
//# sourceMappingURL=clientsidepages.js.map