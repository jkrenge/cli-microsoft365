import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { ClientSidePage } from './clientsidepages.js';
import { getControlTypeDisplayName } from './pageMethods.js';
export const supportedPageLayouts = ['Article', 'Home', 'SingleWebPartAppPage', 'RepostPage', 'HeaderlessSearchResults', 'Spaces', 'Topic'];
export const supportedPromoteAs = ['HomePage', 'NewsPage', 'Template'];
export class Page {
    static async getPage(name, webUrl, logger, debug, verbose) {
        if (verbose) {
            await logger.logToStderr(`Retrieving information about the page...`);
        }
        const pageName = this.getPageNameWithExtension(name);
        const requestOptions = {
            url: `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(webUrl)}/SitePages/${formatting.encodeQueryParameter(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
            headers: {
                'content-type': 'application/json;charset=utf-8',
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
            throw `Page ${name} is not a modern page.`;
        }
        return ClientSidePage.fromHtml(res.ListItemAllFields.CanvasContent1);
    }
    static async checkout(name, webUrl, logger, verbose) {
        if (verbose) {
            await logger.log(`Checking out ${name} page...`);
        }
        const pageName = this.getPageNameWithExtension(name);
        const requestOptions = {
            url: `${webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageName)}')/checkoutpage`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const pageData = await request.post(requestOptions);
        if (!pageData) {
            throw `Page ${name} information not retrieved with the checkout`;
        }
        if (verbose) {
            await logger.log(`Page ${name} is now checked out`);
        }
        return pageData;
    }
    static getControlsInformation(control, isJSONOutput) {
        // remove the column property to be able to serialize the object to JSON
        delete control.column;
        if (!isJSONOutput) {
            control.controlType = getControlTypeDisplayName(control.controlType);
        }
        if (!control.dynamicDataPaths) {
            delete control.dynamicDataPaths;
        }
        if (!control.dynamicDataValues) {
            delete control.dynamicDataValues;
        }
        return control;
    }
    static getColumnsInformation(column, isJSONOutput) {
        const output = {
            factor: column.factor,
            order: column.order
        };
        if (isJSONOutput) {
            output.dataVersion = column.dataVersion;
            output.jsonData = column.jsonData;
        }
        return output;
    }
    static getSectionInformation(section, isJSONOutput) {
        const sectionOutput = {
            order: section.order
        };
        if (this.isVerticalSection(section)) {
            sectionOutput.isVertical = true;
        }
        sectionOutput.columns = section.columns.map(column => this.getColumnsInformation(column, isJSONOutput));
        return sectionOutput;
    }
    /**
     * Publish a modern page in SharePoint Online
     * @param webUrl Absolute URL of the SharePoint site where the page is located
     * @param pageName List relative url of the page to publish
     */
    static async publishPage(webUrl, pageName) {
        const filePath = `${urlUtil.getServerRelativeSiteUrl(webUrl)}/SitePages/${pageName}`;
        const requestOptions = {
            url: `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(filePath)}')/Publish()`,
            headers: {
                accept: 'application/json;odata=nometadata'
            }
        };
        await request.post(requestOptions);
    }
    static getPageNameWithExtension(name) {
        let pageName = name;
        if (pageName.indexOf('.aspx') < 0) {
            pageName += '.aspx';
        }
        return pageName;
    }
    static isVerticalSection(section) {
        return section.layoutIndex === 2 && section?.controlData?.position?.sectionFactor === 12;
    }
}
//# sourceMappingURL=Page.js.map