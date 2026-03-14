var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageSetCommand_instances, _SpoPageSetCommand_initTelemetry, _SpoPageSetCommand_initOptions, _SpoPageSetCommand_initTypes, _SpoPageSetCommand_initValidators;
import { Auth } from '../../../../Auth.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page, supportedPageLayouts, supportedPromoteAs } from './Page.js';
class SpoPageSetCommand extends SpoCommand {
    get name() {
        return commands.PAGE_SET;
    }
    get description() {
        return 'Updates modern page properties';
    }
    constructor() {
        super();
        _SpoPageSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageSetCommand_instances, "m", _SpoPageSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPageSetCommand_instances, "m", _SpoPageSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageSetCommand_instances, "m", _SpoPageSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoPageSetCommand_instances, "m", _SpoPageSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const resource = Auth.getResourceFromUrl(args.options.webUrl);
        let pageName = args.options.name;
        const fileNameWithoutExtension = pageName.replace('.aspx', '');
        let pageTitle = args.options.title || "";
        let pageDescription = args.options.description || "";
        const pageData = {};
        if (!pageName.endsWith('.aspx')) {
            pageName += '.aspx';
        }
        const listServerRelativeUrl = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages`;
        const serverRelativeFileUrl = `${listServerRelativeUrl}/${pageName}`;
        const listUrl = urlUtil.getServerRelativePath(args.options.webUrl, listServerRelativeUrl);
        const requestUrl = `${args.options.webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;
        const needsToSavePage = !!args.options.title || !!args.options.description || !!args.options.content;
        try {
            const requestDigestResult = await spo.getRequestDigest(args.options.webUrl);
            const requestDigest = requestDigestResult.FormDigestValue;
            const page = await Page.checkout(args.options.name, args.options.webUrl, logger, this.verbose);
            pageTitle = pageTitle || page.Title;
            let pageId = page.Id;
            const bannerImageUrl = page.BannerImageUrl;
            const canvasContent1 = args.options.content || page.CanvasContent1;
            const layoutWebpartsContent = page.LayoutWebpartsContent;
            pageDescription = pageDescription || page.Description;
            const topicHeader = page.TopicHeader;
            const authorByline = page.AuthorByline;
            if (args.options.layoutType) {
                const listItemSetOptions = {
                    PageLayoutType: args.options.layoutType
                };
                if (args.options.layoutType === 'Article') {
                    listItemSetOptions.PromotedState = 0;
                    listItemSetOptions.BannerImageUrl = `${resource}/_layouts/15/images/sitepagethumbnail.png, /_layouts/15/images/sitepagethumbnail.png`;
                }
                await spo.systemUpdateListItem(requestUrl, pageId, logger, this.verbose, listItemSetOptions);
            }
            if (args.options.promoteAs) {
                const requestOptions = {
                    responseType: 'json'
                };
                switch (args.options.promoteAs) {
                    case 'HomePage':
                        requestOptions.url = `${args.options.webUrl}/_api/web/rootfolder`;
                        requestOptions.headers = {
                            'X-RequestDigest': requestDigest,
                            'X-HTTP-Method': 'MERGE',
                            'IF-MATCH': '*',
                            'content-type': 'application/json;odata=nometadata',
                            accept: 'application/json;odata=nometadata'
                        };
                        requestOptions.data = {
                            WelcomePage: `SitePages/${pageName}`
                        };
                        await request.post(requestOptions);
                        break;
                    case 'NewsPage': {
                        const listItemSetOptions = {
                            PromotedState: 2,
                            FirstPublishedDate: new Date().toISOString()
                        };
                        await spo.systemUpdateListItem(requestUrl, pageId, logger, this.verbose, listItemSetOptions);
                        break;
                    }
                    case 'Template': {
                        requestOptions.headers = {
                            'X-RequestDigest': requestDigest,
                            'content-type': 'application/json;odata=nometadata',
                            'X-HTTP-Method': 'POST',
                            'IF-MATCH': '*',
                            accept: 'application/json;odata=nometadata'
                        };
                        requestOptions.url = `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsTemplate`;
                        const res = await request.post(requestOptions);
                        if (fileNameWithoutExtension) {
                            pageData.Title = fileNameWithoutExtension;
                        }
                        if (pageDescription) {
                            pageData.Description = pageDescription;
                        }
                        if (res.BannerImageUrl) {
                            pageData.BannerImageUrl = res.BannerImageUrl;
                        }
                        if (res.LayoutWebpartsContent) {
                            pageData.LayoutWebpartsContent = res.LayoutWebpartsContent;
                        }
                        if (res.CanvasContent1) {
                            pageData.CanvasContent1 = res.CanvasContent1;
                        }
                        pageId = res.Id;
                        break;
                    }
                }
            }
            if (args.options.promoteAs !== 'Template') {
                if (pageTitle) {
                    pageData.Title = pageTitle;
                }
                if (pageDescription) {
                    pageData.Description = pageDescription;
                }
                if (bannerImageUrl) {
                    pageData.BannerImageUrl = bannerImageUrl;
                }
                if (canvasContent1) {
                    pageData.CanvasContent1 = canvasContent1;
                }
                if (layoutWebpartsContent) {
                    pageData.LayoutWebpartsContent = layoutWebpartsContent;
                }
                if (topicHeader) {
                    pageData.TopicHeader = topicHeader;
                }
                if (authorByline) {
                    pageData.AuthorByline = authorByline;
                }
            }
            if (needsToSavePage) {
                const requestOptions = {
                    responseType: 'json',
                    url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePage`,
                    headers: {
                        'X-RequestDigest': requestDigest,
                        'X-HTTP-Method': 'MERGE',
                        'IF-MATCH': '*',
                        'content-type': 'application/json;odata=nometadata',
                        accept: 'application/json;odata=nometadata'
                    },
                    data: pageData
                };
                await request.post(requestOptions);
            }
            if (args.options.promoteAs === 'Template') {
                const requestOptions = {
                    responseType: 'json',
                    url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
                    headers: {
                        'X-RequestDigest': requestDigest,
                        'X-HTTP-Method': 'MERGE',
                        'IF-MATCH': '*',
                        'content-type': 'application/json;odata=nometadata',
                        accept: 'application/json;odata=nometadata'
                    },
                    data: pageData
                };
                await request.post(requestOptions);
            }
            if (typeof args.options.commentsEnabled !== 'undefined') {
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${serverRelativeFileUrl}')/ListItemAllFields/SetCommentsDisabled(${args.options.commentsEnabled === false})`,
                    headers: {
                        'X-RequestDigest': requestDigest,
                        'content-type': 'application/json;odata=nometadata',
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
            }
            if (args.options.demoteFrom === 'NewsPage') {
                const listItemSetOptions = {
                    PromotedState: 0
                };
                await spo.systemUpdateListItem(requestUrl, pageId, logger, this.verbose, listItemSetOptions);
            }
            if (args.options.isRetired !== undefined) {
                await spo.systemUpdateListItem(requestUrl, pageId, logger, this.verbose, {
                    _SPIsRetired: args.options.isRetired
                });
            }
            let requestOptions;
            if (!args.options.publish) {
                if (args.options.promoteAs === 'Template' || !pageId) {
                    return;
                }
                requestOptions = {
                    responseType: 'json',
                    url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
                    headers: {
                        'content-type': 'application/json;odata=nometadata',
                        'accept': 'application/json;odata=nometadata'
                    },
                    data: pageData
                };
            }
            else {
                requestOptions = {
                    url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${serverRelativeFileUrl}')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${formatting.encodeQueryParameter(args.options.publishMessage || '')}'&@a2=1`,
                    headers: {
                        'X-RequestDigest': requestDigest,
                        'content-type': 'application/json;odata=nometadata',
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
            }
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageSetCommand_instances = new WeakSet(), _SpoPageSetCommand_initTelemetry = function _SpoPageSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            layoutType: args.options.layoutType || false,
            promoteAs: args.options.promoteAs || false,
            demotefrom: args.options.demoteFrom || false,
            commentsEnabled: args.options.commentsEnabled || false,
            publish: args.options.publish || false,
            publishMessage: typeof args.options.publishMessage !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            content: typeof args.options.content !== 'undefined',
            isRetired: typeof args.options.isRetired !== 'undefined'
        });
    });
}, _SpoPageSetCommand_initOptions = function _SpoPageSetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --layoutType [layoutType]',
        autocomplete: supportedPageLayouts
    }, {
        option: '-p, --promoteAs [promoteAs]',
        autocomplete: supportedPromoteAs
    }, {
        option: '--demoteFrom [demoteFrom]',
        autocomplete: ['NewsPage']
    }, {
        option: '--commentsEnabled [commentsEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--publish'
    }, {
        option: '--publishMessage [publishMessage]'
    }, {
        option: '--description [description]'
    }, {
        option: '--title [title]'
    }, {
        option: '--content [content]'
    }, {
        option: '--isRetired [isRetired]',
        autocomplete: ['true', 'false']
    });
}, _SpoPageSetCommand_initTypes = function _SpoPageSetCommand_initTypes() {
    this.types.boolean.push('commentsEnabled', 'isRetired');
}, _SpoPageSetCommand_initValidators = function _SpoPageSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (!args.options.layoutType && !args.options.promoteAs && !args.options.demoteFrom && args.options.commentsEnabled === undefined &&
            !args.options.publish && !args.options.description && !args.options.title && !args.options.content && args.options.isRetired === undefined) {
            return 'Specify at least one option to update.';
        }
        if (args.options.layoutType &&
            supportedPageLayouts.indexOf(args.options.layoutType) < 0) {
            return `${args.options.layoutType} is not a valid option for layoutType. Allowed values ${supportedPageLayouts.join(', ')}`;
        }
        if (args.options.promoteAs &&
            supportedPromoteAs.indexOf(args.options.promoteAs) < 0) {
            return `${args.options.promoteAs} is not a valid option for promoteAs. Allowed values ${supportedPromoteAs.join(', ')}`;
        }
        if (args.options.demoteFrom &&
            args.options.demoteFrom !== 'NewsPage') {
            return `${args.options.demoteFrom} is not a valid option for demoteFrom. The only allowed value is 'NewsPage'`;
        }
        if (args.options.promoteAs === 'HomePage' && args.options.layoutType !== 'Home') {
            return 'You can only promote home pages as site home page';
        }
        if (args.options.promoteAs === 'NewsPage' && args.options.layoutType && args.options.layoutType !== 'Article') {
            return 'You can only promote article pages as news article';
        }
        if (args.options.content) {
            try {
                JSON.parse(args.options.content);
            }
            catch (e) {
                return `Specified content is not a valid JSON string. Input: ${args.options.content}. Error: ${e}`;
            }
        }
        return true;
    });
};
export default new SpoPageSetCommand();
//# sourceMappingURL=page-set.js.map