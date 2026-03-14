var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageHeaderSetCommand_instances, _SpoPageHeaderSetCommand_initTelemetry, _SpoPageHeaderSetCommand_initOptions, _SpoPageHeaderSetCommand_initValidators, _SpoPageHeaderSetCommand_initTypes;
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page } from './Page.js';
const BannerWebPartId = 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788';
class SpoPageHeaderSetCommand extends SpoCommand {
    get name() {
        return commands.PAGE_HEADER_SET;
    }
    get description() {
        return 'Sets modern page header';
    }
    constructor() {
        super();
        _SpoPageHeaderSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageHeaderSetCommand_instances, "m", _SpoPageHeaderSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPageHeaderSetCommand_instances, "m", _SpoPageHeaderSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageHeaderSetCommand_instances, "m", _SpoPageHeaderSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoPageHeaderSetCommand_instances, "m", _SpoPageHeaderSetCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['imageUrl'];
    }
    async commandAction(logger, args) {
        const noPageHeader = {
            id: BannerWebPartId,
            instanceId: BannerWebPartId,
            title: 'Title Region',
            description: 'Title Region Description',
            serverProcessedContent: {
                htmlStrings: {},
                searchablePlainTexts: {},
                imageSources: {},
                links: {}
            },
            dataVersion: '1.4',
            properties: {
                title: '',
                imageSourceType: 4,
                layoutType: 'NoImage',
                textAlignment: 'Left',
                showTopicHeader: false,
                showPublishDate: false,
                showTimeToRead: false,
                topicHeader: ''
            }
        };
        const defaultPageHeader = {
            id: BannerWebPartId,
            instanceId: BannerWebPartId,
            title: 'Title Region',
            description: 'Title Region Description',
            serverProcessedContent: {
                htmlStrings: {},
                searchablePlainTexts: {},
                imageSources: {},
                links: {}
            },
            dataVersion: '1.4',
            properties: {
                title: '',
                imageSourceType: 4,
                layoutType: 'FullWidthImage',
                textAlignment: 'Left',
                showTopicHeader: false,
                showPublishDate: false,
                showTimeToRead: false,
                topicHeader: ''
            }
        };
        const customPageHeader = {
            id: BannerWebPartId,
            instanceId: BannerWebPartId,
            title: 'Title Region',
            description: 'Title Region Description',
            serverProcessedContent: {
                htmlStrings: {},
                searchablePlainTexts: {},
                imageSources: {
                    imageSource: ''
                },
                links: {},
                customMetadata: {
                    imageSource: {
                        siteId: '',
                        webId: '',
                        listId: '',
                        uniqueId: ''
                    }
                }
            },
            dataVersion: '1.4',
            properties: {
                title: '',
                imageSourceType: 2,
                layoutType: 'FullWidthImage',
                textAlignment: 'Left',
                showTopicHeader: false,
                showPublishDate: false,
                showTimeToRead: false,
                topicHeader: '',
                authors: [],
                altText: '',
                webId: '',
                siteId: '',
                listId: '',
                uniqueId: '',
                translateX: 0,
                translateY: 0
            }
        };
        let pageFullName = args.options.pageName.toLowerCase();
        if (!pageFullName.endsWith('.aspx')) {
            pageFullName += '.aspx';
        }
        let canvasContent = "";
        let bannerImageUrl = "";
        let description = "";
        let authorByline = args.options.authors ? args.options.authors.split(',').map(a => a.trim()) : [];
        let topicHeader = args.options.topicHeader || "";
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about the page...`);
        }
        try {
            let requestOptions = {
                url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')?$select=IsPageCheckedOutToCurrentUser,Title`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const page = await request.get(requestOptions);
            let title = page.Title;
            let pageData;
            if (page.IsPageCheckedOutToCurrentUser) {
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')?$expand=ListItemAllFields`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                pageData = await request.get(requestOptions);
            }
            else {
                pageData = await Page.checkout(pageFullName, args.options.webUrl, logger, this.verbose);
            }
            let header;
            switch (args.options.type) {
                case 'None':
                    header = noPageHeader;
                    break;
                case 'Custom':
                    header = customPageHeader;
                    break;
                case 'Default':
                default:
                    header = defaultPageHeader;
                    break;
            }
            if (pageData) {
                canvasContent = pageData.CanvasContent1;
                authorByline = authorByline.length > 0 ? authorByline : pageData.AuthorByline;
                bannerImageUrl = pageData.BannerImageUrl;
                description = pageData.Description;
                title = pageData.Title;
                topicHeader = topicHeader || pageData.TopicHeader || "";
            }
            const pageControls = JSON.parse(pageData.CanvasContent1);
            //In the new design page header is is a configurable Banner webpart in the first full-width section
            const headerControl = pageControls.find(control => control?.position?.zoneIndex === 1 && control?.position?.sectionFactor === 0 && control?.webPartId === BannerWebPartId);
            const isStandardPageHeader = pageData.LayoutWebpartsContent !== '[]';
            //LayoutWebpartsContent represents standard page header
            if (!isStandardPageHeader) {
                header = headerControl?.webPartData || header;
            }
            header.properties.title = title;
            header.properties.textAlignment = args.options.textAlignment || 'Left';
            header.properties.showTopicHeader = args.options.showTopicHeader || false;
            header.properties.topicHeader = args.options.topicHeader || '';
            header.properties.showPublishDate = args.options.showPublishDate || false;
            header.properties.showTimeToRead = args.options.showTimeToRead || false;
            if (args.options.type !== 'None') {
                header.properties.layoutType = args.options.layout || 'FullWidthImage';
            }
            if (args.options.type === 'Custom') {
                header.serverProcessedContent.imageSources = {
                    imageSource: args.options.imageUrl || ''
                };
                const properties = header.properties;
                properties.altText = args.options.altText || '';
                properties.translateX = args.options.translateX || 0;
                properties.translateY = args.options.translateY || 0;
                header.properties = properties;
                if (!args.options.imageUrl) {
                    header.serverProcessedContent.customMetadata = {
                        imageSource: {
                            siteId: '',
                            webId: '',
                            listId: '',
                            uniqueId: ''
                        }
                    };
                    properties.listId = '';
                    properties.siteId = '';
                    properties.uniqueId = '';
                    properties.webId = '';
                    header.properties = properties;
                }
                else {
                    const res = await Promise.all([
                        spo.getSiteIdBySPApi(args.options.webUrl, logger, this.verbose),
                        spo.getWebId(args.options.webUrl, logger, this.verbose),
                        this.getImageInfo(args.options.webUrl, args.options.imageUrl, this.verbose, logger)
                    ]);
                    header.serverProcessedContent.customMetadata = {
                        imageSource: {
                            siteId: res[0],
                            webId: res[1],
                            listId: res[2].ListId,
                            uniqueId: res[2].UniqueId
                        }
                    };
                    const properties = header.properties;
                    properties.listId = res[2].ListId;
                    properties.siteId = res[0];
                    properties.uniqueId = res[2].UniqueId;
                    properties.webId = res[1];
                    header.properties = properties;
                }
            }
            const requestBody = {
                LayoutWebpartsContent: JSON.stringify([header]),
                CanvasContent1: canvasContent
            };
            if (!isStandardPageHeader) {
                requestBody.LayoutWebpartsContent = '[]';
                header.properties.title = topicHeader;
                if (headerControl) {
                    headerControl.webPartData = header;
                }
                else {
                    for (const pageControl of pageControls) {
                        if (pageControl?.position?.sectionIndex) {
                            pageControl.position.sectionIndex += pageControl.position.sectionIndex;
                        }
                    }
                    pageControls.push({
                        id: BannerWebPartId,
                        controlType: 3,
                        displayMode: 2,
                        emphasis: {},
                        position: {
                            zoneIndex: 1,
                            sectionFactor: 0,
                            layoutIndex: 1,
                            controlIndex: 1,
                            sectionIndex: 1
                        },
                        webPartId: BannerWebPartId,
                        webPartData: header
                    });
                }
                requestBody.CanvasContent1 = JSON.stringify(pageControls);
            }
            if (title) {
                requestBody.Title = title;
            }
            if (topicHeader) {
                requestBody.TopicHeader = topicHeader;
            }
            if (description) {
                requestBody.Description = description;
            }
            if (authorByline) {
                requestBody.AuthorByline = authorByline;
            }
            if (bannerImageUrl) {
                requestBody.BannerImageUrl = bannerImageUrl;
            }
            requestOptions = {
                url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')/SavePageAsDraft`,
                headers: {
                    'X-HTTP-Method': 'MERGE',
                    'IF-MATCH': '*',
                    'content-type': 'application/json;odata=nometadata',
                    accept: 'application/json;odata=nometadata'
                },
                data: requestBody,
                responseType: 'json'
            };
            return request.post(requestOptions);
        }
        catch (err) {
            if (err.status === 404) {
                throw new CommandError(`The specified page '${pageFullName}' does not exist.`);
            }
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getImageInfo(siteUrl, imageUrl, verbose, logger) {
        if (verbose) {
            await logger.logToStderr(`Retrieving information about the header image...`);
        }
        const requestOptions = {
            url: `${siteUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(imageUrl)}')?$select=ListId,UniqueId`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
}
_SpoPageHeaderSetCommand_instances = new WeakSet(), _SpoPageHeaderSetCommand_initTelemetry = function _SpoPageHeaderSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            altText: typeof args.options.altText !== 'undefined',
            authors: typeof args.options.authors !== 'undefined',
            imageUrl: typeof args.options.imageUrl !== 'undefined',
            topicHeader: typeof args.options.topicHeader !== 'undefined',
            layout: args.options.layout,
            showTopicHeader: !!args.options.showTopicHeader,
            showPublishDate: !!args.options.showPublishDate,
            showTimeToRead: !!args.options.showTimeToRead,
            textAlignment: args.options.textAlignment,
            translateX: typeof args.options.translateX !== 'undefined',
            translateY: typeof args.options.translateY !== 'undefined',
            type: args.options.type
        });
    });
}, _SpoPageHeaderSetCommand_initOptions = function _SpoPageHeaderSetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --pageName <pageName>'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --type [type]',
        autocomplete: ['None', 'Default', 'Custom']
    }, {
        option: '--imageUrl [imageUrl]'
    }, {
        option: '--altText [altText]'
    }, {
        option: '-x, --translateX [translateX]'
    }, {
        option: '-y, --translateY [translateY]'
    }, {
        option: '--layout [layout]',
        autocomplete: ['FullWidthImage', 'NoImage', 'ColorBlock', 'CutInShape']
    }, {
        option: '--textAlignment [textAlignment]',
        autocomplete: ['Left', 'Center']
    }, {
        option: '--showTopicHeader [showTopicHeader]',
        autocomplete: ['true', 'false']
    }, {
        option: '--showPublishDate [showPublishDate]',
        autocomplete: ['true', 'false']
    }, {
        option: '--showTimeToRead [showTimeToRead]',
        autocomplete: ['true', 'false']
    }, {
        option: '--topicHeader [topicHeader]'
    }, {
        option: '--authors [authors]'
    });
}, _SpoPageHeaderSetCommand_initValidators = function _SpoPageHeaderSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type && !['None', 'Default', 'Custom'].includes(args.options.type)) {
            return `${args.options.type} is not a valid type value. Allowed values None|Default|Custom`;
        }
        if (args.options.translateX && isNaN(args.options.translateX)) {
            return `${args.options.translateX} is not a valid number`;
        }
        if (args.options.translateY && isNaN(args.options.translateY)) {
            return `${args.options.translateY} is not a valid number`;
        }
        if (args.options.layout && !['FullWidthImage', 'NoImage', 'ColorBlock', 'CutInShape'].includes(args.options.layout)) {
            return `${args.options.layout} is not a valid layout value. Allowed values FullWidthImage|NoImage|ColorBlock|CutInShape`;
        }
        if (args.options.textAlignment && !['Left', 'Center'].includes(args.options.textAlignment)) {
            return `${args.options.textAlignment} is not a valid textAlignment value. Allowed values Left|Center`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoPageHeaderSetCommand_initTypes = function _SpoPageHeaderSetCommand_initTypes() {
    this.types.boolean.push('showTimeToRead', 'showPublishDate', 'showTopicHeader');
};
export default new SpoPageHeaderSetCommand();
//# sourceMappingURL=page-header-set.js.map