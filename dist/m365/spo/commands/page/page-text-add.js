var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageTextAddCommand_instances, _SpoPageTextAddCommand_initTelemetry, _SpoPageTextAddCommand_initOptions, _SpoPageTextAddCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CanvasSection, ClientSideText } from './clientsidepages.js';
import { Page } from './Page.js';
class SpoPageTextAddCommand extends SpoCommand {
    get name() {
        return commands.PAGE_TEXT_ADD;
    }
    get description() {
        return 'Adds text to a modern page';
    }
    constructor() {
        super();
        _SpoPageTextAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageTextAddCommand_instances, "m", _SpoPageTextAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPageTextAddCommand_instances, "m", _SpoPageTextAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageTextAddCommand_instances, "m", _SpoPageTextAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let pageName = args.options.pageName;
        if (args.options.pageName.indexOf('.aspx') < 0) {
            pageName += '.aspx';
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving request digest...`);
        }
        try {
            const reqDigest = await spo.getRequestDigest(args.options.webUrl);
            // Keep the reference of request digest for subsequent requests
            const requestDigest = reqDigest.FormDigestValue;
            if (this.verbose) {
                await logger.logToStderr(`Retrieving modern page ${pageName}...`);
            }
            // Get Client Side Page
            const page = await Page.getPage(pageName, args.options.webUrl, logger, this.debug, this.verbose);
            const section = (args.options.section || 1) - 1;
            const column = (args.options.column || 1) - 1;
            // Add a new section when page does not contain any sections
            if (page.sections.length < 1) {
                const newSection = new CanvasSection(page, 1);
                // eslint-disable-next-line @typescript-eslint/no-unused-expressions
                newSection.defaultColumn;
                page.sections.push(newSection);
            }
            // Make sure the section is in range
            if (section >= page.sections.length) {
                throw new Error(`Invalid section '${section + 1}'`);
            }
            // Make sure the column is in range
            if (column >= page.sections[section].columns.length) {
                throw new Error(`Invalid column '${column + 1}'`);
            }
            const text = new ClientSideText(args.options.text);
            if (typeof args.options.order === 'undefined') {
                page.sections[section].columns[column].addControl(text);
            }
            else {
                const order = args.options.order - 1;
                page.sections[section].columns[column].insertControl(text, order);
            }
            // Save the Client Side Page with updated information
            await this.saveClientSidePage(page, logger, args, pageName, requestDigest);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async saveClientSidePage(clientSidePage, logger, args, pageName, requestDigest) {
        const updatedContent = clientSidePage.toHtml();
        if (this.debug) {
            await logger.logToStderr('Updated canvas content: ');
            await logger.logToStderr(updatedContent);
            await logger.logToStderr('');
        }
        const requestOptions = {
            url: `${args.options
                .webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${pageName}')/ListItemAllFields`,
            headers: {
                'X-RequestDigest': requestDigest,
                'content-type': 'application/json;odata=nometadata',
                'X-HTTP-Method': 'MERGE',
                'IF-MATCH': '*',
                accept: 'application/json;odata=nometadata'
            },
            data: {
                CanvasContent1: updatedContent
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
}
_SpoPageTextAddCommand_instances = new WeakSet(), _SpoPageTextAddCommand_initTelemetry = function _SpoPageTextAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            section: typeof args.options.section !== 'undefined',
            column: typeof args.options.column !== 'undefined',
            order: typeof args.options.order !== 'undefined'
        });
    });
}, _SpoPageTextAddCommand_initOptions = function _SpoPageTextAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --pageName <pageName>'
    }, {
        option: '-t, --text <text>'
    }, {
        option: '--section [section]'
    }, {
        option: '--column [column]'
    }, {
        option: '--order [order]'
    });
}, _SpoPageTextAddCommand_initValidators = function _SpoPageTextAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.section && (!Number.isInteger(args.options.section) || args.options.section < 1)) {
            return 'The value of parameter section must be 1 or higher';
        }
        if (args.options.column && (!Number.isInteger(args.options.column) || args.options.column < 1)) {
            return 'The value of parameter column must be 1 or higher';
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoPageTextAddCommand();
//# sourceMappingURL=page-text-add.js.map