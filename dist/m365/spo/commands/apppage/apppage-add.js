var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppPageAddCommand_instances, _SpoAppPageAddCommand_initTelemetry, _SpoAppPageAddCommand_initOptions, _SpoAppPageAddCommand_initValidators;
import request from '../../../../request.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoAppPageAddCommand extends SpoCommand {
    get name() {
        return commands.APPPAGE_ADD;
    }
    get description() {
        return 'Creates a single-part app page';
    }
    constructor() {
        super();
        _SpoAppPageAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppPageAddCommand_instances, "m", _SpoAppPageAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppPageAddCommand_instances, "m", _SpoAppPageAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppPageAddCommand_instances, "m", _SpoAppPageAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const createPageRequestOptions = {
            url: `${args.options.webUrl}/_api/sitepages/Pages/CreateAppPage`,
            headers: {
                'content-type': 'application/json;odata=nometadata',
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                webPartDataAsJson: args.options.webPartData
            }
        };
        try {
            const page = await request.post(createPageRequestOptions);
            const pageUrl = page.value;
            let requestOptions = {
                url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/${pageUrl}')?$expand=ListItemAllFields`,
                headers: {
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const file = await request.get(requestOptions);
            requestOptions = {
                url: `${args.options.webUrl}/_api/sitepages/Pages/UpdateAppPage`,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: {
                    pageId: file.ListItemAllFields.Id,
                    webPartDataAsJson: args.options.webPartData,
                    title: args.options.title,
                    includeInNavigation: args.options.addToQuickLaunch
                }
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoAppPageAddCommand_instances = new WeakSet(), _SpoAppPageAddCommand_initTelemetry = function _SpoAppPageAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            addToQuickLaunch: args.options.addToQuickLaunch
        });
    });
}, _SpoAppPageAddCommand_initOptions = function _SpoAppPageAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --title <title>'
    }, {
        option: '-d, --webPartData <webPartData>'
    }, {
        option: '--addToQuickLaunch'
    });
}, _SpoAppPageAddCommand_initValidators = function _SpoAppPageAddCommand_initValidators() {
    this.validators.push(async (args) => {
        try {
            JSON.parse(args.options.webPartData);
        }
        catch (e) {
            return `Specified webPartData is not a valid JSON string. Error: ${e}`;
        }
        return true;
    });
};
export default new SpoAppPageAddCommand();
//# sourceMappingURL=apppage-add.js.map