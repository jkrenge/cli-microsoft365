var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignRunListCommand_instances, _SpoSiteDesignRunListCommand_initTelemetry, _SpoSiteDesignRunListCommand_initOptions, _SpoSiteDesignRunListCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignRunListCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_RUN_LIST;
    }
    get description() {
        return 'Lists information about site designs applied to the specified site';
    }
    defaultProperties() {
        return ['ID', 'SiteDesignID', 'SiteDesignTitle', 'StartTime'];
    }
    constructor() {
        super();
        _SpoSiteDesignRunListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRunListCommand_instances, "m", _SpoSiteDesignRunListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRunListCommand_instances, "m", _SpoSiteDesignRunListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRunListCommand_instances, "m", _SpoSiteDesignRunListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const data = {};
        if (args.options.siteDesignId) {
            data.siteDesignId = args.options.siteDesignId;
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;odata=nometadata'
            },
            data: data,
            responseType: 'json'
        };
        try {
            const res = await request.post(requestOptions);
            if (args.options.output !== 'json') {
                res.value.forEach(d => {
                    d.StartTime = new Date(parseInt(d.StartTime)).toLocaleString();
                });
            }
            await logger.log(res.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteDesignRunListCommand_instances = new WeakSet(), _SpoSiteDesignRunListCommand_initTelemetry = function _SpoSiteDesignRunListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            siteDesignId: typeof args.options.siteDesignId !== 'undefined'
        });
    });
}, _SpoSiteDesignRunListCommand_initOptions = function _SpoSiteDesignRunListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --siteDesignId [siteDesignId]'
    });
}, _SpoSiteDesignRunListCommand_initValidators = function _SpoSiteDesignRunListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.siteDesignId) {
            if (!validation.isValidGuid(args.options.siteDesignId)) {
                return `${args.options.siteDesignId} is not a valid GUID`;
            }
        }
        return true;
    });
};
export default new SpoSiteDesignRunListCommand();
//# sourceMappingURL=sitedesign-run-list.js.map