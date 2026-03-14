var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignRunStatusGetCommand_instances, _SpoSiteDesignRunStatusGetCommand_initOptions, _SpoSiteDesignRunStatusGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignRunStatusGetCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_RUN_STATUS_GET;
    }
    get description() {
        return 'Gets information about the site scripts executed for the specified site design';
    }
    constructor() {
        super();
        _SpoSiteDesignRunStatusGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRunStatusGetCommand_instances, "m", _SpoSiteDesignRunStatusGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRunStatusGetCommand_instances, "m", _SpoSiteDesignRunStatusGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const data = {
            runId: args.options.runId
        };
        const requestOptions = {
            url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;odata=nometadata'
            },
            data: data,
            responseType: 'json'
        };
        try {
            const res = await request.post(requestOptions);
            await logger.log(res.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteDesignRunStatusGetCommand_instances = new WeakSet(), _SpoSiteDesignRunStatusGetCommand_initOptions = function _SpoSiteDesignRunStatusGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --runId <runId>'
    });
}, _SpoSiteDesignRunStatusGetCommand_initValidators = function _SpoSiteDesignRunStatusGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (!validation.isValidGuid(args.options.runId)) {
            return `${args.options.runId} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoSiteDesignRunStatusGetCommand();
//# sourceMappingURL=sitedesign-run-status-get.js.map