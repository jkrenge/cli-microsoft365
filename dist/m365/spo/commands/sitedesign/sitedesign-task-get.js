var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignTaskGetCommand_instances, _SpoSiteDesignTaskGetCommand_initOptions, _SpoSiteDesignTaskGetCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignTaskGetCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_TASK_GET;
    }
    get description() {
        return 'Gets information about the specified site design scheduled for execution';
    }
    constructor() {
        super();
        _SpoSiteDesignTaskGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignTaskGetCommand_instances, "m", _SpoSiteDesignTaskGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignTaskGetCommand_instances, "m", _SpoSiteDesignTaskGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                data: {
                    taskId: args.options.id
                },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            if (!res["odata.null"]) {
                await logger.log(res);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteDesignTaskGetCommand_instances = new WeakSet(), _SpoSiteDesignTaskGetCommand_initOptions = function _SpoSiteDesignTaskGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
}, _SpoSiteDesignTaskGetCommand_initValidators = function _SpoSiteDesignTaskGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoSiteDesignTaskGetCommand();
//# sourceMappingURL=sitedesign-task-get.js.map