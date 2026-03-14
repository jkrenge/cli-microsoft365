var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignTaskListCommand_instances, _SpoSiteDesignTaskListCommand_initOptions, _SpoSiteDesignTaskListCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignTaskListCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_TASK_LIST;
    }
    get description() {
        return 'Lists site designs scheduled for execution on the specified site';
    }
    defaultProperties() {
        return ['ID', 'SiteDesignID', 'LogonName'];
    }
    constructor() {
        super();
        _SpoSiteDesignTaskListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignTaskListCommand_instances, "m", _SpoSiteDesignTaskListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignTaskListCommand_instances, "m", _SpoSiteDesignTaskListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
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
_SpoSiteDesignTaskListCommand_instances = new WeakSet(), _SpoSiteDesignTaskListCommand_initOptions = function _SpoSiteDesignTaskListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoSiteDesignTaskListCommand_initValidators = function _SpoSiteDesignTaskListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoSiteDesignTaskListCommand();
//# sourceMappingURL=sitedesign-task-list.js.map