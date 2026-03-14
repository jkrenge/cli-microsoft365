var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignApplyCommand_instances, _SpoSiteDesignApplyCommand_initTelemetry, _SpoSiteDesignApplyCommand_initOptions, _SpoSiteDesignApplyCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignApplyCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_APPLY;
    }
    get description() {
        return 'Applies a site design to an existing site collection';
    }
    constructor() {
        super();
        _SpoSiteDesignApplyCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignApplyCommand_instances, "m", _SpoSiteDesignApplyCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignApplyCommand_instances, "m", _SpoSiteDesignApplyCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignApplyCommand_instances, "m", _SpoSiteDesignApplyCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestBody = {
                siteDesignId: args.options.id,
                webUrl: args.options.webUrl
            };
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.${args.options.asTask ? 'AddSiteDesignTask' : 'ApplySiteDesign'}`,
                headers: {
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: requestBody,
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            if (res.value) {
                await logger.log(res.value);
            }
            else {
                await logger.log(res);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteDesignApplyCommand_instances = new WeakSet(), _SpoSiteDesignApplyCommand_initTelemetry = function _SpoSiteDesignApplyCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asTask: args.options.asTask || false
        });
    });
}, _SpoSiteDesignApplyCommand_initOptions = function _SpoSiteDesignApplyCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--asTask'
    });
}, _SpoSiteDesignApplyCommand_initValidators = function _SpoSiteDesignApplyCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoSiteDesignApplyCommand();
//# sourceMappingURL=sitedesign-apply.js.map