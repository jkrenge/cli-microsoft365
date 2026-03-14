var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteGroupifyCommand_instances, _SpoSiteGroupifyCommand_initTelemetry, _SpoSiteGroupifyCommand_initOptions, _SpoSiteGroupifyCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteGroupifyCommand extends SpoCommand {
    get name() {
        return commands.SITE_GROUPIFY;
    }
    get description() {
        return 'Connects site collection to an Microsoft 365 Group';
    }
    constructor() {
        super();
        _SpoSiteGroupifyCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteGroupifyCommand_instances, "m", _SpoSiteGroupifyCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteGroupifyCommand_instances, "m", _SpoSiteGroupifyCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteGroupifyCommand_instances, "m", _SpoSiteGroupifyCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const optionalParams = {};
        const payload = {
            displayName: args.options.displayName,
            alias: args.options.alias,
            isPublic: args.options.isPublic === true,
            optionalParams: optionalParams
        };
        if (args.options.description) {
            optionalParams.Description = args.options.description;
        }
        if (args.options.classification) {
            optionalParams.Classification = args.options.classification;
        }
        if (args.options.keepOldHomepage) {
            optionalParams.CreationOptions = ["SharePointKeepOldHomepage"];
        }
        const requestOptions = {
            url: `${args.options.url}/_api/GroupSiteManager/CreateGroupForSite`,
            headers: {
                'content-type': 'application/json;odata=nometadata',
                accept: 'application/json;odata=nometadata',
                responseType: 'json'
            },
            data: payload,
            responseType: 'json'
        };
        try {
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteGroupifyCommand_instances = new WeakSet(), _SpoSiteGroupifyCommand_initTelemetry = function _SpoSiteGroupifyCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            classification: typeof args.options.classification !== 'undefined',
            isPublic: args.options.isPublic === true,
            keepOldHomepage: args.options.keepOldHomepage === true
        });
    });
}, _SpoSiteGroupifyCommand_initOptions = function _SpoSiteGroupifyCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-a, --alias <alias>'
    }, {
        option: '-n, --displayName <displayName>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-c, --classification [classification]'
    }, {
        option: '--isPublic'
    }, {
        option: '--keepOldHomepage'
    });
}, _SpoSiteGroupifyCommand_initValidators = function _SpoSiteGroupifyCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoSiteGroupifyCommand();
//# sourceMappingURL=site-groupify.js.map