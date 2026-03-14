var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageCommunityGetCommand_instances, _VivaEngageCommunityGetCommand_initOptions, _VivaEngageCommunityGetCommand_initTypes;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
class VivaEngageCommunityGetCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_COMMUNITY_GET;
    }
    get description() {
        return 'Gets information of a Viva Engage community';
    }
    constructor() {
        super();
        _VivaEngageCommunityGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityGetCommand_instances, "m", _VivaEngageCommunityGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityGetCommand_instances, "m", _VivaEngageCommunityGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Getting the information of Viva Engage community with id '${args.options.id}'...`);
        }
        const requestOptions = {
            url: `${this.resource}/beta/employeeExperience/communities/${args.options.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_VivaEngageCommunityGetCommand_instances = new WeakSet(), _VivaEngageCommunityGetCommand_initOptions = function _VivaEngageCommunityGetCommand_initOptions() {
    this.options.unshift({ option: '-i, --id <id>' });
}, _VivaEngageCommunityGetCommand_initTypes = function _VivaEngageCommunityGetCommand_initTypes() {
    this.types.string.push('id');
};
export default new VivaEngageCommunityGetCommand();
//# sourceMappingURL=engage-community-get.js.map