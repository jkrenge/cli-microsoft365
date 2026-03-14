var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageMessageAddCommand_instances, _VivaEngageMessageAddCommand_initTelemetry, _VivaEngageMessageAddCommand_initOptions, _VivaEngageMessageAddCommand_initValidators;
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageMessageAddCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_MESSAGE_ADD;
    }
    get description() {
        return 'Posts a Viva Engage network message on behalf of the current user';
    }
    constructor() {
        super();
        _VivaEngageMessageAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _VivaEngageMessageAddCommand_instances, "m", _VivaEngageMessageAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageAddCommand_instances, "m", _VivaEngageMessageAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageAddCommand_instances, "m", _VivaEngageMessageAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1/messages.json`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                body: args.options.body,
                replied_to_id: args.options.repliedToId,
                direct_to_user_ids: args.options.directToUserIds,
                group_id: args.options.groupId,
                network_id: args.options.networkId
            }
        };
        try {
            const res = await request.post(requestOptions);
            let result = null;
            if (res.messages && res.messages.length === 1) {
                result = res.messages[0];
            }
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_VivaEngageMessageAddCommand_instances = new WeakSet(), _VivaEngageMessageAddCommand_initTelemetry = function _VivaEngageMessageAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            repliedToId: args.options.repliedToId !== undefined,
            directToUserIds: args.options.directToUserIds !== undefined,
            groupId: args.options.groupId !== undefined,
            networkId: args.options.networkId !== undefined
        });
    });
}, _VivaEngageMessageAddCommand_initOptions = function _VivaEngageMessageAddCommand_initOptions() {
    this.options.unshift({
        option: '-b, --body <body>'
    }, {
        option: '-r, --repliedToId [repliedToId]'
    }, {
        option: '-d, --directToUserIds [directToUserIds]'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--networkId [networkId]'
    });
}, _VivaEngageMessageAddCommand_initValidators = function _VivaEngageMessageAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && typeof args.options.groupId !== 'number') {
            return `${args.options.groupId} is not a number`;
        }
        if (args.options.networkId && typeof args.options.networkId !== 'number') {
            return `${args.options.networkId} is not a number`;
        }
        if (args.options.repliedToId && typeof args.options.repliedToId !== 'number') {
            return `${args.options.repliedToId} is not a number`;
        }
        if (args.options.groupId === undefined &&
            args.options.directToUserIds === undefined &&
            args.options.repliedToId === undefined) {
            return "You must either specify groupId, repliedToId or directToUserIds";
        }
        return true;
    });
};
export default new VivaEngageMessageAddCommand();
//# sourceMappingURL=engage-message-add.js.map