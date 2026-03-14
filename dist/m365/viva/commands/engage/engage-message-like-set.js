var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageMessageLikeSetCommand_instances, _VivaEngageMessageLikeSetCommand_initTelemetry, _VivaEngageMessageLikeSetCommand_initOptions, _VivaEngageMessageLikeSetCommand_initTypes, _VivaEngageMessageLikeSetCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageMessageLikeSetCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_MESSAGE_LIKE_SET;
    }
    get description() {
        return 'Likes or unlikes a Viva Engage message';
    }
    constructor() {
        super();
        _VivaEngageMessageLikeSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _VivaEngageMessageLikeSetCommand_instances, "m", _VivaEngageMessageLikeSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageLikeSetCommand_instances, "m", _VivaEngageMessageLikeSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageLikeSetCommand_instances, "m", _VivaEngageMessageLikeSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageLikeSetCommand_instances, "m", _VivaEngageMessageLikeSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.enable === false) {
            if (args.options.force) {
                await this.executeLikeAction(args.options);
            }
            else {
                const message = `Are you sure you want to unlike message ${args.options.messageId}?`;
                const result = await cli.promptForConfirmation({ message });
                if (result) {
                    await this.executeLikeAction(args.options);
                }
            }
        }
        else {
            await this.executeLikeAction(args.options);
        }
    }
    async executeLikeAction(options) {
        const requestOptions = {
            url: `${this.resource}/v1/messages/liked_by/current.json`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                message_id: options.messageId
            }
        };
        try {
            if (options.enable !== false) {
                await request.post(requestOptions);
            }
            else {
                await request.delete(requestOptions);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_VivaEngageMessageLikeSetCommand_instances = new WeakSet(), _VivaEngageMessageLikeSetCommand_initTelemetry = function _VivaEngageMessageLikeSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            enable: args.options.enable,
            force: (!(!args.options.force)).toString()
        });
    });
}, _VivaEngageMessageLikeSetCommand_initOptions = function _VivaEngageMessageLikeSetCommand_initOptions() {
    this.options.unshift({
        option: '--messageId <messageId>'
    }, {
        option: '--enable [enable]',
        autocomplete: ['true', 'false']
    }, {
        option: '-f, --force'
    });
}, _VivaEngageMessageLikeSetCommand_initTypes = function _VivaEngageMessageLikeSetCommand_initTypes() {
    this.types.boolean.push('enable');
}, _VivaEngageMessageLikeSetCommand_initValidators = function _VivaEngageMessageLikeSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.messageId && typeof args.options.messageId !== 'number') {
            return `${args.options.messageId} is not a number`;
        }
        return true;
    });
};
export default new VivaEngageMessageLikeSetCommand();
//# sourceMappingURL=engage-message-like-set.js.map