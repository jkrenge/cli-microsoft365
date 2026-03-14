var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageMessageListCommand_instances, _a, _VivaEngageMessageListCommand_initTelemetry, _VivaEngageMessageListCommand_initOptions, _VivaEngageMessageListCommand_initValidators;
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageMessageListCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_MESSAGE_LIST;
    }
    get description() {
        return 'Returns all accessible messages from the user\'s Viva Engage network';
    }
    defaultProperties() {
        return ['id', 'replied_to_id', 'thread_id', 'group_id', 'shortBody'];
    }
    constructor() {
        super();
        _VivaEngageMessageListCommand_instances.add(this);
        this.items = [];
        __classPrivateFieldGet(this, _VivaEngageMessageListCommand_instances, "m", _VivaEngageMessageListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageListCommand_instances, "m", _VivaEngageMessageListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageListCommand_instances, "m", _VivaEngageMessageListCommand_initValidators).call(this);
    }
    async getAllItems(logger, args, messageId) {
        let endpoint = `${this.resource}/v1`;
        if (args.options.threadId) {
            endpoint += `/messages/in_thread/${args.options.threadId}.json`;
        }
        else if (args.options.groupId) {
            endpoint += `/messages/in_group/${args.options.groupId}.json`;
        }
        else {
            if (!args.options.feedType) {
                args.options.feedType = "All";
            }
            switch (args.options.feedType) {
                case 'Top':
                    endpoint += `/messages/algo.json`;
                    break;
                case 'My':
                    endpoint += `/messages/my_feed.json`;
                    break;
                case 'Following':
                    endpoint += `/messages/following.json`;
                    break;
                case 'Sent':
                    endpoint += `/messages/sent.json`;
                    break;
                case 'Private':
                    endpoint += `/messages/private.json`;
                    break;
                case 'Received':
                    endpoint += `/messages/received.json`;
                    break;
                default:
                    endpoint += `/messages.json`;
            }
        }
        if (messageId !== -1) {
            endpoint += `?older_than=${messageId}`;
        }
        else if (args.options.olderThanId) {
            endpoint += `?older_than=${args.options.olderThanId}`;
        }
        if (args.options.threaded) {
            if (endpoint.indexOf("?") > -1) {
                endpoint += "&";
            }
            else {
                endpoint += "?";
            }
            endpoint += `threaded=true`;
        }
        const requestOptions = {
            url: endpoint,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        this.items = this.items.concat(res.messages);
        if (args.options.limit && this.items.length > args.options.limit) {
            this.items = this.items.slice(0, args.options.limit);
        }
        else if ((res.meta.older_available === true)) {
            await this.getAllItems(logger, args, this.items[this.items.length - 1].id);
        }
    }
    async commandAction(logger, args) {
        this.items = []; // this will reset the items array in interactive mode
        try {
            await this.getAllItems(logger, args, -1);
            this.items.forEach(m => {
                let shortBody;
                const bodyToProcess = m.body.plain;
                if (bodyToProcess) {
                    let maxLength = 35;
                    let addedDots = "...";
                    if (bodyToProcess.length < maxLength) {
                        maxLength = bodyToProcess.length;
                        addedDots = "";
                    }
                    shortBody = bodyToProcess.replace(/\n/g, ' ').substring(0, maxLength) + addedDots;
                }
                m.shortBody = shortBody;
            });
            await logger.log(this.items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = VivaEngageMessageListCommand, _VivaEngageMessageListCommand_instances = new WeakSet(), _VivaEngageMessageListCommand_initTelemetry = function _VivaEngageMessageListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            olderThanId: args.options.olderThanId !== undefined,
            threaded: args.options.threaded,
            limit: args.options.limit !== undefined,
            feedType: args.options.feedType !== undefined,
            threadId: args.options.threadId !== undefined,
            groupId: args.options.groupId !== undefined
        });
    });
}, _VivaEngageMessageListCommand_initOptions = function _VivaEngageMessageListCommand_initOptions() {
    this.options.unshift({
        option: '--olderThanId [olderThanId]'
    }, {
        option: '--feedType [feedType]',
        autocomplete: _a.feedTypes
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--threadId [threadId]'
    }, {
        option: '--threaded'
    }, {
        option: '--limit [limit]'
    });
}, _VivaEngageMessageListCommand_initValidators = function _VivaEngageMessageListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && args.options.threadId) {
            return `You cannot specify groupId and threadId at the same time`;
        }
        if (args.options.feedType && (args.options.groupId || args.options.threadId)) {
            return `You cannot specify the feedType with groupId or threadId at the same time`;
        }
        if (args.options.feedType && _a.feedTypes.indexOf(args.options.feedType) < 0) {
            return `${args.options.feedType} is not a valid value for the feedType option. Allowed values are ${_a.feedTypes.join(', ')}`;
        }
        if (args.options.olderThanId && typeof args.options.olderThanId !== 'number') {
            return `${args.options.olderThanId} is not a number`;
        }
        if (args.options.groupId && typeof args.options.groupId !== 'number') {
            return `${args.options.groupId} is not a number`;
        }
        if (args.options.threadId && typeof args.options.threadId !== 'number') {
            return `${args.options.threadId} is not a number`;
        }
        if (args.options.limit && typeof args.options.limit !== 'number') {
            return `${args.options.limit} is not a number`;
        }
        return true;
    });
};
VivaEngageMessageListCommand.feedTypes = ['All', 'Top', 'My', 'Following', 'Sent', 'Private', 'Received'];
export default new VivaEngageMessageListCommand();
//# sourceMappingURL=engage-message-list.js.map