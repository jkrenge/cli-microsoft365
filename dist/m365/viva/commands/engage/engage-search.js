var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageSearchCommand_instances, _a, _VivaEngageSearchCommand_initTelemetry, _VivaEngageSearchCommand_initOptions, _VivaEngageSearchCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageSearchCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_SEARCH;
    }
    get description() {
        return 'Returns a list of messages, users, topics and groups that match the specified query.';
    }
    constructor() {
        super();
        _VivaEngageSearchCommand_instances.add(this);
        this.summary = {
            messages: 0,
            groups: 0,
            topics: 0,
            users: 0
        };
        this.messages = [];
        this.groups = [];
        this.topics = [];
        this.users = [];
        __classPrivateFieldGet(this, _VivaEngageSearchCommand_instances, "m", _VivaEngageSearchCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageSearchCommand_instances, "m", _VivaEngageSearchCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageSearchCommand_instances, "m", _VivaEngageSearchCommand_initValidators).call(this);
    }
    async getAllItems(logger, args, page) {
        const endpoint = `${this.resource}/v1/search.json?search=${formatting.encodeQueryParameter(args.options.queryText)}&page=${page}`;
        const requestOptions = {
            url: endpoint,
            responseType: 'json',
            headers: {
                accept: 'application/json;odata=nometadata'
            }
        };
        const results = await request.get(requestOptions);
        // results count should only read once
        if (page === 1) {
            this.summary = {
                messages: results.count.messages,
                topics: results.count.topics,
                users: results.count.users,
                groups: results.count.groups
            };
        }
        const resultMessages = results.messages.messages;
        const resultTopics = results.topics;
        const resultGroups = results.groups;
        const resultUsers = results.users;
        if (resultMessages.length > 0) {
            this.messages = this.messages.concat(resultMessages);
            if (args.options.limit && this.messages.length > args.options.limit) {
                this.messages = this.messages.slice(0, args.options.limit);
            }
        }
        if (resultTopics.length > 0) {
            this.topics = this.topics.concat(resultTopics);
            if (args.options.limit && this.topics.length > args.options.limit) {
                this.topics = this.topics.slice(0, args.options.limit);
            }
        }
        if (resultGroups.length > 0) {
            this.groups = this.groups.concat(resultGroups);
            if (args.options.limit && this.groups.length > args.options.limit) {
                this.groups = this.groups.slice(0, args.options.limit);
            }
        }
        if (resultUsers.length > 0) {
            this.users = this.users.concat(resultUsers);
            if (args.options.limit && this.users.length > args.options.limit) {
                this.users = this.users.slice(0, args.options.limit);
            }
        }
        const continueProcessing = resultMessages.length === 20 ||
            resultUsers.length === 20 ||
            resultGroups.length === 20 ||
            resultTopics.length === 20;
        if (continueProcessing) {
            await this.getAllItems(logger, args, page + 1);
        }
    }
    async commandAction(logger, args) {
        this.summary = {
            messages: 0,
            groups: 0,
            topics: 0,
            users: 0
        };
        this.messages = [];
        this.groups = [];
        this.topics = [];
        this.users = [];
        try {
            await this.getAllItems(logger, args, 1);
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log({
                    summary: this.summary,
                    messages: this.messages,
                    users: this.users,
                    topics: this.topics,
                    groups: this.groups
                });
            }
            else {
                const show = args.options.show?.toLowerCase();
                if (show === "summary") {
                    await logger.log(this.summary);
                }
                else {
                    let results = [];
                    if (show === undefined || show === "messages") {
                        results = [...results, ...this.messages.map((msg) => {
                                let trimmedMessage = msg.content_excerpt;
                                trimmedMessage = trimmedMessage?.length >= 80 ? (trimmedMessage.substring(0, 80) + "...") : trimmedMessage;
                                trimmedMessage = trimmedMessage?.replace(/\n/g, " ");
                                return {
                                    id: msg.id,
                                    description: trimmedMessage,
                                    type: "message",
                                    web_url: msg.web_url
                                };
                            })];
                    }
                    if (show === undefined || show === "topics") {
                        results = [...results, ...this.topics.map((topic) => {
                                return {
                                    id: topic.id,
                                    description: topic.name,
                                    type: "topic",
                                    web_url: topic.web_url
                                };
                            })];
                    }
                    if (show === undefined || show === "users") {
                        results = [...results, ...this.users.map((user) => {
                                return {
                                    id: user.id,
                                    description: user.name,
                                    type: "user",
                                    web_url: user.web_url
                                };
                            })];
                    }
                    if (show === undefined || show === "groups") {
                        results = [...results, ...this.groups.map((group) => {
                                return {
                                    id: group.id,
                                    description: group.name,
                                    type: "group",
                                    web_url: group.web_url
                                };
                            })];
                    }
                    await logger.log(results);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = VivaEngageSearchCommand, _VivaEngageSearchCommand_instances = new WeakSet(), _VivaEngageSearchCommand_initTelemetry = function _VivaEngageSearchCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            show: typeof args.options.show !== 'undefined',
            limit: typeof args.options.limit !== 'undefined'
        });
    });
}, _VivaEngageSearchCommand_initOptions = function _VivaEngageSearchCommand_initOptions() {
    this.options.unshift({
        option: '--queryText <queryText>'
    }, {
        option: '--show [show]',
        autocomplete: _a.showOptions
    }, {
        option: '--limit [limit]'
    });
}, _VivaEngageSearchCommand_initValidators = function _VivaEngageSearchCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.queryText && typeof args.options.queryText !== 'string') {
            return `${args.options.queryText} is not a string`;
        }
        if (args.options.limit && typeof args.options.limit !== 'number') {
            return `${args.options.limit} is not a number`;
        }
        if (args.options.output !== 'json') {
            if (typeof args.options.show !== 'undefined') {
                const scope = args.options.show.toString().toLowerCase();
                if (_a.showOptions.indexOf(scope) < 0) {
                    return `${scope} is not a valid value for show. Allowed values are ${_a.showOptions.join(', ')}`;
                }
            }
        }
        else {
            if (typeof args.options.show !== 'undefined') {
                return `${args.options.show} can't be used when --output set to json`;
            }
        }
        return true;
    });
};
VivaEngageSearchCommand.showOptions = [
    'summary', 'messages', 'users', 'topics', 'groups'
];
export default new VivaEngageSearchCommand();
//# sourceMappingURL=engage-search.js.map