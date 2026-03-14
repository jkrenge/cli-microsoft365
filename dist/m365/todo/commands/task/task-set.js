var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoTaskSetCommand_instances, _TodoTaskSetCommand_initTelemetry, _TodoTaskSetCommand_initOptions, _TodoTaskSetCommand_initValidators, _TodoTaskSetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoTaskSetCommand extends GraphDelegatedCommand {
    get name() {
        return commands.TASK_SET;
    }
    get description() {
        return 'Update a task in a Microsoft To Do task list';
    }
    constructor() {
        super();
        _TodoTaskSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoTaskSetCommand_instances, "m", _TodoTaskSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TodoTaskSetCommand_instances, "m", _TodoTaskSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TodoTaskSetCommand_instances, "m", _TodoTaskSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TodoTaskSetCommand_instances, "m", _TodoTaskSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0`;
        const data = this.mapRequestBody(args.options);
        try {
            const listId = await this.getTodoListId(args.options);
            const requestOptions = {
                url: `${endpoint}/me/todo/lists/${listId}/tasks/${formatting.encodeQueryParameter(args.options.id)}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'Content-Type': 'application/json'
                },
                data: data,
                responseType: 'json'
            };
            const res = await request.patch(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTodoListId(options) {
        if (options.listId) {
            return options.listId;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(options.listName)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const taskList = response.value[0];
        if (!taskList) {
            throw `The specified task list does not exist`;
        }
        return taskList.id;
    }
    getDateTimeTimeZone(dateTime) {
        return {
            dateTime: dateTime,
            timeZone: 'Etc/GMT'
        };
    }
    mapRequestBody(options) {
        const requestBody = {};
        if (options.status) {
            requestBody.status = options.status;
        }
        if (options.title) {
            requestBody.title = options.title;
        }
        if (options.importance) {
            requestBody.importance = options.importance.toLowerCase();
        }
        if (options.bodyContentType || options.bodyContent) {
            requestBody.body = {
                content: options.bodyContent,
                contentType: options.bodyContentType?.toLowerCase() || 'text'
            };
        }
        if (options.dueDateTime) {
            requestBody.dueDateTime = this.getDateTimeTimeZone(options.dueDateTime);
        }
        if (options.reminderDateTime) {
            requestBody.reminderDateTime = this.getDateTimeTimeZone(options.reminderDateTime);
        }
        if (options.categories) {
            requestBody.categories = options.categories.split(',');
        }
        if (options.completedDateTime) {
            requestBody.completedDateTime = this.getDateTimeTimeZone(options.completedDateTime);
        }
        if (options.startDateTime) {
            requestBody.startDateTime = this.getDateTimeTimeZone(options.startDateTime);
        }
        return requestBody;
    }
}
_TodoTaskSetCommand_instances = new WeakSet(), _TodoTaskSetCommand_initTelemetry = function _TodoTaskSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listName: typeof args.options.listName !== 'undefined',
            status: typeof args.options.status !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            bodyContent: typeof args.options.bodyContent !== 'undefined',
            bodyContentType: args.options.bodyContentType,
            dueDateTime: typeof args.options.dueDateTime !== 'undefined',
            importance: args.options.importance,
            reminderDateTime: typeof args.options.reminderDateTime !== 'undefined',
            categories: typeof args.options.categories !== 'undefined',
            completedDateTime: typeof args.options.completedDateTime !== 'undefined',
            startDateTime: typeof args.options.startDateTime !== 'undefined'
        });
    });
}, _TodoTaskSetCommand_initOptions = function _TodoTaskSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-s, --status [status]',
        autocomplete: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']
    }, {
        option: '--listName [listName]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--bodyContent [bodyContent]'
    }, {
        option: '--bodyContentType [bodyContentType]',
        autocomplete: ['text', 'html']
    }, {
        option: '--dueDateTime [dueDateTime]'
    }, {
        option: '--importance [importance]',
        autocomplete: ['low', 'normal', 'high']
    }, {
        option: '--reminderDateTime [reminderDateTime]'
    }, {
        option: '--categories [categories]'
    }, {
        option: '--completedDateTime [completedDateTime]'
    }, {
        option: '--startDateTime [startDateTime]'
    });
}, _TodoTaskSetCommand_initValidators = function _TodoTaskSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.status &&
            args.options.status !== 'notStarted' &&
            args.options.status !== 'inProgress' &&
            args.options.status !== 'completed' &&
            args.options.status !== 'waitingOnOthers' &&
            args.options.status !== 'deferred') {
            return `${args.options.status} is not a valid value. Allowed values are notStarted|inProgress|completed|waitingOnOthers|deferred`;
        }
        if (args.options.bodyContentType && ['text', 'html'].indexOf(args.options.bodyContentType.toLowerCase()) === -1) {
            return `'${args.options.bodyContentType}' is not a valid value for the bodyContentType option. Allowed values are text|html`;
        }
        if (args.options.importance && ['low', 'normal', 'high'].indexOf(args.options.importance.toLowerCase()) === -1) {
            return `'${args.options.importance}' is not a valid value for the importance option. Allowed values are low|normal|high`;
        }
        if (args.options.dueDateTime && !validation.isValidISODateTime(args.options.dueDateTime)) {
            return `'${args.options.dueDateTime}' is not a valid ISO date string`;
        }
        if (args.options.reminderDateTime && !validation.isValidISODateTime(args.options.reminderDateTime)) {
            return `'${args.options.reminderDateTime}' is not a valid ISO date string`;
        }
        if (args.options.completedDateTime && !validation.isValidISODateTime(args.options.completedDateTime)) {
            return `'${args.options.completedDateTime}' is not a valid datetime.`;
        }
        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
            return `'${args.options.startDateTime}' is not a valid datetime.`;
        }
        return true;
    });
}, _TodoTaskSetCommand_initOptionSets = function _TodoTaskSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listName'] });
};
export default new TodoTaskSetCommand();
//# sourceMappingURL=task-set.js.map