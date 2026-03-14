var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoTaskAddCommand_instances, _a, _TodoTaskAddCommand_initTelemetry, _TodoTaskAddCommand_initOptions, _TodoTaskAddCommand_initValidators, _TodoTaskAddCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoTaskAddCommand extends GraphDelegatedCommand {
    get name() {
        return commands.TASK_ADD;
    }
    get description() {
        return 'Add a task to a Microsoft To Do task list';
    }
    constructor() {
        super();
        _TodoTaskAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoTaskAddCommand_instances, "m", _TodoTaskAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TodoTaskAddCommand_instances, "m", _TodoTaskAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TodoTaskAddCommand_instances, "m", _TodoTaskAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TodoTaskAddCommand_instances, "m", _TodoTaskAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0`;
        try {
            const listId = await this.getTodoListId(args);
            const status = args.options.status && _a.allowedStatuses.filter(x => x.toLowerCase() === args.options.status.toLowerCase())[0];
            const requestOptions = {
                url: `${endpoint}/me/todo/lists/${listId}/tasks`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'Content-Type': 'application/json'
                },
                data: {
                    title: args.options.title,
                    body: {
                        content: args.options.bodyContent,
                        contentType: args.options.bodyContentType?.toLowerCase() || 'text'
                    },
                    importance: args.options.importance?.toLowerCase(),
                    dueDateTime: this.getDateTimeTimeZone(args.options.dueDateTime),
                    reminderDateTime: this.getDateTimeTimeZone(args.options.reminderDateTime),
                    categories: args.options.categories?.split(','),
                    completedDateTime: this.getDateTimeTimeZone(args.options.completedDateTime),
                    startDateTime: this.getDateTimeTimeZone(args.options.startDateTime),
                    status: status
                },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getDateTimeTimeZone(dateTime) {
        if (!dateTime) {
            return undefined;
        }
        return {
            dateTime: dateTime,
            timeZone: 'Etc/GMT'
        };
    }
    async getTodoListId(args) {
        if (args.options.listId) {
            return args.options.listId;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.listName)}'`,
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
}
_a = TodoTaskAddCommand, _TodoTaskAddCommand_instances = new WeakSet(), _TodoTaskAddCommand_initTelemetry = function _TodoTaskAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listName: typeof args.options.listName !== 'undefined',
            bodyContent: typeof args.options.bodyContent !== 'undefined',
            bodyContentType: args.options.bodyContentType,
            dueDateTime: typeof args.options.dueDateTime !== 'undefined',
            importance: args.options.importance,
            reminderDateTime: typeof args.options.reminderDateTime !== 'undefined',
            categories: typeof args.options.categories !== 'undefined',
            completedDateTime: typeof args.options.completedDateTime !== 'undefined',
            startDateTime: typeof args.options.startDateTime !== 'undefined',
            status: typeof args.options.status !== 'undefined'
        });
    });
}, _TodoTaskAddCommand_initOptions = function _TodoTaskAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
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
    }, {
        option: '--status [status]',
        autocomplete: _a.allowedStatuses
    });
}, _TodoTaskAddCommand_initValidators = function _TodoTaskAddCommand_initValidators() {
    this.validators.push(async (args) => {
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
        if (args.options.status && _a.allowedStatuses.map(x => x.toLowerCase()).indexOf(args.options.status.toLowerCase()) === -1) {
            return `${args.options.status} is not a valid value for status. Valid values are ${_a.allowedStatuses.join(', ')}`;
        }
        if (args.options.completedDateTime && args.options.status?.toLowerCase() !== 'completed') {
            return `The completedDateTime option can only be used when the status option is set to completed`;
        }
        return true;
    });
}, _TodoTaskAddCommand_initOptionSets = function _TodoTaskAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listName'] });
};
TodoTaskAddCommand.allowedStatuses = ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'];
export default new TodoTaskAddCommand();
//# sourceMappingURL=task-add.js.map