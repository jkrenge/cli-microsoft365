var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskReferenceRemoveCommand_instances, _PlannerTaskReferenceRemoveCommand_initTelemetry, _PlannerTaskReferenceRemoveCommand_initOptions, _PlannerTaskReferenceRemoveCommand_initValidators, _PlannerTaskReferenceRemoveCommand_initOptionSets, _PlannerTaskReferenceRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerTaskReferenceRemoveCommand extends GraphCommand {
    get name() {
        return commands.TASK_REFERENCE_REMOVE;
    }
    get description() {
        return 'Removes the reference from the Planner task';
    }
    constructor() {
        super();
        _PlannerTaskReferenceRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceRemoveCommand_instances, "m", _PlannerTaskReferenceRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceRemoveCommand_instances, "m", _PlannerTaskReferenceRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceRemoveCommand_instances, "m", _PlannerTaskReferenceRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceRemoveCommand_instances, "m", _PlannerTaskReferenceRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceRemoveCommand_instances, "m", _PlannerTaskReferenceRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeReference(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the reference from the Planner task?` });
            if (result) {
                await this.removeReference(logger, args);
            }
        }
    }
    async removeReference(logger, args) {
        try {
            const { etag, url } = await this.getTaskDetailsEtagAndUrl(args.options);
            const requestOptionsTaskDetails = {
                url: `${this.resource}/v1.0/planner/tasks/${args.options.taskId}/details`,
                headers: {
                    'accept': 'application/json;odata.metadata=none',
                    'If-Match': etag,
                    'Prefer': 'return=representation'
                },
                responseType: 'json',
                data: {
                    references: {
                        [formatting.openTypesEncoder(url)]: null
                    }
                }
            };
            await request.patch(requestOptionsTaskDetails);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTaskDetailsEtagAndUrl(options) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(options.taskId)}/details`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        let url = options.url;
        const taskDetails = await request.get(requestOptions);
        if (options.alias) {
            const urls = [];
            if (taskDetails.references) {
                Object.entries(taskDetails.references).forEach((ref) => {
                    if (ref[1].alias?.toLocaleLowerCase() === options.alias.toLocaleLowerCase()) {
                        urls.push(decodeURIComponent(ref[0]));
                    }
                });
            }
            if (urls.length === 0) {
                throw `The specified reference with alias ${options.alias} does not exist`;
            }
            if (urls.length > 1) {
                throw `Multiple references with alias ${options.alias} found. Pass one of the following urls within the "--url" option : ${urls}`;
            }
            url = urls[0];
        }
        return { etag: taskDetails['@odata.etag'], url };
    }
}
_PlannerTaskReferenceRemoveCommand_instances = new WeakSet(), _PlannerTaskReferenceRemoveCommand_initTelemetry = function _PlannerTaskReferenceRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            url: typeof args.options.url !== 'undefined',
            alias: typeof args.options.alias !== 'undefined',
            force: !!args.options.force
        });
    });
}, _PlannerTaskReferenceRemoveCommand_initOptions = function _PlannerTaskReferenceRemoveCommand_initOptions() {
    this.options.unshift({ option: '-u, --url [url]' }, { option: '--alias [alias]' }, { option: '-i, --taskId <taskId>' }, { option: '-f, --force' });
}, _PlannerTaskReferenceRemoveCommand_initValidators = function _PlannerTaskReferenceRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.url && args.options.url.indexOf('https://') !== 0 && args.options.url.indexOf('http://') !== 0) {
            return 'The url option should contain a valid URL. A valid URL starts with http(s)://';
        }
        return true;
    });
}, _PlannerTaskReferenceRemoveCommand_initOptionSets = function _PlannerTaskReferenceRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['url', 'alias'] });
}, _PlannerTaskReferenceRemoveCommand_initTypes = function _PlannerTaskReferenceRemoveCommand_initTypes() {
    this.types.string.push('url', 'taskId', 'alias');
};
export default new PlannerTaskReferenceRemoveCommand();
//# sourceMappingURL=task-reference-remove.js.map