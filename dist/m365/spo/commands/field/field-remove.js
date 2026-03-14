var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFieldRemoveCommand_instances, _SpoFieldRemoveCommand_initTelemetry, _SpoFieldRemoveCommand_initOptions, _SpoFieldRemoveCommand_initValidators, _SpoFieldRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFieldRemoveCommand extends SpoCommand {
    get name() {
        return commands.FIELD_REMOVE;
    }
    get description() {
        return 'Removes the specified list- or site column';
    }
    constructor() {
        super();
        _SpoFieldRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFieldRemoveCommand_instances, "m", _SpoFieldRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFieldRemoveCommand_instances, "m", _SpoFieldRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFieldRemoveCommand_instances, "m", _SpoFieldRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFieldRemoveCommand_instances, "m", _SpoFieldRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let messageEnd;
        if (args.options.listId || args.options.listTitle) {
            messageEnd = `in list ${args.options.listId || args.options.listTitle}`;
        }
        else {
            messageEnd = `in site ${args.options.webUrl}`;
        }
        const removeField = async (listRestUrl, fieldId, title, internalName) => {
            if (this.verbose) {
                await logger.logToStderr(`Removing field ${fieldId || title || internalName} ${messageEnd}...`);
            }
            let fieldRestUrl;
            if (fieldId) {
                fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(fieldId)}')`;
            }
            else {
                fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter(title || internalName)}')`;
            }
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/${listRestUrl}fields${fieldRestUrl}`,
                method: 'POST',
                headers: {
                    'X-HTTP-Method': 'DELETE',
                    'If-Match': '*',
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        };
        const prepareRemoval = async () => {
            let listRestUrl = '';
            if (args.options.listId) {
                listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
            }
            else if (args.options.listTitle) {
                listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
            }
            else if (args.options.listUrl) {
                const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                listRestUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
            }
            if (args.options.group) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving fields assigned to group ${args.options.group}...`);
                }
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/web/${listRestUrl}fields`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                try {
                    const res = await request.get(requestOptions);
                    const filteredResults = res.value.filter((field) => field.Group === args.options.group);
                    if (this.verbose) {
                        await logger.logToStderr(`${filteredResults.length} matches found...`);
                    }
                    const promises = [];
                    for (let index = 0; index < filteredResults.length; index++) {
                        promises.push(removeField(listRestUrl, filteredResults[index].Id));
                    }
                    await Promise.all(promises);
                }
                catch (err) {
                    this.handleRejectedODataJsonPromise(err);
                }
            }
            else {
                try {
                    await removeField(listRestUrl, args.options.id, args.options.title, args.options.internalName);
                }
                catch (err) {
                    this.handleRejectedODataJsonPromise(err);
                }
            }
        };
        if (args.options.force) {
            await prepareRemoval();
        }
        else {
            const confirmMessage = `Are you sure you want to remove the ${args.options.group ? 'fields' : 'field'} ${args.options.id || args.options.title || args.options.internalName || 'from group ' + args.options.group} ${messageEnd}?`;
            const result = await cli.promptForConfirmation({ message: confirmMessage });
            if (result) {
                await prepareRemoval();
            }
        }
    }
}
_SpoFieldRemoveCommand_instances = new WeakSet(), _SpoFieldRemoveCommand_initTelemetry = function _SpoFieldRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            group: typeof args.options.group !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            internalName: typeof args.options.internalName !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoFieldRemoveCommand_initOptions = function _SpoFieldRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--internalName [internalName]'
    }, {
        option: '-g, --group [group]'
    }, {
        option: '-f, --force'
    });
}, _SpoFieldRemoveCommand_initValidators = function _SpoFieldRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFieldRemoveCommand_initOptionSets = function _SpoFieldRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'internalName', 'group'] });
};
export default new SpoFieldRemoveCommand();
//# sourceMappingURL=field-remove.js.map