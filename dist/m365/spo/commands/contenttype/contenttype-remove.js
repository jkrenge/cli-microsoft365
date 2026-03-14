var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeRemoveCommand_instances, _SpoContentTypeRemoveCommand_initTelemetry, _SpoContentTypeRemoveCommand_initOptions, _SpoContentTypeRemoveCommand_initValidators, _SpoContentTypeRemoveCommand_initTypes, _SpoContentTypeRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoContentTypeRemoveCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_REMOVE;
    }
    get description() {
        return 'Deletes site content type';
    }
    constructor() {
        super();
        _SpoContentTypeRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeRemoveCommand_instances, "m", _SpoContentTypeRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeRemoveCommand_instances, "m", _SpoContentTypeRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeRemoveCommand_instances, "m", _SpoContentTypeRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeRemoveCommand_instances, "m", _SpoContentTypeRemoveCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeRemoveCommand_instances, "m", _SpoContentTypeRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let contentTypeId = '';
        const contentTypeIdentifierLabel = args.options.id ?
            `with id ${args.options.id}` :
            `with name ${args.options.name}`;
        const removeContentType = async () => {
            try {
                if (this.debug) {
                    await logger.logToStderr(`Retrieving information about the content type ${contentTypeIdentifierLabel}...`);
                }
                let contentTypeIdResult;
                if (args.options.id) {
                    contentTypeIdResult = { "value": [{ "StringId": args.options.id }] };
                }
                else {
                    if (this.verbose) {
                        await logger.logToStderr(`Looking up the ID of content type ${contentTypeIdentifierLabel}...`);
                    }
                    const requestOptions = {
                        url: `${args.options.webUrl}/_api/web/availableContentTypes?$filter=(Name eq '${formatting.encodeQueryParameter(args.options.name)}')`,
                        headers: {
                            accept: 'application/json;odata=nometadata'
                        },
                        responseType: 'json'
                    };
                    contentTypeIdResult = await request.get(requestOptions);
                }
                let res;
                if (contentTypeIdResult &&
                    contentTypeIdResult.value &&
                    contentTypeIdResult.value.length > 0) {
                    contentTypeId = contentTypeIdResult.value[0].StringId;
                    //execute delete operation
                    const requestOptions = {
                        url: `${args.options.webUrl}/_api/web/contenttypes('${formatting.encodeQueryParameter(contentTypeId)}')`,
                        headers: {
                            'X-HTTP-Method': 'DELETE',
                            'If-Match': '*',
                            'accept': 'application/json;odata=nometadata'
                        },
                        responseType: 'json'
                    };
                    res = await request.post(requestOptions);
                }
                else {
                    res = { "odata.null": true };
                }
                if (res && res["odata.null"] === true) {
                    throw `Content type not found`;
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeContentType();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the content type ${args.options.id || args.options.name}?` });
            if (result) {
                await removeContentType();
            }
        }
    }
}
_SpoContentTypeRemoveCommand_instances = new WeakSet(), _SpoContentTypeRemoveCommand_initTelemetry = function _SpoContentTypeRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoContentTypeRemoveCommand_initOptions = function _SpoContentTypeRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-f, --force'
    });
}, _SpoContentTypeRemoveCommand_initValidators = function _SpoContentTypeRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoContentTypeRemoveCommand_initTypes = function _SpoContentTypeRemoveCommand_initTypes() {
    this.types.string.push('id', 'i');
}, _SpoContentTypeRemoveCommand_initOptionSets = function _SpoContentTypeRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoContentTypeRemoveCommand();
//# sourceMappingURL=contenttype-remove.js.map