var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFieldAddCommand_instances, _SpoFieldAddCommand_initTelemetry, _SpoFieldAddCommand_initOptions, _SpoFieldAddCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFieldAddCommand extends SpoCommand {
    get name() {
        return commands.FIELD_ADD;
    }
    get description() {
        return 'Adds a new list or site column using the CAML field definition';
    }
    constructor() {
        super();
        _SpoFieldAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFieldAddCommand_instances, "m", _SpoFieldAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFieldAddCommand_instances, "m", _SpoFieldAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFieldAddCommand_instances, "m", _SpoFieldAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            let requestUrl = `${args.options.webUrl}/_api/web`;
            if (args.options.listId) {
                requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
            }
            else if (args.options.listTitle) {
                requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
            }
            else if (args.options.listUrl) {
                const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
            }
            const reqDigest = await spo.getRequestDigest(args.options.webUrl);
            const requestOptions = {
                url: `${requestUrl}/fields/CreateFieldAsXml`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue,
                    accept: 'application/json;odata=nometadata'
                },
                data: {
                    parameters: {
                        SchemaXml: args.options.xml,
                        Options: this.getOptions(args.options.options)
                    }
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
    getOptions(options) {
        let optionsValue = 0;
        if (!options) {
            return optionsValue;
        }
        options.split(',').forEach(o => {
            o = o.trim();
            switch (o) {
                case 'DefaultValue':
                    optionsValue += 0;
                    break;
                case 'AddToDefaultContentType':
                    optionsValue += 1;
                    break;
                case 'AddToNoContentType':
                    optionsValue += 2;
                    break;
                case 'AddToAllContentTypes':
                    optionsValue += 4;
                    break;
                case 'AddFieldInternalNameHint':
                    optionsValue += 8;
                    break;
                case 'AddFieldToDefaultView':
                    optionsValue += 16;
                    break;
                case 'AddFieldCheckDisplayName':
                    optionsValue += 32;
                    break;
            }
        });
        return optionsValue;
    }
}
_SpoFieldAddCommand_instances = new WeakSet(), _SpoFieldAddCommand_initTelemetry = function _SpoFieldAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listTitle: typeof args.options.listTitle !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            options: typeof args.options.options !== 'undefined'
        });
    });
}, _SpoFieldAddCommand_initOptions = function _SpoFieldAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-x, --xml <xml>'
    }, {
        option: '--options [options]'
    });
}, _SpoFieldAddCommand_initValidators = function _SpoFieldAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        const listOptions = [args.options.listId, args.options.listTitle, args.options.listUrl];
        if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
            return `Specify either list id or title or list url, but not multiple`;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        if (args.options.options) {
            let optionsError = true;
            const options = ['DefaultValue', 'AddToDefaultContentType', 'AddToNoContentType', 'AddToAllContentTypes', 'AddFieldInternalNameHint', 'AddFieldToDefaultView', 'AddFieldCheckDisplayName'];
            args.options.options.split(',').forEach(o => {
                o = o.trim();
                if (options.indexOf(o) < 0) {
                    optionsError = `${o} is not a valid value for the options argument. Allowed values are DefaultValue|AddToDefaultContentType|AddToNoContentType|AddToAllContentTypes|AddFieldInternalNameHint|AddFieldToDefaultView|AddFieldCheckDisplayName`;
                }
            });
            return optionsError;
        }
        return true;
    });
};
export default new SpoFieldAddCommand();
//# sourceMappingURL=field-add.js.map