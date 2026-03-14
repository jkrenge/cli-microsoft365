var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoGroupGetCommand_instances, _SpoGroupGetCommand_initTelemetry, _SpoGroupGetCommand_initOptions, _SpoGroupGetCommand_initValidators, _SpoGroupGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoGroupGetCommand extends SpoCommand {
    get name() {
        return commands.GROUP_GET;
    }
    get description() {
        return 'Gets site group';
    }
    constructor() {
        super();
        _SpoGroupGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoGroupGetCommand_instances, "m", _SpoGroupGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoGroupGetCommand_instances, "m", _SpoGroupGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoGroupGetCommand_instances, "m", _SpoGroupGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoGroupGetCommand_instances, "m", _SpoGroupGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information for group in site at ${args.options.webUrl}...`);
        }
        let requestUrl = '';
        if (args.options.id) {
            requestUrl = `${args.options.webUrl}/_api/web/sitegroups/GetById('${args.options.id}')`;
        }
        else if (args.options.name) {
            requestUrl = `${args.options.webUrl}/_api/web/sitegroups/GetByName('${formatting.encodeQueryParameter(args.options.name)}')`;
        }
        else if (args.options.associatedGroup) {
            switch (args.options.associatedGroup.toLowerCase()) {
                case 'owner':
                    requestUrl = `${args.options.webUrl}/_api/web/AssociatedOwnerGroup`;
                    break;
                case 'member':
                    requestUrl = `${args.options.webUrl}/_api/web/AssociatedMemberGroup`;
                    break;
                case 'visitor':
                    requestUrl = `${args.options.webUrl}/_api/web/AssociatedVisitorGroup`;
                    break;
            }
        }
        const requestOptions = {
            url: requestUrl,
            method: 'GET',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const groupInstance = await request.get(requestOptions);
            await logger.log(groupInstance);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoGroupGetCommand_instances = new WeakSet(), _SpoGroupGetCommand_initTelemetry = function _SpoGroupGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: (!(!args.options.id)).toString(),
            name: (!(!args.options.name)).toString(),
            associatedGroup: args.options.associatedGroup
        });
    });
}, _SpoGroupGetCommand_initOptions = function _SpoGroupGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--name [name]'
    }, {
        option: '--associatedGroup [associatedGroup]',
        autocomplete: ['Owner', 'Member', 'Visitor']
    });
}, _SpoGroupGetCommand_initValidators = function _SpoGroupGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && isNaN(args.options.id)) {
            return `Specified id ${args.options.id} is not a number`;
        }
        if (args.options.associatedGroup && ['owner', 'member', 'visitor'].indexOf(args.options.associatedGroup.toLowerCase()) === -1) {
            return `${args.options.associatedGroup} is not a valid associatedGroup value. Allowed values are Owner|Member|Visitor.`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoGroupGetCommand_initOptionSets = function _SpoGroupGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name', 'associatedGroup'] });
};
export default new SpoGroupGetCommand();
//# sourceMappingURL=group-get.js.map