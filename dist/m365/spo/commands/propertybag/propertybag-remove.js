var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPropertyBagRemoveCommand_instances, _SpoPropertyBagRemoveCommand_initTelemetry, _SpoPropertyBagRemoveCommand_initOptions, _SpoPropertyBagRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoPropertyBagBaseCommand } from './propertybag-base.js';
class SpoPropertyBagRemoveCommand extends SpoPropertyBagBaseCommand {
    get name() {
        return commands.PROPERTYBAG_REMOVE;
    }
    get description() {
        return 'Removes specified property from the property bag';
    }
    constructor() {
        super();
        _SpoPropertyBagRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPropertyBagRemoveCommand_instances, "m", _SpoPropertyBagRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPropertyBagRemoveCommand_instances, "m", _SpoPropertyBagRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPropertyBagRemoveCommand_instances, "m", _SpoPropertyBagRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeProperty(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the ${args.options.key} property?` });
            if (result) {
                await this.removeProperty(args);
            }
        }
    }
    async removeProperty(args) {
        try {
            const contextResponse = await spo.getRequestDigest(args.options.webUrl);
            this.formDigestValue = contextResponse.FormDigestValue;
            let identityResp = await spo.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
            const opts = args.options;
            if (opts.folder) {
                // get the folder guid instead of the web guid
                identityResp = await spo.getFolderIdentity(identityResp.objectIdentity, opts.webUrl, opts.folder, this.formDigestValue);
            }
            await this.removePropertyWithIdentityResp(identityResp, args.options);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    async removePropertyWithIdentityResp(identityResp, options) {
        let objectType = 'AllProperties';
        if (options.folder) {
            objectType = 'Properties';
        }
        const requestOptions = {
            url: `${options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': this.formDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${formatting.escapeXml(options.key)}</Parameter><Parameter Type="Null" /></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="${objectType}" /><Identity Id="198" Name="${identityResp.objectIdentity}" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const contents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
            throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
        }
        else {
            return res;
        }
    }
}
_SpoPropertyBagRemoveCommand_instances = new WeakSet(), _SpoPropertyBagRemoveCommand_initTelemetry = function _SpoPropertyBagRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folder: (!(!args.options.folder)).toString(),
            force: args.options.force === true
        });
    });
}, _SpoPropertyBagRemoveCommand_initOptions = function _SpoPropertyBagRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-k, --key <key>'
    }, {
        option: '--folder [folder]'
    }, {
        option: '-f, --force'
    });
}, _SpoPropertyBagRemoveCommand_initValidators = function _SpoPropertyBagRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPropertyBagRemoveCommand();
//# sourceMappingURL=propertybag-remove.js.map