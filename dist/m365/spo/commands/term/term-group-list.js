var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermGroupListCommand_instances, _SpoTermGroupListCommand_initTelemetry, _SpoTermGroupListCommand_initOptions, _SpoTermGroupListCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermGroupListCommand extends SpoCommand {
    get name() {
        return commands.TERM_GROUP_LIST;
    }
    get description() {
        return 'Lists taxonomy term groups';
    }
    defaultProperties() {
        return ['Id', 'Name'];
    }
    constructor() {
        super();
        _SpoTermGroupListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermGroupListCommand_instances, "m", _SpoTermGroupListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermGroupListCommand_instances, "m", _SpoTermGroupListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermGroupListCommand_instances, "m", _SpoTermGroupListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoWebUrl = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoWebUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving taxonomy term groups...`);
            }
            const requestOptions = {
                url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="9"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const result = json[json.length - 1];
            if (result._Child_Items_ && result._Child_Items_.length > 0) {
                result._Child_Items_.forEach(t => {
                    t.CreatedDate = new Date(Number(t.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
                    t.Id = t.Id.replace('/Guid(', '').replace(')/', '');
                    t.LastModifiedDate = new Date(Number(t.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
                });
                await logger.log(result._Child_Items_);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoTermGroupListCommand_instances = new WeakSet(), _SpoTermGroupListCommand_initTelemetry = function _SpoTermGroupListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined'
        });
    });
}, _SpoTermGroupListCommand_initOptions = function _SpoTermGroupListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl [webUrl]'
    });
}, _SpoTermGroupListCommand_initValidators = function _SpoTermGroupListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.webUrl) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        return true;
    });
};
export default new SpoTermGroupListCommand();
//# sourceMappingURL=term-group-list.js.map