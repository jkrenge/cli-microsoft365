var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermGroupGetCommand_instances, _SpoTermGroupGetCommand_initTelemetry, _SpoTermGroupGetCommand_initOptions, _SpoTermGroupGetCommand_initValidators, _SpoTermGroupGetCommand_initOptionSets;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermGroupGetCommand extends SpoCommand {
    get name() {
        return commands.TERM_GROUP_GET;
    }
    get description() {
        return 'Gets information about the specified taxonomy term group';
    }
    constructor() {
        super();
        _SpoTermGroupGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermGroupGetCommand_instances, "m", _SpoTermGroupGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermGroupGetCommand_instances, "m", _SpoTermGroupGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermGroupGetCommand_instances, "m", _SpoTermGroupGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTermGroupGetCommand_instances, "m", _SpoTermGroupGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoWebUrl = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoWebUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving taxonomy term groups...`);
            }
            const query = args.options.id ? `<Method Id="32" ParentId="30" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.id)}}</Parameter></Parameters></Method>` : `<Method Id="32" ParentId="30" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter></Parameters></Method>`;
            const requestOptions = {
                url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="25" ObjectPathId="24" /><ObjectIdentityQuery Id="26" ObjectPathId="24" /><ObjectPath Id="28" ObjectPathId="27" /><ObjectIdentityQuery Id="29" ObjectPathId="27" /><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectIdentityQuery Id="34" ObjectPathId="32" /><Query Id="35" ObjectPathId="32"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="24" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="27" ParentId="24" Name="GetDefaultSiteCollectionTermStore" /><Property Id="30" ParentId="27" Name="Groups" />${query}</ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const termGroup = json[json.length - 1];
            delete termGroup._ObjectIdentity_;
            delete termGroup._ObjectType_;
            termGroup.CreatedDate = new Date(Number(termGroup.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
            termGroup.Id = termGroup.Id.replace('/Guid(', '').replace(')/', '');
            termGroup.LastModifiedDate = new Date(Number(termGroup.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
            await logger.log(termGroup);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoTermGroupGetCommand_instances = new WeakSet(), _SpoTermGroupGetCommand_initTelemetry = function _SpoTermGroupGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _SpoTermGroupGetCommand_initOptions = function _SpoTermGroupGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl [webUrl]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    });
}, _SpoTermGroupGetCommand_initValidators = function _SpoTermGroupGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.webUrl) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoTermGroupGetCommand_initOptionSets = function _SpoTermGroupGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoTermGroupGetCommand();
//# sourceMappingURL=term-group-get.js.map