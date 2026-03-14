var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermGroupAddCommand_instances, _SpoTermGroupAddCommand_initTelemetry, _SpoTermGroupAddCommand_initOptions, _SpoTermGroupAddCommand_initValidators;
import { v4 } from 'uuid';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermGroupAddCommand extends SpoCommand {
    get name() {
        return commands.TERM_GROUP_ADD;
    }
    get description() {
        return 'Adds taxonomy term group';
    }
    constructor() {
        super();
        _SpoTermGroupAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermGroupAddCommand_instances, "m", _SpoTermGroupAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermGroupAddCommand_instances, "m", _SpoTermGroupAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermGroupAddCommand_instances, "m", _SpoTermGroupAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let termGroup;
        try {
            const webUrl = args.options.webUrl || await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(webUrl);
            const formDigest = res.FormDigestValue;
            if (this.verbose) {
                await logger.logToStderr(`Getting taxonomy term store...`);
            }
            const requestOptionsPost = {
                url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptionsPost);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const termStore = json[json.length - 1];
            const termGroupId = args.options.id || v4();
            if (this.verbose) {
                await logger.logToStderr(`Adding taxonomy term group...`);
            }
            const requestOptions = {
                url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': formDigest
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter><Parameter Type="Guid">{${termGroupId}}</Parameter></Parameters></Method><Identity Id="6" Name="${termStore._ObjectIdentity_}" /></ObjectPaths></Request>`
            };
            const terms = await request.post(requestOptions);
            const json2 = JSON.parse(terms);
            const response2 = json2[0];
            if (response2.ErrorInfo) {
                throw response2.ErrorInfo.ErrorMessage;
            }
            termGroup = json2[json2.length - 1];
            let termGroups = undefined;
            if (args.options.description) {
                if (this.verbose) {
                    await logger.logToStderr(`Setting taxonomy term group description...`);
                }
                const requestOptionsQuery = {
                    url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': formDigest
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">${formatting.escapeXml(args.options.description)}</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="${termGroup._ObjectIdentity_}" /></ObjectPaths></Request>`
                };
                termGroups = await request.post(requestOptionsQuery);
            }
            if (termGroups) {
                const json = JSON.parse(termGroups);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
            }
            delete termGroup._ObjectIdentity_;
            delete termGroup._ObjectType_;
            termGroup.Id = termGroup.Id.replace('/Guid(', '').replace(')/', '');
            termGroup.Description = args.options.description || '';
            await logger.log(termGroup);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoTermGroupAddCommand_instances = new WeakSet(), _SpoTermGroupAddCommand_initTelemetry = function _SpoTermGroupAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.id !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            webUrl: typeof args.options.webUrl !== 'undefined'
        });
    });
}, _SpoTermGroupAddCommand_initOptions = function _SpoTermGroupAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-u, --webUrl [webUrl]'
    });
}, _SpoTermGroupAddCommand_initValidators = function _SpoTermGroupAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.webUrl) {
            return validation.isValidSharePointUrl(args.options.webUrl);
        }
        return true;
    });
};
export default new SpoTermGroupAddCommand();
//# sourceMappingURL=term-group-add.js.map