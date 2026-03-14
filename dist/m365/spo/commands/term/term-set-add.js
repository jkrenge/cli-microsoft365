var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermSetAddCommand_instances, _SpoTermSetAddCommand_initTelemetry, _SpoTermSetAddCommand_initOptions, _SpoTermSetAddCommand_initValidators, _SpoTermSetAddCommand_initOptionSets;
import { v4 } from 'uuid';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermSetAddCommand extends SpoCommand {
    get name() {
        return commands.TERM_SET_ADD;
    }
    get description() {
        return 'Adds taxonomy term set';
    }
    constructor() {
        super();
        _SpoTermSetAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermSetAddCommand_instances, "m", _SpoTermSetAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermSetAddCommand_instances, "m", _SpoTermSetAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermSetAddCommand_instances, "m", _SpoTermSetAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTermSetAddCommand_instances, "m", _SpoTermSetAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let termSet;
        try {
            const spoWebUrl = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoWebUrl);
            const formDigest = res.FormDigestValue;
            if (this.verbose) {
                await logger.logToStderr(`Adding taxonomy term set...`);
            }
            const termGroupQuery = args.options.termGroupName ?
                `<Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>` :
                `<Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>`;
            const termSetId = args.options.id || v4();
            const requestOptionsQuery = {
                url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" />${termGroupQuery}<Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter><Parameter Type="Guid">{${termSetId}}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptionsQuery);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            termSet = json[json.length - 1];
            let terms = undefined;
            if (!(!args.options.description &&
                !args.options.customProperties)) {
                let termStoreObjectIdentity = '';
                // get term store object identity
                for (let i = 0; i < json.length; i++) {
                    if (json[i] !== 39) {
                        continue;
                    }
                    termStoreObjectIdentity = json[i + 1]._ObjectIdentity_;
                    break;
                }
                if (this.verbose) {
                    await logger.logToStderr(`Setting term set properties...`);
                }
                const properties = [];
                let i = 127;
                if (args.options.description) {
                    properties.push(`<SetProperty Id="${i++}" ObjectPathId="117" Name="Description"><Parameter Type="String">${formatting.escapeXml(args.options.description)}</Parameter></SetProperty>`);
                    termSet.Description = args.options.description;
                }
                if (args.options.customProperties) {
                    const customProperties = JSON.parse(args.options.customProperties);
                    Object.keys(customProperties).forEach(k => {
                        properties.push(`<Method Name="SetCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${formatting.escapeXml(k)}</Parameter><Parameter Type="String">${formatting.escapeXml(customProperties[k])}</Parameter></Parameters></Method>`);
                    });
                    termSet.CustomProperties = customProperties;
                }
                const requestOptions = {
                    url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': formDigest
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${properties.join('')}<Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="${termSet._ObjectIdentity_}" /><Identity Id="109" Name="${termStoreObjectIdentity}" /></ObjectPaths></Request>`
                };
                terms = await request.post(requestOptions);
            }
            if (terms) {
                const json = JSON.parse(terms);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
            }
            delete termSet._ObjectIdentity_;
            delete termSet._ObjectType_;
            termSet.CreatedDate = new Date(Number(termSet.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
            termSet.Id = termSet.Id.replace('/Guid(', '').replace(')/', '');
            termSet.LastModifiedDate = new Date(Number(termSet.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
            await logger.log(termSet);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoTermSetAddCommand_instances = new WeakSet(), _SpoTermSetAddCommand_initTelemetry = function _SpoTermSetAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            customProperties: typeof args.options.customProperties !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            termGroupId: typeof args.options.termGroupId !== 'undefined',
            termGroupName: typeof args.options.termGroupName !== 'undefined'
        });
    });
}, _SpoTermSetAddCommand_initOptions = function _SpoTermSetAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-u, --webUrl [webUrl]'
    }, {
        option: '--termGroupId [termGroupId]'
    }, {
        option: '--termGroupName [termGroupName]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--customProperties [customProperties]'
    });
}, _SpoTermSetAddCommand_initValidators = function _SpoTermSetAddCommand_initValidators() {
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
        if (args.options.termGroupId) {
            if (!validation.isValidGuid(args.options.termGroupId)) {
                return `${args.options.termGroupId} is not a valid GUID`;
            }
        }
        if (args.options.customProperties) {
            try {
                JSON.parse(args.options.customProperties);
            }
            catch (e) {
                return `Error when parsing customProperties JSON: ${e}`;
            }
        }
        return true;
    });
}, _SpoTermSetAddCommand_initOptionSets = function _SpoTermSetAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['termGroupId', 'termGroupName'] });
};
export default new SpoTermSetAddCommand();
//# sourceMappingURL=term-set-add.js.map