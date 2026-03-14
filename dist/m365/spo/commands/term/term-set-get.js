var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermSetGetCommand_instances, _SpoTermSetGetCommand_initTelemetry, _SpoTermSetGetCommand_initOptions, _SpoTermSetGetCommand_initValidators, _SpoTermSetGetCommand_initOptionSets;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermSetGetCommand extends SpoCommand {
    get name() {
        return commands.TERM_SET_GET;
    }
    get description() {
        return 'Gets information about the specified taxonomy term set';
    }
    constructor() {
        super();
        _SpoTermSetGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermSetGetCommand_instances, "m", _SpoTermSetGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermSetGetCommand_instances, "m", _SpoTermSetGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermSetGetCommand_instances, "m", _SpoTermSetGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTermSetGetCommand_instances, "m", _SpoTermSetGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoWebUrl = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoWebUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving taxonomy term set...`);
            }
            const termGroupQuery = args.options.termGroupId ? `<Method Id="62" ParentId="60" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
            const termSetQuery = args.options.id ? `<Method Id="67" ParentId="65" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.id}}</Parameter></Parameters></Method>` : `<Method Id="67" ParentId="65" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter></Parameters></Method>`;
            const requestOptions = {
                url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" />${termGroupQuery}<Property Id="65" ParentId="62" Name="TermSets" />${termSetQuery}</ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const termSet = json[json.length - 1];
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
_SpoTermSetGetCommand_instances = new WeakSet(), _SpoTermSetGetCommand_initTelemetry = function _SpoTermSetGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            termGroupId: typeof args.options.termGroupId !== 'undefined',
            termGroupName: typeof args.options.termGroupName !== 'undefined'
        });
    });
}, _SpoTermSetGetCommand_initOptions = function _SpoTermSetGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl [webUrl]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--termGroupId [termGroupId]'
    }, {
        option: '--termGroupName [termGroupName]'
    });
}, _SpoTermSetGetCommand_initValidators = function _SpoTermSetGetCommand_initValidators() {
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
        return true;
    });
}, _SpoTermSetGetCommand_initOptionSets = function _SpoTermSetGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] }, { options: ['termGroupId', 'termGroupName'] });
};
export default new SpoTermSetGetCommand();
//# sourceMappingURL=term-set-get.js.map