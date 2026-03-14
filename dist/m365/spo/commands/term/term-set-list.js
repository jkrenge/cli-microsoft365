var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermSetListCommand_instances, _SpoTermSetListCommand_initTelemetry, _SpoTermSetListCommand_initOptions, _SpoTermSetListCommand_initValidators, _SpoTermSetListCommand_initOptionSets;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermSetListCommand extends SpoCommand {
    get name() {
        return commands.TERM_SET_LIST;
    }
    get description() {
        return 'Lists taxonomy term sets from the given term group';
    }
    constructor() {
        super();
        _SpoTermSetListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermSetListCommand_instances, "m", _SpoTermSetListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermSetListCommand_instances, "m", _SpoTermSetListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermSetListCommand_instances, "m", _SpoTermSetListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTermSetListCommand_instances, "m", _SpoTermSetListCommand_initOptionSets).call(this);
    }
    defaultProperties() {
        return ['Id', 'Name'];
    }
    async commandAction(logger, args) {
        try {
            const spoWebUrl = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoWebUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving taxonomy term sets...`);
            }
            const query = args.options.termGroupId ? `<Method Id="62" ParentId="60" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
            const requestOptions = {
                url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><Query Id="67" ObjectPathId="65"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" />${query}<Property Id="65" ParentId="62" Name="TermSets" /></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const result = json[json.length - 1];
            if (result._Child_Items_ && result._Child_Items_.length > 0) {
                result._Child_Items_.map(t => {
                    t.CreatedDate = new Date(Number(t.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
                    t.Id = t.Id.replace('/Guid(', '').replace(')/', '');
                    t.LastModifiedDate = new Date(Number(t.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
                    return t;
                });
                await logger.log(result._Child_Items_);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoTermSetListCommand_instances = new WeakSet(), _SpoTermSetListCommand_initTelemetry = function _SpoTermSetListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            termGroupId: typeof args.options.termGroupId !== 'undefined',
            termGroupName: typeof args.options.termGroupName !== 'undefined'
        });
    });
}, _SpoTermSetListCommand_initOptions = function _SpoTermSetListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl [webUrl]'
    }, {
        option: '--termGroupId [termGroupId]'
    }, {
        option: '--termGroupName [termGroupName]'
    });
}, _SpoTermSetListCommand_initValidators = function _SpoTermSetListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.webUrl) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        if (args.options.termGroupId) {
            if (!validation.isValidGuid(args.options.termGroupId)) {
                return `${args.options.termGroupId} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoTermSetListCommand_initOptionSets = function _SpoTermSetListCommand_initOptionSets() {
    this.optionSets.push({ options: ['termGroupId', 'termGroupName'] });
};
export default new SpoTermSetListCommand();
//# sourceMappingURL=term-set-list.js.map