var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteListCommand_instances, _SpoSiteListCommand_initTelemetry, _SpoSiteListCommand_initOptions, _SpoSiteListCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteListCommand extends SpoCommand {
    get name() {
        return commands.SITE_LIST;
    }
    get description() {
        return 'Lists sites of the given type';
    }
    defaultProperties() {
        return ['Title', 'Url'];
    }
    constructor() {
        super();
        _SpoSiteListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteListCommand_instances, "m", _SpoSiteListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteListCommand_instances, "m", _SpoSiteListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteListCommand_instances, "m", _SpoSiteListCommand_initValidators).call(this);
    }
    alias() {
        return [commands.SITE_LIST];
    }
    async commandAction(logger, args) {
        const webTemplate = this.getWebTemplateId(args.options);
        const includeOneDriveSites = args.options.withOneDriveSites || false;
        const personalSite = includeOneDriveSites === false ? '0' : '1';
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving list of site collections...`);
            }
            this.allSites = [];
            await this.getAllSites(spoAdminUrl, formatting.escapeXml(args.options.filter || ''), '0', personalSite, webTemplate, undefined, logger);
            await logger.log(this.allSites);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAllSites(spoAdminUrl, filter, startIndex, personalSite, webTemplate, formDigest, logger) {
        const res = await spo.ensureFormDigest(spoAdminUrl, logger, formDigest, this.debug);
        const requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">${filter}</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">${personalSite}</Property><Property Name="StartIndex" Type="String">${startIndex}</Property><Property Name="Template" Type="String">${webTemplate}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;
        const requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': res.FormDigestValue
            },
            data: requestBody
        };
        const response = await request.post(requestOptions);
        const json = JSON.parse(response);
        const responseContent = json[0];
        if (responseContent.ErrorInfo) {
            throw responseContent.ErrorInfo.ErrorMessage;
        }
        else {
            const sites = json[json.length - 1];
            this.allSites.push(...sites._Child_Items_);
            if (sites.NextStartIndexFromSharePoint) {
                await this.getAllSites(spoAdminUrl, filter, sites.NextStartIndexFromSharePoint, personalSite, webTemplate, formDigest, logger);
            }
            return;
        }
    }
    getWebTemplateId(options) {
        if (options.webTemplate) {
            return options.webTemplate;
        }
        if (options.withOneDriveSites) {
            return '';
        }
        switch (options.type) {
            case "TeamSite":
                return 'GROUP#0';
            case "CommunicationSite":
                return 'SITEPAGEPUBLISHING#0';
            default:
                return '';
        }
    }
}
_SpoSiteListCommand_instances = new WeakSet(), _SpoSiteListCommand_initTelemetry = function _SpoSiteListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webTemplate: args.options.webTemplate,
            type: args.options.type,
            filter: (!(!args.options.filter)).toString(),
            withOneDriveSites: typeof args.options.withOneDriveSites !== 'undefined'
        });
    });
}, _SpoSiteListCommand_initOptions = function _SpoSiteListCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type [type]',
        autocomplete: ['TeamSite', 'CommunicationSite']
    }, {
        option: '--webTemplate [webTemplate]'
    }, {
        option: '--filter [filter]'
    }, {
        option: '--withOneDriveSites'
    });
}, _SpoSiteListCommand_initValidators = function _SpoSiteListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type && args.options.webTemplate) {
            return 'Specify either type or webTemplate, but not both';
        }
        const typeValues = ['TeamSite', 'CommunicationSite'];
        if (args.options.type &&
            typeValues.indexOf(args.options.type) < 0) {
            return `${args.options.type} is not a valid value for the type option. Allowed values are ${typeValues.join('|')}`;
        }
        if (args.options.withOneDriveSites
            && (args.options.type || args.options.webTemplate)) {
            return 'When using withOneDriveSites, don\'t specify the type or webTemplate options';
        }
        return true;
    });
};
export default new SpoSiteListCommand();
//# sourceMappingURL=site-list.js.map