var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpCopilotListCommand_instances, _PpCopilotListCommand_initTelemetry, _PpCopilotListCommand_initOptions;
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpCopilotListCommand extends PowerPlatformCommand {
    get name() {
        return commands.COPILOT_LIST;
    }
    get description() {
        return 'Lists Microsoft Power Platform copilots in the specified Power Platform environment';
    }
    defaultProperties() {
        return ['name', 'botid', 'publishedOn', 'createdOn', 'botModifiedOn'];
    }
    constructor() {
        super();
        _PpCopilotListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpCopilotListCommand_instances, "m", _PpCopilotListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpCopilotListCommand_instances, "m", _PpCopilotListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of copilots for environment '${args.options.environmentName}'.`);
        }
        const fetchXml = `
      <fetch mapping='logical' version='1.0' >
        <entity name='bot'>
          <attribute name='accesscontrolpolicy' alias='accessControlPolicy' />,
          <attribute name='applicationmanifestinformation' alias='applicationManifestInformation' />,
          <attribute name='authenticationmode' alias='authenticationMode' />,
          <attribute name='authenticationtrigger' alias='authenticationTrigger' />,
          <attribute name='authorizedsecuritygroupids' alias='authorizedSecurityGroupIds' />,
          <attribute name='componentidunique' alias='componentIdUnique' />,
          <attribute name='componentstate' alias='componentState' />,
          <attribute name='configuration' alias='configuration' />,
          <attribute name='createdon' alias='createdOn' />,
          <attribute name='importsequencenumber' alias='importSequenceNumber' />,
          <attribute name='ismanaged' alias='isManaged' />,
          <attribute name='language' alias='language' />,
          <attribute name='modifiedon' alias='botModifiedOn' />,
          <attribute name='overriddencreatedon' alias='overriddenCreatedOn' />,
          <attribute name='overwritetime' alias='overwriteTime' />,
          <attribute name='iconbase64' alias='iconBase64' />,
          <attribute name='publishedon' alias='publishedOn' />,
          <attribute name='schemaname' alias='schemaName' />,
          <attribute name='solutionid' alias='solutionId' />,
          <attribute name='statecode' alias='stateCode' />,
          <attribute name='statuscode' alias='statusCode' />,
          <attribute name='timezoneruleversionnumber' alias='timezoneRuleVersionNumber' />,
          <attribute name='utcconversiontimezonecode' alias='utcConversionTimezoneCode' />,
          <attribute name='versionnumber' alias='versionNumber' />,
          <attribute name='name' alias='name' />,
          <attribute name='botid' alias='cdsBotId' />,
          <attribute name='ownerid' alias='ownerId' />,
          <attribute name='synchronizationstatus' alias='synchronizationStatus' />
          <link-entity name='systemuser' to='ownerid' from='systemuserid' link-type='inner' >
            <attribute name='fullname' alias='owner' />
          </link-entity>
          <link-entity name='systemuser' to='modifiedby' from='systemuserid' link-type='inner' >
            <attribute name='fullname' alias='botModifiedBy' />
          </link-entity>
        </entity>
      </fetch>
    `;
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const items = await odata.getAllItems(`${dynamicsApiUrl}/api/data/v9.1/bots?fetchXml=${fetchXml}`);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpCopilotListCommand_instances = new WeakSet(), _PpCopilotListCommand_initTelemetry = function _PpCopilotListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpCopilotListCommand_initOptions = function _PpCopilotListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--asAdmin'
    });
};
export default new PpCopilotListCommand();
//# sourceMappingURL=copilot-list.js.map