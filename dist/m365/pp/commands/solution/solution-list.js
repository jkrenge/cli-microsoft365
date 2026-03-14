var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpSolutionListCommand_instances, _PpSolutionListCommand_initTelemetry, _PpSolutionListCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpSolutionListCommand extends PowerPlatformCommand {
    get name() {
        return commands.SOLUTION_LIST;
    }
    get description() {
        return 'Lists solutions in a given environment.';
    }
    defaultProperties() {
        return ['uniquename', 'version', 'publisher'];
    }
    constructor() {
        super();
        _PpSolutionListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpSolutionListCommand_instances, "m", _PpSolutionListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpSolutionListCommand_instances, "m", _PpSolutionListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of solutions for which the user is an admin...`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const requestUrl = `${dynamicsApiUrl}/api/data/v9.0/solutions?$filter=isvisible eq true&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`;
            const res = await odata.getAllItems(requestUrl);
            if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
                await logger.log(res);
            }
            else {
                //converted to text friendly output
                await logger.log(res.map(i => {
                    return {
                        uniquename: i.uniquename,
                        version: i.version,
                        publisher: i.publisherid.friendlyname
                    };
                }));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpSolutionListCommand_instances = new WeakSet(), _PpSolutionListCommand_initTelemetry = function _PpSolutionListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpSolutionListCommand_initOptions = function _PpSolutionListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--asAdmin'
    });
};
export default new PpSolutionListCommand();
//# sourceMappingURL=solution-list.js.map