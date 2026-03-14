var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppAddCommand_instances, _SpoAppAddCommand_initTelemetry, _SpoAppAddCommand_initOptions, _SpoAppAddCommand_initValidators;
import fs from 'fs';
import path from 'path';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
class SpoAppAddCommand extends SpoAppBaseCommand {
    get name() {
        return commands.APP_ADD;
    }
    get description() {
        return 'Adds an app to the specified SharePoint Online app catalog';
    }
    constructor() {
        super();
        _SpoAppAddCommand_instances.add(this);
        this.appCatalogScopeOptions = ['tenant', 'sitecollection'];
        __classPrivateFieldGet(this, _SpoAppAddCommand_instances, "m", _SpoAppAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppAddCommand_instances, "m", _SpoAppAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppAddCommand_instances, "m", _SpoAppAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const scope = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
        const overwrite = args.options.overwrite || false;
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);
            const fullPath = path.resolve(args.options.filePath);
            if (this.verbose) {
                await logger.logToStderr(`Adding app '${fullPath}' to app catalog...`);
            }
            const fileName = path.basename(fullPath);
            const requestOptions = {
                url: `${appCatalogUrl}/_api/web/${scope}appcatalog/Add(overwrite=${(overwrite.toString().toLowerCase())}, url='${fileName}')`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    binaryStringRequestBody: 'true'
                },
                data: fs.readFileSync(fullPath)
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log(json);
            }
            else {
                await logger.log(json.UniqueId);
            }
        }
        catch (err) {
            this.handleRejectedODataPromise(err);
        }
    }
}
_SpoAppAddCommand_instances = new WeakSet(), _SpoAppAddCommand_initTelemetry = function _SpoAppAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            overwrite: !!args.options.overwrite,
            appCatalogScope: args.options.appCatalogScope || 'tenant',
            appCatalogUrl: typeof args.options.appCatalogUrl !== 'undefined'
        });
    });
}, _SpoAppAddCommand_initOptions = function _SpoAppAddCommand_initOptions() {
    this.options.unshift({
        option: '-p, --filePath <filePath>'
    }, {
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: this.appCatalogScopeOptions
    }, {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
    }, {
        option: '--overwrite [overwrite]'
    });
}, _SpoAppAddCommand_initValidators = function _SpoAppAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appCatalogScope) {
            const appCatalogScope = args.options.appCatalogScope.toLowerCase();
            if (this.appCatalogScopeOptions.indexOf(appCatalogScope) === -1) {
                return `${args.options.appCatalogScope} is not a valid appCatalogScope. Allowed values are: ${this.appCatalogScopeOptions.join(', ')}`;
            }
            if (appCatalogScope === 'sitecollection' && !args.options.appCatalogUrl) {
                return `You must specify appCatalogUrl when appCatalogScope is sitecollection`;
            }
        }
        const fullPath = path.resolve(args.options.filePath);
        if (!fs.existsSync(fullPath)) {
            return `File '${fullPath}' not found`;
        }
        if (fs.lstatSync(fullPath).isDirectory()) {
            return `Path '${fullPath}' points to a directory`;
        }
        if (args.options.appCatalogUrl) {
            return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }
        return true;
    });
};
export default new SpoAppAddCommand();
//# sourceMappingURL=app-add.js.map