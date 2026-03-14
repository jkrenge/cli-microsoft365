var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebSetCommand_instances, _a, _SpoWebSetCommand_initTelemetry, _SpoWebSetCommand_initOptions, _SpoWebSetCommand_initTypes, _SpoWebSetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebSetCommand extends SpoCommand {
    get name() {
        return commands.WEB_SET;
    }
    get description() {
        return 'Updates subsite properties';
    }
    constructor() {
        super();
        _SpoWebSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebSetCommand_instances, "m", _SpoWebSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoWebSetCommand_instances, "m", _SpoWebSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebSetCommand_instances, "m", _SpoWebSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoWebSetCommand_instances, "m", _SpoWebSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const payload = {};
        this.addUnknownOptionsToPayload(payload, args.options);
        if (args.options.title) {
            payload.Title = args.options.title;
        }
        if (args.options.description) {
            payload.Description = args.options.description;
        }
        if (typeof args.options.siteLogoUrl !== 'undefined') {
            payload.SiteLogoUrl = args.options.siteLogoUrl;
        }
        if (typeof args.options.quickLaunchEnabled !== 'undefined') {
            payload.QuickLaunchEnabled = args.options.quickLaunchEnabled;
        }
        if (typeof args.options.headerEmphasis !== 'undefined') {
            payload.HeaderEmphasis = args.options.headerEmphasis;
        }
        if (typeof args.options.headerLayout !== 'undefined') {
            payload.HeaderLayout = args.options.headerLayout === 'standard' ? 1 : 2;
        }
        if (typeof args.options.megaMenuEnabled !== 'undefined') {
            payload.MegaMenuEnabled = args.options.megaMenuEnabled;
        }
        if (typeof args.options.footerEnabled !== 'undefined') {
            payload.FooterEnabled = args.options.footerEnabled;
        }
        if (typeof args.options.navAudienceTargetingEnabled !== 'undefined') {
            payload.NavAudienceTargetingEnabled = args.options.navAudienceTargetingEnabled;
        }
        if (typeof args.options.searchScope !== 'undefined') {
            const searchScope = args.options.searchScope.toLowerCase();
            payload.SearchScope = _a.searchScopeOptions.indexOf(searchScope);
        }
        try {
            const requestOptions = {
                url: `${args.options.url}/_api/web`,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: payload
            };
            if (this.verbose) {
                await logger.logToStderr(`Updating properties of subsite ${args.options.url}...`);
            }
            await request.patch(requestOptions);
            if (typeof args.options.welcomePage !== 'undefined') {
                if (this.verbose) {
                    await logger.logToStderr(`Setting welcome page to: ${args.options.welcomePage}...`);
                }
                const requestOptions = {
                    url: `${args.options.url}/_api/web/RootFolder`,
                    headers: {
                        'content-type': 'application/json;odata=nometadata',
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json',
                    data: {
                        WelcomePage: args.options.welcomePage
                    }
                };
                await request.patch(requestOptions);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    allowUnknownOptions() {
        return true;
    }
}
_a = SpoWebSetCommand, _SpoWebSetCommand_instances = new WeakSet(), _SpoWebSetCommand_initTelemetry = function _SpoWebSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            headerEmphasis: typeof args.options.headerEmphasis !== 'undefined',
            headerLayout: typeof args.options.headerLayout !== 'undefined',
            megaMenuEnabled: typeof args.options.megaMenuEnabled !== 'undefined',
            siteLogoUrl: typeof args.options.siteLogoUrl !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            quickLaunchEnabled: typeof args.options.quickLaunchEnabled !== 'undefined',
            footerEnabled: typeof args.options.footerEnabled !== 'undefined',
            navAudienceTargetingEnabled: typeof args.options.navAudienceTargetingEnabled !== 'undefined',
            searchScope: typeof args.options.searchScope !== 'undefined',
            welcomePage: typeof args.options.welcomePage !== 'undefined'
        });
        this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
}, _SpoWebSetCommand_initOptions = function _SpoWebSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--siteLogoUrl [siteLogoUrl]'
    }, {
        option: '--quickLaunchEnabled [quickLaunchEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--headerLayout [headerLayout]',
        autocomplete: ['standard', 'compact']
    }, {
        option: '--headerEmphasis [headerEmphasis]',
        autocomplete: ['0', '1', '2', '3']
    }, {
        option: '--megaMenuEnabled [megaMenuEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--footerEnabled [footerEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--navAudienceTargetingEnabled [navAudienceTargetingEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--searchScope [searchScope]',
        autocomplete: _a.searchScopeOptions
    }, {
        option: '--welcomePage [welcomePage]'
    });
}, _SpoWebSetCommand_initTypes = function _SpoWebSetCommand_initTypes() {
    this.types.boolean.push('megaMenuEnabled', 'footerEnabled', 'quickLaunchEnabled', 'navAudienceTargetingEnabled');
}, _SpoWebSetCommand_initValidators = function _SpoWebSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (typeof args.options.headerEmphasis !== 'undefined') {
            if (isNaN(args.options.headerEmphasis)) {
                return `${args.options.headerEmphasis} is not a number`;
            }
            if ([0, 1, 2, 3].indexOf(args.options.headerEmphasis) < 0) {
                return `${args.options.headerEmphasis} is not a valid value for headerEmphasis. Allowed values are 0|1|2|3`;
            }
        }
        if (typeof args.options.headerLayout !== 'undefined') {
            if (['standard', 'compact'].indexOf(args.options.headerLayout) < 0) {
                return `${args.options.headerLayout} is not a valid value for headerLayout. Allowed values are standard|compact`;
            }
        }
        if (typeof args.options.searchScope !== 'undefined') {
            const searchScope = args.options.searchScope.toString().toLowerCase();
            if (_a.searchScopeOptions.indexOf(searchScope) < 0) {
                return `${args.options.searchScope} is not a valid value for searchScope. Allowed values are DefaultScope|Tenant|Hub|Site`;
            }
        }
        return true;
    });
};
SpoWebSetCommand.searchScopeOptions = ['defaultscope', 'tenant', 'hub', 'site'];
export default new SpoWebSetCommand();
//# sourceMappingURL=web-set.js.map