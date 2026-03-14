var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteChromeSetCommand_instances, _SpoSiteChromeSetCommand_initTelemetry, _SpoSiteChromeSetCommand_initOptions, _SpoSiteChromeSetCommand_initTypes, _SpoSiteChromeSetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
var HeaderLayout;
(function (HeaderLayout) {
    HeaderLayout[HeaderLayout["Standard"] = 1] = "Standard";
    HeaderLayout[HeaderLayout["Compact"] = 2] = "Compact";
    HeaderLayout[HeaderLayout["Minimal"] = 3] = "Minimal";
    HeaderLayout[HeaderLayout["Extended"] = 4] = "Extended";
})(HeaderLayout || (HeaderLayout = {}));
var FooterLayout;
(function (FooterLayout) {
    FooterLayout[FooterLayout["Simple"] = 1] = "Simple";
    FooterLayout[FooterLayout["Extended"] = 2] = "Extended";
})(FooterLayout || (FooterLayout = {}));
var Alignment;
(function (Alignment) {
    Alignment[Alignment["Left"] = 0] = "Left";
    Alignment[Alignment["Center"] = 1] = "Center";
    Alignment[Alignment["Right"] = 2] = "Right";
})(Alignment || (Alignment = {}));
var Emphasis;
(function (Emphasis) {
    Emphasis[Emphasis["Lightest"] = 0] = "Lightest";
    Emphasis[Emphasis["Light"] = 1] = "Light";
    Emphasis[Emphasis["Dark"] = 2] = "Dark";
    Emphasis[Emphasis["Darkest"] = 3] = "Darkest";
})(Emphasis || (Emphasis = {}));
class SpoSiteChromeSetCommand extends SpoCommand {
    get name() {
        return commands.SITE_CHROME_SET;
    }
    get description() {
        return 'Set the chrome header and footer for the specified site';
    }
    constructor() {
        super();
        _SpoSiteChromeSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteChromeSetCommand_instances, "m", _SpoSiteChromeSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteChromeSetCommand_instances, "m", _SpoSiteChromeSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteChromeSetCommand_instances, "m", _SpoSiteChromeSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoSiteChromeSetCommand_instances, "m", _SpoSiteChromeSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const headerLayout = args.options.headerLayout ? HeaderLayout[args.options.headerLayout] : null;
        const headerEmphasis = args.options.headerEmphasis ? Emphasis[args.options.headerEmphasis] : null;
        const logoAlignment = args.options.logoAlignment ? Alignment[args.options.logoAlignment] : null;
        const footerLayout = args.options.footerLayout ? FooterLayout[args.options.footerLayout] : null;
        const footerEmphasis = args.options.footerEmphasis ? Emphasis[args.options.footerEmphasis] : null;
        const hideTitleInHeader = typeof args.options.hideTitleInHeader !== "undefined" ? args.options.hideTitleInHeader : null;
        const disableMegaMenu = typeof args.options.disableMegaMenu !== 'undefined' ? args.options.disableMegaMenu : null;
        const disableFooter = typeof args.options.disableFooter !== 'undefined' ? args.options.disableFooter : null;
        const body = {};
        if (headerLayout !== null) {
            body["headerLayout"] = headerLayout;
        }
        if (headerEmphasis !== null) {
            body["headerEmphasis"] = headerEmphasis;
        }
        if (logoAlignment !== null) {
            body["logoAlignment"] = logoAlignment;
        }
        if (footerLayout !== null) {
            body["footerLayout"] = footerLayout;
        }
        if (footerEmphasis !== null) {
            body["footerEmphasis"] = 3 - parseInt(footerEmphasis); // Footer is inverted
        }
        if (hideTitleInHeader !== null) {
            body["hideTitleInHeader"] = hideTitleInHeader;
        }
        if (disableMegaMenu !== null) {
            body["megaMenuEnabled"] = !disableMegaMenu;
        }
        if (disableFooter !== null) {
            body["footerEnabled"] = !disableFooter;
        }
        const requestOptions = {
            url: `${args.options.siteUrl}/_api/web/SetChromeOptions`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: body,
            responseType: 'json'
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteChromeSetCommand_instances = new WeakSet(), _SpoSiteChromeSetCommand_initTelemetry = function _SpoSiteChromeSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            headerLayout: args.options.headerLayout,
            headerEmphasis: args.options.headerEmphasis,
            disableMegaMenu: args.options.disableMegaMenu,
            hideTitleInHeader: args.options.hideTitleInHeader,
            logoAlignment: args.options.logoAlignment,
            disableFooter: args.options.disableFooter,
            footerLayout: args.options.footerLayout,
            footerEmphasis: args.options.footerEmphasis
        });
    });
}, _SpoSiteChromeSetCommand_initOptions = function _SpoSiteChromeSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '--headerLayout [headerLayout]',
        autocomplete: ['Standard', 'Compact', 'Minimal', 'Extended']
    }, {
        option: '--headerEmphasis [headerEmphasis]',
        autocomplete: ['Lightest', 'Light', 'Dark', 'Darkest']
    }, {
        option: '--logoAlignment [logoAlignment]',
        autocomplete: ['Left', 'Center', 'Right']
    }, {
        option: '--footerLayout [footerLayout]',
        autocomplete: ['Simple', 'Extended']
    }, {
        option: '--footerEmphasis [footerEmphasis]',
        autocomplete: ['Lightest', 'Light', 'Dark', 'Darkest']
    }, {
        option: '--disableMegaMenu [disableMegaMenu]',
        autocomplete: ['true', 'false']
    }, {
        option: '--hideTitleInHeader [hideTitleInHeader]',
        autocomplete: ['true', 'false']
    }, {
        option: '--disableFooter [disableFooter]',
        autocomplete: ['true', 'false']
    });
}, _SpoSiteChromeSetCommand_initTypes = function _SpoSiteChromeSetCommand_initTypes() {
    this.types.boolean.push('disableMegaMenu', 'hideTitleInHeader', 'disableFooter');
}, _SpoSiteChromeSetCommand_initValidators = function _SpoSiteChromeSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (typeof args.options.footerEmphasis !== "undefined" && !(args.options.footerEmphasis in Emphasis)) {
            return `${args.options.footerEmphasis} is not a valid option for footerEmphasis. Allowed values Lightest|Light|Dark|Darkest`;
        }
        if (typeof args.options.footerLayout !== "undefined" && !(args.options.footerLayout in FooterLayout)) {
            return `${args.options.footerLayout} is not a valid option for footerLayout. Allowed values Simple|Extended`;
        }
        if (typeof args.options.headerEmphasis !== "undefined" && !(args.options.headerEmphasis in Emphasis)) {
            return `${args.options.headerEmphasis} is not a valid option for headerEmphasis. Allowed values Lightest|Light|Dark|Darkest`;
        }
        if (typeof args.options.headerLayout !== "undefined" && !(args.options.headerLayout in HeaderLayout)) {
            return `${args.options.headerLayout} is not a valid option for headerLayout. Allowed values Standard|Compact|Minimal|Extended`;
        }
        if (typeof args.options.logoAlignment !== "undefined" && !(args.options.logoAlignment in Alignment)) {
            return `${args.options.logoAlignment} is not a valid option for logoAlignment. Allowed values Left|Center|Right`;
        }
        return true;
    });
};
export default new SpoSiteChromeSetCommand();
//# sourceMappingURL=site-chrome-set.js.map