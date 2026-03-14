var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileCheckinCommand_instances, _SpoFileCheckinCommand_initTelemetry, _SpoFileCheckinCommand_initOptions, _SpoFileCheckinCommand_initValidators, _SpoFileCheckinCommand_initOptionSets, _SpoFileCheckinCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
var CheckinType;
(function (CheckinType) {
    CheckinType[CheckinType["Minor"] = 0] = "Minor";
    CheckinType[CheckinType["Major"] = 1] = "Major";
    CheckinType[CheckinType["Overwrite"] = 2] = "Overwrite";
})(CheckinType || (CheckinType = {}));
class SpoFileCheckinCommand extends SpoCommand {
    get name() {
        return commands.FILE_CHECKIN;
    }
    get description() {
        return 'Checks in specified file';
    }
    constructor() {
        super();
        _SpoFileCheckinCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileCheckinCommand_instances, "m", _SpoFileCheckinCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckinCommand_instances, "m", _SpoFileCheckinCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckinCommand_instances, "m", _SpoFileCheckinCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckinCommand_instances, "m", _SpoFileCheckinCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckinCommand_instances, "m", _SpoFileCheckinCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        let type = CheckinType.Major;
        if (args.options.type) {
            switch (args.options.type.toLowerCase()) {
                case 'minor':
                    type = CheckinType.Minor;
                    break;
                case 'overwrite':
                    type = CheckinType.Overwrite;
            }
        }
        let comment = '';
        if (args.options.comment) {
            comment = formatting.encodeQueryParameter(args.options.comment);
        }
        let requestUrl = '';
        if (args.options.id) {
            requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(args.options.id)}')/checkin(comment='${comment}',checkintype=${type})`;
        }
        if (args.options.url) {
            const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
            requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/checkin(comment='${comment}',checkintype=${type})`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
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
_SpoFileCheckinCommand_instances = new WeakSet(), _SpoFileCheckinCommand_initTelemetry = function _SpoFileCheckinCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            type: args.options.type || 'Major',
            comment: typeof args.options.comment !== 'undefined'
        });
    });
}, _SpoFileCheckinCommand_initOptions = function _SpoFileCheckinCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--url [url]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --type [type]',
        autocomplete: ['Minor', 'Major', 'Overwrite']
    }, {
        option: '--comment [comment]'
    });
}, _SpoFileCheckinCommand_initValidators = function _SpoFileCheckinCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        if (args.options.comment && args.options.comment.length > 1023) {
            return 'The length of the comment must be less than 1024 letters';
        }
        if (args.options.type) {
            const allowedValues = ['minor', 'major', 'overwrite'];
            const type = args.options.type.toLowerCase();
            if (allowedValues.indexOf(type) === -1) {
                return 'Wrong type specified. Available values are Minor|Major|Overwrite';
            }
        }
        return true;
    });
}, _SpoFileCheckinCommand_initOptionSets = function _SpoFileCheckinCommand_initOptionSets() {
    this.optionSets.push({ options: ['url', 'id'] });
}, _SpoFileCheckinCommand_initTypes = function _SpoFileCheckinCommand_initTypes() {
    this.types.string.push('webUrl', 'url', 'id', 'type', 'comment');
};
export default new SpoFileCheckinCommand();
//# sourceMappingURL=file-checkin.js.map