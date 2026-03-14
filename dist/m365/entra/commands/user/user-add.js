var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserAddCommand_instances, _EntraUserAddCommand_initTelemetry, _EntraUserAddCommand_initOptions, _EntraUserAddCommand_initValidators, _EntraUserAddCommand_initOptionSets, _EntraUserAddCommand_initTypes;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserAddCommand extends GraphCommand {
    get name() {
        return commands.USER_ADD;
    }
    get description() {
        return 'Creates a new user';
    }
    allowUnknownOptions() {
        return true;
    }
    constructor() {
        super();
        _EntraUserAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserAddCommand_instances, "m", _EntraUserAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserAddCommand_instances, "m", _EntraUserAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserAddCommand_instances, "m", _EntraUserAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraUserAddCommand_instances, "m", _EntraUserAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraUserAddCommand_instances, "m", _EntraUserAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding user to AAD with displayName ${args.options.displayName} and userPrincipalName ${args.options.userName}`);
        }
        try {
            const manifest = this.mapRequestBody(args.options);
            const requestOptions = {
                url: `${this.resource}/v1.0/users`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: manifest
            };
            const user = await request.post(requestOptions);
            user.password = requestOptions.data.passwordProfile.password;
            if (args.options.managerUserId || args.options.managerUserName) {
                const managerRequestOptions = {
                    url: `${this.resource}/v1.0/users/${user.id}/manager/$ref`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    data: {
                        '@odata.id': `${this.resource}/v1.0/users/${args.options.managerUserId || args.options.managerUserName}`
                    }
                };
                await request.put(managerRequestOptions);
            }
            await logger.log(user);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapRequestBody(options) {
        const requestBody = {
            accountEnabled: options.accountEnabled ?? true,
            displayName: options.displayName,
            userPrincipalName: options.userName,
            mailNickName: options.mailNickname ?? options.userName.split('@')[0],
            passwordProfile: {
                forceChangePasswordNextSignIn: options.forceChangePasswordNextSignIn || false,
                forceChangePasswordNextSignInWithMfa: options.forceChangePasswordNextSignInWithMfa || false,
                password: options.password ?? this.generatePassword()
            },
            givenName: options.firstName,
            surName: options.lastName,
            usageLocation: options.usageLocation,
            officeLocation: options.officeLocation,
            jobTitle: options.jobTitle,
            companyName: options.companyName,
            department: options.department,
            preferredLanguage: options.preferredLanguage
        };
        this.addUnknownOptionsToPayload(requestBody, options);
        return requestBody;
    }
    /**
     * Generate a password with at least: one digit, one lowercase character, one uppercase character, and a special character.
     */
    generatePassword() {
        const numberChars = '0123456789';
        const upperChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        const lowerChars = 'abcdefghijklmnopqrstuvwxyz';
        const specialChars = '-_@%$#*&';
        const allChars = numberChars + upperChars + lowerChars + specialChars;
        let randPasswordArray = Array(15);
        randPasswordArray[0] = numberChars;
        randPasswordArray[1] = upperChars;
        randPasswordArray[2] = lowerChars;
        randPasswordArray[3] = specialChars;
        randPasswordArray = randPasswordArray.fill(allChars, 4);
        const randomCharacterArray = randPasswordArray.map((charSet) => charSet[Math.floor(Math.random() * charSet.length)]);
        return this.shuffleArray(randomCharacterArray).join('');
    }
    shuffleArray(characterArray) {
        for (let i = characterArray.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            const temp = characterArray[i];
            characterArray[i] = characterArray[j];
            characterArray[j] = temp;
        }
        return characterArray;
    }
}
_EntraUserAddCommand_instances = new WeakSet(), _EntraUserAddCommand_initTelemetry = function _EntraUserAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            accountEnabled: typeof args.options.accountEnabled !== 'undefined',
            mailNickname: typeof args.options.mailNickname !== 'undefined',
            password: typeof args.options.password !== 'undefined',
            firstName: typeof args.options.firstName !== 'undefined',
            lastName: typeof args.options.lastName !== 'undefined',
            forceChangePasswordNextSignIn: !!args.options.forceChangePasswordNextSignIn,
            forceChangePasswordNextSignInWithMfa: !!args.options.forceChangePasswordNextSignInWithMfa,
            usageLocation: typeof args.options.usageLocation !== 'undefined',
            officeLocation: typeof args.options.officeLocation !== 'undefined',
            jobTitle: typeof args.options.jobTitle !== 'undefined',
            companyName: typeof args.options.companyName !== 'undefined',
            department: typeof args.options.department !== 'undefined',
            preferredLanguage: typeof args.options.preferredLanguage !== 'undefined',
            managerUserId: typeof args.options.managerUserId !== 'undefined',
            managerUserName: typeof args.options.managerUserName !== 'undefined'
        });
        this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
}, _EntraUserAddCommand_initOptions = function _EntraUserAddCommand_initOptions() {
    this.options.unshift({
        option: '--displayName <displayName>'
    }, {
        option: '--userName <userName>'
    }, {
        option: '--accountEnabled [accountEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--mailNickname [mailNickname]'
    }, {
        option: '--password [password]'
    }, {
        option: '--firstName [firstName]'
    }, {
        option: '--lastName [lastName]'
    }, {
        option: '--forceChangePasswordNextSignIn'
    }, {
        option: '--forceChangePasswordNextSignInWithMfa'
    }, {
        option: '--usageLocation [usageLocation]'
    }, {
        option: '--officeLocation [officeLocation]'
    }, {
        option: '--jobTitle [jobTitle]'
    }, {
        option: '--companyName [companyName]'
    }, {
        option: '--department [department]'
    }, {
        option: '--preferredLanguage [preferredLanguage]'
    }, {
        option: '--managerUserId [managerUserId]'
    }, {
        option: '--managerUserName [managerUserName]'
    });
}, _EntraUserAddCommand_initValidators = function _EntraUserAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        if (args.options.usageLocation) {
            const regex = new RegExp('^[a-zA-Z]{2}$');
            if (!regex.test(args.options.usageLocation)) {
                return `'${args.options.usageLocation}' is not a valid usageLocation.`;
            }
        }
        if (args.options.preferredLanguage && args.options.preferredLanguage.length < 2) {
            return `'${args.options.preferredLanguage}' is not a valid preferredLanguage`;
        }
        if (args.options.firstName && args.options.firstName.length > 64) {
            return `The maximum amount of characters for 'firstName' is 64.`;
        }
        if (args.options.lastName && args.options.lastName.length > 64) {
            return `The maximum amount of characters for 'lastName' is 64.`;
        }
        if (args.options.jobTitle && args.options.jobTitle.length > 128) {
            return `The maximum amount of characters for 'jobTitle' is 128.`;
        }
        if (args.options.companyName && args.options.companyName.length > 64) {
            return `The maximum amount of characters for 'companyName' is 64.`;
        }
        if (args.options.department && args.options.department.length > 64) {
            return `The maximum amount of characters for 'department' is 64.`;
        }
        if (args.options.managerUserName && !validation.isValidUserPrincipalName(args.options.managerUserName)) {
            return `'${args.options.managerUserName}' is not a valid user principal name.`;
        }
        if (args.options.managerUserId && !validation.isValidGuid(args.options.managerUserId)) {
            return `'${args.options.managerUserId}' is not a valid GUID.`;
        }
        return true;
    });
}, _EntraUserAddCommand_initOptionSets = function _EntraUserAddCommand_initOptionSets() {
    this.optionSets.push({
        options: ['managerUserId', 'managerUserName'],
        runsWhen: (args) => args.options.managerId || args.options.managerUserName
    });
}, _EntraUserAddCommand_initTypes = function _EntraUserAddCommand_initTypes() {
    this.types.boolean.push('accountEnabled');
};
export default new EntraUserAddCommand();
//# sourceMappingURL=user-add.js.map