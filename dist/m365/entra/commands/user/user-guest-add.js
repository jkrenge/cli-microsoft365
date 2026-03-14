var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserGuestAddCommand_instances, _EntraUserGuestAddCommand_initTelemetry, _EntraUserGuestAddCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserGuestAddCommand extends GraphCommand {
    get name() {
        return commands.USER_GUEST_ADD;
    }
    get description() {
        return 'Invite an external user to the organization';
    }
    constructor() {
        super();
        _EntraUserGuestAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserGuestAddCommand_instances, "m", _EntraUserGuestAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserGuestAddCommand_instances, "m", _EntraUserGuestAddCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/invitations`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    invitedUserEmailAddress: args.options.emailAddress,
                    inviteRedirectUrl: args.options.inviteRedirectUrl || 'https://myapplications.microsoft.com',
                    invitedUserDisplayName: args.options.displayName,
                    sendInvitationMessage: args.options.sendInvitationMessage,
                    invitedUserMessageInfo: {
                        customizedMessageBody: args.options.welcomeMessage,
                        messageLanguage: args.options.messageLanguage || 'en-US',
                        ccRecipients: args.options.ccRecipients ? this.mapEmailsToRecipients([args.options.ccRecipients]) : undefined
                    }
                }
            };
            const result = await request.post(requestOptions);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapEmailsToRecipients(emails) {
        return emails.map(mail => ({
            emailAddress: {
                address: mail.trim()
            }
        }));
    }
}
_EntraUserGuestAddCommand_instances = new WeakSet(), _EntraUserGuestAddCommand_initTelemetry = function _EntraUserGuestAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            displayName: typeof args.options.displayName !== 'undefined',
            inviteRedirectUrl: typeof args.options.inviteRedirectUrl !== 'undefined',
            welcomeMessage: typeof args.options.welcomeMessage !== 'undefined',
            messageLanguage: typeof args.options.messageLanguage !== 'undefined',
            ccRecipients: typeof args.options.ccRecipients !== 'undefined',
            sendInvitationMessage: !!args.options.sendInvitationMessage
        });
    });
}, _EntraUserGuestAddCommand_initOptions = function _EntraUserGuestAddCommand_initOptions() {
    this.options.unshift({
        option: '--emailAddress <emailAddress>'
    }, {
        option: '--displayName [displayName]'
    }, {
        option: '--inviteRedirectUrl [inviteRedirectUrl]'
    }, {
        option: '--welcomeMessage [welcomeMessage]'
    }, {
        option: '--messageLanguage [messageLanguage]'
    }, {
        option: '--ccRecipients [ccRecipients]'
    }, {
        option: '--sendInvitationMessage'
    });
};
export default new EntraUserGuestAddCommand();
//# sourceMappingURL=user-guest-add.js.map