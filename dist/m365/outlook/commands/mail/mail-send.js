var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OutlookMailSendCommand_instances, _OutlookMailSendCommand_initTelemetry, _OutlookMailSendCommand_initOptions, _OutlookMailSendCommand_initTypes, _OutlookMailSendCommand_initValidators;
import fs from 'fs';
import path from 'path';
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class OutlookMailSendCommand extends GraphCommand {
    get name() {
        return commands.MAIL_SEND;
    }
    get description() {
        return 'Sends an email';
    }
    constructor() {
        super();
        _OutlookMailSendCommand_instances.add(this);
        __classPrivateFieldGet(this, _OutlookMailSendCommand_instances, "m", _OutlookMailSendCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OutlookMailSendCommand_instances, "m", _OutlookMailSendCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _OutlookMailSendCommand_instances, "m", _OutlookMailSendCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _OutlookMailSendCommand_instances, "m", _OutlookMailSendCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
            if (isAppOnlyAccessToken === true && !args.options.sender) {
                throw `Specify a upn or user id in the 'sender' option when using app only authentication.`;
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/${args.options.sender ? 'users/' + formatting.encodeQueryParameter(args.options.sender) : 'me'}/sendMail`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: this.getRequestBody(args.options)
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapEmailAddressToRecipient(email) {
        if (!email) {
            return undefined;
        }
        return {
            emailAddress: {
                address: email.trim()
            }
        };
    }
    getRequestBody(options) {
        const attachments = typeof options.attachment === 'string' ? [options.attachment] : options.attachment;
        const attachmentContents = attachments?.map(a => ({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: path.basename(a),
            contentBytes: fs.readFileSync(a, { encoding: 'base64' })
        }));
        return ({
            message: {
                subject: options.subject,
                body: {
                    contentType: options.bodyContentType || 'Text',
                    content: options.bodyContents
                },
                from: this.mapEmailAddressToRecipient(options.mailbox),
                toRecipients: options.to.split(',').map(mail => this.mapEmailAddressToRecipient(mail)),
                ccRecipients: options.cc?.split(',').map(mail => this.mapEmailAddressToRecipient(mail)),
                bccRecipients: options.bcc?.split(',').map(mail => this.mapEmailAddressToRecipient(mail)),
                importance: options.importance,
                attachments: attachmentContents
            },
            saveToSentItems: options.saveToSentItems
        });
    }
}
_OutlookMailSendCommand_instances = new WeakSet(), _OutlookMailSendCommand_initTelemetry = function _OutlookMailSendCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cc: typeof args.options.cc !== 'undefined',
            bcc: typeof args.options.bcc !== 'undefined',
            bodyContentType: args.options.bodyContentType,
            saveToSentItems: args.options.saveToSentItems,
            importance: args.options.importance,
            mailbox: typeof args.options.mailbox !== 'undefined',
            sender: typeof args.options.sender !== 'undefined',
            attachment: typeof args.options.attachment !== 'undefined'
        });
    });
}, _OutlookMailSendCommand_initOptions = function _OutlookMailSendCommand_initOptions() {
    this.options.unshift({
        option: '-s, --subject <subject>'
    }, {
        option: '-t, --to <to>'
    }, {
        option: '--cc [cc]'
    }, {
        option: '--bcc [bcc]'
    }, {
        option: '--sender [sender]'
    }, {
        option: '-m, --mailbox [mailbox]'
    }, {
        option: '--bodyContents <bodyContents>'
    }, {
        option: '--bodyContentType [bodyContentType]',
        autocomplete: ['Text', 'HTML']
    }, {
        option: '--importance [importance]',
        autocomplete: ['low', 'normal', 'high']
    }, {
        option: '--attachment [attachment]'
    }, {
        option: '--saveToSentItems [saveToSentItems]',
        autocomplete: ['true', 'false']
    });
}, _OutlookMailSendCommand_initTypes = function _OutlookMailSendCommand_initTypes() {
    this.types.boolean.push('saveToSentItems');
}, _OutlookMailSendCommand_initValidators = function _OutlookMailSendCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.bodyContentType &&
            args.options.bodyContentType !== 'Text' &&
            args.options.bodyContentType !== 'HTML') {
            return `${args.options.bodyContentType} is not a valid value for the bodyContentType option. Allowed values are Text|HTML`;
        }
        if (args.options.importance && ['low', 'normal', 'high'].indexOf(args.options.importance) === -1) {
            return `'${args.options.importance}' is not a valid value for the importance option. Allowed values are low|normal|high`;
        }
        if (args.options.attachment) {
            const attachments = typeof args.options.attachment === 'string' ? [args.options.attachment] : args.options.attachment;
            for (const attachment of attachments) {
                if (!fs.existsSync(attachment)) {
                    return `File with path '${attachment}' was not found.`;
                }
                if (!fs.lstatSync(attachment).isFile()) {
                    return `'${attachment}' is not a file.`;
                }
            }
            const requestBody = this.getRequestBody(args.options);
            // The max body size of the request is 4 194 304 chars before getting a 413 response
            if (JSON.stringify(requestBody).length > 4194304) {
                return 'Exceeded the max total size of attachments which is 3MB.';
            }
        }
        return true;
    });
};
export default new OutlookMailSendCommand();
//# sourceMappingURL=mail-send.js.map