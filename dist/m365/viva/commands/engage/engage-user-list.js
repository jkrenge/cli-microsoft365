var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageUserListCommand_instances, _VivaEngageUserListCommand_initTelemetry, _VivaEngageUserListCommand_initOptions, _VivaEngageUserListCommand_initValidators;
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageUserListCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_USER_LIST;
    }
    get description() {
        return 'Returns users from the current network';
    }
    defaultProperties() {
        return ['id', 'full_name', 'email'];
    }
    constructor() {
        super();
        _VivaEngageUserListCommand_instances.add(this);
        this.items = [];
        __classPrivateFieldGet(this, _VivaEngageUserListCommand_instances, "m", _VivaEngageUserListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageUserListCommand_instances, "m", _VivaEngageUserListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageUserListCommand_instances, "m", _VivaEngageUserListCommand_initValidators).call(this);
    }
    getAllItems(logger, args, page) {
        return new Promise((resolve, reject) => {
            if (page === 1) {
                this.items = [];
            }
            let endPoint = `${this.resource}/v1/users.json`;
            if (args.options.groupId !== undefined) {
                endPoint = `${this.resource}/v1/users/in_group/${args.options.groupId}.json`;
            }
            endPoint += `?page=${page}`;
            if (args.options.reverse !== undefined) {
                endPoint += `&reverse=true`;
            }
            if (args.options.sortBy !== undefined) {
                endPoint += `&sort_by=${args.options.sortBy}`;
            }
            if (args.options.letter !== undefined) {
                endPoint += `&letter=${args.options.letter}`;
            }
            const requestOptions = {
                url: endPoint,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            request
                .get(requestOptions)
                .then((res) => {
                let userOutput = res;
                // groups user retrieval returns a user array containing the user objects
                if (res.users) {
                    userOutput = res.users;
                }
                this.items = this.items.concat(userOutput);
                // this is executed once at the end if the limit operation has been executed
                // we need to return the array of the desired size. The API does not provide such a feature
                if (args.options.limit !== undefined && this.items.length > args.options.limit) {
                    this.items = this.items.slice(0, args.options.limit);
                    resolve();
                }
                else {
                    // if the groups endpoint is used, the more_available will tell if a new retrieval is required
                    // if the user endpoint is used, we need to page by 50 items (hardcoded)
                    if (res.more_available === true || this.items.length % 50 === 0) {
                        this.getAllItems(logger, args, ++page)
                            .then(() => {
                            resolve();
                        }, (err) => {
                            reject(err);
                        });
                    }
                    else {
                        resolve();
                    }
                }
            }, (err) => {
                reject(err);
            });
        });
    }
    async commandAction(logger, args) {
        this.items = []; // this will reset the items array in interactive mode
        try {
            await this.getAllItems(logger, args, 1);
            await logger.log(this.items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_VivaEngageUserListCommand_instances = new WeakSet(), _VivaEngageUserListCommand_initTelemetry = function _VivaEngageUserListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            letter: args.options.letter !== undefined,
            sortBy: args.options.sortBy !== undefined,
            reverse: args.options.reverse !== undefined,
            limit: args.options.limit !== undefined,
            groupId: args.options.groupId !== undefined
        });
    });
}, _VivaEngageUserListCommand_initOptions = function _VivaEngageUserListCommand_initOptions() {
    this.options.unshift({
        option: '-g, --groupId [groupId]'
    }, {
        option: '-l, --letter [letter]'
    }, {
        option: '--reverse'
    }, {
        option: '--limit [limit]'
    }, {
        option: '--sortBy [sortBy]',
        autocomplete: ['messages', 'followers']
    });
}, _VivaEngageUserListCommand_initValidators = function _VivaEngageUserListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && typeof args.options.groupId !== 'number') {
            return `${args.options.groupId} is not a number`;
        }
        if (args.options.limit && typeof args.options.limit !== 'number') {
            return `${args.options.limit} is not a number`;
        }
        if (args.options.sortBy && args.options.sortBy !== 'messages' && args.options.sortBy !== 'followers') {
            return `sortBy accepts only the values "messages" or "followers"`;
        }
        if (args.options.letter && !/^(?!\d)[a-zA-Z]+$/i.test(args.options.letter)) {
            return `Value of 'letter' is invalid. Only characters within the ranges [A - Z], [a - z] are allowed.`;
        }
        if (args.options.letter && args.options.letter.length !== 1) {
            return `Only one char as value of 'letter' accepted.`;
        }
        return true;
    });
};
export default new VivaEngageUserListCommand();
//# sourceMappingURL=engage-user-list.js.map