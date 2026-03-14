var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OutlookRoomListCommand_instances, _OutlookRoomListCommand_initTelemetry, _OutlookRoomListCommand_initOptions;
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class OutlookRoomListCommand extends GraphCommand {
    get name() {
        return commands.ROOM_LIST;
    }
    get description() {
        return 'Get a collection of all available rooms';
    }
    defaultProperties() {
        return ['id', 'displayName', 'phone', 'emailAddress'];
    }
    constructor() {
        super();
        _OutlookRoomListCommand_instances.add(this);
        __classPrivateFieldGet(this, _OutlookRoomListCommand_instances, "m", _OutlookRoomListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OutlookRoomListCommand_instances, "m", _OutlookRoomListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        let endpoint = `${this.resource}/v1.0/places/microsoft.graph.room`;
        if (args.options.roomlistEmail) {
            endpoint = `${this.resource}/v1.0/places/${args.options.roomlistEmail}/microsoft.graph.roomlist/rooms`;
        }
        try {
            const rooms = await odata.getAllItems(endpoint);
            await logger.log(rooms);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_OutlookRoomListCommand_instances = new WeakSet(), _OutlookRoomListCommand_initTelemetry = function _OutlookRoomListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            roomlistEmail: typeof args.options.roomlistEmail !== 'undefined'
        });
    });
}, _OutlookRoomListCommand_initOptions = function _OutlookRoomListCommand_initOptions() {
    this.options.unshift({
        option: '--roomlistEmail [roomlistEmail]'
    });
};
export default new OutlookRoomListCommand();
//# sourceMappingURL=room-list.js.map