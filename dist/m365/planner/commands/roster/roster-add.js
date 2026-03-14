import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerRosterAddCommand extends GraphCommand {
    get name() {
        return commands.ROSTER_ADD;
    }
    get description() {
        return 'Creates a new Microsoft Planner Roster';
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr('Creating a new Microsoft Planner Roster');
        }
        try {
            const requestOptions = {
                url: `${this.resource}/beta/planner/rosters`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: {},
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PlannerRosterAddCommand();
//# sourceMappingURL=roster-add.js.map