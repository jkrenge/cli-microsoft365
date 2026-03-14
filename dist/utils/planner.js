import { cli } from "../cli/cli.js";
import request from "../request.js";
import { formatting } from "./formatting.js";
import { odata } from "./odata.js";
const graphResource = 'https://graph.microsoft.com';
const getRequestOptions = (url, metadata) => ({
    url: url,
    headers: {
        accept: `application/json;odata.metadata=${metadata}`
    },
    responseType: 'json'
});
export const planner = {
    /**
     * Get Planner plan by ID.
     * @param id Planner ID.
     * @param metadata OData metadata level. Default is none.
     * @throws Error when the plan is not found.
     */
    async getPlanById(id, metadata = 'none') {
        const requestOptions = getRequestOptions(`${graphResource}/v1.0/planner/plans/${id}`, metadata);
        try {
            return await request.get(requestOptions);
        }
        catch {
            throw Error(`Planner plan with id '${id}' was not found.`);
        }
    },
    /**
     * Get all Planner plans for a specific group.
     * @param groupId Group ID.
     * @param metadata OData metadata level. Default is none.
     */
    getPlansByGroupId(groupId, metadata = 'none') {
        return odata.getAllItems(`${graphResource}/v1.0/groups/${groupId}/planner/plans`, metadata);
    },
    /**
     * Get the Planner plan for a specific Roster.
     * @param rosterId Roster ID.
     * @param metadata OData metadata level. Default is none.
     * @throws Error when the roster has no plan.
     */
    async getPlanByRosterId(rosterId, metadata = 'none') {
        const plans = await odata.getAllItems(`${graphResource}/beta/planner/rosters/${rosterId}/plans`, metadata);
        if (plans.length === 0) {
            throw Error(`The specified roster '${rosterId}' does not have a plan.`);
        }
        return plans[0];
    },
    /**
     * Get the Planner plan ID for a specific Roster.
     * @param rosterId Roster ID.
     * @throws Error when the roster has no plan.
     */
    async getPlanIdByRosterId(rosterId) {
        const plans = await odata.getAllItems(`${graphResource}/beta/planner/rosters/${rosterId}/plans?$select=id`);
        if (plans.length === 0) {
            throw Error(`The specified roster '${rosterId}' does not have a plan.`);
        }
        return plans[0].id;
    },
    /**
     * Get Planner plan by title in a specific group.
     * @param title Title of the Planner plan. Case insensitive.
     * @param groupId Owner group ID.
     * @param metadata OData metadata level. Default is none.
     * @throws Error when the plan is not found.
     */
    async getPlanByTitle(title, groupId, metadata = 'none') {
        const plans = await this.getPlansByGroupId(groupId, metadata);
        const filteredPlans = plans.filter(p => p.title && p.title.toLowerCase() === title.toLowerCase());
        if (filteredPlans.length === 0) {
            throw Error(`The specified plan '${title}' does not exist.`);
        }
        if (filteredPlans.length > 1) {
            const plansKeyValuePair = formatting.convertArrayToHashTable('id', filteredPlans);
            const plan = await cli.handleMultipleResultsFound(`Multiple plans with title '${title}' found.`, plansKeyValuePair);
            return plan;
        }
        return filteredPlans[0];
    },
    /**
     * Get Planner plan ID by title in a specific group.
     * @param title Title of the Planner plan. Case insensitive.
     * @param groupId Owner group ID.
     * @throws Error when the plan is not found.
     */
    async getPlanIdByTitle(title, groupId) {
        const plans = await odata.getAllItems(`${graphResource}/v1.0/groups/${groupId}/planner/plans?$select=id,title`);
        const filteredPlans = plans.filter(p => p.title && p.title.toLowerCase() === title.toLowerCase());
        if (filteredPlans.length === 0) {
            throw Error(`The specified plan '${title}' does not exist.`);
        }
        if (filteredPlans.length > 1) {
            const plansKeyValuePair = formatting.convertArrayToHashTable('id', filteredPlans);
            const plan = await cli.handleMultipleResultsFound(`Multiple plans with title '${title}' found.`, plansKeyValuePair);
            return plan.id;
        }
        return filteredPlans[0].id;
    },
    /**
     * Get Planner bucket by title in a specific plan.
     * @param title Title of the Planner bucket. Case insensitive.
     * @param planId ID of the plan that contains the bucket.
     * @param metadata OData metadata level. Default is none.
     * @throws Error when the bucket is not found.
     */
    async getBucketByTitle(title, planId, metadata = 'none') {
        const buckets = await odata.getAllItems(`${graphResource}/v1.0/planner/plans/${planId}/buckets`, metadata);
        const filteredBuckets = buckets.filter(b => b.name && b.name.toLowerCase() === title.toLowerCase());
        if (filteredBuckets.length === 0) {
            throw Error(`The specified bucket '${title}' does not exist.`);
        }
        if (filteredBuckets.length > 1) {
            const bucketsKeyValuePair = formatting.convertArrayToHashTable('id', filteredBuckets);
            const bucket = await cli.handleMultipleResultsFound(`Multiple buckets with name '${title}' found.`, bucketsKeyValuePair);
            return bucket;
        }
        return filteredBuckets[0];
    },
    /**
     * Get Planner bucket ID by title in a specific plan.
     * @param title Title of the Planner bucket. Case insensitive.
     * @param planId ID of the plan that contains the bucket.
     * @throws Error when the bucket is not found.
     */
    async getBucketIdByTitle(title, planId) {
        const buckets = await odata.getAllItems(`${graphResource}/v1.0/planner/plans/${planId}/buckets?$select=id,name`);
        const filteredBuckets = buckets.filter(b => b.name && b.name.toLowerCase() === title.toLowerCase());
        if (filteredBuckets.length === 0) {
            throw Error(`The specified bucket '${title}' does not exist.`);
        }
        if (filteredBuckets.length > 1) {
            const bucketsKeyValuePair = formatting.convertArrayToHashTable('id', filteredBuckets);
            const bucket = await cli.handleMultipleResultsFound(`Multiple buckets with name '${title}' found.`, bucketsKeyValuePair);
            return bucket.id;
        }
        return filteredBuckets[0].id;
    }
};
//# sourceMappingURL=planner.js.map