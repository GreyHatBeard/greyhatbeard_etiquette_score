import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types"
import { subscriptionHelper } from '../helpers/subscriptionHelper'

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Starting Azure function!");

    context.log(req.query)
    
    if (req.query['id'] !== null) {
        context.log('ID exists: ' + req.query['id']);
        await subscriptionHelper.deleteSubscription(req.query['id'], context);
    }
    const subs = await subscriptionHelper.listSubscriptions(context);
    context.res = {
        status: 200,
        body: subs
    };
};

export default httpTrigger;