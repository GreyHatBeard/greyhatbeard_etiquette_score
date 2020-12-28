import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types"
import { userHelper } from '../helpers/userHelper';
import { subscriptionHelper } from '../helpers/subscriptionHelper';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Starting Azure function!");

    // TODO: Get user ID from querystring
    
    await subscriptionHelper.addSubscription('48e8a1ab-0d3a-4f9b-b200-e9e9d1437a2b', context);
    context.log("About to initalise user");
    await userHelper.initialiseUserExtension('48e8a1ab-0d3a-4f9b-b200-e9e9d1437a2b', context);
    context.log("Subscription created");
    context.res = {
        status: 200,
        body: {
            result: 'success'
        }
    };
};




export default httpTrigger;