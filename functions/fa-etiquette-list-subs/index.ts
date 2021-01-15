import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types"
import { subscriptionHelper } from '../helpers/subscriptionHelper'
import {defaultClient, setup } from 'applicationinsights';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Starting Azure function!");
    setup("").start();
    context.log(req.query)
    
    if (req.query['id'] !== null) {
        context.log('ID exists: ' + req.query['id']);
        await subscriptionHelper.deleteSubscription(req.query['id']);
    }
    const subs = await subscriptionHelper.listSubscriptions();
    context.res = {
        status: 200,
        body: subs
    };
};

export default httpTrigger;