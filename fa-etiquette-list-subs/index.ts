import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types"
import {NotificationReceiverUrl } from '../secrets'

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Starting Azure function!");

    const subs = await listSubscriptions(context);
    context.res = {
        status: 200,
        body: subs
    };
};

async function listSubscriptions(context: Context) {
    const client = await GraphClient();
    context.log("Adding subscription with url " + NotificationReceiverUrl);

    return client
        .api("/subscriptions")
        .get()
        .then((res) => {
            context.log('Subscriptions received');
            context.log(res);
            return(res);
        })
        .catch((err) => {
            context.log(err);
        });;
}


export default httpTrigger;