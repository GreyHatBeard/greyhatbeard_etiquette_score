import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types"
import {NotificationReceiverUrl } from '../secrets'

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Starting Azure function!");

    // TODO: Get user ID from querystring

    await addSubscription('48e8a1ab-0d3a-4f9b-b200-e9e9d1437a2b', context);
    context.log("Subscription created");
    context.res = {
        status: 200,
        body: {
            result: 'success'
        }
    };
};

async function addSubscription(userID: string, context: Context) {
    const client = await GraphClient();
    context.log("Adding subscription with url " + NotificationReceiverUrl);

    return client
        .api("/subscriptions")
        .post(
            {
                "changeType": "created",
                "expirationDateTime": "2020-12-24T11:23:45.9356913Z",
                "notificationUrl": NotificationReceiverUrl,
                "resource": "users/" + userID + "/events"
            }
        )
        .then((res) => {
            context.log('Result received');
            context.log(res);
            return;
        })
        .catch((err) => {
            context.log(err);
        });;
}


export default httpTrigger;