import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types/microsoft-graph";
import {NotificationReceiverUrl } from '../secrets'

export class subscriptionHelper {

    public static async addSubscription(userID: string, context: Context) {
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

}