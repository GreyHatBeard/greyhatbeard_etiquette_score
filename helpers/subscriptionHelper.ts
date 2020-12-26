import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types/microsoft-graph";
import {NotificationReceiverUrl } from '../secrets'

export class subscriptionHelper {

    public static async addSubscription(userID: string, context: Context) {
        const client = await GraphClient();
        context.log("Adding subscription with url " + NotificationReceiverUrl);
        const currentDate: Date = new Date();
        const subscriptionDate = currentDate.setDate(currentDate.getDate() + 1);

        return client
            .api("/subscriptions")
            .post(
                {
                    "changeType": "created",
                    "expirationDateTime": subscriptionDate,
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