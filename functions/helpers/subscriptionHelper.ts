import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types/microsoft-graph";
import {NotificationReceiverUrl } from '../secrets'

export class subscriptionHelper {

    public static async addEventSubscription(userID: string, context: Context) {
        await this.addSubscription("users/" + userID + "/events", "UserEvent", context);
    }

    public static async addSubscription(resource: string, subscriptionType: string, context: Context) {
        const client = await GraphClient();
        context.log("Adding subscription with url " + NotificationReceiverUrl);
        const currentDate: Date = new Date();
        const subscriptionDate = currentDate.setDate(currentDate.getDate() + 1);

        return client
            .api("/subscriptions")
            .post(
                {
                    "changeType": "created",
                    "clientState": subscriptionType,
                    "expirationDateTime": subscriptionDate,
                    "notificationUrl": NotificationReceiverUrl,
                    "resource": resource
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