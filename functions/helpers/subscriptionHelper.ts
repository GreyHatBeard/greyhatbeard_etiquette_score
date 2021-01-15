import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { Subscription } from "@microsoft/microsoft-graph-types/microsoft-graph";
import {defaultClient, setup } from 'applicationinsights';

export class subscriptionHelper {

    public static async addEventSubscription(userID: string) {
        await this.addSubscription("users/" + userID + "/events", "UserEvent");
    }

    public static async addSubscription(resource: string, subscriptionType: string) {
        const client = await GraphClient();
        defaultClient.trackTrace({message:"Adding subscription with url " + process.env['NotificationReceiverUrl'], severity:3});
        const subscriptionDate: Date = new Date();
        subscriptionDate.setDate(subscriptionDate.getDate() + 1);
        defaultClient.trackTrace({message:"Creating subscription date with expiry " + subscriptionDate.toISOString(), severity:3});
        return client
            .api("/subscriptions")
            .post(
                {
                    "changeType": "created",
                    "clientState": subscriptionType,
                    "expirationDateTime": subscriptionDate.toISOString(),
                    "notificationUrl": process.env['NotificationReceiverUrl'],
                    "resource": resource
                }
            )
            .then((res) => {
                defaultClient.trackTrace({message:'Result received', severity:3});
                // defaultClient.trackTrace({message:res);
                return;
            })
            .catch((err) => {
                defaultClient.trackTrace({message:err, severity:3});
            });;
    }

    public static async listSubscriptions() { // : Promise<Subscription[]>
        const client = await GraphClient();
        defaultClient.trackTrace({message:"Listing subscriptions", severity:3});
    
        return client.api("/subscriptions").get();
    }

    public static async deleteSubscription(subscriptionId: string) {
        const client = await GraphClient();
        defaultClient.trackTrace({message:"Deleting subscription with id " + subscriptionId, severity:3});
    
        await client.api("/subscriptions/" + subscriptionId).delete()
        .catch((err) => {
            defaultClient.trackTrace({message:'Damn, cannot delete: ' + err, severity:3});
        });
    }
}