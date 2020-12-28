import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { eventsHelper } from "../helpers/eventsHelper";
import { userHelper } from "../helpers/userHelper"

export class notificationsHelper {

    public static async handleNotification(notification: any, context: Context) {
        const client = await GraphClient();
        // context.log(notification);
        switch (notification.value[0].clientState.toLowerCase()) {
            case "userevent":
                context.log('Processing UserEvent notification');
                const response: boolean = await eventsHelper.handleEventNotification(notification.value[0].resource, context);
        
                if (response) {
                    context.log('Event agenda valid');
                    let currentUser = notification.value[0].resource;
                    context.log('Current user value 1: ' + currentUser);
                    const startTextPoint = currentUser.indexOf('Users/') + 6;
                    const endTextPoint = currentUser.indexOf('/Events');
                    currentUser = currentUser.substr(startTextPoint, endTextPoint-startTextPoint);
                    context.log('Current user value 2: ' + currentUser);
                    await userHelper.increaseUserScore(currentUser, 5, context);
                } else {
                    context.log('Event agenda not set');
                }
                break;
        }
    }
}
