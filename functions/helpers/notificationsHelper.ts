import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { eventsHelper } from "../helpers/eventsHelper";
import { userHelper } from "../helpers/userHelper";

/*
-----------------------------------------------------------------------------------
notificationsHelper builds out the logic for checking each notification, including the 
calculations for the score/

-----------------------------------------------------------------------------------
*/
export class notificationsHelper {

    public static async handleNotification(notification: any, context: Context) {
        const client = await GraphClient();
        // context.log(notification);
        switch (notification.value[0].clientState.toLowerCase()) {
            case "userevent":
                await this.handleUserEventNotification(notification, context);
                break;
        }
    }

    private static async handleUserEventNotification(notification: any, context: Context) {
        context.log('Processing UserEvent notification');
        let currentUser = await this.getUserIdFromResource(notification.value[0].resource, 'Events');
        await eventsHelper.handleEventNotification(notification.value[0].resource, currentUser, context);
    }

    private static async getUserIdFromResource(resourceText: string, urlStub: string): Promise<string> {
        const startTextPoint = resourceText.indexOf('Users/') + 6;
        const endTextPoint = resourceText.indexOf('/'+urlStub);
        resourceText = resourceText.substr(startTextPoint, endTextPoint-startTextPoint);
        return resourceText;
    }
}
