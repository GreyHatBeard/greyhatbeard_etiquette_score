import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types"
import { eventsHelper } from "../helpers/eventsHelper"
import { userHelper } from "../helpers/userHelper"

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Handling subscription request");
    // context.log(req);
    if (req.query != null && req.query['validationToken'] != null) {
        context.log("Validation token sent: " + req.query['validationToken']);
        
        context.res = {
            status: 200,
            body: req.query['validationToken']
        };
        context.log('response returned');
    }
    else {
        context.log('Notification received');
        const response: boolean = await eventsHelper.handleEventNotification(req.body, context);
        if (response) {
            context.log('Event agenda valid');
            let currentUser = req.body.value[0].resource;
            context.log('Current user value 1: ' + currentUser);
            const startTextPoint = currentUser.indexOf('Users/') + 6;
            const endTextPoint = currentUser.indexOf('/Events');
            currentUser = currentUser.substr(startTextPoint, endTextPoint-startTextPoint);
            context.log('Current user value 2: ' + currentUser);
            await userHelper.increaseUserScore(currentUser, 5, context);
        } else {
            context.log('Event agenda not set');
        }
        context.res = {
            status: 202,
            body: response
        };
    }

    context.log('Finished');
};

export default httpTrigger;