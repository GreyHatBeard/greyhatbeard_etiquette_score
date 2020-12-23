import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types"
import { handleEventNotification } from "../helpers/eventsHelper"

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
        const response: boolean = await handleEventNotification(req.body, context);
        if (response) {
            context.log('Event agenda valid');
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