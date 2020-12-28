import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types"

import { notificationsHelper } from "../helpers/notificationsHelper"

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Handling subscription request");
    // context.log(req);
    if (req.query != null && req.query['validationToken'] != null) {
        context.log("Validation token sent: " + req.query['validationToken']);
        await notificationsHelper.handleNotification(req.body, context);
        context.res = {
            status: 200,
            body: req.query['validationToken']
        };
        context.log('response returned');
    }
    else {
        context.log('Notification received');
        
        context.res = {
            status: 202,
            body: 'Score changed'
        };
    }

    context.log('Finished');
};

export default httpTrigger;