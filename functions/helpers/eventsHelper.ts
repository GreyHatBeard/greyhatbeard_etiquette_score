import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { Event } from "@microsoft/microsoft-graph-types/microsoft-graph";

export class eventsHelper {

    public static async handleEventNotification(resource: string, context: Context): Promise<boolean> {
        const client = await GraphClient();

        // TODO: Check type is created or updated

        return client
            .api(resource)
            .get()
            .then((res) => {
                context.log('Success');
                //context.log(res);
                return this.checkAgendaExists(res, context);
            })
            .catch((err) => {
                context.log('Failed');
                context.log(err);
                throw err;
            });
    }

    public static checkAgendaExists(selectedEvent: Event, context: Context): boolean {
        // TODO: identify further details
        context.log('Checking agenda');
        context.log(selectedEvent.bodyPreview.toLowerCase());
        if (selectedEvent.bodyPreview.toLowerCase().indexOf('agenda') > -1) {
            context.log('Agenda found in event');
            return true;
        }
        context.log('Agenda not found in event');
        return false;
    }

    public static async listUserOrganisedEvents(userID: string, context: Context): Promise<Event[]> {
        const client = await GraphClient();

        return client
            .api('users/'+userID+'/events?$filter=isOrganizer eq true')
            .get()
            .then((res) => {
                context.log('Success');
                //context.log(res);
                return res;
            })
            .catch((err) => {
                context.log('Failed');
                context.log(err);
                throw err;
            });
    }
}
