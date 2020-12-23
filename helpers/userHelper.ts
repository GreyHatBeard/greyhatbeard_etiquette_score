import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types/microsoft-graph";

class userHelper {

    public async initialiseUserExtension(context: Context) {
        const client = await GraphClient();

        return client
            .api('me/extensions')
            .post(
                {
                    "@odata.type":"microsoft.graph.openTypeExtension",
                    "extensionName":"com.greyHatBeard.score",
                    "currentScore":"500"
                }
            )
            .then((res) => {
                context.log('Extensions set');
                //context.log(res);
            })
            .catch((err) => {
                context.log('Failed');
                context.log(err);
                throw err;
            });
    }

    private checkAgendaExists(selectedEvent: Event, context: Context): boolean {
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
}
module.exports = new eventsHelper();