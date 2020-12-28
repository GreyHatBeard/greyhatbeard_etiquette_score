import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { eventsHelper } from "../helpers/eventsHelper"
import { Event } from "@microsoft/microsoft-graph-types/microsoft-graph";
import { userHelper } from "../helpers/userHelper";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Starting Azure function!");

    const statToUpdate = req.query['stat'];
    context.log('Updating stats with stat ${statToUpdate}');

    let returnText = 'Updated: ';
    // TODO: handle multiple things to update
    switch (statToUpdate){
        case "agenda":
            await UpdateAgenda(context);
            break;
    }
    
    context.res = {
        status: 200,
        body: returnText
    };
};

async function UpdateAgenda(context: Context) {
    const client = await GraphClient();
    const userID = '48e8a1ab-0d3a-4f9b-b200-e9e9d1437a2b';
    context.log("Updating agenda for ${userID}");
    var organisedEvents = await eventsHelper.listUserOrganisedEvents(userID, context);
    const totalEventCount = organisedEvents.value.length;
    let eventsWithAgenda = 0;
    context.log(organisedEvents);
    organisedEvents.value.forEach((selectedEvent)=> {
        if (eventsHelper.checkAgendaExists(selectedEvent, context)) {
            eventsWithAgenda++;
        }
    });

    const agendaScore = (eventsWithAgenda/totalEventCount)*10;
    await userHelper.updateScore('agendaScore', userID, agendaScore, context);
}


export default httpTrigger;