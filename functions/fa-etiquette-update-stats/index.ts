import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { eventsHelper } from "../helpers/eventsHelper"
import { Event } from "@microsoft/microsoft-graph-types/microsoft-graph";
import { userHelper } from "../helpers/userHelper";
import {defaultClient, setup } from 'applicationinsights';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    if (context) context.log("Starting Azure function!");
    setup("").start();
    const statToUpdate = req.query['stat'];
    context.log('Updating stats with stat ${statToUpdate}');

    let returnText = 'Updated: ';
    // TODO: handle multiple things to update
    switch (statToUpdate){
        case "agenda":
            await UpdateAgenda();
            break;
    }
    
    context.res = {
        status: 200,
        body: returnText
    };
};

async function UpdateAgenda() {
    const client = await GraphClient();
    const userID = '48e8a1ab-0d3a-4f9b-b200-e9e9d1437a2b';
    defaultClient.trackTrace({message:'Updating agenda for ${userID}', severity:3});
    var organisedEvents = await eventsHelper.listUserOrganisedEvents(userID);
    const totalEventCount = organisedEvents.value.length;
    let eventsWithAgenda = 0;
    
    organisedEvents.value.forEach((selectedEvent)=> {
        if (eventsHelper.checkAgendaExists(selectedEvent)) {
            eventsWithAgenda++;
        }
    });

    const agendaScore = (eventsWithAgenda/totalEventCount)*10;
    await userHelper.updateScore('agendaScore', userID, agendaScore);
}


export default httpTrigger;