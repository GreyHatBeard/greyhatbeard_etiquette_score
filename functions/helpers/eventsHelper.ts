import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { Event, Attendee } from "@microsoft/microsoft-graph-types/microsoft-graph";
import { userHelper } from "../helpers/userHelper";

export class eventsHelper {

    public static async handleEventNotification(resourceText: string, userID: string, currentScore: number, context: Context) {
        const client = await GraphClient();
        context.log('Processing event notification with resource ' + resourceText.toString());
        // TODO: Check type is created or updated

        const currentEvent: Event = await client.api(resourceText).get();
        const isValidAgenda: boolean = this.checkAgendaExists(currentEvent, context);
        await userHelper.setUserScoreInRange(isValidAgenda, 'agendaScore', currentScore, 0, 30, 1, userID, context);

        const attendeesAlreadyBooked: boolean = await this.checkAttendeesNotAlreadyBooked(currentEvent, context);
        await userHelper.setUserScoreInRange(attendeesAlreadyBooked, 'attendeesBookedScore', currentScore, 0,30,1,userID, context);
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

    public static async checkAttendeesNotAlreadyBooked(selectedEvent: Event, context: Context): Promise<boolean> {
        const client = await GraphClient();
        let attendeesAlreadyBooked = false;
        if (selectedEvent.attendees !== null) {
            selectedEvent.attendees.forEach(async (attendee: Attendee)=> {
                // get events for the attendee at this time
                const attendeeEvents = await client.api('users/' + attendee.emailAddress.address + '/calendar/calendarView?startDateTime=' + selectedEvent.start.dateTime + '&endDateTime=' + selectedEvent.end.dateTime).get();
                // TODO: determine if we want to check the number and change the score if more overlapping events
                // let eventCount = 0;
                attendeeEvents.forEach(async (attendeeEvent: Event) => {
                    if (attendeeEvent.id !== selectedEvent.id) {
                        // eventCount++;
                        attendeesAlreadyBooked = true;
                    }
                });
            });
        }
        else {
            return true;
        }
        return attendeesAlreadyBooked;
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
