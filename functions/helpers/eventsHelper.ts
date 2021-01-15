import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { Event, Attendee } from "@microsoft/microsoft-graph-types/microsoft-graph";
import { userHelper } from "../helpers/userHelper";
import { scoreName } from "../model/constants";
import {defaultClient, setup } from 'applicationinsights';

export class eventsHelper {

    public static async handleEventNotification(resourceText: string, userID: string) {
        const client = await GraphClient();
        defaultClient.trackTrace({message:'Processing event notification with resource ' + resourceText.toString(),severity:4});
        // TODO: Check type is created or updated

        const currentEvent: Event = await client.api(resourceText).get();
        const isValidAgenda: boolean = await this.checkAgendaExists(currentEvent);
        const currentAgendaScore = await userHelper.getUserScore(userID, scoreName.agenda);
        defaultClient.trackTrace({message:'Current agenda score is ' + currentAgendaScore,severity:3});
        await userHelper.setUserScoreInRange(isValidAgenda, scoreName.agenda, currentAgendaScore, 0, 30, 1, userID);

        const attendeesAlreadyBooked: boolean = await this.checkAttendeesNotAlreadyBooked(currentEvent);
        if (attendeesAlreadyBooked != null) {
            const currentAttendeeBookingsScore = await userHelper.getUserScore(userID, scoreName.attendeeBookings);
            await userHelper.setUserScoreInRange(attendeesAlreadyBooked, scoreName.attendeeBookings, currentAttendeeBookingsScore, 0,30,1,userID);
        }
    }

    public static checkAgendaExists(selectedEvent: Event): boolean {
        // TODO: identify further details
        defaultClient.trackTrace({message:'Checking agenda',severity:3});
        defaultClient.trackTrace({message:selectedEvent.bodyPreview.toLowerCase(),severity:3});
        if (selectedEvent.bodyPreview.toLowerCase().indexOf('agenda') > -1) {
            defaultClient.trackTrace({message:'Agenda found in event',severity:3});
            return true;
        }
        defaultClient.trackTrace({message:'Agenda not found in event',severity:3});
        return false;
    }

    public static async checkAttendeesNotAlreadyBooked(selectedEvent: Event): Promise<boolean> {
        const client = await GraphClient();
        let attendeesAlreadyBooked = false;
        if (selectedEvent.attendees !== null) {
            defaultClient.trackTrace({message:'Attendee count:'+selectedEvent.attendees.length.toString(),severity:3});
            defaultClient.trackTrace({message:'Attendees found: ' + selectedEvent.attendees.keys.length,severity:3});
            if (selectedEvent.attendees.keys.length === 0) {
                return null;
            }
            selectedEvent.attendees.forEach(async (attendee: Attendee)=> {
                // get events for the attendee at this time
                defaultClient.trackTrace({message:'Retrieving events for attendee ' + attendee.emailAddress.address,severity:3});
                const attendeeEventAPI = 'users/' + attendee.emailAddress.address + '/calendar/calendarView?startDateTime=' + selectedEvent.start.dateTime + '&endDateTime=' + selectedEvent.end.dateTime;
                defaultClient.trackTrace({message:'Retrieval API: ' + attendeeEventAPI,severity:3});
                await client.api(attendeeEventAPI).get()
                .then((attendeeEvents)=>{
                    defaultClient.trackTrace({message:'Event count: ' + attendeeEvents.length.toString(),severity:3});
                    attendeeEvents.forEach(async (attendeeEvent: Event) => {
                        if (attendeeEvent.id !== selectedEvent.id) {
                            // eventCount++;
                            attendeesAlreadyBooked = true;
                        }
                    });
                })
                .catch((err)=>{
                    defaultClient.trackTrace({message:'Failed to retrieve events for ' + attendee.emailAddress.address,severity:3});
                    defaultClient.trackTrace({message:err,severity:3});
                    attendeesAlreadyBooked = null;
                });
                // TODO: determine if we want to check the number and change the score if more overlapping events
                // let eventCount = 0;
                
            });
        }
        else {
            return null;
        }
        return attendeesAlreadyBooked;
    }

    public static async listUserOrganisedEvents(userID: string): Promise<Event[]> {
        const client = await GraphClient();

        return client
            .api('users/'+userID+'/events?$filter=isOrganizer eq true')
            .get()
            .then((res) => {
                defaultClient.trackTrace({message:'Success',severity:3});
                //defaultClient.trackTrace({message:res);
                return res;
            })
            .catch((err) => {
                defaultClient.trackTrace({message:'Failed',severity:3});
                defaultClient.trackTrace({message:err,severity:3});
                throw err;
            });
    }
}
