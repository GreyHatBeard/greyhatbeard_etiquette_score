import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types/microsoft-graph";
import { scoreName } from '../model/constants';
import {defaultClient, setup } from 'applicationinsights';

export class userHelper {

    public static async initialiseUserExtension(userId: string) {
        
        defaultClient.trackTrace({message:'Initialising user extension',severity:4});
        await this.initialiseScoreExtension(userId, scoreName.current,500);
        await this.initialiseScoreExtension(userId, scoreName.agenda, 15);
        await this.initialiseScoreExtension(userId, scoreName.attendeeBookings, 15);
    }

    private static async initialiseScoreExtension(userId: string, scoreName: string, score: number) {
        const client = await GraphClient();
        await client
        .api('users/' + userId + '/extensions/com.greyhatbeard.etiquettescores.' + scoreName)
        .get()
        .then(async (res) => {
            defaultClient.trackTrace({message:'Extension value exists so resetting user value',severity:4});
            await client
                .api('users/' + userId + '/extensions/com.greyhatbeard.etiquettescores.' + scoreName)
                .patch(
                    {
                        "@odata.type":"microsoft.graph.openTypeExtension",
                        "extensionName":"com.greyhatbeard.etiquettescores." + scoreName,
                        "score":score
                    }
                )
                .then((res) => {
                    defaultClient.trackTrace({message:'Current score extensions set for ' + scoreName, severity:3});
                    //defaultClient.trackTrace({message:res);
                })
                .catch((err) => {
                    defaultClient.trackTrace({message:'Error patching extension for score ' + scoreName, severity:3});
                });
        })
        .catch(async (err) => {
            defaultClient.trackTrace({message:'Extension does not exist for ' + scoreName + ' so creating',severity:4});
            client
                .api('users/' + userId + '/extensions')
                .post(
                    {
                        "@odata.type":"microsoft.graph.openTypeExtension",
                        "extensionName":"com.greyhatbeard.etiquettescores." + scoreName,
                        "score":score
                    }
                )
                .then((res) => {
                    defaultClient.trackTrace({message:'Extensions created for ' + scoreName, severity:3});
                })
                .catch((err) => {
                    defaultClient.trackTrace({message:'Failed to create extension for ' + scoreName + ': ' + err, severity:3});
                });
        });
    }

    public static async updateScore(scoreName: string, userId: string, updatedScore: number) {
        const client = await GraphClient();
        defaultClient.trackTrace({message:'Updating user score ' + scoreName, severity:3});
        // TODO: check if exists already

        return client
            .api('users/' + userId + '/extensions/com.greyhatbeard.etiquettescores.' + scoreName)
            .patch(
                {
                    "@odata.type":"microsoft.graph.openTypeExtension",
                    "extensionName":"com.greyhatbeard.etiquettescores." + scoreName,
                    "score":updatedScore
                }
            )
            .then((res) => {
                defaultClient.trackTrace({message:'Score updated for ' + scoreName, severity:3});
                //defaultClient.trackTrace({message:res);
            })
            .catch((err) => {
                defaultClient.trackTrace({message:'Failed to update score ' + scoreName, severity:3});
                defaultClient.trackTrace({message:err, severity:3});
                throw err;
            });
    }

    public static async getUserScore(userId: string, scoreName: string): Promise<number> {
        const client = await GraphClient();
        defaultClient.trackTrace({message:'Retrieving user score for ' + scoreName, severity:3});
        return client
            .api('users/' + userId + '/extensions/com.greyhatbeard.etiquettescores.' + scoreName)
            .get()
            .then((res) => {
                defaultClient.trackTrace({message:'Extensions set',severity:4});
                defaultClient.trackTrace({message:res, severity:3});
                const currentScore: number = res.score;
                return currentScore;
            })
            .catch((err) => {
                defaultClient.trackTrace({message:'Failed to load score ' + scoreName, severity:3});
                defaultClient.trackTrace({message:err, severity:3});
                throw err;
            });
    }

    public static async setUserScoreInRange(isValid: boolean, scoreName: string, 
            currentScore: number, lowRangeScore: number, highRangeScore: number, 
            incrementValue: number, currentUser: string) {

        defaultClient.trackTrace({message:'Setting user score to for ' + scoreName, severity:3});
        if (isValid) {
            defaultClient.trackTrace({message:'Increment score ' + scoreName, severity:3});
            
            if (currentScore >= highRangeScore) {
                await userHelper.updateScore(scoreName,currentUser, highRangeScore);
            } else {
                await userHelper.updateScore(scoreName,currentUser, currentScore+incrementValue);
            }

        } else {
            defaultClient.trackTrace({message:'Decrement score ' + scoreName, severity:3});

            if (currentScore <= lowRangeScore) {
                await userHelper.updateScore(scoreName,currentUser, lowRangeScore);
            } else {
                await userHelper.updateScore(scoreName,currentUser, currentScore-incrementValue);
            }
        }
    }
}
