import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types/microsoft-graph";
import { scoreName } from '../model/constants';

export class userHelper {

    public static async initialiseUserExtension(userId: string, context: Context) {
        
        context.log('Initialising user extension');
        await this.initialiseScoreExtension(userId, scoreName.current,500, context);
        await this.initialiseScoreExtension(userId, scoreName.agenda, 15, context);
        await this.initialiseScoreExtension(userId, scoreName.attendeeBookings, 15, context);
    }

    private static async initialiseScoreExtension(userId: string, scoreName: string, score: number, context: Context) {
        const client = await GraphClient();
        await client
        .api('users/' + userId + '/extensions/com.greyhatbeard.etiquettescores.' + scoreName)
        .get()
        .then(async (res) => {
            context.log('Extension value exists so resetting user value');
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
                    context.log('Current score extensions set for ' + scoreName);
                    //context.log(res);
                })
                .catch((err) => {
                    context.log('Error patching extension for score ' + scoreName);
                });
        })
        .catch(async (err) => {
            context.log('Extension does not exist for ' + scoreName + ' so creating');
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
                    context.log('Extensions created for ' + scoreName);
                })
                .catch((err) => {
                    context.log('Failed to create extension for ' + scoreName + ': ' + err);
                });
        });
    }

    public static async updateScore(scoreName: string, userId: string, updatedScore: number, context: Context) {
        const client = await GraphClient();
        context.log('Updating user score ' + scoreName);
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
                context.log('Score updated for ' + scoreName);
                //context.log(res);
            })
            .catch((err) => {
                context.log('Failed to update score ' + scoreName);
                context.log(err);
                throw err;
            });
    }

    public static async getUserScore(userId: string, scoreName: string, context: Context): Promise<number> {
        const client = await GraphClient();
        context.log('Retrieving user score for ' + scoreName);
        return client
            .api('users/' + userId + '/extensions/com.greyhatbeard.etiquettescores.' + scoreName)
            .get()
            .then((res) => {
                context.log('Extensions set');
                context.log(res);
                const currentScore: number = res.score;
                return currentScore;
            })
            .catch((err) => {
                context.log('Failed to load score ' + scoreName);
                context.log(err);
                throw err;
            });
    }

    public static async setUserScoreInRange(isValid: boolean, scoreName: string, 
            currentScore: number, lowRangeScore: number, highRangeScore: number, 
            incrementValue: number, currentUser: string, context:Context) {

        context.log('Setting user score to for ' + scoreName);
        if (isValid) {
            context.log('Increment score ' + scoreName);
            
            if (currentScore >= highRangeScore) {
                await userHelper.updateScore(scoreName,currentUser, highRangeScore,context);
            } else {
                await userHelper.updateScore(scoreName,currentUser, currentScore+incrementValue,context);
            }

        } else {
            context.log('Decrement score ' + scoreName);

            if (currentScore <= lowRangeScore) {
                await userHelper.updateScore(scoreName,currentUser, lowRangeScore,context);
            } else {
                await userHelper.updateScore(scoreName,currentUser, currentScore-incrementValue,context);
            }
        }
    }
}
