import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { GraphClient } from "../authHelpers";
import { User } from "@microsoft/microsoft-graph-types/microsoft-graph";

export class userHelper {

    public static async initialiseUserExtension(userId: string, context: Context) {
        const client = await GraphClient();
        context.log('Initialising user extension');
        return client
            .api('users/' + userId + '/extensions')
            .post(
                {
                    "@odata.type":"microsoft.graph.openTypeExtension",
                    "extensionName":"com.greyHatBeard.score",
                    "currentScore":500
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

    public static async updateUserScore(userId: string, updatedScore: number, context: Context) {
        await this.updateScore('currentScore', userId, updatedScore, context);
    }

    public static async updateScore(scoreName: string, userId: string, updatedScore: number, context: Context) {
        const client = await GraphClient();
        context.log('Updating user score');
        return client
            .api('users/' + userId + '/extensions/com.greyHatBeard.score')
            .patch(
                {
                    "@odata.type":"microsoft.graph.openTypeExtension",
                    "extensionName":"com.greyHatBeard.score",
                    [scoreName]:updatedScore
                }
            )
            .then((res) => {
                context.log('Score updated');
                //context.log(res);
            })
            .catch((err) => {
                context.log('Failed');
                context.log(err);
                throw err;
            });
    }

    public static async getUserScore(userId: string, context: Context): Promise<number> {
        const client = await GraphClient();
        context.log('Retrieving user score');
        return client
            .api('users/' + userId + '/extensions')
            .get()
            .then((res) => {
                context.log('Extensions set');
                context.log(res);
                const currentScore: number = res.value[0].currentScore;
                return currentScore;
            })
            .catch((err) => {
                context.log('Failed');
                context.log(err);
                throw err;
            });
    }

    public static async increaseUserScore(userId: string, increment: number, context: Context) {
        context.log('Increasing user score');
        const client = await GraphClient();
        const currentScore: number = await this.getUserScore(userId, context);
        context.log('Current score is ' + currentScore);
        const newScore:number = +increment + +currentScore;
        context.log('New score is ' + newScore);
        context.log('Updating user score');
        await this.updateUserScore(userId, newScore, context);
        context.log('Updated user score');
    }
}
