import { SPFI, spfi ,SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/folders/web";
import "@pnp/sp/files/folder";
import "@pnp/sp/items/list";
import "@pnp/sp/fields/list";
import "@pnp/sp/views/list";
import "@pnp/sp/site-users/web";
import { IList } from "@pnp/sp/lists";
import "@pnp/sp/items";
import { IUserInfo, IResponseDetails, IQuestionDetails } from "../models";
import { Fields, ListsUrl } from "./enumHelper";
import { CalendarType, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType } from "@pnp/sp/fields/types";

export default class SPHelper {
    private selectFields: string[] = ["ID", "Title", "QuestionID", "UserResponse"];
    private _list: IList = null;
    private lst_response: string = "";
    private questions_lst: string = "";

    private _sp:SPFI | undefined = undefined;

    public constructor(context: any) {
        this.lst_response = "QuickPoll";
        this.questions_lst = ListsUrl.Questions;
        this._sp = spfi().using(SPFx(context));

        this._list = this._sp.web.lists.getByTitle(this.lst_response);
    }
    /**
     * Get the current logged in user information
     */
    public getCurrentUserInfo = async (): Promise<IUserInfo> => {
        let userinfo: IUserInfo | null = null;
        const currentUserInfo = await this._sp.web.currentUser();
        userinfo = {
            ID: currentUserInfo.Id.toString(),
            Email: currentUserInfo.Email,
            LoginName: currentUserInfo.LoginName,
            DisplayName: currentUserInfo.Title,
            Picture: '/_layouts/15/userphoto.aspx?size=S&username=' + currentUserInfo.UserPrincipalName,
        };
        return userinfo;
    }
    /**
     * Get the poll response based on the question id.
     */
    public getPollResponse = async (questionId: string | undefined) => {

        const questionResponse = this._list && await this._list.items.select(this.selectFields.join(',')).filter(`QuestionID eq '${questionId}'`).expand('FieldValuesAsText')();
        if (questionResponse && questionResponse.length > 0) {
            var tmpResponse = questionResponse[0].UserResponse;
            // Decode the HTML entities
            const decodedString = tmpResponse.replace(/&#123;/g, '{')
            .replace(/&#125;/g, '}')
            .replace(/&quot;/g, '"')
            .replace(/&#58;/g, ':');
            if (decodedString != undefined && decodedString != null && decodedString !== "") {
                var jsonQResponse = JSON.parse(decodedString);
                return jsonQResponse;
            } else return [];
        } else return [];
    }
    /**
     * Add the user response.
     */
    public addPollResponse = async (userResponse: IResponseDetails, allUserResponse: any): Promise<any> => {
        let addedresponse = this._list && await this._list.items.add({
            Title: userResponse.PollQuestion,
            QuestionID: userResponse.PollQuestionId.toString(),
            UserResponse: JSON.stringify(allUserResponse)
        });
        return addedresponse;
    }
    /**
     * Update the over all response based on the end user response.
     */
    public updatePollResponse = async (questionId: string | undefined, allUserResponse: any) => {
        var response = this._list && await this._list.items.select(this.selectFields.join(','))
            .filter(`QuestionID eq '${questionId}'`).expand('FieldValuesAsText')();
        if (response.length > 0) {
            if (allUserResponse.length > 0) {
                let updatedResponse = this._list && await this._list.items.getById(response[0].ID).update({
                    UserResponse: JSON.stringify(allUserResponse)
                });
                return updatedResponse;
            } else return this._list && await this._list.items.getById(response[0].ID).delete();
        }
    }
    /**
     * Submit the user response.
     */
    public submitResponse = async (userResponse: IResponseDetails): Promise<boolean> => {
        try {
            let allUserResponse = await this.getPollResponse(userResponse.PollQuestionId);
            if (allUserResponse.length > 0) {
                // Remove the existing response for the given UserID and add the new response
                allUserResponse = allUserResponse.filter((response: any) => response.UserID !== userResponse.UserID);

               
                allUserResponse.push({
                    UserID: userResponse.UserID,
                    UserName: userResponse.UserDisplayName,
                    Response: userResponse.PollResponse,
                    MultiResponse: userResponse.PollMultiResponse,
                });
                // Update the user response
                await this.updatePollResponse(userResponse.PollQuestionId, allUserResponse);
            } else {
                allUserResponse.push({
                    UserID: userResponse.UserID,
                    UserName: userResponse.UserDisplayName,
                    Response: userResponse.PollResponse,
                    MultiResponse: userResponse.PollMultiResponse,
                });
                // Add the user response
                await this.addPollResponse(userResponse, allUserResponse);
            }
            return true;
        } catch (err) {
            console.log(err);
            return false;
        }
    }
        // Private method to map a question
    private async mapQuestion(item: any): Promise<IQuestionDetails> {

        const question: IQuestionDetails = {
            Id: item.ID,
            DisplayName: item.DisplayName,
            Choices:  item.Choices,
            DisplayNameFr: item.DisplayNameFr,
            ChoicesFr:  item.ChoicesFr,
            MultiChoice: item.MultiChoice ? item.MultiChoice : false,
            StartDate: new Date(item.StartDate),
            EndDate: new Date(item.EndDate),
            SortIdx: item.SortIdx,
        };
        return question;
    }
    

     // Public method to get a question and map it
    public async getQuestions () {
        const questions: IQuestionDetails[] = [];

        const items = await this._sp.web.lists.getByTitle(`${ListsUrl.Questions}`)
            .items.select(
            `${Fields.Question},${Fields.QuestionFr},${Fields.AnswersOptions},${Fields.AnswersOptionsFr},${Fields.StartDate},${Fields.EndDate},${Fields.isMultiChoice},${Fields.ID},${Fields.SortIdx},${Fields.UseDate},${Fields.Title}`
            )
            .top(3)
            .orderBy(Fields.ID, true)();

        if (items && items.length > 0) {
            for (const item of items) {
            questions.push(await this.mapQuestion(item));
            }
        }
        return questions.length > 0? questions : null;
    }
    
    
    /**
     * Check and create the User response list.
     */
    public checkListExists = async (): Promise<boolean> => {
        try {
            // Check if the response list exists
            await this._sp.web.lists.getByTitle(this.lst_response)();
            
            // Check if the questions list exists
            await this._sp.web.lists.getByTitle(this.questions_lst)();
            
            // If both lists exist, return true
            return true;
        } catch (err) {
            // If any of the lists do not exist, create them
            console.log("One or both lists do not exist, creating them...");
    
            // Ensure both lists exist
            const responseList = await this._sp.web.lists.ensure(this.lst_response);
            const questionsList = await this._sp.web.lists.ensure(this.questions_lst);
    
            // Add fields to the Questions List
            await this.addFieldsToQuestionsList(questionsList.list);
    
            // Add fields to the Response List
            await this.addFieldsToResponseList(responseList.list);
    
            // Return true indicating the lists were created and fields added
            return true;
        }
    };

    // Helper method to add fields to the Questions List
private addFieldsToQuestionsList = async (questionsList: any) => {
    await questionsList.fields.addMultilineText('DisplayName', {
        Required: false,
        Description: '',
        AppendOnly: false,
        RestrictedMode: false
    });
    await questionsList.fields.addMultilineText('DisplayNameFr', {
        Required: false,
        Description: '',
        AppendOnly: false,
        RestrictedMode: false
    });

    await questionsList.fields.addMultilineText('Choices', {
        Required: false,
        Description: '',
        AppendOnly: false,
        RestrictedMode: false
    });
    await questionsList.fields.addMultilineText('ChoicesFr', {
        Required: false,
        Description: '',
        AppendOnly: false,
        RestrictedMode: false
    });

    await questionsList.fields.addBoolean('MultiChoice', {
        Required: false,
        Description: '',
    });

    await questionsList.fields.addDateTime("StartDate", { 
        DisplayFormat: DateTimeFieldFormatType.DateOnly, 
        DateTimeCalendarType: CalendarType.Gregorian, 
        FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Disabled 
    });

    await questionsList.fields.addDateTime("EndDate", { 
        DisplayFormat: DateTimeFieldFormatType.DateOnly, 
        DateTimeCalendarType: CalendarType.Gregorian, 
        FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Disabled 
    });

    await questionsList.fields.addDateTime("UseDate", { 
        DisplayFormat: DateTimeFieldFormatType.DateOnly, 
        DateTimeCalendarType: CalendarType.Gregorian, 
        FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Disabled
    });
    await questionsList.fields.addBoolean('IsQuestionActive', {
        Required: false,
        Description: '',
    });

    await questionsList.fields.addBoolean('IsFeaturedPoll', {
        Required: false,
        Description: '',
    });

    await questionsList.fields.addNumber("SortIdx", { MinimumValue: 1, MaximumValue: 100 });

    // Add fields to the 'All Items' view
    const allQuestionsListItemsView = await questionsList.views.getByTitle('All Items');
    await allQuestionsListItemsView.fields.add('DisplayName');
    await allQuestionsListItemsView.fields.add('DisplayNameFr');
    await allQuestionsListItemsView.fields.add('Choices');
    await allQuestionsListItemsView.fields.add('ChoicesFr');
    await allQuestionsListItemsView.fields.add('MultiChoice');
    await allQuestionsListItemsView.fields.add('StartDate');
    await allQuestionsListItemsView.fields.add('EndDate');
    await allQuestionsListItemsView.fields.add('UseDate');
    await allQuestionsListItemsView.fields.add('SortIdx');
    await allQuestionsListItemsView.fields.add('IsQuestionActive');
    await allQuestionsListItemsView.fields.add('IsFeaturedPoll');

};

// Helper method to add fields to the Response List
private addFieldsToResponseList = async (responseList: any) => {
    await responseList.fields.addText('QuestionID', { MaxLength: 255, Required: true, Description: '' });
    
    await responseList.fields.addMultilineText('UserResponse', {
        Required: false,
        Description: '',
        AppendOnly: false,
        RestrictedMode: false
    });

    // Add fields to the 'All Items' view
    const allItemsView = await responseList.views.getByTitle('All Items');
    await allItemsView.fields.add('QuestionID');
    await allItemsView.fields.add('UserResponse');
};
    
}