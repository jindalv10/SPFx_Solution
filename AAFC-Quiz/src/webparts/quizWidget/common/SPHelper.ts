import { SPFI, spfi ,SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/folders/web";
import "@pnp/sp/files/folder";
import "@pnp/sp/items/list";
import "@pnp/sp/fields/list";
import "@pnp/sp/views/list";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items";
import { IUserInfo, IQuizDetails } from "../models";
import { Fields, ListsUrl } from "./enumHelper";
import { CalendarType, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType } from "@pnp/sp/fields/types";

export default class SPHelper {
    
    private questions_lst: string = "";

    private _sp:SPFI | undefined = undefined;

    public constructor(context: any) {
        this.questions_lst = ListsUrl.QuizQuestions;
        this._sp = spfi().using(SPFx(context));

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
   
   
    // Private method to map a question
    private async mapQuestion(item: any): Promise<IQuizDetails> {

        const question: IQuizDetails = {
            Id: item.ID,
            DisplayName: item.DisplayName,
            Choices:  item.Choices,
            DisplayNameFr: item.DisplayNameFr,
            ChoicesFr:  item.ChoicesFr,
            MultiChoice: item.MultiChoice ? item.MultiChoice : false,
            StartDate: new Date(item.StartDate),
            EndDate: new Date(item.EndDate),
            SortIdx: item.SortIdx,
            Answers: item.Answers,
            AnswersFr: item.AnswersFr,
            Explanations: item.Explanations,
            ExplanationsFr: item.ExplanationsFr,
        };
        return question;
    }
    

     // Public method to get a question and map it
    public async getQuestions () {
        const questions: IQuizDetails[] = [];

        const items = await this._sp.web.lists.getByTitle(`${ListsUrl.QuizQuestions}`)
            .items.select(
            `${Fields.IsFeaturedQuiz},${Fields.Question},${Fields.QuestionFr},${Fields.AnswersOptions},${Fields.AnswersOptionsFr},${Fields.Answers},${Fields.AnswersFr},${Fields.Explanations},${Fields.ExplanationsFr},${Fields.StartDate},${Fields.EndDate},${Fields.isMultiChoice},${Fields.ID},${Fields.SortIdx},${Fields.UseDate},${Fields.Title}`
            ).filter(`${Fields.IsFeaturedQuiz}` + " eq '" + 1 + "'")
            .top(5000)
            .orderBy(Fields.ID, true)();
        if (items && items.length > 0) {
            for(const item of items) {
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
            //await this._sp.web.lists.getByTitle(this.lst_response)();
            
            // Check if the questions list exists
            await this._sp.web.lists.getByTitle(this.questions_lst)();
            
            // If both lists exist, return true
            return true;
        } catch (err) {
            // If any of the lists do not exist, create them
            console.log("One or both lists do not exist, creating them...");
    
            // Ensure both lists exist
            //const responseList = await this._sp.web.lists.ensure(this.lst_response);
            const questionsList = await this._sp.web.lists.ensure(this.questions_lst);
    
            // Add fields to the Questions List
            await this.addFieldsToQuestionsList(questionsList.list);
    
            // Add fields to the Response List
            //await this.addFieldsToResponseList(responseList.list);
    
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
        await questionsList.fields.addMultilineText('Answers', {
            Required: false,
            Description: '',
            AppendOnly: false,
            RestrictedMode: false
        });
        await questionsList.fields.addMultilineText('AnswersFr', {
            Required: false,
            Description: '',
            AppendOnly: false,
            RestrictedMode: false
        });
        await questionsList.fields.addMultilineText('Explanations', {
            Required: false,
            Description: '',
            AppendOnly: false,
            RestrictedMode: false
        });
        await questionsList.fields.addMultilineText('ExplanationsFr', {
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
        await questionsList.fields.addBoolean('IsFeaturedQuiz', {
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
        await allQuestionsListItemsView.fields.add('Answers');
        await allQuestionsListItemsView.fields.add('AnswersFr');
        await allQuestionsListItemsView.fields.add('Explanations');
        await allQuestionsListItemsView.fields.add('ExplanationsFr');
        await allQuestionsListItemsView.fields.add('MultiChoice');
        await allQuestionsListItemsView.fields.add('StartDate');
        await allQuestionsListItemsView.fields.add('EndDate');
        await allQuestionsListItemsView.fields.add('UseDate');
        await allQuestionsListItemsView.fields.add('SortIdx');
        await allQuestionsListItemsView.fields.add('IsQuestionActive');
        await allQuestionsListItemsView.fields.add('IsFeaturedQuiz');

    };
    
}