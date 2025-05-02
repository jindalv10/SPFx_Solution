export interface IResponseDetails {
    UserID: string;
    UserDisplayName: string;
    UserLoginName?: string;
    PollResponse?: string;
    PollMultiResponse?: string[];
    PollQuestion: string | undefined;
    PollQuestionId: string | undefined;
    IsMulti: boolean;
    Response?: string;

}