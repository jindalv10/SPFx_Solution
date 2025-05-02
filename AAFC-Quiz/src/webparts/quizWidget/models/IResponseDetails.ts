export interface IResponseDetails {
    UserID: string | undefined;
    UserDisplayName: string | undefined;
    UserLoginName?: string;
    QuizResponse?: string;
    QuizMultiResponse?: string[];
    QuizQuestion: string | undefined;
    QuizQuestionId: string | undefined;
    IsMulti: boolean;
    Response?: string;
}