export enum MessageScope{
    Success = 0,
    Failure = 1,
    Warning = 2,
    Info = 3
}

export enum PanelColors{
    Success = 0,
    Danger = 1,
    Warning = 2,
    Info = 3
}

export enum FilterViews{
    MostPopular,
    Latest
}

export class Fields
{
    static readonly Question= 'DisplayName';
    static readonly QuestionFr= 'DisplayNameFr';
    static readonly AnswersOptions= 'Choices';
    static readonly AnswersOptionsFr= 'ChoicesFr';
    static readonly isMultiChoice= 'MultiChoice';
    static readonly ID= 'ID';
    static readonly IsQuestionActive= 'IsQuestionActive';
    static readonly Title= 'Title';
    static readonly UserEmail= 'UserEmail';
    static readonly StartDate= 'StartDate';
    static readonly EndDate= 'EndDate';
    static readonly SortIdx= 'SortIdx';
    static readonly UseDate= 'UseDate';

}

export class ListsUrl
{
    static readonly Questions = 'PollQuestions'
    static readonly Answers = 'PollAnswers'
}