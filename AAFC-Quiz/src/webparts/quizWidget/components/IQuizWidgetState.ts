import { IQuizDetails, IResponseDetails } from "../models";

export interface IQuizWidgetState {
	listExists: boolean;
	QuizQuestions: IQuizDetails[];
	UserResponse: IResponseDetails[];
	displayQuestionId: string;
	displayQuestion: IQuizDetails | null;  // Allow null here;
	enableSubmit: boolean;
	enableChoices: boolean;
	showOptions: boolean;
	showProgress: boolean;
	showMessage: boolean;
	showAnswer: boolean;
	isError: boolean;
	MsgContent: string;
	showSubmissionProgress: boolean;
	currentQuizResponse: {AnswerVal: string; isEqual: boolean};
}