
export interface IQuizDetails {
	Id: string;
	DisplayName: string;
	DisplayNameFr?:string;
	Choices?: string |undefined;
	ChoicesFr?: string |undefined;
	MultiChoice?: boolean;
	StartDate: Date;
	EndDate: Date;
	UseDate?: boolean;
	SortIdx: number;
	Answers?: string;
	AnswersFr?: string;
	Explanations?: string;
	ExplanationsFr?: string;
}
