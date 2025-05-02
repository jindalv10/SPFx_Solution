import { IPollAnalyticsInfo } from "./IPollAnalyticsInfo";

export interface IQuestionDetails {
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
	PollAnalytics?: IPollAnalyticsInfo;
	UserResponse?: {
		Response?: string, // The current user's response
		FullResponses?: any[] // All responses for the question
	  },
	  
}
