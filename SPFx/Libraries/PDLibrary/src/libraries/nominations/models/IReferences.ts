import { ISpUser } from "./ISpUser";

export interface IReferences {
  id: number;
  referencesUser?:ISpUser;
  referencesTrackVal?: string;
}
