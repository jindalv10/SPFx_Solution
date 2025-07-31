import { ICustomLogMessage } from "./ICustomLogMessage";

export interface ICustomLogger {
    Log(logMessage: ICustomLogMessage): void;
    Warn(logMessage: ICustomLogMessage): void;
    Verbose(logMessage: ICustomLogMessage): void;
    Error(logMessage: ICustomLogMessage): void;
}


