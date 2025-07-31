import {ICustomLogger} from '../../models/ICustomLogger';
import {ICustomLogMessage} from '../../models/ICustomLogMessage';
import { sp } from '@pnp/sp/presets/all';
import SPService from '../SPService';

export default class CustomLogger extends SPService implements ICustomLogger {

  constructor(context: any) {
    super(context);
  }
    public Log = async (logMessage: ICustomLogMessage) => {
        try {
            this.saveLogs(logMessage, "Log");
        } catch (error) {
            //Can't do anything
            console.error(error.Message);
        }
    }

    public Warn = async (logMessage: ICustomLogMessage) => {
        try {
            this.saveLogs(logMessage, "Warn");
        } catch (error) {
            //Can't do anything
            console.error(error.Message);
        }
    }

    public Verbose = async (logMessage: ICustomLogMessage) => {
        try {
            this.saveLogs(logMessage, "Verbose");
        } catch (error) {
            //Can't do anything
            console.error(error.Message);
        }
    }

    public Error = async (logMessage: ICustomLogMessage) => {
        try {
            console.error(logMessage.Message);
            this.saveLogs(logMessage, "Error");
        } catch (error) {
            //Can't do anything
            console.error(error.Message);
        }
    }

    private saveLogs = async (logMessage: ICustomLogMessage, logType: string) => {
        sp.web.lists.getByTitle('ExceptionLogs').items.add({
            WebPartName: logMessage.WebPartName,
            ComponentName: logMessage.ComponentName,
            MethodName: logMessage.MethodName,
            Message: logMessage.Message,
        });
    }
}
