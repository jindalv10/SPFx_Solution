import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { Text } from '@fluentui/react';
import styles from './MessageContainer.module.scss';
import { MessageScope } from '../../common/enumHelper';

export interface IMessageContainerProps {
    Message?: string;
    LongMessage?: string;
    MessageScope: MessageScope;
}

export default function MessageContainer(props: IMessageContainerProps) {
    return (
        <div className={styles.MessageContainer}>
            {
                props.MessageScope === MessageScope.Success &&
                <MessageBar className={styles.errorMessage} messageBarType={MessageBarType.success}>
                    <Text block variant={"mediumPlus"}>{props.Message}</Text>
                   
                    {
                    props.LongMessage &&
                        <>
                            <br />
                            <br />
                            <Text block variant={"mediumPlus"}>{props.LongMessage}</Text>
                        </>
                    }
                </MessageBar>
            }
            {
                props.MessageScope === MessageScope.Failure &&
                <MessageBar  className={styles.successMessage} messageBarType={MessageBarType.error}>
                    <Text block variant={"mediumPlus"}>{props.Message}</Text>
                    {
                    props.LongMessage &&
                        <>
                            <br />
                            <br />
                            <Text block variant={"mediumPlus"}>{props.LongMessage}</Text>
                        </>
                    }
                </MessageBar>
            }
            {
                props.MessageScope === MessageScope.Warning &&
                <MessageBar messageBarType={MessageBarType.warning}>
                    <Text block variant={"mediumPlus"}>{props.Message}</Text>
                    {
                    props.LongMessage &&
                        <>
                            <br />
                            <br />
                            <Text block variant={"mediumPlus"}>{props.LongMessage}</Text>
                        </>
                    }

                </MessageBar>
            }
            {
                props.MessageScope === MessageScope.Info &&
                <MessageBar className={styles.infoMessage}>
                    <Text block variant={"mediumPlus"}>{props.Message}</Text>
                    {
                    props.LongMessage &&
                        <>
                            <br />
                            <br />
                            <Text block variant={"mediumPlus"}>{props.LongMessage}</Text>
                        </>
                    }
                </MessageBar>
            }
        </div>
    );
}