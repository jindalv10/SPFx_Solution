import * as React from 'react';
import { Spinner, SpinnerSize, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import styles from "./spinnerComponent.module.scss";

interface ISpinnerComponentProps {
    text: string;
    spinnerSize?: SpinnerSize;
    minHeightInPixel?: number;
}


export default function SpinnerComponent(props: ISpinnerComponentProps) {
    return (
        <div className={styles.formLoadingScreen} style={props.minHeightInPixel ? { minHeight: `${props.minHeightInPixel}px` } : {}}>
            <Spinner className={styles.spinner}
                size={props.spinnerSize ? props.spinnerSize : SpinnerSize.large} 
            
            />
            {props.text}
        </div>
    );
}
