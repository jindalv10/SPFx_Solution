import * as React from 'react';
import styles from './Tabs.module.scss';
import type { ITabsProps } from './ITabsProps';
import { DisplayMode } from '@microsoft/sp-core-library';
// @ts-ignore:
import * as $ from 'jquery';  // Ensure jQuery is installed

export interface ITabsState {
  activeTab: number;
}


export default class Tabs extends React.Component<ITabsProps, ITabsState> {

  constructor(props: ITabsProps) {
    super(props);
    this.state = {
      activeTab: 0,
    };
  }


  private handleTabClick(index: number) {
    this.setState({ activeTab: index });
  }

   // Function to render tabs and content
   private  renderTabsAndContents() {
    require('./AddTabs.js');
    require('./AddTabs.css');
    const tabData = this.props.tabData;
    const { activeTab } = this.state;

    const tabs = tabData.map((tab, index) => (
      <div
        key={index}
        className={`tab ${index === activeTab ? 'active' : ''}`}
        onClick={() => this.handleTabClick(index)}
      >
        {tab.TabLabel}
      </div>
    ));
    // @ts-ignore:

    const tabWebPartID = $(this.domElement).closest("div." + this.props.webPartClass).attr("id");
    // @ts-ignore:
       
    
    const tabsDiv = `${tabWebPartID}tabs`;
    const contentsDiv = `${tabWebPartID}Contents`;

    const contents = tabData.map((tab, index) => (
      <div key={index} style={{ display: index === activeTab ? 'block' : 'none' }}>
        {/* Render the corresponding content, assuming tab.WebPartID corresponds to a component */}
        <div id={tab.WebPartID}>{}</div>
      </div>
    ));

    return (
      <div data-addui="tabs">
        <div role="tabs">{tabs}</div>
        <div role="contents">{contents}</div>
      </div>
    );
  }


  public render(): React.ReactElement<ITabsProps> {
    const {
      hasTeamsContext,
      displayMode,
    } = this.props;

    return (
      <section className={`${styles.tabs} ${hasTeamsContext}`}>
        {
        displayMode === DisplayMode.Read ?
          this.renderTabsAndContents()
        :  
        <div className={styles.tabs}>
          <div className={ styles.container}>
            <div className={ styles.row }>
              <div className={ styles.column}>
                <span className={styles.title}>Modern Hillbilly Tabs By Mark Rackley</span>
                <p className={styles.subTitle}>Place Web Parts into Tabs.</p>
                <p className={styles.description}>To use Modern Hillbilly Tabs: 
                  <ul>
                    <li>Place this web part in the same section of the page as the web parts you would like to put into tabs.</li> 
                    <li>Add the web parts to the section and then edit the properties of this web part.</li>
                    <li>Click on the button to 'Manage Tab Labels' and then specify the labels for each web part using the property control.</li>
                  </ul> 
                  The other two Web Part Properties are used to identify sections/web parts on the screen. Do not change these values unless you know what you are doing.</p>
                <a href="https://github.com/mrackley/Modern_Hillbilly_Tabs" className="${ styles.button }">
                  <span className="${ styles.label }">View Source on GitHub</span>
                </a>
                <a href="http://www.markrackley.net/2022/06/29/the-return-of-hillbilly-tabs/" className="${ styles.button }">
                  <span className="${ styles.label }">View Blog Post</span>
                </a>
              </div>
            </div>
          </div>
        </div>
      }
      </section>
    );
  }
}
