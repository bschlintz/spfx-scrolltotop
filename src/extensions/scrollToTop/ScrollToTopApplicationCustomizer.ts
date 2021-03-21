import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';

import styles from './ScrollToTop.module.scss';

const LOG_SOURCE: string = 'ScrollToTopButton';

const SCROLL_TO_TARGET_SELECTORS = [
  "[data-automation-id='pageHeader']",
  "[data-automation-id='contentScrollRegion'] > *",
  ".webPartContainer"
];

export interface IScrollToTopApplicationCustomizerProperties {
}

export default class ScrollToTopApplicationCustomizer
  extends BaseApplicationCustomizer<IScrollToTopApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent;
  private _isScrolling: boolean = false;

  @override
  public async onInit(): Promise<void> {
    this.context.application.navigatedEvent.add(this, this.onNavigated);
  }

  /**
   * Event handler that fires on every page load
   */
  private async onNavigated(): Promise<void> {
    // Only show the button when on a site page and NOT in IE (elem.scrollIntoView function not available in IE)
    if (this.isOnSitePage() && !this.isIE()) {
      this.renderButton();
    }
    // Else remove the button if isn't already disposed
    else if (this._bottomPlaceholder) {
      this._bottomPlaceholder.dispose();
    }
  }

  /**
   * Determine whether we're currently on a Site Page
   */
  private isOnSitePage = (): boolean => {
    return !!this.context.pageContext.list && // We're on a list
           this.context.pageContext.list.title === "Site Pages" && // And its title is "Site Pages"
           !!this.context.pageContext.listItem; // And we have an item (i.e. Page)
  }

  /**
   * Render footer React component
   */
  private renderButton(): void {
    if (!this._bottomPlaceholder || this._bottomPlaceholder.isDisposed) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);

      if (!this._bottomPlaceholder) {
        Log.error(LOG_SOURCE, new Error(`Unable to render Bottom placeholder`));
        return;
      }

      this._bottomPlaceholder.domElement.innerHTML = `
        <div class="${styles.container}" title="Back to top">
          <div class="${styles.icon}"></div>          
        </div>
      `;
  
      this._bottomPlaceholder.domElement.addEventListener('click', this.handleButtonClick);
    }
  }

  /**
   * Get scroll target
   */
  private getScrollToTarget(): Element {
    // Each selector in SCROLL_TO_TARGET_SELECTORS will be tried in order. 
    // As soon as we find a valid element, we'll use it.
    // There are multiple selectors as fallbacks for when the DOM changes.
    let targetElement = null;
    for (let idx = 0; idx < SCROLL_TO_TARGET_SELECTORS.length; idx++) {
      const element = document.querySelector(SCROLL_TO_TARGET_SELECTORS[idx]);
      if (element) {
        targetElement = element;
        break;
      }
    }
    if (!targetElement) {
      Log.error(LOG_SOURCE, new Error(`No scroll target found. Tried ${SCROLL_TO_TARGET_SELECTORS.length} selectors.`));
    }
    return targetElement;
  }

  /**
   * Button click handler
   */
  private handleButtonClick = (): void => {
    const scrollTarget = this.getScrollToTarget();
    if (scrollTarget && !this._isScrolling) {
      this._isScrolling = true;
      scrollTarget.scrollIntoView({ behavior: 'smooth' });
      // The SharePoint header has an expansion delay when scrolling up
      // After about half a second, trigger the scroll again to get to the *VERY* top
      setTimeout(() =>  {
        scrollTarget.scrollIntoView({ behavior: 'smooth' });
        this._isScrolling = false;
      }, 600);
    }
  }

  private isIE = (): boolean => {
    var old_ie = window.navigator.userAgent.indexOf('MSIE ');
    var less_old_ie = window.navigator.userAgent.indexOf('Trident/');
    return (old_ie > -1) || (less_old_ie > -1);
  }
}
