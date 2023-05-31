import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'InternalJobsAdaptiveCardExtensionStrings';
import { IInternalJobsAdaptiveCardExtensionProps, IInternalJobsAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../InternalJobsAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IInternalJobsAdaptiveCardExtensionProps, IInternalJobsAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IBasicCardParameters {
    return {
      title: this.properties.title,
      primaryText: (this.state.listTitle)
        ? `View ${this.state.length} job openings in Creospark!`
        : `Missing list ID`,
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: {
        view: QUICK_VIEW_REGISTRY_ID
      }
    };
  }
};