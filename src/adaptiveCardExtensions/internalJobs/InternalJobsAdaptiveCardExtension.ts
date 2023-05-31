import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { InternalJobsPropertyPane } from './InternalJobsPropertyPane';
import {
  fetchListItems,
  fetchListTitle,
  fetchListLength,
  IListItem
} from './sp.service';

export interface IInternalJobsAdaptiveCardExtensionProps {
  title: string;
  listId: string;
}

export interface IInternalJobsAdaptiveCardExtensionState {
  listTitle: string;
  listItems: IListItem[];
  currentIndex: number;
  length: number;
  jobUrl: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'InternalJobs_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'InternalJobs_QUICK_VIEW';

export default class InternalJobsAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IInternalJobsAdaptiveCardExtensionProps,
  IInternalJobsAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: InternalJobsPropertyPane | undefined;

  public async onInit(): Promise<void> {
    this.state = {
      currentIndex: 0,
      listTitle: '',
      listItems: [],
      length: 0,
      jobUrl: ''
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    if (this.properties.listId) {
      Promise.all([
        this.setState({ listTitle: await fetchListTitle(this.context, this.properties.listId) }),
        this.setState({ listItems: await fetchListItems(this.context, this.properties.listId) }),
        this.setState({ length: await fetchListLength(this.context, this.properties.listId) }),
      ]);
    }

    return Promise.resolve();
  }

  protected async loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'InternalJobs-property-pane'*/
      './InternalJobsPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.InternalJobsPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
  
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listId' && newValue !== oldValue) {
      if (newValue) {
        (async () => {
          console.log('fetching');
          this.setState({ listTitle: await fetchListTitle(this.context, newValue) });
          this.setState({ listItems: await fetchListItems(this.context, newValue) });
          this.setState({ length: await fetchListLength(this.context, newValue) });
        })();
      } else {
        this.setState({ listTitle: '' });
        this.setState({ listItems: [] });
        this.setState({ length: 0 });
      }
    }
  }
}