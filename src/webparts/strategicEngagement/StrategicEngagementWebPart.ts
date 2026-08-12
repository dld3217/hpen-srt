import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { SseFrontDoor, ISseFrontDoorProps } from './components/SseFrontDoor';

export interface IStrategicEngagementWebPartProps {
  description: string;
}

export default class StrategicEngagementWebPart extends BaseClientSideWebPart<IStrategicEngagementWebPartProps> {
  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    const element: React.ReactElement<ISseFrontDoorProps> = React.createElement(SseFrontDoor, {
      sp: this._sp,
      context: this.context,
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Strategic Engagement Request Settings' },
        groups: [],
      }],
    };
  }
}
