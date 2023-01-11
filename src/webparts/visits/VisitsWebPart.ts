import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "VisitsWebPartStrings";
import Visits from "./components/Visits";
import { HttpService, IHttpService } from "./services/HttpService";
import { IVisitsProps } from "./Types/Types";

export interface IVisitsWebPartProps {
  description: string;
}

export default class VisitsWebPart extends BaseClientSideWebPart<IVisitsWebPartProps> {
  private HttpService: IHttpService;
  public render(): void {
    const element: React.ReactElement<IVisitsProps> = React.createElement(
      Visits,
      {
        context: this.context,
        HttpService: this.HttpService,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.HttpService = this.context.serviceScope.consume<IHttpService>(
      HttpService.serviceKey
    );
    return super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
