import { IReadonlyTheme } from '@microsoft/sp-component-base'
import { Version } from '@microsoft/sp-core-library'
import {
	IPropertyPaneConfiguration,
	PropertyPaneTextField
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import * as strings from 'PnPjsRepro2304WebPartStrings'
import * as React from 'react'
import * as ReactDom from 'react-dom'
import { IPnPjsRepro2304Props } from './components/IPnPjsRepro2304Props'
import { QueryWrapper } from './components/PnPjsRepro2304'


export interface IPnPjsRepro2304WebPartProps {
	description: string
}

export default class PnPjsRepro2304WebPart extends BaseClientSideWebPart<IPnPjsRepro2304WebPartProps> {

	private _isDarkTheme: boolean = false;
	private _environmentMessage: string = '';

	protected async onInit(): Promise<void> {
		this._environmentMessage = this._getEnvironmentMessage()

		await super.onInit()
	}

	public render(): void {
		const element: React.ReactElement<IPnPjsRepro2304Props> = React.createElement(
			QueryWrapper,
			{
				description: this.properties.description,
				isDarkTheme: this._isDarkTheme,
				environmentMessage: this._environmentMessage,
				hasTeamsContext: !!this.context.sdks.microsoftTeams,
				userDisplayName: this.context.pageContext.user.displayName,
				context: this.context,
			}
		)

		ReactDom.render(element, this.domElement)
	}

	private _getEnvironmentMessage(): string {
		if (!!this.context.sdks.microsoftTeams) { // running in Teams
			return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment
		}

		return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment
	}

	protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
		if (!currentTheme) {
			return
		}

		this._isDarkTheme = !!currentTheme.isInverted
		const {
			semanticColors
		} = currentTheme
		this.domElement.style.setProperty('--bodyText', semanticColors.bodyText)
		this.domElement.style.setProperty('--link', semanticColors.link)
		this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered)

	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement)
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0')
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField('description', {
									label: strings.DescriptionFieldLabel
								})
							]
						}
					]
				}
			]
		}
	}
}
