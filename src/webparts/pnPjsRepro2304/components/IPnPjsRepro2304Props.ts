import { WebPartContext } from '@microsoft/sp-webpart-base'

export interface IPnPjsRepro2304Props {
	description: string
	isDarkTheme: boolean
	environmentMessage: string
	hasTeamsContext: boolean
	userDisplayName: string
	context: WebPartContext
}
