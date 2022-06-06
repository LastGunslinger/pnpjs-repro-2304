import { escape } from '@microsoft/sp-lodash-subset'
import '@pnp/sp'
import '@pnp/sp/security'
import '@pnp/sp/site-users'
import '@pnp/sp/webs'
import * as React from 'react'
import { useState } from 'react'
import { QueryClient, QueryClientProvider } from 'react-query'
import { useCurrentUserPrincipal } from '../hooks/users'
import { IPnPjsRepro2304Props } from './IPnPjsRepro2304Props'
import styles from './PnPjsRepro2304.module.scss'


const PnPjsRepro2304: React.FC<IPnPjsRepro2304Props> = (props) => {
	const { hasTeamsContext, isDarkTheme, userDisplayName, environmentMessage, description } = props
	const currentUser = useCurrentUserPrincipal(props.context)


	return (
		<section className={`${styles.pnPjsRepro2304} ${hasTeamsContext ? styles.teams : ''}`}>
			<div className={styles.welcome}>
				<img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
				<h2>Well done, {escape(userDisplayName)}!</h2>
				<div>{environmentMessage}</div>
				<div>Web part property value: <strong>{escape(description)}</strong></div>
			</div>
			<div>
				<h3>Welcome to SharePoint Framework!</h3>
				<p>
					The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
				</p>
				<h4>Learn more about SPFx development:</h4>
				<ul className={styles.links}>
					<li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
					<li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
					<li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
					<li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
					<li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
					<li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
					<li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
				</ul>
				<p>UserPrincipal DisplayName: {currentUser?.DisplayName ?? ''}</p>
			</div>
		</section >
	)
}

export const QueryWrapper: React.FC<IPnPjsRepro2304Props> = (props) => {
	const [client] = useState(new QueryClient({
		defaultOptions: {
			queries: {
				refetchOnMount: true,
				refetchOnReconnect: true,
				refetchOnWindowFocus: true,
				staleTime: 1000 * 60 * 5,  // 5 minute stale time
				structuralSharing: true,
			}
		}
	}))

	return <QueryClientProvider client={client}>
		<PnPjsRepro2304 {...props} />
	</QueryClientProvider>
}