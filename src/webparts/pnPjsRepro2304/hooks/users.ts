import { WebPartContext } from '@microsoft/sp-webpart-base'
import '@pnp/sp'
import { PrincipalSource, PrincipalType, spfi, SPFx } from '@pnp/sp'
import '@pnp/sp/security'
import '@pnp/sp/site-users'
import '@pnp/sp/sputilities'
import '@pnp/sp/webs'
import { useQuery } from 'react-query'


export const useCurrentUserPrincipal = (context: WebPartContext) => {

	const userInfo = useQuery(
		['currentUser'],
		async () => {
			const sp = spfi().using(SPFx(context))
			const user = await sp.web.currentUser()
			const userPrincipal = await sp.utility.resolvePrincipal(
				user.LoginName,
				PrincipalType.All,
				PrincipalSource.All,
				false,
				true,
				true,
			)

			return userPrincipal
		}
	)

	return userInfo.data
}