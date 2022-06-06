import { QueryClient, useQuery } from 'react-query'
import { PrincipalSource, PrincipalType, spfi, SPFx } from '@pnp/sp'
import '@pnp/sp'
import '@pnp/sp/webs'
import '@pnp/sp/site-users'
import '@pnp/sp/security'
import '@pnp/sp/sputilities'
import { WebPartContext } from '@microsoft/sp-webpart-base'
import { PermissionKind } from '@pnp/sp/security'


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

			const userIsOwner = await sp.web.currentUserHasPermissions(PermissionKind.ManageWeb)

			return { ...userPrincipal, userIsOwner }
		}
	)

	return userInfo.data
}