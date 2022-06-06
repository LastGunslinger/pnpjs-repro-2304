import { QueryClient, useQuery } from 'react-query'
import { spfi, SPFx } from '@pnp/sp'
import '@pnp/sp'
import '@pnp/sp/webs'
import '@pnp/sp/site-users'
import '@pnp/sp/security'
import { WebPartContext } from '@microsoft/sp-webpart-base'
import { PermissionKind } from '@pnp/sp/security'


export const useCurrentUser = (context: WebPartContext) => {

	const userInfo = useQuery(
		['currentUser'],
		async () => {
			const sp = spfi().using(SPFx(context))
			const user = await sp.web.currentUser()

			const userIsOwner = await sp.web.currentUserHasPermissions(PermissionKind.ManageWeb)

			return { ...user, userIsOwner }
		}
	)

	return userInfo.data
}