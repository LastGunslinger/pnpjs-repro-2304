import { WebPartContext } from '@microsoft/sp-webpart-base'
import '@pnp/sp'
import { sp } from '@pnp/sp'
import '@pnp/sp/security'
import { PermissionKind } from '@pnp/sp/security'
import '@pnp/sp/site-users'
import '@pnp/sp/webs'
import { useQuery } from 'react-query'


export const useCurrentUser = () => {

	const userInfo = useQuery(
		['currentUser'],
		async () => {
			const user = await sp.web.currentUser()

			const userIsOwner = await sp.web.currentUserHasPermissions(PermissionKind.ManageWeb)

			return { ...user, userIsOwner }
		}
	)

	return userInfo.data
}