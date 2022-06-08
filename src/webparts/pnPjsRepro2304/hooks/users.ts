import '@pnp/sp'
import { PrincipalSource, PrincipalType, SPFI } from '@pnp/sp'
import '@pnp/sp/security'
import '@pnp/sp/site-users'
import '@pnp/sp/sputilities'
import '@pnp/sp/webs'
import { useQuery } from 'react-query'


export const useCurrentUserPrincipal = (spfi: SPFI) => {

	const userInfo = useQuery(
		['currentUser'],
		async () => {
			const user = await spfi.web.currentUser()
			const userPrincipal = await spfi.utility.resolvePrincipal(
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