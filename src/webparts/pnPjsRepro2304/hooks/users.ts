import '@pnp/sp'
import { PrincipalSource, PrincipalType, sp } from '@pnp/sp'
import '@pnp/sp/security'
import '@pnp/sp/site-users'
import '@pnp/sp/sputilities'
import '@pnp/sp/webs'
import { useQuery } from 'react-query'


export const useCurrentUser = () => {

	const userInfo = useQuery(
		['currentUser'],
		async () => {
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