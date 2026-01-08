import { useMsal } from "@azure/msal-react";

export const useIsAdmin = () => {
  const { accounts } = useMsal();
  
  if (accounts.length > 0) {
    const account = accounts[0];
    const roles = (account.idTokenClaims as any)?.roles || [];
    return roles.includes("Totem.Admin");
  }

  return false;
};
