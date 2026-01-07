import { useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import { Client } from "@microsoft/microsoft-graph-client";
import { loginRequest } from "../auth/authConfig";

export function useAuthenticatedGraphClient() {
  const { instance, accounts } = useMsal();

  return useCallback(async () => {
    const account = accounts[0];
    if (!account) {
      throw new Error("Nessun account attivo trovato. Effettua il login.");
    }

    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account,
    });

    return Client.init({
      authProvider: (done) => done(null, response.accessToken),
      defaultVersion: "v1.0",
    });
  }, [accounts, instance]);
}
