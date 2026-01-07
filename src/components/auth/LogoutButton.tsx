import { useMsal } from "@azure/msal-react";

export function LogoutButton() {
  const { instance, accounts } = useMsal();

  const handleLogout = async () => {
    const activeAccount = accounts[0] || instance.getActiveAccount();
    await instance.logoutRedirect({ account: activeAccount ?? undefined });
  };

  return (
    <button className="btn secondary" onClick={handleLogout} type="button">
      Esci
    </button>
  );
}
