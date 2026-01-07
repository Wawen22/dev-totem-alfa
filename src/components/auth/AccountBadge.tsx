import { useMsal } from "@azure/msal-react";

export function AccountBadge() {
  const { accounts } = useMsal();
  const account = accounts[0];

  if (!account) {
    return <span className="pill">Non autenticato</span>;
  }

  const claims = (account.idTokenClaims || {}) as any;
  const email = (claims?.preferred_username || account.username || "").toString();

  return (
    <div className="pill">
      <span style={{ width: 10, height: 10, borderRadius: "50%", background: "#22d3ee" }} />
      <span>{email}</span>
    </div>
  );
}
