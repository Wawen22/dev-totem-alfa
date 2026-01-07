import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../../auth/authConfig";

export function LoginButton() {
  const { instance } = useMsal();

  const handleLogin = async () => {
    await instance.loginRedirect(loginRequest);
  };

  return (
    <button className="btn" onClick={handleLogin} type="button">
      Accedi con Microsoft
    </button>
  );
}
