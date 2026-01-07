import { useState } from "react";
import { protectedResources } from "../../auth/authConfig";
import { PowerAutomateService } from "../../services/powerAutomateService";

export function PowerAutomatePanel() {
  const [status, setStatus] = useState<string>("");
  const [error, setError] = useState<string>("");
  const service = new PowerAutomateService();

  const pingFlow = async () => {
    setStatus("Invio...");
    setError("");
    try {
      await service.invoke({ action: "send", source: "blueprint", message: "hello" });
      setStatus("Flow chiamato (controlla run history)");
    } catch (err: any) {
      setError(err?.message || "Errore chiamata flow");
    }
  };

  return (
    <div className="panel">
      <h3>Power Automate</h3>
      <p>Endpoint HTTP trigger per email/OTP. Definisci lo schema in Power Automate e incolla l'URL nell'env.</p>
      {!protectedResources.powerAutomateFlowUrl && (
        <div className="card-highlight">Imposta VITE_PA_FLOW_URL per abilitare questo test.</div>
      )}
      <div className="button-row" style={{ marginBottom: 12 }}>
        <button className="btn" onClick={pingFlow} disabled={!protectedResources.powerAutomateFlowUrl} type="button">
          Test HTTP Flow
        </button>
        <span className="tag">{protectedResources.powerAutomateFlowUrl ? "URL pronto" : "Manca URL"}</span>
      </div>
      {status && <div className="card-highlight">{status}</div>}
      {error && <div className="card-highlight" style={{ borderColor: "#ef4444", color: "#fca5a5" }}>{error}</div>}
    </div>
  );
}
