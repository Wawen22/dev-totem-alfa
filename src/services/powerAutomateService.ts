export interface FlowPayload {
  action: "send" | "resend" | "otp" | string;
  [key: string]: unknown;
}

export class PowerAutomateService {
  constructor(private readonly flowUrl = import.meta.env.VITE_PA_FLOW_URL) {}

  async invoke(payload: FlowPayload) {
    if (!this.flowUrl) {
      throw new Error("VITE_PA_FLOW_URL non configurato");
    }

    const response = await fetch(this.flowUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Flow HTTP ${response.status}: ${text}`);
    }

    return response.json().catch(() => ({}));
  }
}
