import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useAuthenticatedGraphClient } from "../../hooks/useAuthenticatedGraphClient";

// --- Icons ---
const FolderIcon = () => (
  <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ color: "#3b82f6" }}>
    <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
  </svg>
);

const PdfIcon = () => (
  <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ color: "#ef4444" }}>
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
    <polyline points="14 2 14 8 20 8"></polyline>
    <line x1="16" y1="13" x2="8" y2="13"></line>
    <line x1="16" y1="17" x2="8" y2="17"></line>
    <polyline points="10 9 9 9 8 9"></polyline>
  </svg>
);

const FileIcon = () => (
  <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ color: "#94a3b8" }}>
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
    <polyline points="14 2 14 8 20 8"></polyline>
  </svg>
);

const BackIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="15 18 9 12 15 6"></polyline>
  </svg>
);

const HomeIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path>
    <polyline points="9 22 9 12 15 12 15 22"></polyline>
  </svg>
);

const CloseIcon = () => (
  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <line x1="18" y1="6" x2="6" y2="18"></line>
    <line x1="6" y1="6" x2="18" y2="18"></line>
  </svg>
);

const RefreshIcon = ({ spinning }: { spinning?: boolean }) => (
  <svg 
    width="18" 
    height="18" 
    viewBox="0 0 24 24" 
    fill="none" 
    stroke="currentColor" 
    strokeWidth="2" 
    strokeLinecap="round" 
    strokeLinejoin="round"
    style={{ animation: spinning ? "spin 1s linear infinite" : "none" }}
  >
    <polyline points="23 4 23 10 17 10"></polyline>
    <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"></path>
  </svg>
);

const encodePathSegments = (path: string) =>
  path.split("/").map((seg) => encodeURIComponent(seg)).join("/");

const isPdf = (name?: string, mime?: string) => {
  const lower = (name || "").toLowerCase();
  if (mime && mime.toLowerCase() === "application/pdf") return true;
  return lower.endsWith(".pdf");
};

type DriveItem = {
  id: string;
  name: string;
  webUrl: string;
  size?: number;
  folder?: { childCount: number };
  file?: { mimeType?: string };
  lastModifiedDateTime?: string;
  "@microsoft.graph.downloadUrl"?: string;
};

export function DocumentBrowser({
  siteId,
  driveId,
  initialPath,
}: {
  siteId?: string;
  driveId?: string;
  initialPath?: string;
}) {
  const getClient = useAuthenticatedGraphClient();
  const [path, setPath] = useState(initialPath?.trim() || "");
  const [items, setItems] = useState<DriveItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [viewer, setViewer] = useState<{ name: string; blobUrl?: string; loading: boolean } | null>(null);
  const blobUrlRef = useRef<string | null>(null);

  const breadcrumb = useMemo(() => {
    if (!path) return [] as string[];
    // Remove initial slash if present to avoid empty first segment
    const cleanPath = path.startsWith('/') ? path.slice(1) : path;
    const segments = cleanPath.split("/").filter(Boolean);
    // If we are deeper than initial path but path contains initial path, we might want to trim it for display? 
    // For now simple split is fine.
    return segments;
  }, [path]);

  const fetchItems = useCallback(async () => {
    if (!siteId || !driveId) {
      setError("Configurazione SharePoint mancante (ID sito o drive non trovati).");
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const client = await getClient();
      const base = `/sites/${siteId}/drives/${driveId}`;
      const api = path
        ? `${base}/root:/${encodePathSegments(path)}:/children`
        : `${base}/root/children`;
      
      // Select removed to ensure @microsoft.graph.downloadUrl is present
      const response = await client.api(api).get();
      setItems(response.value || []);
    } catch (err: any) {
      console.error("Errore fetch documenti", err);
      // More friendly error
      if (err.statusCode === 404) {
         setError("Cartella non trovata o accesso negato.");
      } else {
         setError("Impossibile caricare i documenti. Controlla la connessione.");
      }
    } finally {
      setLoading(false);
    }
  }, [driveId, getClient, path, siteId]);

  useEffect(() => {
    fetchItems();
  }, [fetchItems]);

  const handleOpenFolder = (name: string) => {
    setPath((prev) => (prev ? `${prev}/${name}` : name));
  };

  const handleGoUp = () => {
    if (!path) return;
    const parts = path.split("/").filter(Boolean);
    parts.pop();
    setPath(parts.join("/"));
  };

  const handleOpenFile = async (item: DriveItem) => {
    if (!isPdf(item.name, item.file?.mimeType)) {
      setError("Al momento è possibile aprire solo file PDF.");
      return;
    }
    const downloadUrl = item["@microsoft.graph.downloadUrl"];
    if (!downloadUrl) {
      setError("Link di download non disponibile.");
      return;
    }

    try {
      setError(null);
      if (blobUrlRef.current) {
        URL.revokeObjectURL(blobUrlRef.current);
        blobUrlRef.current = null;
      }
      setViewer({ name: item.name, loading: true });
      
      const resp = await fetch(downloadUrl);
      if (!resp.ok) throw new Error("Errore nel download del file");
      
      const blob = await resp.blob();
      const blobUrl = URL.createObjectURL(blob);
      blobUrlRef.current = blobUrl;
      
      setViewer({ name: item.name, blobUrl, loading: false });
    } catch (err: any) {
      console.error("Errore apertura PDF", err);
      setViewer(null);
      setError("Impossibile aprire il documento. Riprova.");
    }
  };

  const handleBreadcrumbClick = (index: number) => {
    if (index === -1) {
      setPath(initialPath || ""); 
      return;
    }
    const next = breadcrumb.slice(0, index + 1).join("/");
    setPath(next);
  };

  // Cleanup blob on unmount
  useEffect(() => {
    return () => {
      if (blobUrlRef.current) {
        URL.revokeObjectURL(blobUrlRef.current);
        blobUrlRef.current = null;
      }
    };
  }, []);

  return (
    <>
      <style>{`
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        .doc-browser-container {
          animation: fadeIn 0.4s ease-out;
        }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        
        .doc-card {
          transition: all 0.2s cubic-bezier(0.25, 0.8, 0.25, 1);
          border: 1px solid #e2e8f0;
        }
        .doc-card:hover {
          transform: translateY(-4px);
          box-shadow: 0 12px 24px -10px rgba(0, 0, 0, 0.15);
          border-color: #cbd5e1;
        }
        .doc-card:active {
          transform: scale(0.98);
        }
        .breadcrumb-item:hover {
          background: rgba(255,255,255,0.15) !important;
        }
        /* Custom Scrollbar for Grid */
        .doc-grid::-webkit-scrollbar {
          width: 6px;
        }
        .doc-grid::-webkit-scrollbar-thumb {
          background-color: #cbd5e1;
          border-radius: 4px;
        }
      `}</style>

      <div className="doc-browser-container" style={{ width: "100%", display: "flex", flexDirection: "column", gap: 24, paddingBottom: 40 }}>
        
        {/* Header Section */}
        <div style={{ 
          background: "linear-gradient(120deg, #1e293b 0%, #0f172a 100%)", 
          borderRadius: 24, 
          padding: "24px 32px",
          color: "white",
          boxShadow: "0 20px 40px -10px rgba(15, 23, 42, 0.3)",
          position: "relative",
          overflow: "hidden"
        }}>
          {/* Decorative Circle */}
          <div style={{ position: "absolute", top: -20, right: -20, width: 140, height: 140, background: "radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%)", borderRadius: "50%" }} />

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 20 }}>
            <div>
              <h2 style={{ fontSize: 28, fontWeight: 800, margin: 0, letterSpacing: "-0.03em" }}>Documentazione</h2>
              <div style={{ opacity: 0.7, marginTop: 6, fontSize: 15 }}>Esplora archivio PDF e procedure</div>
            </div>
            
            <div style={{ display: "flex", gap: 12 }}>
               {path !== (initialPath || "") && path !== "" && (
                  <button 
                  onClick={handleGoUp}
                  className="btn-reset"
                  style={{ 
                    display: "flex", 
                    alignItems: "center", 
                    gap: 8, 
                    background: "rgba(255,255,255,0.1)", 
                    border: "1px solid rgba(255,255,255,0.2)", 
                    padding: "10px 16px", 
                    borderRadius: 12,
                    color: "white",
                    cursor: "pointer",
                    fontWeight: 600,
                    fontSize: 14
                  }}
                >
                  <BackIcon /> Indietro
                </button>
               )}
               <button 
                onClick={fetchItems}
                disabled={loading}
                className="btn-reset"
                style={{ 
                  display: "flex", 
                  alignItems: "center", 
                  gap: 8, 
                  background: "rgba(255,255,255,0.1)", 
                  border: "1px solid rgba(255,255,255,0.2)", 
                  padding: "10px 16px", 
                  borderRadius: 12,
                  color: "white",
                  cursor: loading ? "wait" : "pointer",
                  fontWeight: 600,
                  fontSize: 14
                }}
              >
                <RefreshIcon spinning={loading} />
                {loading ? "Aggiornamento..." : "Aggiorna"}
              </button>
            </div>
          </div>

          {/* Breadcrumbs */}
          <div style={{ marginTop: 24, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", fontSize: 13 }}>
             <button 
                onClick={() => handleBreadcrumbClick(-1)}
                className="breadcrumb-item"
                style={{ 
                  background: "rgba(255,255,255,0.08)", 
                  border: "none", 
                  borderRadius: 8, 
                  padding: "6px 10px", 
                  color: "#e2e8f0", display: "flex", alignItems: "center", cursor: "pointer", transition: "background 0.2s" }}
             >
                <HomeIcon /> 
                <span style={{ marginLeft: 6 }}>Home</span>
             </button>
             
             {breadcrumb.map((seg, i) => (
                <React.Fragment key={i}>
                  <span style={{ opacity: 0.4 }}>/</span>
                  <button 
                    onClick={() => handleBreadcrumbClick(i)}
                    className="breadcrumb-item"
                    style={{ 
                      background: "rgba(255,255,255,0.08)", 
                      border: "none", 
                      borderRadius: 8, 
                      padding: "6px 12px", 
                      color: i === breadcrumb.length - 1 ? "#fff" : "#cbd5e1", 
                      fontWeight: i === breadcrumb.length - 1 ? 600 : 400,
                      cursor: "pointer",
                      transition: "background 0.2s"
                    }}
                  >
                    {decodeURIComponent(seg)}
                  </button>
                </React.Fragment>
             ))}
          </div>
        </div>

        {/* Dynamic Content Area */}
        <div style={{ minHeight: 400 }}>
          {error && (
            <div style={{ padding: 20, background: "#fef2f2", color: "#991b1b", borderRadius: 12, border: "1px solid #fecaca", display: "flex", alignItems: "center", gap: 12 }}>
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              {error}
            </div>
          )}

          {!loading && !error && items.length === 0 && (
            <div style={{ textAlign: "center", padding: "60px 0", opacity: 0.6 }}>
              <div style={{ fontSize: 48, marginBottom: 16 }}>📂</div>
              <div style={{ fontSize: 18, fontWeight: 500 }}>Questa cartella è vuota</div>
            </div>
          )}

          <div className="doc-grid" style={{ 
            display: "grid", 
            gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", 
            gap: 20,
            marginTop: error ? 24 : 0
          }}>
            {items.map((item) => {
               const isFolder = Boolean(item.folder);
               const isPdfFile = isPdf(item.name, item.file?.mimeType);
               
               return (
                 <button
                    key={item.id}
                    className="doc-card"
                    onClick={() => isFolder ? handleOpenFolder(item.name) : handleOpenFile(item)}
                    style={{
                      background: "white",
                      borderRadius: 20,
                      padding: 24,
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "flex-start",
                      justifyContent: "space-between",
                      cursor: "pointer",
                      height: 220, // Taller for better look
                      position: "relative",
                      textAlign: "left"
                    }}
                 >
                    <div style={{ width: "100%", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                      <div style={{ 
                        padding: 14, 
                        background: isFolder ? "#eff6ff" : isPdfFile ? "#fef2f2" : "#f1f5f9", 
                        borderRadius: 16,
                        display: "grid", placeItems: "center" 
                      }}>
                        {isFolder ? <FolderIcon /> : isPdfFile ? <PdfIcon /> : <FileIcon />}
                      </div>
                      
                      {!isFolder && (
                         <span style={{ 
                           fontSize: 11, fontWeight: 700, 
                           color: isPdfFile ? "#ef4444" : "#64748b",
                           background: isPdfFile ? "#FEF2F2" : "#F1F5F9",
                           padding: "4px 8px", borderRadius: 99
                         }}>
                           {item.name.split('.').pop()?.toUpperCase()}
                         </span>
                      )}
                    </div>

                    <div style={{ width: "100%" }}>
                       <div style={{ 
                         fontSize: 16, 
                         fontWeight: 700, 
                         color: "#1e293b", 
                         marginBottom: 6,
                         display: "-webkit-box",
                         WebkitLineClamp: 3,
                         WebkitBoxOrient: "vertical",
                         overflow: "hidden",
                         lineHeight: 1.3
                       }}>
                         {item.name}
                       </div>
                       
                       <div style={{ fontSize: 12, color: "#94a3b8", fontWeight: 500 }}>
                          {isFolder ? `${item.folder?.childCount} oggetti` : item.size ? `${(item.size / 1024).toFixed(0)} KB` : ''}
                       </div>
                    </div>
                 </button>
               )
            })}
          </div>
        </div>

      </div>

      {/* PDF Viewer Mobile/Modal Overlay */}
      {viewer && (
        <div style={{ 
          position: "fixed", inset: 0, zIndex: 9999, 
          background: "rgba(15, 23, 42, 0.9)", 
          backdropFilter: "blur(8px)",
          display: "flex", flexDirection: "column",
          animation: "fadeIn 0.2s ease-out"
        }}>
          {/* Viewer Toolbar */}
          <div style={{ 
             height: 80, 
             background: "white", 
             display: "flex", alignItems: "center", justifyContent: "space-between", 
             padding: "0 32px",
             boxShadow: "0 4px 20px rgba(0,0,0,0.1)"
          }}>
             <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
                <div style={{ padding: 10, background: "#fef2f2", borderRadius: 10 }}>
                   <PdfIcon />
                </div>
                <div style={{ display: "flex", flexDirection: "column" }}>
                   <span style={{ fontWeight: 700, color: "#0f172a", fontSize: 18 }}>{viewer.name}</span>
                   <span style={{ fontSize: 13, color: "#64748b" }}>Anteprima Documento</span>
                </div>
             </div>

             <button 
               onClick={() => setViewer(null)}
               style={{ 
                 background: "#eff6ff", 
                 border: "none", 
                 padding: "12px 24px", 
                 borderRadius: 14, 
                 color: "#1e40af", 
                 fontWeight: 700, 
                 cursor: "pointer",
                 fontSize: 16,
                 display: "flex", alignItems: "center", gap: 10
               }}
             >
               Chiudi <CloseIcon />
             </button>
          </div>

          <div style={{ flex: 1, padding: 32, display: "flex", justifyContent: "center" }}>
             {viewer.loading ? (
                <div style={{ color: "white", fontSize: 20, display: "flex", gap: 12, alignItems: "center" }}>
                   <RefreshIcon spinning={true} /> Caricamento documento in corso...
                </div>
             ) : (
                <iframe 
                  src={viewer.blobUrl} 
                  style={{ 
                    width: "100%", 
                    height: "100%", 
                    maxWidth: 1400,
                    border: "none", 
                    borderRadius: 16, 
                    boxShadow: "0 20px 60px rgba(0,0,0,0.5)", 
                    background: "white" 
                  }} 
                  title="PDF Viewer"
                />
             )}
          </div>
        </div>
      )}
    </>
  );
}
