import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useAuthenticatedGraphClient } from "../../hooks/useAuthenticatedGraphClient";

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

const videoExtensions = new Set([".mp4", ".webm", ".mov", ".m4v"]);

const encodePathSegments = (path: string) =>
  path
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");

const formatFileSize = (value?: number) => {
  if (!value || value <= 0) return "Dimensione non disponibile";
  const units = ["B", "KB", "MB", "GB"];
  let size = value;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }
  const decimals = unitIndex === 0 ? 0 : 1;
  return `${size.toFixed(decimals)} ${units[unitIndex]}`;
};

const formatDateTime = (value?: string) => {
  if (!value) return "Data non disponibile";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Data non disponibile";
  return new Intl.DateTimeFormat("it-IT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
};

const isSupportedVideo = (item: DriveItem) => {
  const mime = (item.file?.mimeType || "").toLowerCase();
  if (mime.startsWith("video/")) return true;
  const lowerName = item.name.toLowerCase();
  for (const extension of videoExtensions) {
    if (lowerName.endsWith(extension)) return true;
  }
  return false;
};

const sortByName = (left: DriveItem, right: DriveItem) =>
  left.name.localeCompare(right.name, "it", { numeric: true, sensitivity: "base" });

const FolderIcon = () => (
  <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">
    <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
  </svg>
);

const VideoIcon = () => (
  <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">
    <polygon points="23 7 16 12 23 17 23 7"></polygon>
    <rect x="1" y="5" width="15" height="14" rx="2" ry="2"></rect>
  </svg>
);

type VideoBrowserProps = {
  siteId?: string;
  driveId?: string;
  initialPath?: string;
  onExitToHome: () => void;
};

export function VideoBrowser({
  siteId,
  driveId,
  initialPath,
  onExitToHome,
}: VideoBrowserProps) {
  const getClient = useAuthenticatedGraphClient();
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const [categories, setCategories] = useState<DriveItem[]>([]);
  const [videos, setVideos] = useState<DriveItem[]>([]);
  const [selectedCategory, setSelectedCategory] = useState<DriveItem | null>(null);
  const [activeVideo, setActiveVideo] = useState<DriveItem | null>(null);
  const [loadingCategories, setLoadingCategories] = useState(false);
  const [loadingVideos, setLoadingVideos] = useState(false);
  const [openingVideoId, setOpeningVideoId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [playbackError, setPlaybackError] = useState<string | null>(null);

  const rootPath = useMemo(() => (initialPath || "").trim().replace(/^\/+|\/+$/g, ""), [initialPath]);

  const listChildren = useCallback(
    async (path: string) => {
      if (!siteId || !driveId) {
        throw new Error("Configurazione SharePoint mancante (ID sito o drive non trovati).");
      }

      const client = await getClient();
      const base = `/sites/${siteId}/drives/${driveId}`;
      const normalizedPath = path.replace(/^\/+|\/+$/g, "");
      const api = normalizedPath
        ? `${base}/root:/${encodePathSegments(normalizedPath)}:/children`
        : `${base}/root/children`;
      const response = await client.api(api).get();
      return (response.value || []) as DriveItem[];
    },
    [driveId, getClient, siteId]
  );

  const hydrateItem = useCallback(
    async (itemId: string) => {
      if (!siteId || !driveId) {
        throw new Error("Configurazione SharePoint mancante (ID sito o drive non trovati).");
      }

      const client = await getClient();
      const response = await client.api(`/sites/${siteId}/drives/${driveId}/items/${itemId}`).get();
      return response as DriveItem;
    },
    [driveId, getClient, siteId]
  );

  const loadCategories = useCallback(async () => {
    setLoadingCategories(true);
    setError(null);
    try {
      const items = await listChildren(rootPath);
      const nextCategories = items.filter((item) => item.folder).sort(sortByName);
      setCategories(nextCategories);
    } catch (err: any) {
      console.error("Errore caricamento categorie video", err);
      setError(err?.message || "Impossibile caricare le categorie video.");
    } finally {
      setLoadingCategories(false);
    }
  }, [listChildren, rootPath]);

  useEffect(() => {
    loadCategories();
  }, [loadCategories]);

  const loadVideos = useCallback(
    async (category: DriveItem) => {
      setLoadingVideos(true);
      setError(null);
      setActiveVideo(null);
      setPlaybackError(null);

      try {
        const categoryPath = rootPath ? `${rootPath}/${category.name}` : category.name;
        const items = await listChildren(categoryPath);
        const nextVideos = items.filter((item) => !item.folder && isSupportedVideo(item)).sort(sortByName);
        setSelectedCategory(category);
        setVideos(nextVideos);
      } catch (err: any) {
        console.error("Errore caricamento video categoria", err);
        setError(err?.message || "Impossibile caricare i video della categoria.");
      } finally {
        setLoadingVideos(false);
      }
    },
    [listChildren, rootPath]
  );

  const handleOpenVideo = useCallback(
    async (item: DriveItem) => {
      setOpeningVideoId(item.id);
      setError(null);
      setPlaybackError(null);

      try {
        const hydrated = item["@microsoft.graph.downloadUrl"] ? item : await hydrateItem(item.id);
        if (!hydrated["@microsoft.graph.downloadUrl"]) {
          throw new Error("Link video non disponibile.");
        }
        setActiveVideo(hydrated);
      } catch (err: any) {
        console.error("Errore apertura video", err);
        setError(err?.message || "Impossibile aprire il video selezionato.");
      } finally {
        setOpeningVideoId(null);
      }
    },
    [hydrateItem]
  );

  useEffect(() => {
    if (!activeVideo || !videoRef.current) return;

    const video = videoRef.current;
    const playPromise = video.play();
    if (playPromise && typeof playPromise.catch === "function") {
      playPromise.catch((err) => {
        console.warn("Autoplay video non riuscito", err);
        setPlaybackError("La riproduzione automatica non è partita. Premi Riprova per avviare il video.");
      });
    }
  }, [activeVideo]);

  useEffect(() => {
    if (!activeVideo) return;

    const previousBodyOverflow = document.body.style.overflow;
    const previousHtmlOverflow = document.documentElement.style.overflow;

    document.body.style.overflow = "hidden";
    document.documentElement.style.overflow = "hidden";

    return () => {
      document.body.style.overflow = previousBodyOverflow;
      document.documentElement.style.overflow = previousHtmlOverflow;
    };
  }, [activeVideo]);

  const categoryLabel = selectedCategory?.name || "Categorie";
  const hasCategories = categories.length > 0;
  const hasVideos = videos.length > 0;
  const retryPlayback = async (event: React.MouseEvent<HTMLButtonElement>) => {
    event.stopPropagation();
    if (!videoRef.current) return;
    try {
      await videoRef.current.play();
      setPlaybackError(null);
    } catch (err) {
      console.warn("Riproduzione manuale non riuscita", err);
      setPlaybackError("Impossibile avviare il video. Verifica il formato del file.");
    }
  };

  if (activeVideo?.["@microsoft.graph.downloadUrl"]) {
    return (
      <div
        onClick={onExitToHome}
        style={{
          position: "fixed",
          inset: 0,
          width: "100vw",
          height: "100vh",
          overflow: "hidden",
          background: "radial-gradient(circle at top, #1f3554 0%, #09111d 58%, #02050a 100%)",
          boxShadow: "0 28px 80px rgba(2, 6, 23, 0.45)",
          cursor: "pointer",
          zIndex: 9999,
        }}
      >
        <video
          ref={videoRef}
          src={activeVideo["@microsoft.graph.downloadUrl"]}
          autoPlay
          playsInline
          onError={() => setPlaybackError("Impossibile riprodurre il video selezionato.")}
          style={{ width: "100%", height: "100%", objectFit: "contain", background: "#000" }}
        />

        <div
          style={{
            position: "absolute",
            top: 24,
            left: 24,
            right: 24,
            display: "flex",
            justifyContent: "space-between",
            gap: 12,
            alignItems: "center",
          }}
        >
          <div
            style={{
              background: "rgba(15, 23, 42, 0.72)",
              color: "#fff",
              padding: "14px 18px",
              borderRadius: 18,
              backdropFilter: "blur(10px)",
              maxWidth: "60%",
            }}
          >
            <div style={{ fontSize: 14, opacity: 0.8, marginBottom: 4 }}>{categoryLabel}</div>
            <div style={{ fontSize: 26, fontWeight: 800, lineHeight: 1.1 }}>{activeVideo.name}</div>
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            <button
              type="button"
              onClick={(event) => {
                event.stopPropagation();
                setActiveVideo(null);
                setPlaybackError(null);
              }}
              style={overlayButtonStyle}
            >
              Indietro
            </button>
            <button
              type="button"
              onClick={(event) => {
                event.stopPropagation();
                onExitToHome();
              }}
              style={overlayButtonStyle}
            >
              Home
            </button>
          </div>
        </div>

        <div
          style={{
            position: "absolute",
            bottom: 24,
            left: "50%",
            transform: "translateX(-50%)",
            background: "rgba(15, 23, 42, 0.72)",
            color: "#fff",
            padding: "14px 20px",
            borderRadius: 999,
            fontSize: 16,
            fontWeight: 700,
            letterSpacing: "0.01em",
            backdropFilter: "blur(10px)",
          }}
        >
          Tocca lo schermo per uscire e tornare alla home
        </div>

        {playbackError && (
          <div
            style={{
              position: "absolute",
              bottom: 88,
              left: 24,
              right: 24,
              padding: "16px 18px",
              borderRadius: 18,
              background: "rgba(185, 28, 28, 0.88)",
              color: "#fff",
              fontWeight: 700,
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              gap: 16,
            }}
          >
            <span>{playbackError}</span>
            <button
              type="button"
              onClick={retryPlayback}
              style={{
                border: "1px solid rgba(255,255,255,0.32)",
                background: "rgba(255,255,255,0.14)",
                color: "#fff",
                padding: "12px 16px",
                borderRadius: 14,
                fontSize: 15,
                fontWeight: 800,
                cursor: "pointer",
                whiteSpace: "nowrap",
              }}
            >
              Riprova
            </button>
          </div>
        )}
      </div>
    );
  }

  return (
    <div style={{ width: "100%", display: "flex", flexDirection: "column", gap: 24, paddingBottom: 32 }}>
      <div
        style={{
          borderRadius: 28,
          padding: "26px 30px",
          color: "#fff",
          background: "linear-gradient(135deg, #0f172a 0%, #12345b 45%, #1d4ed8 100%)",
          boxShadow: "0 26px 70px rgba(15, 23, 42, 0.22)",
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", gap: 20, alignItems: "flex-start", flexWrap: "wrap" }}>
          <div>
            <div style={{ fontSize: 14, textTransform: "uppercase", letterSpacing: "0.08em", opacity: 0.7 }}>SharePoint / VIDEO</div>
            <h2 style={{ fontSize: 34, lineHeight: 1, margin: "10px 0 12px", fontWeight: 900 }}>
              {selectedCategory ? `Categoria: ${selectedCategory.name}` : "Raccolta Video"}
            </h2>
            <p style={{ margin: 0, fontSize: 17, maxWidth: 720, opacity: 0.92 }}>
              {selectedCategory
                ? "Scegli un video da avviare a schermo pieno. Durante la riproduzione un tocco riporta subito alla home."
                : "Le categorie vengono lette in automatico da SharePoint. Aggiungere una nuova cartella o un nuovo file lo rende subito disponibile sul totem."}
            </p>
          </div>

          <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
            {selectedCategory && (
              <button
                type="button"
                onClick={() => {
                  setSelectedCategory(null);
                  setVideos([]);
                  setActiveVideo(null);
                  setPlaybackError(null);
                  setError(null);
                }}
                style={headerButtonStyle}
              >
                Indietro
              </button>
            )}
            <button type="button" onClick={loadCategories} style={headerButtonStyle}>
              Aggiorna
            </button>
            <button type="button" onClick={onExitToHome} style={headerButtonStyle}>
              Home
            </button>
          </div>
        </div>
      </div>

      {error && (
        <div
          style={{
            borderRadius: 20,
            padding: "18px 22px",
            background: "#fff1f2",
            border: "1px solid #fecdd3",
            color: "#9f1239",
            fontSize: 16,
            fontWeight: 700,
          }}
        >
          {error}
        </div>
      )}

      {!selectedCategory ? (
        <section style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 20 }}>
          {loadingCategories && <StatusCard message="Caricamento categorie video..." />}
          {!loadingCategories && !hasCategories && (
            <StatusCard message={`Nessuna categoria trovata nella cartella ${rootPath || "VIDEO"}.`} />
          )}
          {!loadingCategories &&
            categories.map((category) => (
              <button
                key={category.id}
                type="button"
                onClick={() => loadVideos(category)}
                style={tileButtonStyle}
              >
                <div style={iconWrapStyle}>
                  <FolderIcon />
                </div>
                <div style={{ textAlign: "left" }}>
                  <div style={{ fontSize: 28, fontWeight: 900, color: "#0f172a", lineHeight: 1.05 }}>{category.name}</div>
                  <div style={{ marginTop: 10, fontSize: 15, color: "#475569" }}>
                    {category.folder?.childCount ?? 0} elementi disponibili
                  </div>
                </div>
                <div style={{ marginLeft: "auto", fontSize: 36, fontWeight: 900, color: "#2563eb", opacity: 0.5 }}>→</div>
              </button>
            ))}
        </section>
      ) : (
        <section style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 20 }}>
          {loadingVideos && <StatusCard message="Caricamento video della categoria..." />}
          {!loadingVideos && !hasVideos && (
            <StatusCard message={`Nessun video supportato presente in ${selectedCategory.name}.`} />
          )}
          {!loadingVideos &&
            videos.map((video) => (
              <button
                key={video.id}
                type="button"
                onClick={() => handleOpenVideo(video)}
                disabled={openingVideoId === video.id}
                style={{
                  ...tileButtonStyle,
                  opacity: openingVideoId && openingVideoId !== video.id ? 0.72 : 1,
                  cursor: openingVideoId ? "wait" : "pointer",
                }}
              >
                <div style={{ display: "flex", flexDirection: "column", gap: 16, width: "100%" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 16, alignItems: "center" }}>
                    <div style={{ ...iconWrapStyle, background: "linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%)", color: "#1d4ed8" }}>
                      <VideoIcon />
                    </div>
                    <div
                      style={{
                        padding: "8px 12px",
                        borderRadius: 999,
                        background: "#eff6ff",
                        color: "#1d4ed8",
                        fontWeight: 800,
                        fontSize: 13,
                        textTransform: "uppercase",
                        letterSpacing: "0.05em",
                      }}
                    >
                      {openingVideoId === video.id ? "Apertura..." : "Riproduci"}
                    </div>
                  </div>

                  <div style={{ textAlign: "left" }}>
                    <div style={{ fontSize: 24, fontWeight: 900, color: "#0f172a", lineHeight: 1.12 }}>{video.name}</div>
                    <div style={{ marginTop: 12, color: "#475569", fontSize: 15 }}>
                      Aggiornato il {formatDateTime(video.lastModifiedDateTime)}
                    </div>
                    <div style={{ marginTop: 4, color: "#64748b", fontSize: 15 }}>{formatFileSize(video.size)}</div>
                  </div>
                </div>
              </button>
            ))}
        </section>
      )}
    </div>
  );
}

function StatusCard({ message }: { message: string }) {
  return (
    <div
      style={{
        gridColumn: "1 / -1",
        minHeight: 180,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 24,
        textAlign: "center",
        borderRadius: 26,
        border: "1px dashed #93c5fd",
        background: "linear-gradient(145deg, #ffffff 0%, #eff6ff 100%)",
        color: "#1e3a8a",
        fontSize: 20,
        fontWeight: 800,
      }}
    >
      {message}
    </div>
  );
}

const headerButtonStyle: React.CSSProperties = {
  border: "1px solid rgba(255,255,255,0.22)",
  background: "rgba(255,255,255,0.12)",
  color: "#fff",
  padding: "14px 18px",
  borderRadius: 16,
  fontSize: 16,
  fontWeight: 800,
  cursor: "pointer",
  backdropFilter: "blur(10px)",
};

const overlayButtonStyle: React.CSSProperties = {
  border: "1px solid rgba(255,255,255,0.22)",
  background: "rgba(255,255,255,0.12)",
  color: "#fff",
  padding: "14px 18px",
  borderRadius: 16,
  fontSize: 16,
  fontWeight: 800,
  cursor: "pointer",
  backdropFilter: "blur(10px)",
};

const tileButtonStyle: React.CSSProperties = {
  minHeight: 220,
  width: "100%",
  display: "flex",
  alignItems: "center",
  gap: 18,
  padding: "24px 26px",
  borderRadius: 28,
  border: "1px solid #dbeafe",
  background: "linear-gradient(145deg, #ffffff 0%, #f8fbff 100%)",
  boxShadow: "0 16px 42px rgba(37, 99, 235, 0.08)",
  cursor: "pointer",
};

const iconWrapStyle: React.CSSProperties = {
  width: 72,
  height: 72,
  flexShrink: 0,
  borderRadius: 22,
  display: "grid",
  placeItems: "center",
  background: "linear-gradient(135deg, #e0f2fe 0%, #dbeafe 100%)",
  color: "#2563eb",
};
