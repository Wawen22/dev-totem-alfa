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

type VideoCategory = DriveItem & {
  previewImageUrl?: string;
};

type VideoItem = DriveItem & {
  previewImageUrl?: string;
};

const videoExtensions = new Set([".mp4", ".webm", ".mov", ".m4v"]);
const previewFieldNeedle = "anteprima";

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

const isAbsoluteUrl = (value: string) => /^https?:\/\//i.test(value);

const joinUrl = (base?: string, relative?: string) => {
  if (!base || !relative || !relative.startsWith("/")) return undefined;
  return `${base.replace(/\/+$/g, "")}${relative}`;
};

const extractPreviewImageUrl = (value: unknown): string | undefined => {
  if (!value) return undefined;

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return undefined;
    if (isAbsoluteUrl(trimmed) || trimmed.startsWith("/")) return trimmed;
    if (
      (trimmed.startsWith("{") && trimmed.endsWith("}")) ||
      (trimmed.startsWith("[") && trimmed.endsWith("]"))
    ) {
      try {
        return extractPreviewImageUrl(JSON.parse(trimmed));
      } catch {
        return undefined;
      }
    }
    return undefined;
  }

  if (Array.isArray(value)) {
    for (const item of value) {
      const extracted = extractPreviewImageUrl(item);
      if (extracted) return extracted;
    }
    return undefined;
  }

  if (typeof value === "object") {
    const record = value as Record<string, unknown>;

    for (const key of ["thumbnailUrl", "url", "src", "imageUrl", "previewUrl", "absoluteUrl"]) {
      const extracted = extractPreviewImageUrl(record[key]);
      if (extracted) return extracted;
    }

    const serverUrl = typeof record.serverUrl === "string" ? record.serverUrl.trim() : undefined;
    for (const key of ["serverRelativeUrl", "siteRelativeUrl", "fileRef", "path"]) {
      const candidate = record[key];
      if (typeof candidate !== "string") continue;
      const trimmed = candidate.trim();
      if (!trimmed) continue;
      if (isAbsoluteUrl(trimmed)) return trimmed;
      if (trimmed.startsWith("/")) {
        return joinUrl(serverUrl, trimmed) || trimmed;
      }
    }

    for (const key of ["file", "image", "preview", "media", "value"]) {
      const extracted = extractPreviewImageUrl(record[key]);
      if (extracted) return extracted;
    }
  }

  return undefined;
};

const getPreviewFieldValue = (fields: Record<string, unknown>) => {
  const match = Object.keys(fields).find((key) => key.toLowerCase().includes(previewFieldNeedle));
  return match ? fields[match] : undefined;
};

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
  hiddenCategoryNames?: string[];
  onExitToHome: () => void;
};

export function VideoBrowser({
  siteId,
  driveId,
  initialPath,
  hiddenCategoryNames = [],
  onExitToHome,
}: VideoBrowserProps) {
  const getClient = useAuthenticatedGraphClient();
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const [categories, setCategories] = useState<VideoCategory[]>([]);
  const [videos, setVideos] = useState<VideoItem[]>([]);
  const [selectedCategory, setSelectedCategory] = useState<VideoCategory | null>(null);
  const [activeVideo, setActiveVideo] = useState<DriveItem | null>(null);
  const [loadingCategories, setLoadingCategories] = useState(false);
  const [loadingVideos, setLoadingVideos] = useState(false);
  const [openingVideoId, setOpeningVideoId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [playbackError, setPlaybackError] = useState<string | null>(null);

  const rootPath = useMemo(() => (initialPath || "").trim().replace(/^\/+|\/+$/g, ""), [initialPath]);
  const hiddenCategoryNamesSet = useMemo(
    () => new Set(hiddenCategoryNames.map((name) => name.trim().toLowerCase()).filter(Boolean)),
    [hiddenCategoryNames]
  );

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

  const fetchItemPreview = useCallback(
    async (itemId: string) => {
      if (!siteId || !driveId) return undefined;

      try {
        const client = await getClient();
        const fields = (await client
          .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}/listItem/fields`)
          .get()) as Record<string, unknown>;

        return extractPreviewImageUrl(getPreviewFieldValue(fields));
      } catch (err) {
        console.warn("Anteprima elemento non disponibile", err);
        return undefined;
      }
    },
    [driveId, getClient, siteId]
  );

  const loadCategories = useCallback(async () => {
    setLoadingCategories(true);
    setError(null);
    try {
      const items = await listChildren(rootPath);
      const folderItems = items
        .filter((item) => item.folder && !hiddenCategoryNamesSet.has(item.name.trim().toLowerCase()))
        .sort(sortByName);
      const nextCategories = await Promise.all(
        folderItems.map(async (item) => ({
          ...item,
          previewImageUrl: await fetchItemPreview(item.id),
        }))
      );
      setCategories(nextCategories);
    } catch (err: any) {
      console.error("Errore caricamento categorie video", err);
      setError(err?.message || "Impossibile caricare le categorie video.");
    } finally {
      setLoadingCategories(false);
    }
  }, [fetchItemPreview, hiddenCategoryNamesSet, listChildren, rootPath]);

  useEffect(() => {
    loadCategories();
  }, [loadCategories]);

  const loadVideos = useCallback(
    async (category: VideoCategory) => {
      setLoadingVideos(true);
      setError(null);
      setActiveVideo(null);
      setPlaybackError(null);

      try {
        const categoryPath = rootPath ? `${rootPath}/${category.name}` : category.name;
        const items = await listChildren(categoryPath);
        const supportedVideos = items.filter((item) => !item.folder && isSupportedVideo(item)).sort(sortByName);
        const nextVideos = await Promise.all(
          supportedVideos.map(async (item) => ({
            ...item,
            previewImageUrl: await fetchItemPreview(item.id),
          }))
        );
        setSelectedCategory(category);
        setVideos(nextVideos);
      } catch (err: any) {
        console.error("Errore caricamento video categoria", err);
        setError(err?.message || "Impossibile caricare i video della categoria.");
      } finally {
        setLoadingVideos(false);
      }
    },
    [fetchItemPreview, listChildren, rootPath]
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
        style={{
          position: "fixed",
          inset: 0,
          width: "100vw",
          height: "100vh",
          overflow: "hidden",
          background: "radial-gradient(circle at top, #1f3554 0%, #09111d 58%, #02050a 100%)",
          boxShadow: "0 28px 80px rgba(2, 6, 23, 0.45)",
          cursor: "default",
          zIndex: 9999,
        }}
      >
        <video
          ref={videoRef}
          src={activeVideo["@microsoft.graph.downloadUrl"]}
          autoPlay
          controls
          controlsList="nodownload"
          playsInline
          onClick={(event) => event.stopPropagation()}
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
          position: "relative",
          overflow: "hidden",
          borderRadius: 28,
          padding: "26px 30px",
          color: "#fff",
          background: "linear-gradient(135deg, #0f172a 0%, #12345b 45%, #1d4ed8 100%)",
          boxShadow: "0 26px 70px rgba(15, 23, 42, 0.22)",
        }}
      >
        {selectedCategory?.previewImageUrl && (
          <>
            <div
              style={{
                position: "absolute",
                inset: 0,
                backgroundImage: `url("${selectedCategory.previewImageUrl}")`,
                backgroundPosition: "center",
                backgroundSize: "cover",
                backgroundRepeat: "no-repeat",
                transform: "scale(1.04)",
              }}
            />
            <div
              style={{
                position: "absolute",
                inset: 0,
                background:
                  "linear-gradient(110deg, rgba(2, 6, 23, 0.82) 0%, rgba(15, 23, 42, 0.60) 38%, rgba(29, 78, 216, 0.48) 100%)",
              }}
            />
          </>
        )}

        <div
          style={{
            position: "relative",
            zIndex: 1,
            display: "flex",
            justifyContent: "space-between",
            gap: 20,
            alignItems: "flex-start",
            flexWrap: "wrap",
          }}
        >
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
        <section style={{ display: "flex", flexDirection: "column", gap: 18 }}>
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
                style={categoryButtonStyle}
              >
                <div style={categoryPreviewStyle}>
                  {category.previewImageUrl && (
                    <div
                      style={{
                        position: "absolute",
                        inset: 0,
                        backgroundImage: `url("${category.previewImageUrl}")`,
                        backgroundPosition: "center",
                        backgroundSize: "cover",
                        backgroundRepeat: "no-repeat",
                        transform: "scale(1.02)",
                      }}
                    />
                  )}
                  <div
                    style={{
                      position: "absolute",
                      inset: 0,
                      background: category.previewImageUrl
                        ? "linear-gradient(135deg, rgba(15, 23, 42, 0.14) 0%, rgba(15, 23, 42, 0.68) 100%)"
                        : "linear-gradient(135deg, rgba(255,255,255,0.16) 0%, rgba(255,255,255,0.04) 100%)",
                    }}
                  />
                  <div
                    style={{
                      position: "relative",
                      zIndex: 1,
                      height: "100%",
                      padding: 22,
                      display: "flex",
                      flexDirection: "column",
                      justifyContent: "space-between",
                      gap: 16,
                      color: "#fff",
                    }}
                  >
                    <div
                      style={{
                        alignSelf: "flex-start",
                        padding: "8px 12px",
                        borderRadius: 999,
                        background: "rgba(255,255,255,0.16)",
                        border: "1px solid rgba(255,255,255,0.18)",
                        backdropFilter: "blur(8px)",
                        fontSize: 12,
                        fontWeight: 800,
                        letterSpacing: "0.08em",
                        textTransform: "uppercase",
                      }}
                    >
                      {category.previewImageUrl ? "Anteprima SharePoint" : "Raccolta Video"}
                    </div>

                    <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                      <div
                        style={{
                          ...iconWrapStyle,
                          width: 64,
                          height: 64,
                          borderRadius: 20,
                          background: "rgba(255,255,255,0.15)",
                          color: "#fff",
                          backdropFilter: "blur(10px)",
                        }}
                      >
                        <FolderIcon />
                      </div>
                      <div style={{ textAlign: "left" }}>
                        <div style={{ fontSize: 12, textTransform: "uppercase", letterSpacing: "0.08em", opacity: 0.82 }}>
                          Categoria
                        </div>
                        <div style={{ marginTop: 6, fontSize: 18, fontWeight: 800, lineHeight: 1.12 }}>
                          {category.previewImageUrl ? "Visuale personalizzata attiva" : "Fallback grafico automatico"}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div
                  style={{
                    flex: "1 1 420px",
                    minWidth: 0,
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 18,
                    padding: "28px 30px",
                    flexWrap: "wrap",
                  }}
                >
                  <div style={{ textAlign: "left", minWidth: 0, flex: "1 1 360px" }}>
                    <div
                      style={{
                        fontSize: 13,
                        fontWeight: 800,
                        textTransform: "uppercase",
                        letterSpacing: "0.08em",
                        color: "#2563eb",
                      }}
                    >
                      Raccolta Video
                    </div>
                    <div
                      style={{
                        marginTop: 10,
                        fontSize: "clamp(28px, 3vw, 40px)",
                        fontWeight: 900,
                        color: "#0f172a",
                        lineHeight: 1.02,
                        wordBreak: "break-word",
                      }}
                    >
                      {category.name}
                    </div>
                    <div style={{ marginTop: 12, fontSize: 16, color: "#475569", maxWidth: 760 }}>
                      Apri la categoria per visualizzare i video associati e avviare subito la riproduzione fullscreen sul totem.
                    </div>
                    <div style={{ marginTop: 18, display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <span style={metaChipStyle}>
                        {category.folder?.childCount ?? 0} elementi disponibili
                      </span>
                    </div>
                  </div>

                  <div style={{ display: "flex", alignItems: "center", gap: 14, marginLeft: "auto" }}>
                    <div
                      style={{
                        padding: "12px 16px",
                        borderRadius: 999,
                        background: "#eff6ff",
                        color: "#1d4ed8",
                        fontSize: 14,
                        fontWeight: 800,
                        letterSpacing: "0.04em",
                        textTransform: "uppercase",
                        whiteSpace: "nowrap",
                      }}
                    >
                      Apri categoria
                    </div>
                    <div
                      style={{
                        width: 56,
                        height: 56,
                        borderRadius: 18,
                        display: "grid",
                        placeItems: "center",
                        background: "linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)",
                        color: "#fff",
                        fontSize: 28,
                        fontWeight: 900,
                        boxShadow: "0 16px 28px rgba(37, 99, 235, 0.24)",
                        flexShrink: 0,
                      }}
                    >
                      →
                    </div>
                  </div>
                </div>
              </button>
            ))}
        </section>
      ) : (
        <section style={{ display: "flex", flexDirection: "column", gap: 18 }}>
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
                  ...videoCardStyle,
                  opacity: openingVideoId && openingVideoId !== video.id ? 0.72 : 1,
                  cursor: openingVideoId ? "wait" : "pointer",
                }}
              >
                <div style={videoCoverStyle}>
                  {video.previewImageUrl && (
                    <div
                      style={{
                        position: "absolute",
                        inset: 0,
                        backgroundImage: `url("${video.previewImageUrl}")`,
                        backgroundPosition: "center",
                        backgroundSize: "cover",
                        backgroundRepeat: "no-repeat",
                        transform: "scale(1.02)",
                      }}
                    />
                  )}
                  <div
                    style={{
                      position: "absolute",
                      inset: 0,
                      background: video.previewImageUrl
                        ? "linear-gradient(135deg, rgba(15, 23, 42, 0.14) 0%, rgba(15, 23, 42, 0.68) 100%)"
                        : "linear-gradient(135deg, rgba(255,255,255,0.16) 0%, rgba(255,255,255,0.04) 100%)",
                    }}
                  />
                  <div
                    style={{
                      position: "relative",
                      zIndex: 1,
                      height: "100%",
                      padding: 22,
                      display: "flex",
                      flexDirection: "column",
                      justifyContent: "space-between",
                      gap: 16,
                      color: "#fff",
                    }}
                  >
                    <div
                      style={{
                        alignSelf: "flex-start",
                        padding: "8px 12px",
                        borderRadius: 999,
                        background: "rgba(255,255,255,0.16)",
                        border: "1px solid rgba(255,255,255,0.18)",
                        backdropFilter: "blur(8px)",
                        fontSize: 12,
                        fontWeight: 800,
                        letterSpacing: "0.08em",
                        textTransform: "uppercase",
                      }}
                    >
                      {video.previewImageUrl ? "Anteprima SharePoint" : "Raccolta Video"}
                    </div>

                    <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                      <div
                        style={{
                          ...iconWrapStyle,
                          width: 64,
                          height: 64,
                          borderRadius: 20,
                          background: "rgba(255,255,255,0.15)",
                          color: "#fff",
                          backdropFilter: "blur(10px)",
                        }}
                      >
                        <VideoIcon />
                      </div>
                      <div style={{ textAlign: "left" }}>
                        <div style={{ fontSize: 12, textTransform: "uppercase", letterSpacing: "0.08em", opacity: 0.82 }}>
                          Video
                        </div>
                        <div style={{ marginTop: 6, fontSize: 18, fontWeight: 800, lineHeight: 1.12 }}>
                          {video.previewImageUrl ? "Anteprima personalizzata attiva" : "Fallback grafico automatico"}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div style={videoCardContentStyle}>
                  <div style={{ textAlign: "left", minWidth: 0, flex: "1 1 360px" }}>
                    <div
                      style={{
                        fontSize: 13,
                        fontWeight: 800,
                        textTransform: "uppercase",
                        letterSpacing: "0.08em",
                        color: "#2563eb",
                      }}
                    >
                      Video
                    </div>
                    <div
                      style={{
                        marginTop: 10,
                        fontSize: "clamp(28px, 3vw, 40px)",
                        fontWeight: 900,
                        color: "#0f172a",
                        lineHeight: 1.02,
                        wordBreak: "break-word",
                      }}
                    >
                      {video.name}
                    </div>
                    <div style={{ marginTop: 12, color: "#475569", fontSize: 16 }}>
                      Aggiornato il {formatDateTime(video.lastModifiedDateTime)}
                    </div>
                    <div style={{ marginTop: 18, display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <span style={metaChipStyle}>{formatFileSize(video.size)}</span>
                    </div>
                  </div>

                  <div style={{ display: "flex", alignItems: "center", gap: 14, marginLeft: "auto" }}>
                    <div
                      style={{
                        padding: "12px 16px",
                        borderRadius: 999,
                        background: "#eff6ff",
                        color: "#1d4ed8",
                        fontSize: 14,
                        fontWeight: 800,
                        letterSpacing: "0.04em",
                        textTransform: "uppercase",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {openingVideoId === video.id ? "Apertura..." : "Riproduci"}
                    </div>
                    <div
                      style={{
                        width: 56,
                        height: 56,
                        borderRadius: 18,
                        display: "grid",
                        placeItems: "center",
                        background: "linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)",
                        color: "#fff",
                        fontSize: 22,
                        fontWeight: 900,
                        boxShadow: "0 16px 28px rgba(37, 99, 235, 0.24)",
                        flexShrink: 0,
                      }}
                    >
                      ▶
                    </div>
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

const videoCardStyle: React.CSSProperties = {
  width: "100%",
  display: "flex",
  alignItems: "stretch",
  gap: 0,
  padding: 0,
  borderRadius: 32,
  overflow: "hidden",
  border: "1px solid rgba(191, 219, 254, 0.95)",
  background: "linear-gradient(145deg, #ffffff 0%, #f8fbff 100%)",
  boxShadow: "0 22px 52px rgba(37, 99, 235, 0.10)",
  cursor: "pointer",
  flexWrap: "wrap",
};

const videoCoverStyle: React.CSSProperties = {
  position: "relative",
  flex: "0 0 clamp(260px, 28vw, 360px)",
  minHeight: 220,
  background: "linear-gradient(135deg, #1e3a8a 0%, #2563eb 55%, #38bdf8 100%)",
  overflow: "hidden",
};

const videoCardContentStyle: React.CSSProperties = {
  flex: "1 1 420px",
  minWidth: 0,
  display: "flex",
  justifyContent: "space-between",
  alignItems: "center",
  gap: 18,
  padding: "28px 30px",
  flexWrap: "wrap",
};

const categoryButtonStyle: React.CSSProperties = {
  width: "100%",
  display: "flex",
  alignItems: "stretch",
  gap: 0,
  padding: 0,
  borderRadius: 32,
  overflow: "hidden",
  border: "1px solid rgba(191, 219, 254, 0.95)",
  background: "linear-gradient(145deg, #ffffff 0%, #f8fbff 100%)",
  boxShadow: "0 22px 52px rgba(37, 99, 235, 0.10)",
  cursor: "pointer",
  flexWrap: "wrap",
};

const categoryPreviewStyle: React.CSSProperties = {
  position: "relative",
  flex: "0 0 clamp(260px, 28vw, 360px)",
  minHeight: 220,
  background: "linear-gradient(135deg, #1e3a8a 0%, #2563eb 55%, #38bdf8 100%)",
  overflow: "hidden",
};

const metaChipStyle: React.CSSProperties = {
  display: "inline-flex",
  alignItems: "center",
  padding: "8px 12px",
  borderRadius: 999,
  background: "#eff6ff",
  color: "#1d4ed8",
  fontSize: 14,
  fontWeight: 800,
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
