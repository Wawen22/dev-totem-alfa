import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useAuthenticatedGraphClient } from "../../hooks/useAuthenticatedGraphClient";

type DriveItem = {
  id: string;
  name: string;
  webUrl?: string;
  folder?: { childCount: number };
  file?: { mimeType?: string };
  lastModifiedDateTime?: string;
  "@microsoft.graph.downloadUrl"?: string;
};

type IdleScreensaverProps = {
  siteId?: string;
  driveId?: string;
  rootPath?: string;
  categoryName?: string;
  idleTimeoutMs?: number;
  enabled?: boolean;
  onWakeUp?: () => void;
};

const videoExtensions = new Set([".mp4", ".webm", ".mov", ".m4v"]);

const encodePathSegments = (path: string) =>
  path
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");

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

const hasActiveForegroundVideo = (screensaverVideo: HTMLVideoElement | null) =>
  Array.from(document.querySelectorAll("video")).some((video) => {
    if (video === screensaverVideo) return false;
    return !video.paused && !video.ended && video.readyState > 2;
  });

export function IdleScreensaver({
  siteId,
  driveId,
  rootPath,
  categoryName = "Screensaver",
  idleTimeoutMs = 5 * 60 * 1000,
  enabled = true,
  onWakeUp,
}: IdleScreensaverProps) {
  const getClient = useAuthenticatedGraphClient();
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const idleTimerRef = useRef<number | null>(null);
  const [playlist, setPlaylist] = useState<DriveItem[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);

  const isActive = playlist.length > 0;
  const activeVideo = playlist[currentIndex] ?? null;

  const normalizedRootPath = useMemo(
    () => (rootPath || "").trim().replace(/^\/+|\/+$/g, ""),
    [rootPath]
  );

  const clearIdleTimer = useCallback(() => {
    if (idleTimerRef.current !== null) {
      window.clearTimeout(idleTimerRef.current);
      idleTimerRef.current = null;
    }
  }, []);

  const listChildren = useCallback(
    async (path: string) => {
      if (!siteId || !driveId) {
        throw new Error("Configurazione SharePoint video mancante.");
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
        throw new Error("Configurazione SharePoint video mancante.");
      }

      const client = await getClient();
      const response = await client.api(`/sites/${siteId}/drives/${driveId}/items/${itemId}`).get();
      return response as DriveItem;
    },
    [driveId, getClient, siteId]
  );

  const loadScreensaverPlaylist = useCallback(async () => {
    const screensaverPath = normalizedRootPath
      ? `${normalizedRootPath}/${categoryName}`
      : categoryName;
    const items = await listChildren(screensaverPath);
    const candidates = items.filter((item) => !item.folder && isSupportedVideo(item)).sort(sortByName);

    if (candidates.length === 0) {
      throw new Error(`Nessun video supportato trovato nella cartella ${screensaverPath}.`);
    }

    const settled = await Promise.allSettled(candidates.map((item) => hydrateItem(item.id)));
    const playable = settled
      .filter((result): result is PromiseFulfilledResult<DriveItem> => result.status === "fulfilled")
      .map((result) => result.value)
      .filter((item) => item["@microsoft.graph.downloadUrl"]);

    if (playable.length === 0) {
      throw new Error("I video screensaver non espongono un downloadUrl.");
    }

    return playable;
  }, [categoryName, hydrateItem, listChildren, normalizedRootPath]);

  const wakeUp = useCallback(() => {
    clearIdleTimer();
    setPlaylist([]);
    setCurrentIndex(0);
    onWakeUp?.();
  }, [clearIdleTimer, onWakeUp]);

  const armIdleTimer = useCallback(() => {
    clearIdleTimer();

    if (!enabled || isActive) return;

    idleTimerRef.current = window.setTimeout(async () => {
      if (document.hidden) {
        armIdleTimer();
        return;
      }

      if (hasActiveForegroundVideo(videoRef.current)) {
        armIdleTimer();
        return;
      }

      try {
        const nextPlaylist = await loadScreensaverPlaylist();
        setPlaylist(nextPlaylist);
        setCurrentIndex(0);
      } catch (err) {
        console.warn("Screensaver non avviato", err);
        armIdleTimer();
      }
    }, idleTimeoutMs);
  }, [clearIdleTimer, enabled, idleTimeoutMs, isActive, loadScreensaverPlaylist]);

  useEffect(() => {
    if (!enabled) {
      clearIdleTimer();
      setPlaylist([]);
      setCurrentIndex(0);
      return;
    }

    const resetEvents = ["pointerdown", "pointermove", "touchstart", "keydown", "wheel"] as const;
    const wakeEvents = new Set<string>(["pointerdown", "touchstart", "keydown", "wheel"]);

    const handleActivity = (event: Event) => {
      if (isActive) {
        if (wakeEvents.has(event.type)) {
          wakeUp();
        }
        return;
      }

      armIdleTimer();
    };

    armIdleTimer();
    resetEvents.forEach((eventName) => {
      window.addEventListener(eventName, handleActivity, { passive: true });
    });

    return () => {
      resetEvents.forEach((eventName) => {
        window.removeEventListener(eventName, handleActivity);
      });
      clearIdleTimer();
    };
  }, [armIdleTimer, clearIdleTimer, enabled, isActive, wakeUp]);

  useEffect(() => {
    if (!activeVideo || !videoRef.current) return;

    const playPromise = videoRef.current.play();
    if (playPromise && typeof playPromise.catch === "function") {
      playPromise.catch((err) => {
        console.warn("Autoplay screensaver non riuscito", err);
      });
    }
  }, [activeVideo]);

  useEffect(() => {
    return () => {
      clearIdleTimer();
    };
  }, [clearIdleTimer]);

  if (!activeVideo?.["@microsoft.graph.downloadUrl"]) {
    return null;
  }

  const isSingleVideo = playlist.length <= 1;

  return (
    <div
      onClick={wakeUp}
      style={{
        position: "fixed",
        inset: 0,
        width: "100vw",
        height: "100vh",
        overflow: "hidden",
        background: "#020617",
        cursor: "pointer",
        zIndex: 10000,
      }}
    >
      <video
        key={activeVideo.id}
        ref={videoRef}
        src={activeVideo["@microsoft.graph.downloadUrl"]}
        autoPlay
        muted
        loop={isSingleVideo}
        playsInline
        onEnded={() => {
          if (isSingleVideo) return;
          setCurrentIndex((prev) => (prev + 1) % playlist.length);
        }}
        style={{
          width: "100%",
          height: "100%",
          objectFit: "contain",
          background: "#000",
        }}
      />

      <div
        style={{
          position: "absolute",
          inset: 0,
          background:
            "linear-gradient(180deg, rgba(2, 6, 23, 0.20) 0%, rgba(2, 6, 23, 0.42) 100%)",
        }}
      />

      <div
        style={{
          position: "absolute",
          left: 24,
          right: 24,
          bottom: 24,
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          gap: 16,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            padding: "16px 20px",
            borderRadius: 18,
            background: "rgba(15, 23, 42, 0.72)",
            backdropFilter: "blur(10px)",
            color: "#fff",
            maxWidth: 720,
          }}
        >
          <div style={{ fontSize: 14, opacity: 0.75, textTransform: "uppercase", letterSpacing: "0.08em" }}>
            Screensaver{playlist.length > 1 ? ` · ${currentIndex + 1}/${playlist.length}` : ""}
          </div>
          <div style={{ marginTop: 8, fontSize: 28, lineHeight: 1.05, fontWeight: 900 }}>
            {activeVideo.name}
          </div>
        </div>

        <div
          style={{
            padding: "14px 18px",
            borderRadius: 999,
            background: "rgba(15, 23, 42, 0.72)",
            backdropFilter: "blur(10px)",
            color: "#fff",
            fontSize: 16,
            fontWeight: 800,
            whiteSpace: "nowrap",
          }}
        >
          Tocca lo schermo per tornare alla home
        </div>
      </div>
    </div>
  );
}
