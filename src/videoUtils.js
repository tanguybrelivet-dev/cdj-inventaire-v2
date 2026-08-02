import { supabase } from "./supabaseClient.js";

// ---------------------------------------------------------------------------
// Filigrane vidéo : on rejoue la vidéo source dans un <video> caché, on
// dessine chaque image sur un <canvas> avec le même bandeau "© SCP BRELIVET
// Tanguy" + badge que les photos, on capture le flux du canvas et on
// ré-enregistre le tout avec MediaRecorder. C'est plus lourd qu'un simple
// filtre image, mais ça permet d'avoir un vrai filigrane "brûlé" dans le
// fichier vidéo final (pas juste une légende affichée à la lecture).
//
// Le support de MediaRecorder + captureStream varie selon les navigateurs
// (notamment Safari iOS) : isVideoWatermarkSupported() permet de vérifier
// avant de se lancer, et watermarkVideo() renvoie null si le filigrane
// échoue en cours de route, pour laisser l'appelant se replier sur la
// vidéo brute plutôt que de bloquer l'utilisateur.
// ---------------------------------------------------------------------------

export function isVideoWatermarkSupported() {
  return (
    typeof window !== "undefined" &&
    typeof window.MediaRecorder !== "undefined" &&
    typeof HTMLCanvasElement !== "undefined" &&
    typeof HTMLCanvasElement.prototype.captureStream === "function"
  );
}

function pickMimeType() {
  const candidates = ["video/mp4;codecs=h264", "video/webm;codecs=vp9", "video/webm;codecs=vp8", "video/webm"];
  for (const type of candidates) {
    if (window.MediaRecorder.isTypeSupported && window.MediaRecorder.isTypeSupported(type)) return type;
  }
  return "";
}

export function watermarkVideo(file, onProgress) {
  return new Promise((resolve) => {
    if (!isVideoWatermarkSupported()) {
      resolve(null);
      return;
    }

    const objectUrl = URL.createObjectURL(file);
    const video = document.createElement("video");
    video.src = objectUrl;
    video.muted = true; // évite l'écho pendant le traitement ; l'audio original est capté séparément
    video.playsInline = true;

    video.onloadedmetadata = async () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = video.videoWidth || 1280;
        canvas.height = video.videoHeight || 720;
        const ctx = canvas.getContext("2d");

        const scale = canvas.width / 1000;
        const fontSize = Math.max(16, 24 * scale);
        const pad = Math.max(10, 16 * scale);
        const label = "© SCP BRELIVET Tanguy";

        const canvasStream = canvas.captureStream(25);

        // Récupère la piste audio d'origine si présente (via captureStream du <video>)
        let audioTracks = [];
        try {
          const srcStream = video.captureStream ? video.captureStream() : video.mozCaptureStream();
          audioTracks = srcStream.getAudioTracks();
        } catch {
          audioTracks = [];
        }
        const combined = new MediaStream([...canvasStream.getVideoTracks(), ...audioTracks]);

        const mimeType = pickMimeType();
        const recorder = new MediaRecorder(combined, mimeType ? { mimeType } : undefined);
        const chunks = [];
        recorder.ondataavailable = (e) => {
          if (e.data.size > 0) chunks.push(e.data);
        };

        const done = new Promise((res) => {
          recorder.onstop = () => res(new Blob(chunks, { type: mimeType || "video/webm" }));
        });

        let rafId;
        function drawFrame() {
          ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

          ctx.font = `bold ${fontSize}px Arial, sans-serif`;
          const textWidth = ctx.measureText(label).width;
          const barHeight = fontSize + pad * 1.2;
          ctx.fillStyle = "rgba(20,18,16,0.55)";
          ctx.fillRect(0, canvas.height - barHeight, textWidth + pad * 2, barHeight);
          ctx.fillStyle = "#FFFFFF";
          ctx.textBaseline = "middle";
          ctx.fillText(label, pad, canvas.height - barHeight / 2);

          const r = Math.max(18, 28 * scale);
          const cx = canvas.width - r - pad;
          const cy = canvas.height - r - pad;
          ctx.beginPath();
          ctx.arc(cx, cy, r, 0, Math.PI * 2);
          ctx.fillStyle = "#D9622C";
          ctx.fill();
          ctx.fillStyle = "#FFFFFF";
          ctx.font = `bold ${r * 0.8}px Arial, sans-serif`;
          ctx.textAlign = "center";
          ctx.fillText("HJ", cx, cy);
          ctx.textAlign = "left";

          if (onProgress && video.duration) onProgress(Math.min(1, video.currentTime / video.duration));

          if (!video.ended) {
            rafId = requestAnimationFrame(drawFrame);
          }
        }

        video.onended = () => {
          cancelAnimationFrame(rafId);
          recorder.stop();
        };

        recorder.start();
        await video.play();
        drawFrame();

        const blob = await done;
        URL.revokeObjectURL(objectUrl);
        resolve(blob);
      } catch {
        URL.revokeObjectURL(objectUrl);
        resolve(null);
      }
    };
    video.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(null);
    };
  });
}

export async function uploadMedia(bucket, blob, filename, defaultType) {
  if (!supabase) return { url: null, error: new Error("Supabase non configuré") };
  const path = `${Date.now()}_${filename.replace(/[^a-z0-9._-]+/gi, "_")}`;
  const { error } = await supabase.storage.from(bucket).upload(path, blob, {
    contentType: blob.type || defaultType,
    upsert: false,
  });
  if (error) return { url: null, error };
  const { data } = supabase.storage.from(bucket).getPublicUrl(path);
  return { url: data?.publicUrl || null, error: null };
}

export async function uploadVideo(blob, filename) {
  return uploadMedia("videos", blob, filename, "video/mp4");
}

export async function uploadAudio(blob, filename) {
  return uploadMedia("audios", blob, filename, "audio/webm");
}
