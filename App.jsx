import { useState, useMemo, useRef, useEffect } from "react";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ImageRun,
  AlignmentType,
  HeadingLevel,
  ShadingType,
  VerticalAlign,
  PageBreak,
} from "docx";
import { saveDossier, listDossiers, loadDossier } from "./dossiers.js";
import { supabaseConfigured } from "./supabaseClient.js";
import {
  Camera,
  Plus,
  Trash2,
  Stamp,
  MapPin,
  Clock,
  X,
  Check,
  Pencil,
  Settings2,
  ImageOff,
  FilePlus2,
  Building2,
  Gavel,
  ChevronRight,
  Search,
  Sparkles,
  FileDown,
  Mic,
  ClipboardList,
  TrafficCone,
  Home,
  Lock,
  ArrowLeft,
  FileQuestion,
  Cloud,
  LayoutDashboard,
} from "lucide-react";

// ---------------------------------------------------------------------------
// Design tokens — navy + huissier orange, pulled from the real letterhead
// ---------------------------------------------------------------------------
const INK = "#1A1A18";
const NAVY = "#1F3864";
const RUST = "#D9622C";
const PAPER = "#F6F3EC";
const PAPER_DEEP = "#EDE8DB";
const LINE = "#DCD5C4";
const MUTED = "#8A8371";

const SEED_CATEGORIES = [
  { key: "materiel", label: "Matériel d'exploitation" },
  { key: "location", label: "Location" },
  { key: "incorporel", label: "Biens incorporels" },
];

const SEED_ZONES = ["SALLE DE BAR", "CUISINE", "PLONGE", "RESERVE", "EXTERIEUR"];

const SEED_LOTS = [
  { id: 1, zone: "SALLE DE BAR", cat: "incorporel", desig: "UNE LICENCE IV", value: 8000, photos: ["seed"] },
  { id: 2, zone: "SALLE DE BAR", cat: "location", desig: "UNE BOX APPARTENANT A CANAL PLUS", value: null, photos: [] },
  { id: 3, zone: "CUISINE", cat: "materiel", desig: "FOUR A CONVECTION AT211-MDI", value: 200, photos: ["seed", "seed"] },
];

// ---------------------------------------------------------------------------
// Menu: choix du type de constat
// ---------------------------------------------------------------------------
const ACT_TYPES = [
  {
    key: "inventaire_lj",
    label: "Inventaire (liquidation judiciaire)",
    desc: "Inventaire prisatif de l'actif mobilier et incorporel",
    icon: ClipboardList,
    available: true,
  },
  {
    key: "affichage",
    label: "Constat d'affichage",
    desc: "Constat de présence ou d'absence d'un affichage",
    icon: Camera,
    available: false,
  },
  {
    key: "voirie",
    label: "Constat de voirie",
    desc: "État de la voie publique, dégradations, stationnement",
    icon: TrafficCone,
    available: false,
  },
  {
    key: "etat_lieux",
    label: "État des lieux",
    desc: "Entrée, sortie ou état descriptif d'un local",
    icon: Home,
    available: false,
  },
  {
    key: "autre",
    label: "Autre constat",
    desc: "Tout autre type de constat non répertorié",
    icon: FileQuestion,
    available: false,
  },
];

function MenuScreen({ onSelect, onOpenDashboard }) {
  return (
    <div className="flex h-full flex-col">
      <div className="shrink-0 px-5 pb-5 pt-6" style={{ background: NAVY, color: PAPER }}>
        <div className="flex h-11 w-11 items-center justify-center rounded-full" style={{ background: RUST }}>
          <Stamp size={20} color={PAPER} strokeWidth={2} />
        </div>
        <h1 className="mt-3 text-[19px] font-semibold leading-tight">Nouveau constat</h1>
        <p className="mt-1 text-[12px]" style={{ color: "#C3CEDE" }}>
          Choisissez le type d'acte à établir
        </p>
      </div>
      <div className="flex-1 overflow-y-auto px-5 py-4">
        {ACT_TYPES.map((t) => {
          const Icon = t.icon;
          return (
            <button
              key={t.key}
              type="button"
              onClick={() => t.available && onSelect(t.key)}
              disabled={!t.available}
              className="mb-3 flex w-full items-center gap-3 rounded-2xl border px-4 py-3.5 text-left transition-colors disabled:opacity-45"
              style={{ borderColor: t.available ? LINE : LINE, background: "white" }}
            >
              <div
                className="flex h-11 w-11 shrink-0 items-center justify-center rounded-full"
                style={{ background: t.available ? "rgba(217,98,44,0.1)" : PAPER_DEEP, color: t.available ? RUST : MUTED }}
              >
                <Icon size={19} />
              </div>
              <div className="min-w-0 flex-1">
                <p className="text-[13.5px] font-semibold" style={{ color: INK }}>{t.label}</p>
                <p className="truncate text-[11px]" style={{ color: MUTED }}>{t.desc}</p>
              </div>
              {t.available ? (
                <ChevronRight size={18} style={{ color: MUTED }} />
              ) : (
                <span
                  className="flex shrink-0 items-center gap-1 rounded-full px-2 py-1 font-mono text-[9px] font-semibold uppercase tracking-wide"
                  style={{ background: PAPER_DEEP, color: MUTED }}
                >
                  <Lock size={9} />
                  Bientôt
                </span>
              )}
            </button>
          );
        })}
        <p className="mt-2 px-1 text-[11px]" style={{ color: MUTED }}>
          Les autres types de constat arrivent prochainement — dites-nous lesquels prioriser.
        </p>
        <button
          type="button"
          onClick={onOpenDashboard}
          className="mt-4 flex w-full items-center justify-center gap-2 rounded-2xl border-2 border-dashed py-3 text-[13px] font-semibold"
          style={{ borderColor: NAVY, color: NAVY }}
        >
          <LayoutDashboard size={16} />
          Tableau de bord des dossiers
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Tableau de bord — liste des dossiers enregistrés dans le cloud (Supabase)
// ---------------------------------------------------------------------------
function Dashboard({ onOpenDossier, onBack, embedded }) {
  const [dossiers, setDossiers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    let cancelled = false;
    async function load() {
      setLoading(true);
      const { data, error } = await listDossiers();
      if (!cancelled) {
        if (error) setError(error.message);
        setDossiers(data);
        setLoading(false);
      }
    }
    load();
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <div className="flex h-full flex-col" style={{ background: PAPER }}>
      <div className="shrink-0 px-5 pb-5 pt-6" style={{ background: NAVY, color: PAPER }}>
        {!embedded && (
          <button onClick={onBack} className="mb-2 flex items-center gap-1 text-[11px] font-semibold" style={{ color: "#C3CEDE" }}>
            <ArrowLeft size={13} />
            Retour
          </button>
        )}
        <div className="flex items-center gap-2">
          <LayoutDashboard size={20} />
          <h1 className="text-[19px] font-semibold leading-tight">Tableau de bord</h1>
        </div>
        <p className="mt-1 text-[12px]" style={{ color: "#C3CEDE" }}>
          {dossiers.length} dossier{dossiers.length > 1 ? "s" : ""} enregistré{dossiers.length > 1 ? "s" : ""}
        </p>
      </div>

      <div className="flex-1 overflow-y-auto px-5 py-4">
        {!supabaseConfigured && (
          <div className="rounded-2xl border px-4 py-3.5" style={{ borderColor: LINE, background: "white" }}>
            <p className="text-[13px]" style={{ color: INK }}>
              Supabase n'est pas encore configuré. Complétez <code>src/supabaseClient.js</code> avec l'URL et la
              clé de votre projet pour activer le tableau de bord (voir README.md).
            </p>
          </div>
        )}
        {supabaseConfigured && loading && (
          <p className="text-center text-[13px]" style={{ color: MUTED }}>Chargement…</p>
        )}
        {supabaseConfigured && error && (
          <p className="rounded-xl px-3 py-2 text-[12px]" style={{ background: "rgba(181,48,27,0.08)", color: "#B5301B" }}>{error}</p>
        )}
        {supabaseConfigured && !loading && !error && dossiers.length === 0 && (
          <p className="text-center text-[13px]" style={{ color: MUTED }}>
            Aucun dossier enregistré pour l'instant.
          </p>
        )}
        {dossiers.map((d) => (
          <button
            key={d.id}
            onClick={() => onOpenDossier(d.id)}
            className="mb-3 flex w-full items-center gap-3 rounded-2xl border px-4 py-3.5 text-left"
            style={{ borderColor: LINE, background: "white" }}
          >
            <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-full" style={{ background: "rgba(31,56,100,0.08)", color: NAVY }}>
              <ClipboardList size={18} />
            </div>
            <div className="min-w-0 flex-1">
              <p className="truncate text-[13.5px] font-semibold" style={{ color: INK }}>{d.name}</p>
              <p className="truncate text-[11px]" style={{ color: MUTED }}>{d.address}</p>
              <p className="mt-0.5 text-[10.5px]" style={{ color: MUTED }}>
                {d.lot_count} lot{d.lot_count > 1 ? "s" : ""} · {euros(d.total)} · maj {new Date(d.updated_at).toLocaleDateString("fr-FR")}
              </p>
            </div>
            <ChevronRight size={18} style={{ color: MUTED }} />
          </button>
        ))}
      </div>
    </div>
  );
}

const SEED_REQUERANTS = [
  {
    id: "r1",
    denomination: "SELARL EP & ASSOCIÉS",
    capital: "200 000",
    rcsVille: "BREST",
    rcsNumero: "808 072 821",
    adresse: "62A Quai de l'Odet, 29000 Quimper",
    nom: "Jordi PAGANI",
    qualite: "Gérant",
  },
];

const emptyRequerant = {
  denomination: "",
  capital: "",
  rcsVille: "",
  rcsNumero: "",
  adresse: "",
  nom: "",
  qualite: "Gérant",
};

// ---------------------------------------------------------------------------
// Simulated Pappers lookup — in production this calls api.pappers.fr/v2/recherche
// through a backend relay (the API key must never ship inside the mobile app).
// ---------------------------------------------------------------------------
const MOCK_PAPPERS = [
  {
    denomination: "SELARL EP & ASSOCIÉS",
    capital: "200 000",
    rcsVille: "BREST",
    rcsNumero: "808 072 821",
    adresse: "62A Quai de l'Odet, 29000 Quimper",
    nom: "Jordi PAGANI",
    qualite: "Gérant",
  },
  {
    denomination: "SELARL AVEIRO & PARTENAIRES",
    capital: "150 000",
    rcsVille: "RENNES",
    rcsNumero: "512 344 998",
    adresse: "4 Rue de la Monnaie, 35000 Rennes",
    nom: "Camille AVEIRO",
    qualite: "Gérante",
  },
  {
    denomination: "SCP LERAY MANDATAIRES JUDICIAIRES",
    capital: "80 000",
    rcsVille: "NANTES",
    rcsNumero: "398 221 764",
    adresse: "12 Cours des Cinquante Otages, 44000 Nantes",
    nom: "Thibault LERAY",
    qualite: "Associé gérant",
  },
  {
    denomination: "SELARL BODIN JUDICIAIRE",
    capital: "100 000",
    rcsVille: "QUIMPER",
    rcsNumero: "441 087 320",
    adresse: "9 Place Saint-Corentin, 29000 Quimper",
    nom: "Anne BODIN",
    qualite: "Gérante",
  },
];

// ---------------------------------------------------------------------------
// Simulated SIREN lookup for the défendeur — in production this calls the
// INSEE Sirene API (api.insee.fr/entreprises/sirene), the official free
// registry for SIREN lookups, through the same backend relay used for
// the requérant search above (no API key needed for Sirene, but a relay
// avoids CORS issues and lets you add caching).
// ---------------------------------------------------------------------------
const MOCK_SIREN = {
  "979626306": {
    formeJuridique: "SARL",
    denomination: "IZOA",
    capital: "1 000",
    rcsVille: "Quimper",
    rcsNumero: "979 626 306",
    adresseSiege: "108 Avenue de la Gare, 29900 Concarneau",
    representant: "Emmelyne GOBIN",
    qualiteRepresentant: "Gérante",
  },
  "441087320": {
    formeJuridique: "SARL",
    denomination: "PASTRY JUSTREG",
    capital: "5 000",
    rcsVille: "Quimper",
    rcsNumero: "441 087 320",
    adresseSiege: "104 Avenue de la Gare, 29900 Concarneau",
    representant: "Justine REGNAULT",
    qualiteRepresentant: "Gérante",
  },
};

function SirenLookup({ onFound }) {
  const [siren, setSiren] = useState("");
  const [status, setStatus] = useState("idle"); // idle | loading | found | notfound

  function digitsOnly(v) {
    return v.replace(/\D/g, "").slice(0, 9);
  }

  function lookup(value) {
    const clean = digitsOnly(value);
    setSiren(clean);
    if (clean.length !== 9) {
      setStatus("idle");
      return;
    }
    setStatus("loading");
    window.setTimeout(() => {
      const match = MOCK_SIREN[clean];
      if (match) {
        onFound(match);
        setStatus("found");
      } else {
        setStatus("notfound");
      }
    }, 400);
  }

  return (
    <div className="mb-4">
      <label className="mb-1.5 flex items-center gap-1.5 font-mono text-[10px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>
        N&deg; SIREN (optionnel, remplit le reste automatiquement)
        <span className="flex items-center gap-0.5 rounded-full px-1.5 py-0.5" style={{ background: "rgba(31,56,100,0.1)", color: NAVY }}>
          <Sparkles size={9} />
          <span className="text-[8.5px] normal-case tracking-normal">via Sirene (INSEE)</span>
        </span>
      </label>
      <div className="relative">
        <Search size={14} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
        <input
          value={formatGrouped3(siren)}
          onChange={(e) => lookup(e.target.value)}
          inputMode="numeric"
          placeholder="ex. 979 626 306"
          maxLength={11}
          className="w-full rounded-xl border py-2.5 pl-9 pr-3.5 font-mono text-[14px] tracking-wide outline-none"
          style={{
            borderColor: status === "found" ? "#2D6A4F" : status === "notfound" ? "#B5301B" : LINE,
            background: "white",
            color: INK,
          }}
        />
      </div>
      {status === "loading" && (
        <p className="mt-1 text-[11px]" style={{ color: MUTED }}>Recherche en cours…</p>
      )}
      {status === "found" && (
        <p className="mt-1 text-[11px]" style={{ color: "#2D6A4F" }}>Entreprise trouvée — champs remplis automatiquement, à vérifier.</p>
      )}
      {status === "notfound" && (
        <p className="mt-1 text-[11px]" style={{ color: "#B5301B" }}>Aucune entreprise trouvée pour ce SIREN — saisissez les informations manuellement.</p>
      )}
    </div>
  );
}

function PappersAutocomplete({ value, onChange, onSelect }) {
  const [open, setOpen] = useState(false);
  const results =
    value.trim().length >= 2
      ? MOCK_PAPPERS.filter((c) => c.denomination.toLowerCase().includes(value.trim().toLowerCase()))
      : [];

  return (
    <div className="relative mb-4">
      <label className="mb-1.5 flex items-center gap-1.5 font-mono text-[10px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>
        Dénomination (ex. SELARL...)
        <span className="flex items-center gap-0.5 rounded-full px-1.5 py-0.5" style={{ background: "rgba(217,98,44,0.12)", color: RUST }}>
          <Sparkles size={9} />
          <span className="text-[8.5px] normal-case tracking-normal">via Pappers</span>
        </span>
      </label>
      <div className="relative">
        <Search size={14} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
        <input
          value={value}
          onChange={(e) => { onChange(e.target.value); setOpen(true); }}
          onFocus={() => setOpen(true)}
          placeholder="Commencez à taper le nom de l'étude..."
          className="w-full rounded-xl border py-2.5 pl-9 pr-3.5 text-[14px] outline-none"
          style={{ borderColor: LINE, background: "white", color: INK }}
        />
      </div>
      {open && results.length > 0 && (
        <div
          className="absolute inset-x-0 top-full z-10 mt-1.5 max-h-56 overflow-y-auto rounded-xl border shadow-lg"
          style={{ borderColor: LINE, background: "white" }}
        >
          {results.map((r) => (
            <button
              key={r.denomination}
              type="button"
              onClick={() => { onSelect(r); setOpen(false); }}
              className="flex w-full flex-col items-start border-b px-3.5 py-2.5 text-left last:border-0"
              style={{ borderColor: LINE }}
            >
              <span className="text-[13px] font-semibold" style={{ color: INK }}>{r.denomination}</span>
              <span className="text-[11px]" style={{ color: MUTED }}>
                RCS {r.rcsVille} {r.rcsNumero} · {r.adresse}
              </span>
            </button>
          ))}
        </div>
      )}
      {open && value.trim().length >= 2 && results.length === 0 && (
        <div className="absolute inset-x-0 top-full z-10 mt-1.5 rounded-xl border px-3.5 py-2.5 shadow-lg" style={{ borderColor: LINE, background: "white" }}>
          <span className="text-[12px]" style={{ color: MUTED }}>Aucun résultat — saisissez les informations manuellement.</span>
        </div>
      )}
    </div>
  );
}

const emptyCaseDraft = {
  formeJuridique: "SARL",
  denomination: "",
  capital: "",
  rcsVille: "",
  rcsNumero: "",
  adresseSiege: "",
  representant: "",
  qualiteRepresentant: "Gérant(e)",
  tribunal: "",
  dateJugement: "",
  adresseLieu: "",
  dateInventaire: "",
  heureDebut: "",
};

function formatGrouped3(v) {
  // Groups digits by 3 with spaces (e.g. "979626306" -> "979 626 306"),
  // keeping non-digit separators the user already typed out of the way.
  const digits = v.replace(/\D/g, "");
  return digits.replace(/(\d{3})(?=\d)/g, "$1 ");
}

function euros(n) {
  if (n === null || n === undefined) return "Mémoire";
  return n.toLocaleString("fr-FR") + " €";
}

function StampBurst({ show }) {
  if (!show) return null;
  return (
    <div className="pointer-events-none absolute inset-0 z-40 flex items-center justify-center">
      <div
        className="flex h-24 w-24 rotate-[-12deg] items-center justify-center rounded-full border-[3px] text-center font-mono text-[11px] font-bold uppercase tracking-wide"
        style={{ borderColor: RUST, color: RUST, animation: "stampIn 480ms cubic-bezier(.2,1.4,.4,1) forwards" }}
      >
        Lot
        <br />
        constaté
      </div>
    </div>
  );
}

function Sheet({ open, onClose, title, eyebrow, children, footer, z = 30 }) {
  return (
    <>
      <div
        className={`absolute inset-x-0 bottom-0 flex flex-col rounded-t-[2rem] border-t shadow-[0_-8px_30px_rgba(0,0,0,0.18)] transition-transform duration-300 ease-out ${
          open ? "translate-y-0" : "translate-y-full"
        }`}
        style={{ background: PAPER, borderColor: LINE, height: "88%", zIndex: z }}
      >
        <div className="flex shrink-0 items-center justify-between border-b px-5 py-4" style={{ borderColor: LINE }}>
          <div>
            {eyebrow && (
              <p className="font-mono text-[10px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>
                {eyebrow}
              </p>
            )}
            <h3 className="text-[16px] font-semibold" style={{ color: INK }}>{title}</h3>
          </div>
          <button onClick={onClose} className="flex h-8 w-8 items-center justify-center rounded-full" style={{ background: PAPER_DEEP, color: INK }} aria-label="Fermer">
            <X size={16} />
          </button>
        </div>
        <div className="flex-1 overflow-y-auto px-5 py-4">{children}</div>
        {footer && <div className="shrink-0 border-t px-5 py-4" style={{ borderColor: LINE }}>{footer}</div>}
      </div>
      {open && (
        <button aria-label="Fermer" onClick={onClose} className="absolute inset-0" style={{ background: "rgba(20,18,16,0.35)", zIndex: z - 1 }} />
      )}
    </>
  );
}

function Field({ label, ...props }) {
  return (
    <div className="mb-4">
      <label className="mb-1.5 block font-mono text-[10px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>
        {label}
      </label>
      <input
        {...props}
        className="w-full rounded-xl border px-3.5 py-2.5 text-[14px] outline-none"
        style={{ borderColor: LINE, background: "white", color: INK }}
      />
    </div>
  );
}

function SectionCard({ icon: Icon, title, children }) {
  return (
    <div className="mb-4 rounded-2xl border p-4" style={{ borderColor: LINE, background: "white" }}>
      <div className="mb-3 flex items-center gap-2">
        <div className="flex h-7 w-7 items-center justify-center rounded-full" style={{ background: PAPER_DEEP, color: NAVY }}>
          <Icon size={14} />
        </div>
        <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>{title}</h3>
      </div>
      {children}
    </div>
  );
}

function ZonePicker({ zones, value, onChange, onManage }) {
  return (
    <div className="flex items-center gap-2">
      <div className="flex flex-1 gap-2 overflow-x-auto pb-1 scrollbar-none">
        {zones.map((z) => {
          const active = value === z;
          return (
            <button
              key={z}
              type="button"
              onClick={() => onChange(z)}
              className="shrink-0 whitespace-nowrap rounded-full border px-3.5 py-2 font-mono text-[11px] font-semibold uppercase tracking-wide transition-colors"
              style={{ borderColor: active ? NAVY : LINE, background: active ? NAVY : "transparent", color: active ? PAPER : "#6B6558" }}
            >
              {z}
            </button>
          );
        })}
      </div>
      <button type="button" onClick={onManage} aria-label="Gérer les emplacements" className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full border" style={{ borderColor: LINE, color: MUTED }}>
        <Settings2 size={14} />
      </button>
    </div>
  );
}

function CategoryPicker({ categories, value, onChange, onManage }) {
  return (
    <div>
      <div className="flex flex-wrap gap-2">
        {categories.map((c) => {
          const active = value === c.key;
          return (
            <button
              key={c.key}
              type="button"
              onClick={() => onChange(c.key)}
              className="rounded-xl border px-3 py-2.5 text-center text-[12px] font-medium leading-tight transition-colors"
              style={{ borderColor: active ? RUST : LINE, background: active ? "rgba(217,98,44,0.08)" : "white", color: active ? RUST : INK }}
            >
              {c.label}
            </button>
          );
        })}
      </div>
      <button
        type="button"
        onClick={onManage}
        className="mt-2 flex items-center gap-1 text-[11px] font-semibold"
        style={{ color: NAVY }}
      >
        <Settings2 size={12} />
        Gérer les catégories
      </button>
    </div>
  );
}

function PhotoThumb({ src, onRemove }) {
  return (
    <div className="relative h-16 w-16 shrink-0 overflow-hidden rounded-lg border" style={{ borderColor: LINE }}>
      {src === "seed" ? (
        <div className="flex h-full w-full items-center justify-center" style={{ background: PAPER_DEEP, color: MUTED }}>
          <ImageOff size={18} />
        </div>
      ) : (
        <img src={src} alt="" className="h-full w-full object-cover" />
      )}
      {onRemove && (
        <button onClick={onRemove} aria-label="Retirer la photo" className="absolute right-0.5 top-0.5 flex h-4 w-4 items-center justify-center rounded-full" style={{ background: "rgba(20,18,16,0.65)" }}>
          <X size={10} color="white" />
        </button>
      )}
    </div>
  );
}

function LotRow({ lot, index, onEdit, onDelete }) {
  return (
    <button
      type="button"
      onClick={() => onEdit(lot)}
      className="group relative flex w-full items-center gap-3 border-b py-3 text-left last:border-0"
      style={{ borderColor: LINE }}
    >
      <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full font-mono text-[12px] font-bold" style={{ background: PAPER_DEEP, color: NAVY }}>
        {index + 1}
      </div>
      <div className="min-w-0 flex-1">
        <div className="flex items-center gap-1.5">
          <span className="font-mono text-[9.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>{lot.zone}</span>
          {lot.photos.length > 0 && (
            <span className="flex items-center gap-0.5" style={{ color: MUTED }}>
              <Camera size={11} strokeWidth={2.25} />
              <span className="font-mono text-[9.5px]">{lot.photos.length}</span>
            </span>
          )}
        </div>
        <p className="truncate text-[13.5px] font-medium" style={{ color: INK }}>{lot.desig}</p>
      </div>
      <div className="shrink-0 font-mono text-[13.5px] font-semibold tabular-nums" style={{ color: lot.value === null ? MUTED : NAVY }}>
        {euros(lot.value)}
      </div>
      <Pencil size={13} strokeWidth={2} className="shrink-0 opacity-40" style={{ color: MUTED }} />
      <span
        role="button"
        tabIndex={0}
        onClick={(e) => { e.stopPropagation(); onDelete(lot.id); }}
        onKeyDown={(e) => { if (e.key === "Enter") { e.stopPropagation(); onDelete(lot.id); } }}
        aria-label="Supprimer le lot"
        className="shrink-0 rounded-full p-1.5"
        style={{ color: "#B5301B" }}
      >
        <Trash2 size={15} strokeWidth={2} />
      </span>
    </button>
  );
}

// ---------------------------------------------------------------------------
// Tab: Nouveau dossier
// ---------------------------------------------------------------------------
function NewCaseTab({ requerants, onAddRequerant, onCreateCase }) {
  const [selectedReq, setSelectedReq] = useState(requerants[0]?.id ?? null);
  const [reqSheetOpen, setReqSheetOpen] = useState(false);
  const [reqDraft, setReqDraft] = useState(emptyRequerant);

  const [c, setC] = useState(emptyCaseDraft);
  const setField = (k) => (e) => setC((prev) => ({ ...prev, [k]: e.target.value }));

  const canCreate =
    selectedReq && c.denomination.trim() && c.adresseSiege.trim() && c.adresseLieu.trim();

  function saveRequerant() {
    if (!reqDraft.denomination.trim()) return;
    const id = "r" + Date.now();
    onAddRequerant({ ...reqDraft, id });
    setSelectedReq(id);
    setReqDraft(emptyRequerant);
    setReqSheetOpen(false);
  }

  return (
    <div className="flex-1 overflow-y-auto px-5 py-4">
      <SectionCard icon={Gavel} title="Requérant (mandataire liquidateur)">
        <div className="mb-3 space-y-2">
          {requerants.map((r) => {
            const active = selectedReq === r.id;
            return (
              <button
                key={r.id}
                onClick={() => setSelectedReq(r.id)}
                className="flex w-full items-center justify-between rounded-xl border px-3.5 py-2.5 text-left"
                style={{ borderColor: active ? NAVY : LINE, background: active ? "rgba(31,56,100,0.06)" : "white" }}
              >
                <div className="min-w-0">
                  <p className="truncate text-[13px] font-semibold" style={{ color: INK }}>{r.denomination}</p>
                  <p className="truncate text-[11px]" style={{ color: MUTED }}>{r.nom} · {r.adresse}</p>
                </div>
                {active && <Check size={16} style={{ color: NAVY }} className="shrink-0" />}
              </button>
            );
          })}
        </div>
        <button
          onClick={() => setReqSheetOpen(true)}
          className="flex w-full items-center justify-center gap-1.5 rounded-xl border-2 border-dashed py-2.5 text-[12.5px] font-semibold"
          style={{ borderColor: LINE, color: NAVY }}
        >
          <Plus size={15} />
          Nouveau requérant
        </button>
      </SectionCard>

      <SectionCard icon={Building2} title="Défendeur">
        <SirenLookup onFound={(data) => setC((prev) => ({ ...prev, ...data }))} />
        <div className="grid grid-cols-2 gap-x-3">
          <Field label="Forme juridique" value={c.formeJuridique} onChange={setField("formeJuridique")} placeholder="SARL" />
          <Field label="Capital" value={c.capital} onChange={setField("capital")} placeholder="1 000" />
        </div>
        <Field label="Dénomination sociale" value={c.denomination} onChange={setField("denomination")} placeholder="Ex. IZOA" />
        <div className="grid grid-cols-2 gap-x-3">
          <Field label="Ville RCS" value={c.rcsVille} onChange={setField("rcsVille")} placeholder="Quimper" />
          <Field label="N° RCS / SIREN" value={c.rcsNumero} onChange={(e) => setC((p) => ({ ...p, rcsNumero: formatGrouped3(e.target.value) }))} placeholder="979 626 306" />
        </div>
        <Field label="Adresse du siège social" value={c.adresseSiege} onChange={setField("adresseSiege")} placeholder="108 Avenue de la Gare, 29900 Concarneau" />
        <div className="grid grid-cols-2 gap-x-3">
          <Field label="Représentant légal" value={c.representant} onChange={setField("representant")} placeholder="Prénom Nom" />
          <Field label="Qualité" value={c.qualiteRepresentant} onChange={setField("qualiteRepresentant")} placeholder="Gérant(e)" />
        </div>
      </SectionCard>

      <SectionCard icon={Gavel} title="Jugement">
        <div className="grid grid-cols-2 gap-x-3">
          <Field label="Tribunal de commerce de" value={c.tribunal} onChange={setField("tribunal")} placeholder="Quimper" />
          <Field label="Date du jugement" type="date" value={c.dateJugement} onChange={setField("dateJugement")} />
        </div>
      </SectionCard>

      <SectionCard icon={MapPin} title="Inventaire">
        <Field label="Adresse du lieu de l'inventaire" value={c.adresseLieu} onChange={setField("adresseLieu")} placeholder="Idem siège social si vide" />
        <div className="grid grid-cols-2 gap-x-3">
          <Field label="Date" type="date" value={c.dateInventaire} onChange={setField("dateInventaire")} />
          <Field label="Heure de début" type="time" value={c.heureDebut} onChange={setField("heureDebut")} />
        </div>
      </SectionCard>

      <button
        onClick={() => canCreate && onCreateCase(c, requerants.find((r) => r.id === selectedReq))}
        disabled={!canCreate}
        className="mt-2 flex w-full items-center justify-center gap-2 rounded-2xl py-3.5 text-[14.5px] font-semibold transition-opacity active:scale-[0.98] disabled:opacity-35"
        style={{ background: RUST, color: PAPER }}
      >
        <FilePlus2 size={17} />
        Créer le dossier et démarrer l'inventaire
      </button>
      <div className="h-6" />

      <Sheet
        open={reqSheetOpen}
        onClose={() => setReqSheetOpen(false)}
        eyebrow="Requérant"
        title="Nouveau requérant"
        z={40}
        footer={
          <button
            onClick={saveRequerant}
            disabled={!reqDraft.denomination.trim()}
            className="flex w-full items-center justify-center gap-2 rounded-2xl py-3.5 text-[14.5px] font-semibold active:scale-[0.98] transition-transform disabled:opacity-35"
            style={{ background: NAVY, color: PAPER }}
          >
            <Check size={17} />
            Enregistrer ce requérant
          </button>
        }
      >
        <PappersAutocomplete
          value={reqDraft.denomination}
          onChange={(v) => setReqDraft((p) => ({ ...p, denomination: v }))}
          onSelect={(r) => setReqDraft({ ...r })}
        />
        <p className="-mt-2 mb-4 text-[11px]" style={{ color: MUTED }}>
          Sélectionnez une étude dans la liste pour remplir automatiquement le capital, le RCS, l'adresse et le représentant. Vous pouvez ensuite tout corriger à la main.
        </p>
        <div className="grid grid-cols-2 gap-x-3">
          <Field label="Capital (€)" value={reqDraft.capital} onChange={(e) => setReqDraft((p) => ({ ...p, capital: e.target.value }))} placeholder="200 000" />
          <Field label="Ville RCS" value={reqDraft.rcsVille} onChange={(e) => setReqDraft((p) => ({ ...p, rcsVille: e.target.value }))} placeholder="Brest" />
        </div>
        <Field label="N° RCS" value={reqDraft.rcsNumero} onChange={(e) => setReqDraft((p) => ({ ...p, rcsNumero: formatGrouped3(e.target.value) }))} placeholder="808 072 821" />
        <Field label="Adresse du siège" value={reqDraft.adresse} onChange={(e) => setReqDraft((p) => ({ ...p, adresse: e.target.value }))} placeholder="62A Quai de l'Odet, 29000 Quimper" />
        <div className="grid grid-cols-2 gap-x-3">
          <Field label="Représentant" value={reqDraft.nom} onChange={(e) => setReqDraft((p) => ({ ...p, nom: e.target.value }))} placeholder="Prénom Nom" />
          <Field label="Qualité" value={reqDraft.qualite} onChange={(e) => setReqDraft((p) => ({ ...p, qualite: e.target.value }))} placeholder="Gérant" />
        </div>
      </Sheet>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Tab: Inventaire (field capture)
// ---------------------------------------------------------------------------
function InventoryTab({ caseInfo, onEditCase, initialData, dossierId, onSaved }) {
  const [zones, setZones] = useState(initialData?.zones || SEED_ZONES);
  const [zoneSheetOpen, setZoneSheetOpen] = useState(false);
  const [newZone, setNewZone] = useState("");

  const [categories, setCategories] = useState(initialData?.categories || SEED_CATEGORIES);
  const [catSheetOpen, setCatSheetOpen] = useState(false);
  const [newCatLabel, setNewCatLabel] = useState("");

  const [lots, setLots] = useState(initialData?.lots || SEED_LOTS);
  const [zone, setZone] = useState((initialData?.zones || SEED_ZONES)[0]);
  const [cat, setCat] = useState("materiel");
  const [desig, setDesig] = useState("");
  const [value, setValue] = useState("");
  const [memoire, setMemoire] = useState(false);
  const [photos, setPhotos] = useState([]);
  const [processingPhotos, setProcessingPhotos] = useState(false);
  const [showStamp, setShowStamp] = useState(false);
  const [sheetOpen, setSheetOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [savingDossier, setSavingDossier] = useState(false);
  const [saveMessage, setSaveMessage] = useState("");
  const nextId = useRef((initialData?.lots?.length || SEED_LOTS.length) + 1);
  const fileInputRef = useRef(null);
  const recognitionRef = useRef(null);
  const [listening, setListening] = useState(false);
  const [dictationSupported] = useState(
    () => typeof window !== "undefined" && (window.SpeechRecognition || window.webkitSpeechRecognition)
  );

  function toggleDictation() {
    if (!dictationSupported) return;

    if (listening) {
      recognitionRef.current?.stop();
      return;
    }

    const SpeechRecognitionCtor = window.SpeechRecognition || window.webkitSpeechRecognition;
    const recognition = new SpeechRecognitionCtor();
    recognition.lang = "fr-FR";
    recognition.continuous = true;
    recognition.interimResults = false;

    recognition.onresult = (event) => {
      let finalText = "";
      for (let i = event.resultIndex; i < event.results.length; i++) {
        if (event.results[i].isFinal) finalText += event.results[i][0].transcript;
      }
      if (finalText.trim()) {
        setDesig((prev) => (prev ? `${prev.trim()} ${finalText.trim()}` : finalText.trim()).toUpperCase());
      }
    };
    recognition.onerror = () => setListening(false);
    recognition.onend = () => setListening(false);

    recognitionRef.current = recognition;
    recognition.start();
    setListening(true);
  }

  useEffect(() => {
    return () => recognitionRef.current?.stop();
  }, []);

  function openNewLot() {
    setEditingId(null);
    setZone(zones[0]);
    setCat("materiel");
    setDesig("");
    setValue("");
    setMemoire(false);
    setPhotos([]);
    setSheetOpen(true);
  }

  function openEditLot(lot) {
    setEditingId(lot.id);
    setZone(lot.zone);
    setCat(lot.cat);
    setDesig(lot.desig);
    setValue(lot.value === null ? "" : String(lot.value));
    setMemoire(lot.value === null);
    setPhotos(lot.photos);
    setSheetOpen(true);
  }

  const totals = useMemo(() => {
    const byCat = Object.fromEntries(categories.map((c) => [c.key, 0]));
    let grand = 0;
    for (const l of lots) {
      if (l.value !== null) {
        byCat[l.cat] = (byCat[l.cat] ?? 0) + l.value;
        grand += l.value;
      }
    }
    return { byCat, grand };
  }, [lots, categories]);

  const canAdd = desig.trim().length > 0 && (memoire || value !== "");
  const isEditing = editingId !== null;

  function saveLot() {
    if (!canAdd) return;
    const payload = { zone, cat, desig: desig.trim(), value: memoire ? null : Number(value), photos };
    if (isEditing) {
      setLots((prev) => prev.map((l) => (l.id === editingId ? { ...l, ...payload } : l)));
    } else {
      setLots((prev) => [...prev, { id: nextId.current++, ...payload }]);
    }
    setSheetOpen(false);
    setEditingId(null);
    setShowStamp(true);
    window.setTimeout(() => setShowStamp(false), 520);
  }

  function deleteLot(id) {
    setLots((prev) => prev.filter((l) => l.id !== id));
  }

  function watermarkImage(file) {
    return new Promise((resolve) => {
      const objectUrl = URL.createObjectURL(file);
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0);

        const scale = img.width / 1000; // scale text/badge relative to a 1000px-wide reference
        const fontSize = Math.max(14, 22 * scale);
        const pad = Math.max(8, 14 * scale);
        const label = "© SCP BRELIVET Tanguy";

        // Bottom-left dark bar with the study's name
        ctx.font = `bold ${fontSize}px Arial, sans-serif`;
        const textWidth = ctx.measureText(label).width;
        const barHeight = fontSize + pad * 1.2;
        ctx.fillStyle = "rgba(20,18,16,0.55)";
        ctx.fillRect(0, img.height - barHeight, textWidth + pad * 2, barHeight);
        ctx.fillStyle = "#FFFFFF";
        ctx.textBaseline = "middle";
        ctx.fillText(label, pad, img.height - barHeight / 2);

        // Bottom-right orange circular stamp badge (stylised — not a pixel copy of the real logo)
        const r = Math.max(16, 26 * scale);
        const cx = img.width - r - pad;
        const cy = img.height - r - pad;
        ctx.beginPath();
        ctx.arc(cx, cy, r, 0, Math.PI * 2);
        ctx.fillStyle = "#D9622C";
        ctx.fill();
        ctx.fillStyle = "#FFFFFF";
        ctx.font = `bold ${r * 0.8}px Arial, sans-serif`;
        ctx.textAlign = "center";
        ctx.fillText("HJ", cx, cy);
        ctx.textAlign = "left";

        canvas.toBlob(
          (blob) => {
            URL.revokeObjectURL(objectUrl);
            resolve(URL.createObjectURL(blob));
          },
          "image/jpeg",
          0.9
        );
      };
      img.src = objectUrl;
    });
  }

  async function handleFiles(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = "";
    setProcessingPhotos(true);
    try {
      const urls = await Promise.all(files.map(watermarkImage));
      setPhotos((prev) => [...prev, ...urls]);
    } finally {
      setProcessingPhotos(false);
    }
  }

  function removePhoto(idx) {
    setPhotos((prev) => prev.filter((_, i) => i !== idx));
  }

  function addZone() {
    const z = newZone.trim().toUpperCase();
    if (!z || zones.includes(z)) return;
    setZones((prev) => [...prev, z]);
    setNewZone("");
  }

  function removeZone(z) {
    setZones((prev) => prev.filter((x) => x !== z));
    if (zone === z && zones.length > 1) setZone(zones.find((x) => x !== z));
  }

  function addCategory() {
    const label = newCatLabel.trim();
    if (!label) return;
    const key = label
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
    if (!key || categories.some((c) => c.key === key)) return;
    setCategories((prev) => [...prev, { key, label }]);
    setNewCatLabel("");
  }

  function removeCategory(key) {
    setCategories((prev) => prev.filter((c) => c.key !== key));
    if (cat === key && categories.length > 1) {
      setCat(categories.find((c) => c.key !== key).key);
    }
  }

  const catLabel = (key) => categories.find((c) => c.key === key)?.label ?? key;

  const [exporting, setExporting] = useState(false);

  async function saveToCloud() {
    setSavingDossier(true);
    setSaveMessage("");
    const { data, error } = await saveDossier({ id: dossierId, caseInfo, lots, zones, categories });
    setSavingDossier(false);
    if (error) {
      setSaveMessage("Échec de l'enregistrement : " + error.message);
      return;
    }
    setSaveMessage("Dossier enregistré (" + new Date().toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" }) + ")");
    onSaved?.(data?.id);
  }
  const [exportError, setExportError] = useState("");
  const [exportedFile, setExportedFile] = useState(null); // { url, filename }

  async function exportToWord() {
    setExporting(true);
    setExportError("");
    setExportedFile(null);
    try {
      const req = caseInfo.requerant || {};
      const today = new Date().toLocaleDateString("fr-FR", { day: "numeric", month: "long", year: "numeric" });
      const dateInv = caseInfo.dateInventaire
        ? new Date(caseInfo.dateInventaire).toLocaleDateString("fr-FR", { day: "numeric", month: "long", year: "numeric" })
        : today;
      const dateJug = caseInfo.dateJugement
        ? new Date(caseInfo.dateJugement).toLocaleDateString("fr-FR", { day: "numeric", month: "long", year: "numeric" })
        : "[DATE DU JUGEMENT]";

      const NAVY_HEX = "1F3864";
      const LINE_HEX = "DCD5C4";
      const RECAP_HEX = "D9E2F3";

      // Pre-fetch every real (non-seed) photo as raw bytes for embedding.
      const photoBytesCache = {};
      for (const l of lots) {
        for (const src of l.photos) {
          if (src !== "seed" && !photoBytesCache[src]) {
            try {
              const res = await fetch(src);
              photoBytesCache[src] = new Uint8Array(await res.arrayBuffer());
            } catch {
              photoBytesCache[src] = null;
            }
          }
        }
      }

      const P = (text, opts = {}) =>
        new Paragraph({
          alignment: opts.align || AlignmentType.JUSTIFIED,
          spacing: { after: 160 },
          children: [new TextRun({ text, bold: !!opts.bold, italics: !!opts.italics, size: 22 })],
        });

      const cell = (text, opts = {}) =>
        new TableCell({
          width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
          shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined,
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: opts.align || AlignmentType.LEFT,
              children: [
                new TextRun({
                  text,
                  bold: !!opts.bold,
                  italics: !!opts.italics,
                  color: opts.color,
                  size: 19,
                }),
              ],
            }),
          ],
        });

      const lotTableRows = [
        new TableRow({
          tableHeader: true,
          children: [
            cell("N° de lot", { fill: NAVY_HEX, color: "FFFFFF", bold: true, width: 8, align: AlignmentType.CENTER }),
            cell("Catégorie", { fill: NAVY_HEX, color: "FFFFFF", bold: true, width: 20 }),
            cell("Désignation", { fill: NAVY_HEX, color: "FFFFFF", bold: true, width: 52 }),
            cell("Valeur de réalisation", { fill: NAVY_HEX, color: "FFFFFF", bold: true, width: 20, align: AlignmentType.RIGHT }),
          ],
        }),
        ...(lots.length
          ? lots.map(
              (l, i) =>
                new TableRow({
                  children: [
                    cell(String(i + 1), { align: AlignmentType.CENTER, bold: true }),
                    cell(catLabel(l.cat), { italics: true }),
                    cell(
                      `[${l.zone}] ${l.desig}` +
                        (l.photos.length ? `  (${l.photos.length} photo${l.photos.length > 1 ? "s" : ""}, voir annexe)` : "")
                    ),
                    cell(euros(l.value), { align: AlignmentType.RIGHT, bold: true }),
                  ],
                })
            )
          : [
              new TableRow({
                children: [
                  new TableCell({
                    columnSpan: 4,
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("Aucun lot constaté")] })],
                  }),
                ],
              }),
            ]),
      ];

      const recapTableRows = [
        new TableRow({
          tableHeader: true,
          children: [
            cell("Rubrique", { fill: NAVY_HEX, color: "FFFFFF", bold: true, width: 70 }),
            cell("Valeur de réalisation", { fill: NAVY_HEX, color: "FFFFFF", bold: true, width: 30, align: AlignmentType.RIGHT }),
          ],
        }),
        ...categories.map(
          (c) =>
            new TableRow({
              children: [
                cell(c.label, { fill: RECAP_HEX }),
                cell(euros(totals.byCat[c.key]), { fill: RECAP_HEX, bold: true, align: AlignmentType.RIGHT }),
              ],
            })
        ),
        new TableRow({
          children: [
            cell("TOTAL GENERAL", { fill: NAVY_HEX, color: "FFFFFF", bold: true }),
            cell(euros(totals.grand), { fill: NAVY_HEX, color: "FFFFFF", bold: true, align: AlignmentType.RIGHT }),
          ],
        }),
      ];

      const costTableRows = [
        new TableRow({ children: [cell("Emoluments R444-3"), cell("[A CALCULER]", { align: AlignmentType.RIGHT })] }),
        new TableRow({ children: [cell("TVA"), cell("[A CALCULER]", { align: AlignmentType.RIGHT })] }),
        new TableRow({
          children: [
            cell("TOTAL TTC", { fill: NAVY_HEX, color: "FFFFFF", bold: true }),
            cell("[A CALCULER]", { fill: NAVY_HEX, color: "FFFFFF", bold: true, align: AlignmentType.RIGHT }),
          ],
        }),
      ];

      // Photo annex — one image + caption per block, 2 per page.
      const photoAnnexChildren = [];
      if (lots.some((l) => l.photos.length)) {
        photoAnnexChildren.push(
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 240 },
            children: [new TextRun({ text: "ANNEXE PHOTOGRAPHIQUE", bold: true, size: 26, color: NAVY_HEX })],
          })
        );
        let photoCount = 0;
        lots.forEach((l, li) => {
          l.photos.forEach((src, pi) => {
            const bytes = src === "seed" ? null : photoBytesCache[src];
            photoCount++;
            if (bytes) {
              photoAnnexChildren.push(
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 80 },
                  children: [new ImageRun({ data: bytes, transformation: { width: 320, height: 320 } })],
                })
              );
            } else {
              photoAnnexChildren.push(P("[photo d'exemple non incluse]", { align: AlignmentType.CENTER, italics: true }));
            }
            photoAnnexChildren.push(
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 240 },
                children: [
                  new TextRun({
                    text: `Lot n°${li + 1} — ${l.desig} — photo ${pi + 1}/${l.photos.length}`,
                    italics: true,
                    size: 16,
                    color: "8A8371",
                  }),
                ],
              })
            );
            if (photoCount % 2 === 0) photoAnnexChildren.push(new Paragraph({ children: [new PageBreak()] }));
          });
        });
      }

      const doc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 80 },
                children: [new TextRun({ text: "PROCES-VERBAL DE CONSTAT D'INVENTAIRE", bold: true, size: 30 })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 240 },
                children: [
                  new TextRun({
                    text: `Document généré le ${today} — brouillon issu de l'application de saisie terrain`,
                    italics: true,
                    size: 18,
                  }),
                ],
              }),
              P("A LA REQUETE DE :", { bold: true }),
              P(
                `${req.denomination || "[DENOMINATION DU MANDATAIRE]"}, au capital de ${req.capital || "[MONTANT]"} euros, inscrite au registre du commerce et des sociétés de ${req.rcsVille || "[VILLE RCS]"} sous le numéro ${req.rcsNumero || "[N° RCS]"}, dont le siège social est situé ${req.adresse || "[ADRESSE DU MANDATAIRE]"}, agissant poursuites et diligences de son représentant légal ${req.nom || "[REPRESENTANT]"}, ${req.qualite || "[QUALITE]"},`
              ),
              P(
                `Agissant en qualité de mandataire liquidateur dans le cadre de la liquidation judiciaire de la ${caseInfo.formeJuridique || "[FORME]"} ${caseInfo.denomination || "[DENOMINATION]"}, au capital de ${caseInfo.capital || "[CAPITAL]"} euros, inscrite au RCS de ${caseInfo.rcsVille || "[VILLE]"} sous le numéro ${caseInfo.rcsNumero || "[N°]"}, dont le siège social est situé ${caseInfo.adresseSiege || "[ADRESSE]"}, représentée par ${caseInfo.representant || "[REPRESENTANT]"}, ${caseInfo.qualiteRepresentant || "[QUALITE]"},`
              ),
              P(
                `En vertu d'un jugement rendu par le Tribunal de Commerce de ${caseInfo.tribunal || "[VILLE]"}, en date du ${dateJug}, prononçant la liquidation judiciaire susvisée, et des articles L.641-1 et suivants du Code de commerce,`
              ),
              P("Qu'elle me requiert afin de dresser un procès-verbal de constat et toutes constatations utiles à la sauvegarde de ses droits,"),
              P(`Me suis rendu, ${caseInfo.address}, le ${dateInv} à ${caseInfo.heureDebut || "[HEURE]"},`),
              P("Et j'ai constaté ce qui suit :"),
              P(`Activité exercée : [A COMPLETER]. Statut : ${caseInfo.formeJuridique || "[FORME]"}, société de capitaux. Capital : ${caseInfo.capital || "[CAPITAL]"} €.`),

              new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "Copie des pièces reçues :", bold: true })] }),
              new Paragraph({ bullet: { level: 0 }, children: [new TextRun("Extrait Kbis — "), new TextRun({ text: "NON OBTENU", bold: true, highlight: "yellow" })] }),
              new Paragraph({ bullet: { level: 0 }, children: [new TextRun("Bilan et compte de résultat — "), new TextRun({ text: "NON OBTENU", bold: true, highlight: "yellow" })] }),
              new Paragraph({ bullet: { level: 0 }, spacing: { after: 240 }, children: [new TextRun("Etat des privilèges et nantissements — "), new TextRun({ text: "NON OBTENU", bold: true, highlight: "yellow" })] }),

              new Paragraph({ spacing: { before: 120, after: 120 }, children: [new TextRun({ text: "INVENTAIRE DES LOTS", bold: true, size: 26, color: NAVY_HEX })] }),
              new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: lotTableRows }),

              new Paragraph({ spacing: { before: 240, after: 120 }, children: [new TextRun({ text: "RECAPITULATIF PAR RUBRIQUE", bold: true, size: 26, color: NAVY_HEX })] }),
              new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: recapTableRows }),

              new Paragraph({
                spacing: { before: 240 },
                children: [
                  new TextRun(
                    "Ayant satisfait à la mission pour laquelle j'avais été requis, j'ai suspendu mes opérations et de tout ce qui précède, j'ai dressé le présent procès-verbal de constat pour servir et valoir ce que de droit à ma requérante."
                  ),
                ],
              }),

              new Paragraph({ spacing: { before: 240, after: 120 }, children: [new TextRun({ text: "Coût :", bold: true })] }),
              new Table({ width: { size: 60, type: WidthType.PERCENTAGE }, rows: costTableRows }),

              ...photoAnnexChildren,

              new Paragraph({
                spacing: { before: 480 },
                children: [
                  new TextRun({
                    text: "Export automatique depuis l'application de saisie terrain — document de travail à relire et compléter (pièces reçues, coûts, mentions légales, signature) avant remise officielle dans le modèle de l'étude.",
                    size: 16,
                    color: "8A8371",
                  }),
                ],
              }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      const safeName = (caseInfo.denomination || "inventaire").replace(/[^a-z0-9]+/gi, "_");
      const filename = `PV_inventaire_${safeName}.docx`;
      const file = new File([blob], filename, {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      // Always keep a manual link ready — the iOS Share Sheet gives no
      // confirmation once a target is chosen, so if "Save to Files" /
      // "Open in Word" silently fails or the person isn't sure it worked,
      // they still have a guaranteed way to grab the file themselves.
      const url = URL.createObjectURL(blob);
      setExportedFile({ url, filename });

      // iOS (Safari and installed PWAs) handles the native Share Sheet far
      // more reliably than a simulated <a download> click, which is often
      // silently swallowed in standalone/home-screen mode. Desktop browsers
      // fall back to a normal download.
      let shared = false;
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        try {
          await navigator.share({
            files: [file],
            title: filename,
            text: "Procès-verbal d'inventaire (brouillon)",
          });
          shared = true;
          setExportError("");
        } catch (shareErr) {
          // AbortError = the person just cancelled the share sheet — not a real error.
          if (shareErr?.name === "AbortError") shared = true;
        }
      }

      if (!shared) {
        try {
          const a = document.createElement("a");
          a.href = url;
          a.download = filename;
          a.rel = "noopener";
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        } catch {
          // Manual link below covers this case.
        }
      }
    } catch (err) {
      setExportError(
        "Échec de la génération du document (" + (err?.message || "erreur inconnue") + "). Réessayez ; si ça persiste, réduisez le nombre de photos et retentez."
      );
    } finally {
      setExporting(false);
    }
  }

  return (
    <>
      <StampBurst show={showStamp} />

      <div className="shrink-0 px-5 pb-4 pt-5" style={{ background: NAVY, color: PAPER }}>
        <div className="flex items-center justify-between">
          <div className="min-w-0">
            <p className="font-mono text-[10px] font-semibold uppercase tracking-[0.14em]" style={{ color: "#9FB0C9" }}>Inventaire en cours</p>
            <h1 className="mt-0.5 truncate text-[18px] font-semibold leading-tight">{caseInfo.name}</h1>
          </div>
          <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-full" style={{ background: RUST }}>
            <Stamp size={18} color={PAPER} strokeWidth={2} />
          </div>
        </div>
        <button onClick={onEditCase} className="mt-3 flex items-center gap-1.5 rounded-full py-1 text-[11.5px]" style={{ color: "#C3CEDE" }}>
          <MapPin size={12} />
          <span className="truncate">{caseInfo.address}</span>
          <Pencil size={11} className="shrink-0 opacity-70" />
        </button>
        <span className="mt-1 flex items-center gap-1 text-[11.5px]" style={{ color: "#C3CEDE" }}>
          <Clock size={12} /> {caseInfo.heureDebut || "14:02"}
        </span>
      </div>

      <div className="flex shrink-0 divide-x overflow-x-auto border-b scrollbar-none" style={{ borderColor: LINE, background: PAPER_DEEP }}>
        {categories.map((c) => (
          <div key={c.key} className="shrink-0 px-3 py-2.5 text-center" style={{ minWidth: `${100 / Math.min(categories.length, 3)}%` }}>
            <p className="truncate font-mono text-[8.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>{c.label}</p>
            <p className="mt-0.5 font-mono text-[13px] font-bold tabular-nums" style={{ color: NAVY }}>{euros(totals.byCat[c.key])}</p>
          </div>
        ))}
      </div>

      <div className="flex-1 overflow-y-auto px-5 pt-3">
        <div className="mb-2 flex items-baseline justify-between">
          <h2 className="font-mono text-[10.5px] font-semibold uppercase tracking-[0.12em]" style={{ color: MUTED }}>Lots constatés</h2>
          <span className="font-mono text-[10.5px]" style={{ color: MUTED }}>{lots.length} lot{lots.length > 1 ? "s" : ""}</span>
        </div>

        {lots.length === 0 ? (
          <div className="mt-10 rounded-2xl border border-dashed px-4 py-8 text-center" style={{ borderColor: LINE }}>
            <p className="text-[13px]" style={{ color: MUTED }}>Aucun lot constaté pour l'instant. Touchez « Ajouter un lot » pour commencer l'inventaire.</p>
          </div>
        ) : (
          <div>
            {lots.map((lot, i) => (
              <LotRow key={lot.id} lot={lot} index={i} onEdit={openEditLot} onDelete={deleteLot} />
            ))}
          </div>
        )}
        <div className="h-4" />
      </div>

      <div className="shrink-0 border-t px-5 pb-4 pt-3" style={{ borderColor: LINE, background: PAPER }}>
        <div className="mb-3 flex items-center justify-between">
          <span className="font-mono text-[11px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>Total général</span>
          <span className="font-mono text-[19px] font-bold tabular-nums" style={{ color: NAVY }}>{euros(totals.grand)}</span>
        </div>
        <button onClick={openNewLot} className="flex w-full items-center justify-center gap-2 rounded-2xl py-3.5 text-[14.5px] font-semibold shadow-sm active:scale-[0.98] transition-transform" style={{ background: RUST, color: PAPER }}>
          <Plus size={18} strokeWidth={2.5} />
          Ajouter un lot
        </button>
        {supabaseConfigured && (
          <button
            onClick={saveToCloud}
            disabled={savingDossier}
            className="mt-2 flex w-full items-center justify-center gap-2 rounded-2xl border py-3 text-[13.5px] font-semibold transition-opacity active:scale-[0.98] disabled:opacity-35"
            style={{ borderColor: "#2D6A4F", color: "#2D6A4F" }}
          >
            <Cloud size={16} />
            {savingDossier ? "Enregistrement…" : dossierId ? "Mettre à jour le dossier" : "Enregistrer le dossier"}
          </button>
        )}
        {saveMessage && (
          <p className="mt-1.5 text-center text-[11px]" style={{ color: "#2D6A4F" }}>{saveMessage}</p>
        )}
        <button
          onClick={exportToWord}
          disabled={lots.length === 0 || exporting}
          className="mt-2 flex w-full items-center justify-center gap-2 rounded-2xl border py-3 text-[13.5px] font-semibold transition-opacity active:scale-[0.98] disabled:opacity-35"
          style={{ borderColor: NAVY, color: NAVY }}
        >
          <FileDown size={16} />
          {exporting ? "Préparation du document…" : "Exporter vers le modèle Word"}
        </button>
        {exportError && (
          <p className="mt-2 rounded-xl px-3 py-2 text-[11.5px]" style={{ background: "rgba(181,48,27,0.08)", color: "#B5301B" }}>
            {exportError}
          </p>
        )}
        {exportedFile && (
          <div className="mt-2 rounded-xl px-3 py-2.5" style={{ background: "rgba(45,106,79,0.08)" }}>
            <p className="mb-1.5 text-[11.5px]" style={{ color: "#2D6A4F" }}>
              Document prêt. Si le partage n'a rien donné de visible, ouvrez-le directement ici :
            </p>
            <a
              href={exportedFile.url}
              download={exportedFile.filename}
              className="flex items-center justify-center gap-1.5 rounded-lg py-2 text-[12.5px] font-semibold"
              style={{ background: "#2D6A4F", color: "white" }}
            >
              <FileDown size={14} />
              Ouvrir / enregistrer {exportedFile.filename}
            </a>
          </div>
        )}
      </div>

      <Sheet
        open={sheetOpen}
        onClose={() => { setSheetOpen(false); setEditingId(null); }}
        eyebrow={isEditing ? "Modifier le lot" : "Nouveau lot"}
        title={isEditing ? `N° ${lots.findIndex((l) => l.id === editingId) + 1}` : `N° ${lots.length + 1}`}
        footer={
          <button onClick={saveLot} disabled={!canAdd} className="flex w-full items-center justify-center gap-2 rounded-2xl py-3.5 text-[14.5px] font-semibold transition-opacity active:scale-[0.98] disabled:opacity-35" style={{ background: RUST, color: PAPER }}>
            <Stamp size={17} />
            {isEditing ? "Enregistrer les modifications" : `Constater ce lot — ${catLabel(cat)}`}
          </button>
        }
      >
        <label className="mb-1.5 block font-mono text-[10.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>Emplacement</label>
        <ZonePicker zones={zones} value={zone} onChange={setZone} onManage={() => setZoneSheetOpen(true)} />

        <label className="mb-1.5 mt-5 block font-mono text-[10.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>Catégorie</label>
        <CategoryPicker categories={categories} value={cat} onChange={setCat} onManage={() => setCatSheetOpen(true)} />

        <label className="mb-1.5 mt-5 flex items-center justify-between font-mono text-[10.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>
          <span>Désignation</span>
          {dictationSupported && (
            <button
              type="button"
              onClick={toggleDictation}
              className="flex items-center gap-1 rounded-full px-2.5 py-1 normal-case tracking-normal transition-colors"
              style={{
                background: listening ? "#B5301B" : PAPER_DEEP,
                color: listening ? "white" : NAVY,
              }}
            >
              <Mic size={12} className={listening ? "animate-pulse" : ""} />
              <span className="text-[10.5px] font-semibold">{listening ? "Écoute…" : "Dicter"}</span>
            </button>
          )}
        </label>
        {!dictationSupported && (
          <p className="mb-1.5 -mt-1 text-[10.5px]" style={{ color: MUTED }}>
            Dictée vocale non disponible dans ce navigateur (fonctionne sur Chrome/Edge et dans une vraie app mobile).
          </p>
        )}
        <textarea
          value={desig}
          onChange={(e) => setDesig(e.target.value)}
          placeholder="Ex. UNE PLAQUE DE CUISSON PERFOREE BARTSCHER"
          rows={2}
          className="w-full resize-none rounded-xl border px-3.5 py-3 text-[14px] outline-none"
          style={{ borderColor: LINE, background: "white", color: INK }}
        />

        <label className="mb-1.5 mt-5 block font-mono text-[10.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>Valeur de réalisation</label>
        <div className="flex items-center gap-2">
          <div className="relative flex-1">
            <input
              type="number"
              inputMode="numeric"
              value={value}
              onChange={(e) => setValue(e.target.value)}
              disabled={memoire}
              placeholder="0"
              className="w-full rounded-xl border px-3.5 py-3 pr-9 text-[15px] font-mono outline-none tabular-nums disabled:opacity-40"
              style={{ borderColor: LINE, background: "white", color: INK }}
            />
            <span className="pointer-events-none absolute right-3.5 top-1/2 -translate-y-1/2 font-mono text-[14px]" style={{ color: MUTED }}>€</span>
          </div>
          <button
            type="button"
            onClick={() => { setMemoire((m) => !m); setValue(""); }}
            className="shrink-0 rounded-xl border px-3.5 py-3 text-[12.5px] font-semibold"
            style={{ borderColor: memoire ? NAVY : LINE, background: memoire ? NAVY : "white", color: memoire ? PAPER : MUTED }}
          >
            Mémoire
          </button>
        </div>

        <label className="mb-1.5 mt-5 block font-mono text-[10.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>Photos du lot</label>
        <div className="flex flex-wrap gap-2">
          {photos.map((src, i) => (
            <PhotoThumb key={src + i} src={src} onRemove={() => removePhoto(i)} />
          ))}
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={processingPhotos}
            className="flex h-16 w-16 shrink-0 flex-col items-center justify-center gap-1 rounded-lg border-2 border-dashed disabled:opacity-50"
            style={{ borderColor: LINE, color: MUTED }}
            aria-label="Ajouter une photo"
          >
            <Camera size={17} />
            <span className="text-[9px] font-medium text-center leading-tight">{processingPhotos ? "Filigrane…" : "Ajouter"}</span>
          </button>
          <input ref={fileInputRef} type="file" accept="image/*" capture="environment" multiple onChange={handleFiles} className="hidden" />
        </div>
        <p className="mt-1.5 text-[10px]" style={{ color: MUTED }}>
          © SCP Brelivet Tanguy est incrusté automatiquement sur chaque photo.
        </p>
      </Sheet>

      <Sheet open={zoneSheetOpen} onClose={() => setZoneSheetOpen(false)} eyebrow="Configuration" title="Emplacements">
        <p className="mb-4 text-[13px]" style={{ color: MUTED }}>
          Ajoutez ou retirez les zones utilisées pour classer les lots de ce dossier.
        </p>
        <div className="mb-5 space-y-2">
          {zones.map((z) => (
            <div key={z} className="flex items-center justify-between rounded-xl border px-3.5 py-2.5" style={{ borderColor: LINE, background: "white" }}>
              <span className="font-mono text-[12px] font-semibold uppercase tracking-wide" style={{ color: INK }}>{z}</span>
              <button onClick={() => removeZone(z)} disabled={zones.length <= 1} aria-label={`Retirer ${z}`} className="rounded-full p-1 disabled:opacity-30" style={{ color: "#B5301B" }}>
                <Trash2 size={14} />
              </button>
            </div>
          ))}
        </div>
        <label className="mb-1.5 block font-mono text-[10.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>Nouvel emplacement</label>
        <div className="flex gap-2">
          <input
            type="text"
            value={newZone}
            onChange={(e) => setNewZone(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && addZone()}
            placeholder="Ex. TERRASSE"
            className="flex-1 rounded-xl border px-3.5 py-2.5 text-[13.5px] outline-none"
            style={{ borderColor: LINE, background: "white", color: INK }}
          />
          <button onClick={addZone} className="flex items-center justify-center gap-1 rounded-xl px-4 text-[13px] font-semibold" style={{ background: NAVY, color: PAPER }}>
            <Plus size={15} />
            Ajouter
          </button>
        </div>
      </Sheet>

      <Sheet open={catSheetOpen} onClose={() => setCatSheetOpen(false)} eyebrow="Configuration" title="Catégories">
        <p className="mb-4 text-[13px]" style={{ color: MUTED }}>
          Ajoutez ou retirez les catégories utilisées pour classer les lots de ce dossier.
        </p>
        <div className="mb-5 space-y-2">
          {categories.map((c) => (
            <div key={c.key} className="flex items-center justify-between rounded-xl border px-3.5 py-2.5" style={{ borderColor: LINE, background: "white" }}>
              <span className="text-[13px] font-medium" style={{ color: INK }}>{c.label}</span>
              <button onClick={() => removeCategory(c.key)} disabled={categories.length <= 1} aria-label={`Retirer ${c.label}`} className="rounded-full p-1 disabled:opacity-30" style={{ color: "#B5301B" }}>
                <Trash2 size={14} />
              </button>
            </div>
          ))}
        </div>
        <label className="mb-1.5 block font-mono text-[10.5px] font-semibold uppercase tracking-wide" style={{ color: MUTED }}>Nouvelle catégorie</label>
        <div className="flex gap-2">
          <input
            type="text"
            value={newCatLabel}
            onChange={(e) => setNewCatLabel(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && addCategory()}
            placeholder="Ex. Véhicules"
            className="flex-1 rounded-xl border px-3.5 py-2.5 text-[13.5px] outline-none"
            style={{ borderColor: LINE, background: "white", color: INK }}
          />
          <button onClick={addCategory} className="flex items-center justify-center gap-1 rounded-xl px-4 text-[13px] font-semibold" style={{ background: NAVY, color: PAPER }}>
            <Plus size={15} />
            Ajouter
          </button>
        </div>
      </Sheet>
    </>
  );
}

// ---------------------------------------------------------------------------
// Root
// ---------------------------------------------------------------------------
export default function InventoryFieldApp() {
  const [screen, setScreen] = useState("menu"); // "menu" | "app"
  const [tab, setTab] = useState("nouveau");
  const [requerants, setRequerants] = useState(SEED_REQUERANTS);
  const [caseInfo, setCaseInfo] = useState({
    name: "SARL IZOA — La Caserne",
    address: "108 Av. de la Gare, Concarneau",
    heureDebut: "14:02",
    formeJuridique: "SARL",
    denomination: "IZOA",
    capital: "1 000",
    rcsVille: "Quimper",
    rcsNumero: "979 626 306",
    adresseSiege: "108 Avenue de la Gare, 29900 Concarneau",
    representant: "Emmelyne GOBIN",
    qualiteRepresentant: "Gérante",
    tribunal: "Quimper",
    dateJugement: "",
    dateInventaire: "",
    requerant: SEED_REQUERANTS[0],
  });
  const [caseEditOpen, setCaseEditOpen] = useState(false);
  const [draftName, setDraftName] = useState(caseInfo.name);
  const [draftAddress, setDraftAddress] = useState(caseInfo.address);
  const [inventoryKey, setInventoryKey] = useState(0);
  const [dossierId, setDossierId] = useState(null);
  const [loadedInitialData, setLoadedInitialData] = useState(null);
  const [embedded, setEmbedded] = useState(false);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get("view") === "dashboard") {
      setEmbedded(true);
      setScreen("dashboard");
    }
  }, []);

  async function handleOpenDossier(id) {
    const { data, error } = await loadDossier(id);
    if (error || !data) return;
    const payload = data.payload || {};
    setCaseInfo(payload.caseInfo || {});
    setLoadedInitialData({ lots: payload.lots || [], zones: payload.zones, categories: payload.categories });
    setDossierId(id);
    setInventoryKey((k) => k + 1);
    setScreen("app");
    setTab("inventaire");
  }

  function handleCreateCase(c, requerant) {
    setCaseInfo({
      name: `${c.formeJuridique} ${c.denomination}`.trim(),
      address: c.adresseLieu || c.adresseSiege,
      heureDebut: c.heureDebut,
      formeJuridique: c.formeJuridique,
      denomination: c.denomination,
      capital: c.capital,
      rcsVille: c.rcsVille,
      rcsNumero: c.rcsNumero,
      adresseSiege: c.adresseSiege,
      representant: c.representant,
      qualiteRepresentant: c.qualiteRepresentant,
      tribunal: c.tribunal,
      dateJugement: c.dateJugement,
      dateInventaire: c.dateInventaire,
      requerant,
    });
    setDossierId(null);
    setLoadedInitialData(null);
    setInventoryKey((k) => k + 1); // remounts InventoryTab with a clean slate
    setTab("inventaire");
  }

  return (
    <div className="flex min-h-screen items-center justify-center" style={{ background: PAPER }}>
      <style>{`
        @keyframes stampIn {
          0% { opacity: 0; transform: scale(1.8) rotate(-12deg); }
          55% { opacity: 1; transform: scale(0.94) rotate(-12deg); }
          75% { transform: scale(1.04) rotate(-12deg); }
          100% { opacity: 0; transform: scale(1) rotate(-12deg); }
        }
        .scrollbar-none::-webkit-scrollbar { display: none; }
        .scrollbar-none { -ms-overflow-style: none; scrollbar-width: none; }
      `}</style>

      <div className="relative flex h-[100dvh] w-full max-w-[520px] flex-col overflow-hidden" style={{ background: PAPER }}>
        {screen === "dashboard" ? (
          <Dashboard onOpenDossier={handleOpenDossier} onBack={() => setScreen("menu")} embedded={embedded} />
        ) : screen === "menu" ? (
          <MenuScreen onSelect={() => { setScreen("app"); setTab("nouveau"); }} onOpenDashboard={() => setScreen("dashboard")} />
        ) : (
        <>
        <div className="flex min-h-0 flex-1 flex-col">
          {tab === "inventaire" ? (
            <InventoryTab
              key={inventoryKey}
              caseInfo={caseInfo}
              onEditCase={() => { setDraftName(caseInfo.name); setDraftAddress(caseInfo.address); setCaseEditOpen(true); }}
              initialData={loadedInitialData}
              dossierId={dossierId}
              onSaved={(id) => setDossierId(id)}
            />
          ) : (
            <>
              <div className="shrink-0 px-5 pb-4 pt-5" style={{ background: NAVY, color: PAPER }}>
                <button
                  onClick={() => setScreen("menu")}
                  className="mb-2 flex items-center gap-1 text-[11px] font-semibold"
                  style={{ color: "#C3CEDE" }}
                >
                  <ArrowLeft size={13} />
                  Changer de type de constat
                </button>
                <p className="font-mono text-[10px] font-semibold uppercase tracking-[0.14em]" style={{ color: "#9FB0C9" }}>Nouveau dossier</p>
                <h1 className="mt-0.5 text-[18px] font-semibold leading-tight">Ouvrir un inventaire</h1>
                <p className="mt-1 text-[11.5px]" style={{ color: "#C3CEDE" }}>Requérant, défendeur, jugement et lieu</p>
              </div>
              <NewCaseTab
                requerants={requerants}
                onAddRequerant={(r) => setRequerants((prev) => [...prev, r])}
                onCreateCase={handleCreateCase}
              />
            </>
          )}
        </div>

        {/* Bottom tab bar */}
        <div className="grid shrink-0 grid-cols-2 border-t" style={{ borderColor: LINE, background: PAPER }}>
          <button
            onClick={() => setTab("inventaire")}
            className="flex flex-col items-center gap-1 py-2.5"
            style={{ color: tab === "inventaire" ? NAVY : MUTED }}
          >
            <Stamp size={19} strokeWidth={tab === "inventaire" ? 2.4 : 1.8} />
            <span className="text-[10.5px] font-semibold">Inventaire</span>
          </button>
          <button
            onClick={() => setTab("nouveau")}
            className="flex flex-col items-center gap-1 py-2.5"
            style={{ color: tab === "nouveau" ? NAVY : MUTED }}
          >
            <FilePlus2 size={19} strokeWidth={tab === "nouveau" ? 2.4 : 1.8} />
            <span className="text-[10.5px] font-semibold">Nouveau dossier</span>
          </button>
        </div>

        <Sheet
          open={caseEditOpen}
          onClose={() => setCaseEditOpen(false)}
          eyebrow="Dossier"
          title="Modifier le dossier"
          z={40}
          footer={
            <button
              onClick={() => {
                setCaseInfo((p) => ({ ...p, name: draftName.trim() || p.name, address: draftAddress.trim() || p.address }));
                setCaseEditOpen(false);
              }}
              className="flex w-full items-center justify-center gap-2 rounded-2xl py-3.5 text-[14.5px] font-semibold active:scale-[0.98] transition-transform"
              style={{ background: NAVY, color: PAPER }}
            >
              <Check size={17} />
              Enregistrer
            </button>
          }
        >
          <Field label="Nom du dossier" value={draftName} onChange={(e) => setDraftName(e.target.value)} placeholder="Ex. SARL IZOA — La Caserne" />
          <Field label="Adresse de l'inventaire" value={draftAddress} onChange={(e) => setDraftAddress(e.target.value)} placeholder="Ex. 108 Avenue de la Gare, 29900 Concarneau" />
        </Sheet>
        </>
        )}
      </div>
    </div>
  );
}
