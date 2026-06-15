import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  ArrowLeft,
  ArrowLeftRight,
  BookOpen,
  Check,
  CloudUpload,
  FileSpreadsheet,
  Layers,
  RotateCcw,
  Shuffle,
  ToggleLeft,
  ToggleRight,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";
import "./styles.css";

type Section = "home" | "words" | "definitions";
type Direction = "front-back" | "back-front" | "mixed";

type BaseCard = {
  id: string;
  front: string;
  back: string;
  extra?: string;
};

type StudyCard = BaseCard & {
  frontSide: "front" | "back";
};

type StudyConfig = {
  section: Exclude<Section, "home">;
  title: string;
  subtitle: string;
  frontName: string;
  backName: string;
  fileHelp: string;
  sampleCards: BaseCard[];
};

type SavedCards = {
  cards: BaseCard[];
  fileName: string;
  savedAt: string;
};

type SharedCards = SavedCards;

const studyConfigs: Record<Exclude<Section, "home">, StudyConfig> = {
  words: {
    section: "words",
    title: "Kelime Kartları",
    subtitle: "İngilizce ve Türkçe kelime çiftleriyle çalış.",
    frontName: "İngilizce",
    backName: "Türkçe",
    fileHelp: "Excel'de ilk iki sütuna İngilizce ve Türkçe kelimeleri yaz.",
    sampleCards: [
      { id: "word-sample-1", front: "curious", back: "meraklı" },
      { id: "word-sample-2", front: "gentle", back: "nazik" },
      { id: "word-sample-3", front: "practice", back: "alıştırma" },
      { id: "word-sample-4", front: "remember", back: "hatırlamak" },
    ],
  },
  definitions: {
    section: "definitions",
    title: "Tanım Kartları",
    subtitle: "Kelimeyi gör, tanımını çevirerek kontrol et.",
    frontName: "Kelime",
    backName: "Tanım",
    fileHelp: "Excel'de ilk iki sütuna kelime ve tanımını yaz.",
    sampleCards: [
      { id: "definition-sample-1", front: "noun", back: "Kişi, yer, nesne veya kavram adı." },
      { id: "definition-sample-2", front: "adjective", back: "Bir ismi niteleyen ya da açıklayan kelime." },
      { id: "definition-sample-3", front: "synonym", back: "Anlamı aynı ya da çok yakın olan kelime." },
      { id: "definition-sample-4", front: "sentence", back: "Tam bir düşünce anlatan kelime grubu." },
    ],
  },
};

const normalize = (value: unknown) => String(value ?? "").trim();

const slug = (value: string) =>
  value
    .toLocaleLowerCase("tr-TR")
    .replace(/\s+/g, "-")
    .replace(/[^\p{L}\p{N}-]+/gu, "");

const makeId = (front: string, back: string, index: number) => `${slug(front)}-${slug(back)}-${index}`;

const shuffleCards = <T,>(items: T[]) => {
  const copy = [...items];
  for (let index = copy.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [copy[index], copy[swapIndex]] = [copy[swapIndex], copy[index]];
  }
  return copy;
};

const detectColumns = (rows: unknown[][], config: StudyConfig) => {
  const firstRow = rows[0]?.map((cell) => normalize(cell).toLocaleLowerCase("tr-TR")) ?? [];
  const frontLabels =
    config.section === "words" ? ["english", "ingilizce", "en", "word", "kelime"] : ["word", "kelime", "terim"];
  const backLabels =
    config.section === "words"
      ? ["turkish", "türkçe", "turkce", "tr", "anlam", "karşılık", "karsilik"]
      : ["definition", "tanım", "tanim", "açıklama", "aciklama", "anlam"];
  const extraLabels = ["türkçe karşılığı", "turkce karsiligi", "türkçe karşılık", "turkce karsilik", "türkçe", "turkce", "tr"];

  const frontIndex = firstRow.findIndex((cell) => frontLabels.includes(cell));
  const backIndex = firstRow.findIndex((cell) => backLabels.includes(cell));
  const extraIndex =
    config.section === "definitions"
      ? firstRow.findIndex((cell, index) => extraLabels.includes(cell) && index !== frontIndex && index !== backIndex)
      : -1;

  if (frontIndex >= 0 && backIndex >= 0 && frontIndex !== backIndex) {
    return { frontIndex, backIndex, extraIndex, startRow: 1 };
  }

  const width = Math.max(...rows.map((row) => row.length), 0);
  const columns = Array.from({ length: width }, (_, columnIndex) => ({
    index: columnIndex,
    count: rows.reduce((count, row) => count + (normalize(row[columnIndex]) ? 1 : 0), 0),
  }))
    .filter((column) => column.count > 0)
    .sort((a, b) => b.count - a.count || a.index - b.index);

  return {
    frontIndex: columns[0]?.index ?? 0,
    backIndex: columns[1]?.index ?? 1,
    extraIndex: config.section === "definitions" ? (columns[2]?.index ?? -1) : -1,
    startRow: 0,
  };
};

const parseRowsToCards = (rows: unknown[][], config: StudyConfig) => {
  if (!rows.length) return [];

  const { frontIndex, backIndex, extraIndex, startRow } = detectColumns(rows, config);

  return rows
    .slice(startRow)
    .map((row, index) => {
      const front = normalize(row[frontIndex]);
      const back = normalize(row[backIndex]);
      const extra = extraIndex >= 0 ? normalize(row[extraIndex]) : "";
      return { id: makeId(front, back, index), front, back, extra };
    })
    .filter((card) => card.front && card.back);
};

function buildStudyDeck(cards: BaseCard[], direction: Direction): StudyCard[] {
  return shuffleCards(
    cards.map((card) => ({
      ...card,
      frontSide:
        direction === "mixed"
          ? Math.random() > 0.5
            ? "front"
            : "back"
          : direction === "front-back"
            ? "front"
            : "back",
    })),
  );
}

const storageKey = (section: StudyConfig["section"]) => `kori:${section}:last-excel`;

const loadSavedCards = (config: StudyConfig): SavedCards | null => {
  try {
    const raw = window.localStorage.getItem(storageKey(config.section));
    if (!raw) return null;
    const saved = JSON.parse(raw) as SavedCards;
    if (!Array.isArray(saved.cards) || !saved.cards.length) return null;
    return saved;
  } catch {
    return null;
  }
};

const saveCards = (config: StudyConfig, cards: BaseCard[], fileName: string) => {
  const saved: SavedCards = {
    cards,
    fileName,
    savedAt: new Date().toISOString(),
  };
  window.localStorage.setItem(storageKey(config.section), JSON.stringify(saved));
};

const clearSavedCards = (config: StudyConfig) => {
  window.localStorage.removeItem(storageKey(config.section));
};

async function fetchSharedCards(config: StudyConfig): Promise<SharedCards | null> {
  const response = await fetch(`/api/shared-cards?section=${config.section}`, {
    headers: { Accept: "application/json" },
  });

  if (!response.ok) return null;

  const shared = (await response.json()) as SharedCards;
  if (!Array.isArray(shared.cards) || !shared.cards.length) return null;
  return shared;
}

async function publishSharedCards(config: StudyConfig, payload: SavedCards, adminPin: string) {
  const response = await fetch(`/api/shared-cards?section=${config.section}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      ...(adminPin.trim() ? { "x-admin-pin": adminPin.trim() } : {}),
    },
    body: JSON.stringify(payload),
  });

  const result = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(result.error || "Paylaşım başarısız oldu.");
  }

  return result as SharedCards;
}

function HomeScreen({ onOpen }: { onOpen: (section: Exclude<Section, "home">) => void }) {
  return (
    <main className="app-shell">
      <section className="home-layout">
        <div className="home-header">
          <div className="brand large">
            <div className="brand-mark">K</div>
            <div>
              <h1>Kori ile İngilizce</h1>
              <p>Kendi Excel dosyalarınla kart çalış.</p>
            </div>
          </div>
        </div>

        <div className="mode-grid">
          <button className="mode-card" type="button" onClick={() => onOpen("words")}>
            <BookOpen size={34} />
            <span>Kelime Kartları</span>
            <small>İngilizce-Türkçe kelime eşleştirme</small>
          </button>

          <button className="mode-card accent" type="button" onClick={() => onOpen("definitions")}>
            <Layers size={34} />
            <span>Tanım Kartları</span>
            <small>Kelimeler ve açıklamalarıyla çalışma</small>
          </button>
        </div>
      </section>
    </main>
  );
}

function StudyScreen({ config, onBack }: { config: StudyConfig; onBack: () => void }) {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const savedCards = useMemo(() => loadSavedCards(config), [config]);
  const initialCards = savedCards?.cards ?? config.sampleCards;
  const [cards, setCards] = useState<BaseCard[]>(initialCards);
  const [deck, setDeck] = useState<StudyCard[]>(() => buildStudyDeck(initialCards, "front-back"));
  const [direction, setDirection] = useState<Direction>("front-back");
  const [currentIndex, setCurrentIndex] = useState(0);
  const [isFlipped, setIsFlipped] = useState(false);
  const [removeKnown, setRemoveKnown] = useState(true);
  const [knownIds, setKnownIds] = useState<Set<string>>(new Set());
  const [missedIds, setMissedIds] = useState<Set<string>>(new Set());
  const [savedFileName, setSavedFileName] = useState(savedCards?.fileName ?? "");
  const [sharedCards, setSharedCards] = useState<SharedCards | null>(null);
  const [adminPin, setAdminPin] = useState("");
  const [isPublishing, setIsPublishing] = useState(false);
  const [status, setStatus] = useState(
    savedCards
      ? `${savedCards.fileName} kayıtlı. Kaldığın yerden çalışabilir veya sıfırlayabilirsin.`
      : "Örnek kartlarla başlayabilir veya kendi Excel dosyanı yükleyebilirsin.",
  );

  useEffect(() => {
    let isCurrent = true;

    fetchSharedCards(config)
      .then((shared) => {
        if (!isCurrent || !shared) return;

        setSharedCards(shared);

        if (!savedCards) {
          setCards(shared.cards);
          setDeck(buildStudyDeck(shared.cards, "front-back"));
          setCurrentIndex(0);
          setIsFlipped(false);
          setKnownIds(new Set());
          setMissedIds(new Set());
          setStatus(`${shared.fileName} paylaşılan kartlardan yüklendi.`);
        }
      })
      .catch(() => {
        if (!savedCards) {
          setStatus("Örnek kartlarla başlayabilir veya kendi Excel dosyanı yükleyebilirsin.");
        }
      });

    return () => {
      isCurrent = false;
    };
  }, [config, savedCards]);

  const currentCard = deck[currentIndex];
  const isFinished = deck.length === 0 || currentIndex >= deck.length;

  const stats = useMemo(
    () => ({
      total: cards.length,
      deck: deck.length,
      known: knownIds.size,
      missed: missedIds.size,
    }),
    [cards.length, deck.length, knownIds.size, missedIds.size],
  );

  const resetStudy = (nextCards = cards, nextDirection = direction, onlyMissed = false) => {
    const sourceCards = onlyMissed ? nextCards.filter((card) => missedIds.has(card.id)) : nextCards;
    setDeck(buildStudyDeck(sourceCards, nextDirection));
    setCurrentIndex(0);
    setIsFlipped(false);
    if (!onlyMissed) {
      setKnownIds(new Set());
      setMissedIds(new Set());
    }
  };

  const handleDirection = (nextDirection: Direction) => {
    setDirection(nextDirection);
    resetStudy(cards, nextDirection);
  };

  const handleFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, blankrows: false });
      const parsedCards = parseRowsToCards(rows, config);

      if (!parsedCards.length) {
        setStatus("Dosyada iki dolu sütun bulamadım. İlk iki sütunu doldurup tekrar dene.");
        return;
      }

      setCards(parsedCards);
      setDeck(buildStudyDeck(parsedCards, direction));
      setCurrentIndex(0);
      setIsFlipped(false);
      setKnownIds(new Set());
      setMissedIds(new Set());
      setSavedFileName(file.name);
      saveCards(config, parsedCards, file.name);
      setStatus(`${file.name} içinden ${parsedCards.length} kart yüklendi.`);
    } catch {
      setStatus("Dosyayı okuyamadım. .xlsx, .xls veya .csv formatında tekrar yüklemeyi dene.");
    } finally {
      event.target.value = "";
    }
  };

  const resetSavedExcel = () => {
    const fallbackCards = sharedCards?.cards ?? config.sampleCards;
    clearSavedCards(config);
    setCards(fallbackCards);
    setDeck(buildStudyDeck(fallbackCards, direction));
    setCurrentIndex(0);
    setIsFlipped(false);
    setKnownIds(new Set());
    setMissedIds(new Set());
    setSavedFileName("");
    setStatus(sharedCards ? `${sharedCards.fileName} paylaşılan kartlarına dönüldü.` : "Kayıtlı Excel temizlendi. Örnek kartlara dönüldü.");
  };

  const publishCurrentCards = async () => {
    if (!cards.length) return;

    setIsPublishing(true);
    try {
      const shared = await publishSharedCards(
        config,
        {
          cards,
          fileName: savedFileName || sharedCards?.fileName || `${config.title}.xlsx`,
          savedAt: new Date().toISOString(),
        },
        adminPin,
      );
      setSharedCards(shared);
      setStatus(`${shared.fileName} herkes için paylaşılan kart olarak yayınlandı.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Paylaşım başarısız oldu.");
    } finally {
      setIsPublishing(false);
    }
  };

  const markCard = (result: "known" | "missed") => {
    if (!currentCard) return;

    if (result === "known") {
      setKnownIds((previous) => new Set(previous).add(currentCard.id));
      setMissedIds((previous) => {
        const next = new Set(previous);
        next.delete(currentCard.id);
        return next;
      });
    } else {
      setMissedIds((previous) => new Set(previous).add(currentCard.id));
      setKnownIds((previous) => {
        const next = new Set(previous);
        next.delete(currentCard.id);
        return next;
      });
    }

    setIsFlipped(false);
    setCurrentIndex((index) => index + 1);
  };

  const repeatMissed = () => {
    const missedCards = cards.filter((card) => missedIds.has(card.id));
    if (!missedCards.length) return;
    setDeck(buildStudyDeck(missedCards, direction));
    setCurrentIndex(0);
    setIsFlipped(false);
  };

  const repeatAll = () => resetStudy(cards, direction);

  const frontText = currentCard?.frontSide === "front" ? currentCard.front : currentCard?.back;
  const backText = currentCard?.frontSide === "front" ? currentCard.back : currentCard?.front;
  const frontLabel = currentCard?.frontSide === "front" ? config.frontName : config.backName;
  const backLabel = currentCard?.frontSide === "front" ? config.backName : config.frontName;
  const visibleText = isFlipped ? backText : frontText;
  const showDefinitionBack = config.section === "definitions" && isFlipped;
  const definitionLength = visibleText?.length ?? 0;
  const definitionSizeClass =
    definitionLength > 260 ? "long" : definitionLength > 150 ? "medium" : definitionLength > 90 ? "short" : "";
  const showExtraTranslation = config.section === "definitions" && isFlipped && currentCard?.frontSide === "front" && currentCard.extra;

  return (
    <main className="app-shell">
      <section className="workspace">
        <aside className="sidebar">
          <button className="back-button" type="button" onClick={onBack}>
            <ArrowLeft size={18} />
            Ana ekran
          </button>

          <div className="brand">
            <div className="brand-mark">K</div>
            <div>
              <h1>{config.title}</h1>
              <p>{config.subtitle}</p>
            </div>
          </div>

          <button className="upload-button" type="button" onClick={() => fileInputRef.current?.click()}>
            <Upload size={20} />
            Excel / CSV yükle
          </button>
          <input
            ref={fileInputRef}
            className="visually-hidden"
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFile}
          />

          {savedFileName ? (
            <button className="clear-button" type="button" onClick={resetSavedExcel}>
              <Trash2 size={18} />
              Kayıtlı Excel'i sıfırla
            </button>
          ) : null}

          <div className="publish-box">
            <input
              value={adminPin}
              onChange={(event) => setAdminPin(event.target.value)}
              placeholder="Yayın PIN"
              type="password"
            />
            <button className="publish-button" type="button" onClick={publishCurrentCards} disabled={isPublishing || !cards.length}>
              <CloudUpload size={18} />
              {isPublishing ? "Yayınlanıyor" : "Paylaşılan olarak yayınla"}
            </button>
          </div>

          <div className="status-box">
            <FileSpreadsheet size={18} />
            <p>
              {status}
              <span>{config.fileHelp}</span>
            </p>
          </div>

          <div className="control-group">
            <span>Çalışma modu</span>
            <div className="segmented-control" aria-label="Çalışma modu">
              <button
                className={direction === "front-back" ? "active" : ""}
                type="button"
                onClick={() => handleDirection("front-back")}
              >
                {config.frontName} → {config.backName}
              </button>
              <button
                className={direction === "back-front" ? "active" : ""}
                type="button"
                onClick={() => handleDirection("back-front")}
              >
                {config.backName} → {config.frontName}
              </button>
              <button className={direction === "mixed" ? "active" : ""} type="button" onClick={() => handleDirection("mixed")}>
                Karışık
              </button>
            </div>
          </div>

          <button className={`toggle-row ${removeKnown ? "enabled" : ""}`} type="button" onClick={() => setRemoveKnown((value) => !value)}>
            {removeKnown ? <ToggleRight size={28} /> : <ToggleLeft size={28} />}
            <span>Bildiklerimi tekrar çıkarma</span>
          </button>

          <div className="stats-grid" aria-label="Çalışma özeti">
            <div>
              <strong>{stats.total}</strong>
              <span>Toplam</span>
            </div>
            <div>
              <strong>{stats.known}</strong>
              <span>Bildim</span>
            </div>
            <div>
              <strong>{stats.missed}</strong>
              <span>Tekrar</span>
            </div>
          </div>
        </aside>

        <section className="study-area">
          <div className="topbar">
            <button className="icon-button" type="button" title="Kartları karıştır" onClick={() => resetStudy()}>
              <Shuffle size={20} />
            </button>
            <div className="progress">
              <span>
                {isFinished ? deck.length : currentIndex + 1} / {deck.length}
              </span>
              <div>
                <i style={{ width: `${deck.length ? Math.min(((currentIndex + 1) / deck.length) * 100, 100) : 0}%` }} />
              </div>
            </div>
            <button className="icon-button" type="button" title="Baştan başlat" onClick={repeatAll}>
              <RotateCcw size={20} />
            </button>
          </div>

          {isFinished ? (
            <div className="finished-panel">
              <h2>Bu tur bitti</h2>
              {removeKnown && missedIds.size > 0 ? (
                <p>Bildiklerin çıkarıldı. Şimdi sadece bilmediklerini tekrar çalışabilirsin.</p>
              ) : (
                <p>Bilmediğin kart kalmadı. İstersen tüm kartları baştan tekrar et.</p>
              )}
              <div className="action-row">
                {removeKnown && missedIds.size > 0 ? (
                  <button className="primary-action" type="button" onClick={repeatMissed}>
                    <RotateCcw size={20} />
                    Bilmediklerimi çalış
                  </button>
                ) : null}
                <button className="secondary-action neutral" type="button" onClick={repeatAll}>
                  <RotateCcw size={20} />
                  Baştan tekrar et
                </button>
              </div>
            </div>
          ) : (
            <>
              <button className={`flashcard ${isFlipped ? "flipped" : ""}`} type="button" onClick={() => setIsFlipped((value) => !value)}>
                <span className="card-mode">
                  <ArrowLeftRight size={18} />
                  {isFlipped ? backLabel : frontLabel}
                </span>
                <span className={showDefinitionBack ? `card-definition ${definitionSizeClass}` : "card-word"}>
                  {visibleText}
                </span>
                {showExtraTranslation ? <span className="card-extra">{currentCard.extra}</span> : null}
                <span className="card-hint">Kartı çevirmek için tıkla</span>
              </button>

              <div className="action-row">
                <button className="secondary-action" type="button" onClick={() => markCard("missed")}>
                  <X size={20} />
                  Bilemedim
                </button>
                <button className="primary-action" type="button" onClick={() => markCard("known")}>
                  <Check size={20} />
                  Bildim
                </button>
              </div>
            </>
          )}
        </section>
      </section>
    </main>
  );
}

function App() {
  const [section, setSection] = useState<Section>("home");

  if (section === "home") {
    return <HomeScreen onOpen={setSection} />;
  }

  return <StudyScreen key={section} config={studyConfigs[section]} onBack={() => setSection("home")} />;
}

createRoot(document.getElementById("root")!).render(<App />);
