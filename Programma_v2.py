import os
import re
import io
import json
import time
import base64
import queue
import threading
import unicodedata
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image as PILImage, ImageOps
from groq import Groq, GroqError
from pydantic import BaseModel, Field, field_validator
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

# Kraken è opzionale: se non è installato o non c'è un modello, il programma
# funziona comunque solo con Groq, senza bloccarsi.
try:
    from kraken import pageseg, rpred, binarization
    from kraken.lib import models as kraken_models
    KRAKEN_DISPONIBILE = True
except ImportError:
    KRAKEN_DISPONIBILE = False

# =====================================================================
# CONFIGURAZIONI GLOBALI
# =====================================================================
GROQ_API_KEY = "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"
FILE_EXCEL = "archivio_schede.xlsx"
MODELLO_VISION = "qwen/qwen3.6-27b"
LATO_MASSIMO_IMMAGINE = 2600
MAX_RETRIES = 3  # Tentativi automatici in caso di errore di rete/API
SALVA_OGNI_N_SCHEDE = 10  # Checkpoint: salva su Excel ogni N schede per non perdere il lavoro in caso di crash
MAX_TOKENS_RISPOSTA = 2000  # Margine ampio: con il campo "note" a lunghezza variabile, un limite troppo
                             # basso tronca il JSON a metà e lo rende non parsabile -> causa nota di schede "fallite"
                             # NB: se si riattiva il reasoning (reasoning_effort="default" più sotto),
                             # conviene alzare questo valore a 3000-4000, perché il "pensiero" del modello
                             # consuma token prima ancora di iniziare a scrivere il JSON.

PERCORSO_MODELLO_KRAKEN = "McCATMuS_nfd_nofix_V1.mlmodel"  
CAMPI_DA_CONFRONTARE = ["numero", "cognome", "nome", "di", "citta", "data", "tipo", "categoria", "dimensioni", "note"]

# =====================================================================
# 1. SCHEMA DEI DATI (Pydantic v2)
# =====================================================================
class SchedaBibliografica(BaseModel):
    numero: str = Field(default="", description="Codice numerico del fondo")
    cognome: str = Field(default="", description="Cognome principale")
    nome: str = Field(default="", description="Nome di battesimo")
    di: str = Field(default="", description="Riferimento persona/ente")
    citta: str = Field(default="", description="Città o luogo")
    data: str = Field(default="", description="Data completa o anno")
    tipo: str = Field(default="", description="Lettera di classificazione")
    categoria: str = Field(default="", description="Categoria o sezione d'archivio")
    dimensioni: str = Field(default="", description="Dimensioni della scheda, formato o supporto (es. in-8, in-4, 12x18 cm, ecc.)")
    note: str = Field(default="", description="Tutto il resto del testo o informazioni extra presenti sulla scheda non rientranti nei campi precedenti")
    confidenza: str = Field(default="media", description="Auto-valutazione del modello: alta, media o bassa, in base a quanto è leggibile la scheda nel suo complesso")

    @field_validator('*', mode='before')
    @classmethod
    def pulisci_stringhe(cls, valore):
        if isinstance(valore, list):
            valore = ", ".join(str(v) for v in valore)
        
        if isinstance(valore, str):
            valore_pulito = re.sub(r'\s+', ' ', valore).strip()
            if valore_pulito.lower() in ["n/a", "none", "null", "non presente", "non indicato", ""]:
                return ""
            return valore_pulito
        return str(valore) if valore is not None else ""

# =====================================================================
# 2. MOTORE DI ELABORAZIONE
# =====================================================================
ULTIMO_ERRORE_KRAKEN = ""  # Salva l'ultimo errore incontrato da Kraken, per poterlo mostrare nel log della GUI
                            # invece di nasconderlo silenziosamente (in precedenza qualsiasi errore veniva
                            # ignorato e il programma proseguiva come se Kraken non fosse configurato,
                            # rendendo impossibile capire perché il doppio controllo non scattava).

def leggi_con_kraken(percorso_immagine: str) -> str:
    """Esegue segmentazione + riconoscimento HTR con Kraken, restituisce testo grezzo riga per riga.
    Restituisce stringa vuota se Kraken non è disponibile o il modello non è configurato."""
    global ULTIMO_ERRORE_KRAKEN
    if not KRAKEN_DISPONIBILE or not PERCORSO_MODELLO_KRAKEN:
        return ""

    if not os.path.isfile(PERCORSO_MODELLO_KRAKEN):
        ULTIMO_ERRORE_KRAKEN = f"File modello non trovato al percorso: '{os.path.abspath(PERCORSO_MODELLO_KRAKEN)}'"
        return ""

    try:
        img = PILImage.open(percorso_immagine).convert("L")  # scala di grigi, richiesto da Kraken
        img_binarizzata = binarization.nlbin(img)  # Kraken richiede un'immagine bi-livello (bianco/nero puro)
                                                    # per la segmentazione: la scala di grigi da sola non basta
        modello = kraken_models.load_any(PERCORSO_MODELLO_KRAKEN)
        righe_segmentate = pageseg.segment(img_binarizzata)

        righe_testo = [record.prediction for record in rpred.rpred(modello, img_binarizzata, righe_segmentate)]
        return "\n".join(righe_testo)
    except Exception as e:
        ULTIMO_ERRORE_KRAKEN = f"{type(e).__name__}: {e}"
        return ""

def prepara_immagine_per_invio(percorso_immagine: str) -> tuple[str, str]:
    """Ottimizza l'immagine riducendo la risoluzione e applicando autocontrasto."""
    with PILImage.open(percorso_immagine) as img:
        img = img.convert("RGB")
        
        larg, alt = img.size
        lato_maggiore = max(larg, alt)
        if lato_maggiore > LATO_MASSIMO_IMMAGINE:
            fattore = LATO_MASSIMO_IMMAGINE / lato_maggiore
            nuova_dim = (int(larg * fattore), int(alt * fattore))
            img = img.resize(nuova_dim, PILImage.LANCZOS)

        img = ImageOps.autocontrast(img, cutoff=1)

        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=95)
        base64_str = base64.b64encode(buffer.getvalue()).decode("utf-8")
        
    return base64_str, "image/jpeg"

def _estrai_json_pulito(testo_grezzo: str) -> str:
    """Rende il parsing robusto anche se il modello aggiunge testo/backtick attorno al JSON
    (es. blocchi ```json ... ``` o frasi introduttive), cosa più probabile ora che i campi
    sono meno rigidi. Isola la porzione tra la prima '{' e l'ultima '}'."""
    testo = testo_grezzo.strip()
    testo = re.sub(r'^```(?:json)?', '', testo).strip()
    testo = re.sub(r'```$', '', testo).strip()

    inizio = testo.find('{')
    fine = testo.rfind('}')
    if inizio != -1 and fine != -1 and fine > inizio:
        testo = testo[inizio:fine + 1]
    return testo

def analizza_scheda_con_groq_vision(client: Groq, percorso_immagine: str, bozza_ocr: str = "") -> dict:
    """Invia l'immagine a Groq con schema JSON semplificato per evitare errori di validazione su .exe"""
    base64_immagine, mime_type = prepara_immagine_per_invio(percorso_immagine)

    blocco_bozza = ""
    if bozza_ocr.strip():
        blocco_bozza = (
            "\nBOZZA OCR (supporto con possibili errori):\n"
            f"{bozza_ocr.strip()}\n"
        )

    # Schema JSON semplificato senza metadati di Pydantic che causano il fallimento in Groq
    schema_json_pulito = {
        "type": "object",
        "properties": {
            "numero": {"type": "string"},
            "cognome": {"type": "string"},
            "nome": {"type": "string"},
            "di": {"type": "string"},
            "citta": {"type": "string"},
            "data": {"type": "string"},
            "tipo": {"type": "string"},
            "categoria": {"type": "string"},
            "dimensioni": {"type": "string"},
            "note": {"type": "string"},
            "confidenza": {"type": "string", "enum": ["alta", "media", "bassa"]}
        },
        "required": ["numero", "cognome", "nome", "di", "citta", "data", "tipo", "categoria", "dimensioni", "note", "confidenza"]
    }

    schema_json_str = json.dumps(schema_json_pulito, indent=2)

    # Prompt ottimizzato per mettere Kraken al centro della lettura
    if bozza_ocr.strip():
        istruzione_principio = (
            "RUOLO: Sei uno specialista di strutturazione dati per archivi storici.\n"
            "FONTE PRIMARIA DEL TESTO: La trascrizione HTR qui sotto è stata generata dal modello Kraken, "
            "specializzato sulla scrittura di questa collezione. Considera questo testo come la LETTURA DI RIFERIMENTO.\n\n"
            f"--- TRASCRIZIONE KRAKEN ---\n{bozza_ocr.strip()}\n---------------------------\n\n"
            "IL TUO COMPITO:\n"
            "1. Analizza l'immagine della scheda solo per capire la disposizione spaziale e la struttura dei dati.\n"
            "2. Prendi il testo fornito da Kraken e distribuiscilo nei giusti campi JSON (numero, cognome, nome, di, citta, data, tipo, categoria, dimensioni, note).\n"
            "3. Non modificare le parole trascritte da Kraken a meno che nell'immagine non ci sia un'evidente correzione visiva o un errore palese di segmentazione.\n"
            "4. Se una riga o parola di Kraken non appartiene ai campi specifici, inseriscila nel campo 'note'.\n"
        )
    else:
        istruzione_principio = (
            "Sei un archivista storico. Leggi la FOTO di questa scheda d'archivio e trascrivi i dati nei campi JSON.\n"
        )

    prompt = (
        f"{istruzione_principio}\n"
        "SIGLE RICORRENTI: 'coSS', 'prior', 'gal'. Il numero che segue una di queste sigle è il campo 'numero'.\n"
        "REGOLE:\n"
        "- Estrai numero, data, cognome, nome, di, citta, tipo, categoria.\n"
        "- CAMPO DIMENSIONI: estrai formato o misure (es. 'in-8', '12x18 cm').\n"
        "- CAMPO NOTE: inserisci qualsiasi altro testo o annotazione non rientrante nei campi precedenti.\n"
        "- Se un campo non è presente, usa \"\". Non inventare mai dati.\n\n"
        "CONFIDENZA: valuta con 'alta', 'media' o 'bassa'.\n\n"
        f"Devi rispondere RISPETTANDO RIGOROSAMENTE questo schema JSON:\n{schema_json_str}\n"
        "Restituisci SOLO ed ESCLUSIVAMENTE il codice JSON senza alcun testo prima o dopo."
    )

    risposta_raw = None
    for tentativo in range(1, MAX_RETRIES + 1):
        try:
            chat_completion = client.chat.completions.create(
                model=MODELLO_VISION,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:{mime_type};base64,{base64_immagine}"},
                            },
                        ],
                    }
                ],
                response_format={"type": "json_object"},
                temperature=0.0,
                max_tokens=MAX_TOKENS_RISPOSTA,
                reasoning_effort="none",   # disattiva il "ragionamento" interno del modello:
                                           # su schede difficili da leggere il modello consumava
                                           # tutto il budget di token a "pensare" prima di scrivere
                                           # il JSON, causando l'errore json_validate_failed
                reasoning_format="hidden", # non far comparire eventuale testo di ragionamento
                                           # nel contenuto della risposta (richiesto insieme a JSON mode)
            )

            risposta_raw = chat_completion.choices[0].message.content
            risposta_pulita = _estrai_json_pulito(risposta_raw)
            dati_validati = SchedaBibliografica.model_validate_json(risposta_pulita)
            return dati_validati.model_dump()

        except (GroqError, Exception) as e:
            if tentativo == MAX_RETRIES:
                # Includo un estratto della risposta grezza: senza questo è impossibile capire
                # se il modello ha troncato il JSON, aggiunto testo extra, o altro.
                anteprima = (risposta_raw[:300] + "...") if risposta_raw else "(nessuna risposta testuale ricevuta - vedi 'Errore' qui sopra per la causa reale)"
                raise RuntimeError(
                    f"Fallito dopo {MAX_RETRIES} tentativi. Errore: {str(e)} | Risposta grezza: {anteprima}"
                )
            time.sleep(2 ** tentativo)
def _normalizza(testo: str) -> str:
    """Rimuove accenti/maiuscole per confronti tolleranti tra letture diverse."""
    testo = unicodedata.normalize("NFKD", testo).encode("ascii", "ignore").decode()
    return re.sub(r'\s+', ' ', testo).strip().lower()

def analizza_scheda_completa(client: Groq, percorso_immagine: str) -> dict:
    """Orchestratore: legge la scheda con Kraken (se disponibile) e con Groq due volte."""
    testo_kraken = leggi_con_kraken(percorso_immagine)

    risultato_indipendente = analizza_scheda_con_groq_vision(client, percorso_immagine, bozza_ocr="")

    if testo_kraken.strip():
        risultato_assistito = analizza_scheda_con_groq_vision(client, percorso_immagine, bozza_ocr=testo_kraken)

        campi_divergenti = [
            campo for campo in CAMPI_DA_CONFRONTARE
            if _normalizza(risultato_indipendente.get(campo, "")) != _normalizza(risultato_assistito.get(campo, ""))
        ]

        risultato_finale = risultato_assistito
        risultato_finale["divergenza_kraken"] = (
            "OK" if not campi_divergenti else f"discordanza su: {', '.join(campi_divergenti)}"
        )
        risultato_finale["testo_grezzo_kraken"] = testo_kraken
    else:
        risultato_finale = risultato_indipendente
        risultato_finale["divergenza_kraken"] = "n/d (Kraken non disponibile)"
        risultato_finale["testo_grezzo_kraken"] = ""

    return risultato_finale

CAMPI_CHIAVE = ["numero", "cognome"]

def _riga_da_rivedere(riga: dict) -> bool:
    """Determina se una scheda necessita di revisione umana."""
    if str(riga.get("confidenza", "")).lower() == "bassa":
        return True
    if any(not str(riga.get(campo, "")).strip() for campo in CAMPI_CHIAVE):
        return True
    divergenza = str(riga.get("divergenza_kraken", ""))
    if divergenza.startswith("discordanza"):
        return True
    return False

def salva_lista_su_excel(lista_dati: list[dict], percorso_file_excel: str):
    """Salva la lista completa dei risultati in Excel, evidenziando in giallo le schede da rivedere manualmente."""
    if not lista_dati:
        return

    df_nuovo = pd.DataFrame(lista_dati)

    if os.path.exists(percorso_file_excel):
        try:
            df_esistente = pd.read_excel(percorso_file_excel).dropna(how='all')
            df_finale = pd.concat([df_esistente, df_nuovo], ignore_index=True)
        except Exception:
            df_finale = df_nuovo
    else:
        df_finale = df_nuovo

    colonne_ordinate = ["numero", "cognome", "nome", "di", "citta", "data", "tipo", "categoria",
                        "dimensioni", "note", "confidenza", "divergenza_kraken", "file_origine", "testo_grezzo_kraken"]
    colonne_presenti = [c for c in colonne_ordinate if c in df_finale.columns]
    colonne_extra = [c for c in df_finale.columns if c not in colonne_presenti]
    df_finale = df_finale[colonne_presenti + colonne_extra]

    df_finale.to_excel(percorso_file_excel, index=False)

    fill_giallo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    wb = load_workbook(percorso_file_excel)
    ws = wb.active
    intestazioni = [cell.value for cell in ws[1]]
    n_colonne = len(intestazioni)

    for indice_riga, riga_dict in enumerate(df_finale.to_dict(orient="records"), start=2):
        if _riga_da_rivedere(riga_dict):
            for col in range(1, n_colonne + 1):
                ws.cell(row=indice_riga, column=col).fill = fill_giallo

    wb.save(percorso_file_excel)

# =====================================================================
# 3. INTERFACCIA GRAFICA (Thread-Safe via Queue)
# =====================================================================
class AppEstrattore(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AI Document Extractor & Archivist v1.3")
        self.geometry("750x620")
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.cartella_selezionata = ""
        self.gui_queue = queue.Queue()

        self._crea_widgets()
        self._check_queue()

    def _crea_widgets(self):
        self.title_label = ctk.CTkLabel(self, text="📚 Estrattore AI Schede Bibliografiche", font=ctk.CTkFont(size=22, weight="bold"))
        self.title_label.pack(padx=20, pady=15)

        self.frame_cartella = ctk.CTkFrame(self)
        self.frame_cartella.pack(fill="x", padx=30, pady=10)
        self.btn_sfoglia = ctk.CTkButton(self.frame_cartella, text="Seleziona Cartella Foto", command=self.seleziona_cartella)
        self.btn_sfoglia.pack(side="left", padx=10, pady=10)
        self.lbl_percorso = ctk.CTkLabel(self.frame_cartella, text="Nessuna cartella selezionata...", text_color="gray")
        self.lbl_percorso.pack(side="left", fill="x", expand=True, padx=10, pady=10)

        self.txt_log = ctk.CTkTextbox(self, width=500, height=280, font=ctk.CTkFont(family="Consolas", size=11))
        self.txt_log.pack(fill="both", expand=True, padx=30, pady=10)
        self.scrivi_log("Pronto. Seleziona una cartella e avvia l'elaborazione.")

        self.progress_bar = ctk.CTkProgressBar(self, width=500)
        self.progress_bar.pack(padx=30, pady=5)
        self.progress_bar.set(0)

        self.btn_avvia = ctk.CTkButton(
            self, text="🚀 AVVIA ANALISI AUTOMATICA", 
            font=ctk.CTkFont(size=14, weight="bold"), 
            state="disabled", 
            command=self.avvia_thread_elaborazione
        )
        self.btn_avvia.pack(padx=30, pady=15)

    def scrivi_log(self, testo: str):
        self.txt_log.insert("end", testo + "\n")
        self.txt_log.see("end")

    def seleziona_cartella(self):
        percorso = filedialog.askdirectory()
        if percorso:
            self.cartella_selezionata = percorso
            self.lbl_percorso.configure(text=os.path.basename(percorso), text_color="white")
            self.btn_avvia.configure(state="normal")
            self.scrivi_log(f"Cartella caricata: {percorso}")
        else:
            self.cartella_selezionata = ""
            self.lbl_percorso.configure(text="Nessuna cartella selezionata...", text_color="gray")
            self.btn_avvia.configure(state="disabled")

    def avvia_thread_elaborazione(self):
        if not GROQ_API_KEY.strip():
            self.scrivi_log("❌ Errore: Chiave API non impostata nel codice!")
            return

        self.btn_sfoglia.configure(state="disabled")
        self.btn_avvia.configure(state="disabled")
        self.progress_bar.set(0)

        thread = threading.Thread(target=self.elabora_schede_worker, daemon=True)
        thread.start()

    def _check_queue(self):
        try:
            while True:
                msg_type, payload = self.gui_queue.get_nowait()
                if msg_type == "LOG":
                    self.scrivi_log(payload)
                elif msg_type == "PROGRESS":
                    self.progress_bar.set(payload)
                elif msg_type == "FINISHED":
                    self.scrivi_log(payload)
                    self.ripristina_pulsanti()
                elif msg_type == "ERROR":
                    self.scrivi_log(f"🛑 ERRORE CRITICO: {payload}")
                    self.ripristina_pulsanti()
        except queue.Empty:
            pass
        finally:
            self.after(100, self._check_queue)

    def elabora_schede_worker(self):
        try:
            client = Groq(api_key=GROQ_API_KEY, timeout=90.0)
            estensioni_valide = (".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp")
            file_immagini = [f for f in os.listdir(self.cartella_selezionata) if f.lower().endswith(estensioni_valide)]

            if not file_immagini:
                self.gui_queue.put(("LOG", "❌ Errore: Nessuna immagine trovata nella cartella."))
                self.gui_queue.put(("FINISHED", "Operazione annullata."))
                return

            totale = len(file_immagini)
            self.gui_queue.put(("LOG", f"\n⏳ Trovate {totale} immagini. Inizio elaborazione..."))

            if KRAKEN_DISPONIBILE and PERCORSO_MODELLO_KRAKEN:
                self.gui_queue.put(("LOG", "🔎 Kraken attivo: doppia lettura + confronto per ogni scheda."))
            else:
                self.gui_queue.put(("LOG", "ℹ️ Kraken non attivo (non installato o modello non configurato): solo Groq."))

            risultati = []
            totale_salvate = 0
            falliti = []  # nomi file falliti al primo giro: verranno ritentati a fine lotto
            avviso_kraken_mostrato = False  # mostra l'eventuale errore reale di Kraken una sola volta

            for indice, nome_file in enumerate(file_immagini, start=1):
                percorso_completo = os.path.join(self.cartella_selezionata, nome_file)
                self.gui_queue.put(("LOG", f"🔄 [{indice}/{totale}] Analisi: {nome_file}..."))

                try:
                    dati_puliti = analizza_scheda_completa(client, percorso_completo)
                    dati_puliti["file_origine"] = nome_file
                    risultati.append(dati_puliti)

                    if not avviso_kraken_mostrato and str(dati_puliti.get("divergenza_kraken", "")).startswith("n/d"):
                        avviso_kraken_mostrato = True
                        causa = ULTIMO_ERRORE_KRAKEN or "causa sconosciuta (nessuna eccezione registrata)"
                        self.gui_queue.put(("LOG", f"⚠️ Kraken non ha prodotto testo per questa scheda. Causa: {causa}"))

                    divergenza = dati_puliti.get("divergenza_kraken", "")
                    suffisso = f" ⚠️ {divergenza}" if divergenza.startswith("discordanza") else ""
                    self.gui_queue.put(("LOG", f"  ✔️ Estratto: {dati_puliti.get('numero', '')} - {dati_puliti.get('cognome', '')}{suffisso}"))
                except Exception as e:
                    self.gui_queue.put(("LOG", f"  ❌ Fallito {nome_file}: {e}"))
                    falliti.append(nome_file)

                self.gui_queue.put(("PROGRESS", indice / totale))

                if risultati and (indice % SALVA_OGNI_N_SCHEDE == 0 or indice == totale):
                    self.gui_queue.put(("LOG", f"💾 Checkpoint: salvataggio di {len(risultati)} schede su '{FILE_EXCEL}'..."))
                    self._salva_con_retry(risultati)
                    totale_salvate += len(risultati)
                    risultati = []

            if falliti:
                self.gui_queue.put(("LOG", f"\n🔁 {len(falliti)} scheda/e fallite per problemi di connessione. Pausa di 10s e nuovo tentativo..."))
                time.sleep(10)

                ancora_falliti = []
                for indice_recupero, nome_file in enumerate(falliti, start=1):
                    percorso_completo = os.path.join(self.cartella_selezionata, nome_file)
                    self.gui_queue.put(("LOG", f"🔄 Recupero [{indice_recupero}/{len(falliti)}]: {nome_file}..."))
                    try:
                        dati_puliti = analizza_scheda_completa(client, percorso_completo)
                        dati_puliti["file_origine"] = nome_file
                        risultati.append(dati_puliti)
                        self.gui_queue.put(("LOG", f"  ✔️ Recuperata: {dati_puliti.get('numero', '')} - {dati_puliti.get('cognome', '')}"))
                    except Exception as e:
                        self.gui_queue.put(("LOG", f"  ❌ Ancora fallita {nome_file}: {e}"))
                        ancora_falliti.append(nome_file)

                if risultati:
                    self.gui_queue.put(("LOG", f"💾 Salvataggio finale di {len(risultati)} schede recuperate su '{FILE_EXCEL}'..."))
                    self._salva_con_retry(risultati)
                    totale_salvate += len(risultati)
                    risultati = []

                if ancora_falliti:
                    self.gui_queue.put(("LOG", f"⚠️ {len(ancora_falliti)} scheda/e non recuperabili: {', '.join(ancora_falliti)}. Riprova manualmente più tardi."))

            if totale_salvate > 0:
                self.gui_queue.put(("FINISHED", f"\n🎉 Completato! Salvate {totale_salvate} schede in '{FILE_EXCEL}' (righe gialle = da rivedere manualmente)."))
            else:
                self.gui_queue.put(("FINISHED", "\n⚠️ Nessun dato estratto con successo."))

        except Exception as e:
            self.gui_queue.put(("ERROR", str(e)))

    def _salva_con_retry(self, risultati: list[dict]):
        salvato = False
        while not salvato:
            try:
                salva_lista_su_excel(risultati, FILE_EXCEL)
                salvato = True
            except PermissionError:
                self.gui_queue.put(("LOG", f"⚠️ Impossibile salvare! Il file '{FILE_EXCEL}' è aperto in Excel. Chiudilo! Riprovo tra 3s..."))
                time.sleep(3)

    def ripristina_pulsanti(self):
        self.btn_sfoglia.configure(state="normal")
        self.btn_avvia.configure(state="normal")

if __name__ == "__main__":
    app = AppEstrattore()
    app.mainloop()