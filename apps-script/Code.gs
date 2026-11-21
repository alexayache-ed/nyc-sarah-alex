/**
 * Relais Google Apps Script — Notes du site NYC → onglet "Notes site" du Google Doc.
 *
 * Installation (1 seule fois, ~2 min) :
 * 1. Ouvre le Google Doc → Extensions → Apps Script.
 * 2. Remplace le contenu par ce fichier et enregistre (Ctrl/Cmd + S).
 * 3. Sélectionne `testSetup` dans la barre du haut et clique ▶ Exécuter (pour autoriser l'accès au Doc).
 * 4. Clique Déployer → Nouveau déploiement → Type : Application Web
 *    - Exécuter en tant que : Moi
 *    - Qui a accès : Tout le monde
 * 5. Copie l'URL https://script.google.com/macros/s/.../exec et donne-la dans le chat (ou colle-la dans le site).
 */

const DOC_ID = '1tB_BDRB2Zen7F0OpNSfrmaOrOA6rRx62FM5mypNRlOI';
const TAB_ID = 't.8yeojcqkb01m';
const TAB_TITLE = 'Notes site';
const PROP_KEY = 'NYC_SITE_NOTES_V1';

const SCOPES = [
  { id: 'jeu-3',    label: 'Jeu 3 déc · Midtown & Noël' },
  { id: 'ven-4',    label: 'Ven 4 déc · Central Park & musées' },
  { id: 'sam-5',    label: 'Sam 5 déc · Brooklyn' },
  { id: 'dim-6',    label: 'Dim 6 déc · Harlem, Flatiron & SoHo' },
  { id: 'lun-7',    label: 'Lun 7 déc · Downtown & Dyker Heights' },
  { id: 'mar-8',    label: 'Mar 8 déc · Village & Chelsea' },
  { id: 'reserver', label: 'À réserver' },
  { id: 'adresses', label: 'Adresses & restos' },
  { id: 'general',  label: 'Général' }
];
const VALID_SCOPES = new Set(SCOPES.map(s => s.id));

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readNotes() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function writeNotes(notes) {
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify(notes));
  syncToDocTab(notes);
}

function getNotesBody() {
  const doc = DocumentApp.openById(DOC_ID);
  let tab = null;
  if (typeof doc.getTab === 'function') {
    tab = doc.getTab(TAB_ID);
  }
  if (!tab && typeof doc.getTabs === 'function') {
    const tabs = doc.getTabs();
    tab = tabs.find(t => t.getId() === TAB_ID || t.getTitle() === TAB_TITLE) || null;
  }
  if (!tab) {
    throw new Error('Onglet "' + TAB_TITLE + '" introuvable (' + TAB_ID + ').');
  }
  return tab.asDocumentTab().getBody();
}

function syncToDocTab(notes) {
  const body = getNotesBody();
  body.clear();

  const title = body.getChild(0).asParagraph();
  title.setText('Notes du site — New York (2 → 9 déc)');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const nowStr = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
  const meta = body.appendParagraph('Synchronisé automatiquement depuis le site · Dernière mise à jour : ' + nowStr + ' (Paris)');
  meta.editAsText().setItalic(true);

  if (!notes.length) {
    body.appendParagraph('Aucune note en attente.');
    return;
  }

  SCOPES.forEach(scope => {
    const items = notes.filter(n => n.scope === scope.id);
    if (!items.length) return;
    const heading = body.appendParagraph(scope.label);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    items.forEach(n => {
      const prefix = n.author ? '[' + n.author + '] ' : '';
      const li = body.appendListItem(prefix + n.text);
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
    });
  });
}

function normalizeNote(raw) {
  const text = String((raw && raw.text) || '').trim().slice(0, 500);
  if (!text) return null;
  const scope = VALID_SCOPES.has(raw.scope) ? raw.scope : 'general';
  const author = raw.author === 'Sarah' ? 'Sarah' : (raw.author === 'Alex' ? 'Alex' : '');
  const id = (typeof raw.id === 'string' && raw.id.trim())
    ? raw.id.trim().slice(0, 40)
    : ('n_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 7));
  return { id: id, scope: scope, author: author, text: text, createdAt: Date.now() };
}

function doGet(e) {
  try {
    return jsonOut({ ok: true, notes: readNotes() });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = payload.action;
    let notes = readNotes();

    if (action === 'add') {
      if (notes.length >= 100) return jsonOut({ ok: false, error: 'Limite de 100 notes atteinte' });
      const note = normalizeNote(payload);
      if (!note) return jsonOut({ ok: false, error: 'Note vide' });
      if (!notes.some(n => n.id === note.id)) {
        notes.push(note);
        writeNotes(notes);
      }
    } else if (action === 'update') {
      const text = String(payload.text || '').trim().slice(0, 500);
      const target = notes.find(n => n.id === payload.id);
      if (target && text) {
        target.text = text;
        writeNotes(notes);
      }
    } else if (action === 'delete') {
      const next = notes.filter(n => n.id !== payload.id);
      if (next.length !== notes.length) {
        notes = next;
        writeNotes(notes);
      }
    } else if (action === 'clear') {
      notes = [];
      writeNotes(notes);
    } else if (action === 'sync') {
      syncToDocTab(notes);
    } else {
      return jsonOut({ ok: false, error: 'Action inconnue' });
    }

    return jsonOut({ ok: true, notes: notes });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function testSetup() {
  syncToDocTab(readNotes());
}
