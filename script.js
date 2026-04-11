/* ═══════════════════════════════════════════════════════════════════════
   Digital Ink Signer — script.js  (v9)
   ─────────────────────────────────────────────────────────────────────
   §0  Document Organizer  (NEW)
         ┌── Organizer Pipeline ───────────────────────────────────────┐
         │  Upload (multiple PDFs) → pdf.js thumbnail render           │
         │  → HTML5 drag-and-drop reorder → page delete                │
         │  → "Proceed to Sign" → pdf-lib merge into single ArrayBuffer│
         │  → hand off to §4 signing canvas as if one PDF was uploaded │
         └─────────────────────────────────────────────────────────────┘
   §1  PIN screen
   §2  State & DOM refs
   §3  Pagination bar (injected)
   §4  PDF load from bytes → pdf.js render  (now called by §0 merger)
   §5  Fabric.js canvas  (single instance, never disposed)
   §6  Image-based Stamps  (Ink Signature / Designation Seal / Office Seal)
         ┌── Image Processing & Storage Pipeline (v6) ─────────────────┐
         │  ① Auto Background Removal  — white/light-grey → transparent│
         │  ② Auto-Crop (Bounding Box) — trim invisible padding        │
         │  ③a Downscale              — hard cap at 200 px             │
         │  ③b Compress to JPEG       — quality 0.75, target <15 KB    │
         │  ④  Persistent localStorage                                  │
         │  ⑤  Auto-Load on Start     — sidebar preview + trash button │
         └─────────────────────────────────────────────────────────────┘
   §7  Freehand draw tool
   §8  Delete  (button + Delete/Backspace key)
   §9  Download  — pdf-lib, rotation-aware, BlendMode.Multiply on top
   ═══════════════════════════════════════════════════════════════════════ */

document.addEventListener('DOMContentLoaded', () => {

  /* ════════════════════════════════════════════════════════════════════
     §1  PIN SCREEN
  ════════════════════════════════════════════════════════════════════ */
  const pinInput     = document.getElementById('pin-input');
  const verifyBtn    = document.getElementById('verify-btn');
  const pinScreen    = document.getElementById('pin-screen');
  const appDashboard = document.getElementById('app-dashboard');
  const dotContainer = document.getElementById('dot-container');
  const pinTitle     = document.getElementById('pin-title');
  const pinSubtitle  = document.getElementById('pin-subtitle');
  const btnText      = document.getElementById('btn-text');

  let savedPin     = localStorage.getItem('ink_signer_pin');
  let isSettingPin = !savedPin;

  if (isSettingPin) {
    pinTitle.innerText    = 'Set New PIN';
    pinSubtitle.innerText = 'സുരക്ഷയ്ക്കായി ഒരു പുതിയ PIN സൃഷ്ടിക്കുക';
    btnText.innerText     = 'Save PIN';
  }

  pinInput.addEventListener('input', () => {
    pinInput.value = pinInput.value.replace(/[^0-9]/g, '');
    const len  = pinInput.value.length;
    const dots = dotContainer.children;
    for (let i = 0; i < 4; i++) {
      dots[i].className = i < len
        ? 'pin-dot bg-primary-container shadow-sm'
        : 'pin-dot bg-surface-container-highest border border-outline-variant/30';
    }
  });

  verifyBtn.addEventListener('click', processPIN);
  pinInput.addEventListener('keypress', e => { if (e.key === 'Enter') processPIN(); });

  function processPIN() {
    const val = pinInput.value;
    if (val.length !== 4) { alert('ദയവായി 4 അക്ക PIN നൽകുക!'); return; }
    if (isSettingPin) {
      localStorage.setItem('ink_signer_pin', val);
      alert('PIN വിജയകരമായി സേവ് ചെയ്തു!');
      unlockApp();
    } else {
      if (val === savedPin) {
        unlockApp();
      } else {
        alert('തെറ്റായ PIN! വീണ്ടും ശ്രമിക്കുക.');
        pinInput.value = '';
        Array.from(dotContainer.children).forEach(d =>
          (d.className = 'pin-dot bg-surface-container-highest border border-outline-variant/30'));
      }
    }
  }

  function unlockApp() {
    pinScreen.style.opacity    = '0';
    pinScreen.style.transition = 'opacity 0.5s';
    setTimeout(() => (pinScreen.style.display = 'none'), 500);
    appDashboard.classList.remove('blur-md', 'pointer-events-none', 'grayscale-[0.2]', 'opacity-40');
  }
// PIN സ്ക്രീനിലെ Reset ബട്ടൺ
  document.getElementById('btn-forgot-pin')?.addEventListener('click', () => {
    if (!confirm('ആപ്പ് പൂർണ്ണമായും റിസെറ്റ് ചെയ്യണോ?\n\nനിങ്ങളുടെ പഴയ PIN ഉം സേവ് ചെയ്ത ഒപ്പുകളും ഡിലീറ്റ് ആകുന്നതാണ്.')) return;
    
    ['ink_signer_pin', 'saved_ink', 'saved_desig_seal', 'saved_office_seal'].forEach(k => localStorage.removeItem(k));
    location.reload();
  });

  /* ════════════════════════════════════════════════════════════════════
     §2  STATE & DOM REFS
  ════════════════════════════════════════════════════════════════════ */
  const uploadArea        = document.getElementById('upload-area');
  const fileInput         = document.getElementById('pdf-file-input');
  const pdfCanvas         = document.getElementById('pdf-render-canvas');
  const pdfPlaceholder    = document.getElementById('pdf-placeholder');
  const pdfWrapper        = document.getElementById('pdf-wrapper');
  const pageInfo          = document.getElementById('page-info');
  const deleteBtnEl       = document.getElementById('btn-delete-sel');
  const organizerScreen   = document.getElementById('organizer-screen');
  const organizerControls = document.getElementById('organizer-controls');
  const stampControls     = document.getElementById('stamp-controls');
  const thumbGrid         = document.getElementById('thumb-grid');
  const orgSpinner        = document.getElementById('org-spinner');
  const orgEmpty          = document.getElementById('org-empty');
  const orgPageCount      = document.getElementById('org-page-count');

  const RENDER_SCALE      = 1.8;
  const THUMB_SCALE       = 0.25;   // scale used to render organizer thumbnails

  let pdfDoc       = null;
  let pdfBytes     = null;
  let currentPage  = 1;
  let totalPages   = 1;
  let fabricCanvas = null;

  const pageObjects    = new Map();
  const pageDimensions = new Map();

  // Organizer state — each entry describes one page from any uploaded PDF
  // { srcBytes: ArrayBuffer, srcPageIndex: number (0-based), label: string, thumbCanvas: HTMLCanvasElement }
  let orgPages = [];

  let dragSrcIdx = null;   // index being dragged in orgPages


  /* ════════════════════════════════════════════════════════════════════
     §2·5  GLOBAL LOADING OVERLAY  +  RESET APP
     ──────────────────────────────────────────────────────────────────
     showLoader(msg) / hideLoader() wrap every heavy async operation so
     the user sees immediate feedback instead of a frozen UI.

     The setTimeout(..., 10) before each heavy call is critical: it yields
     the main thread back to the browser for one paint cycle so the overlay
     is actually rendered before the CPU-intensive work begins.

     resetApp() clears all localStorage keys (PIN + saved stamps) and
     reloads the page — the simplest possible full reset.
  ════════════════════════════════════════════════════════════════════ */
  const loadingOverlay = document.getElementById('loading-overlay');
  const loadingText    = document.getElementById('loading-text');

  function showLoader(msg = 'Please wait…') {
    loadingText.textContent = msg;
    loadingOverlay.classList.add('active');
  }
  function hideLoader() {
    loadingOverlay.classList.remove('active');
  }

  // Reset App button — clears PIN + all saved stamps then reloads
  document.getElementById('btn-reset-app').addEventListener('click', () => {
    if (!confirm(
      'Reset the app?\n\n' +
      'This will permanently delete your saved PIN and all stored ' +
      'Ink Signatures / Designation Seals / Office Seals.\n\n' +
      'The page will reload.'
    )) return;

    // Clear every key this app ever writes to localStorage
    ['ink_signer_pin', 'saved_ink', 'saved_desig_seal', 'saved_office_seal']
      .forEach(k => localStorage.removeItem(k));

    location.reload();
  });


  /* ════════════════════════════════════════════════════════════════════
     §3  PAGINATION BAR  (injected below #pdf-wrapper by JS)
  ════════════════════════════════════════════════════════════════════ */
  const paginationBar = document.createElement('div');
  paginationBar.id    = 'pagination-bar';
  Object.assign(paginationBar.style, {
    display: 'none', alignItems: 'center', justifyContent: 'center',
    gap: '14px', fontFamily: 'Inter, sans-serif',
  });

  const _btnStyle = `padding:6px 20px;border-radius:8px;background:#005bbf;color:#fff;
    font-size:12px;font-weight:700;border:none;cursor:pointer;transition:opacity .2s,transform .1s;`;
  paginationBar.innerHTML = `
    <button id="btn-prev-page" style="${_btnStyle}opacity:.32;pointer-events:none;">← Prev</button>
    <span id="page-label" style="font-size:12px;font-weight:600;color:#414754;min-width:96px;text-align:center;">
      Page 1 of 1
    </span>
    <button id="btn-next-page" style="${_btnStyle}">Next →</button>
  `;
  pdfWrapper.parentNode.insertBefore(paginationBar, pdfWrapper.nextSibling);

  const btnPrev = document.getElementById('btn-prev-page');
  const btnNext = document.getElementById('btn-next-page');
  const pageLbl = document.getElementById('page-label');

  function updatePaginationUI() {
    pageLbl.textContent  = `Page ${currentPage} of ${totalPages}`;
    pageInfo.textContent = `Page ${currentPage} of ${totalPages}`;
    btnPrev.style.opacity       = currentPage <= 1            ? '0.32' : '1';
    btnPrev.style.pointerEvents = currentPage <= 1            ? 'none'  : 'auto';
    btnNext.style.opacity       = currentPage >= totalPages   ? '0.32' : '1';
    btnNext.style.pointerEvents = currentPage >= totalPages   ? 'none'  : 'auto';
  }

  btnPrev.addEventListener('click', () => { if (currentPage > 1)          navigateTo(currentPage - 1); });
  btnNext.addEventListener('click', () => { if (currentPage < totalPages) navigateTo(currentPage + 1); });

  async function navigateTo(newPage) {
    saveCurrentPageObjects();
    setDrawingMode(false);
    currentPage = newPage;
    await renderPage(currentPage);
    restorePageObjects(currentPage);
    updatePaginationUI();
  }


  /* ════════════════════════════════════════════════════════════════════
     §0  DOCUMENT ORGANIZER
     ──────────────────────────────────────────────────────────────────
     Flow:
       uploadArea click → fileInput (multiple) → processUploadedFiles()
         → for each PDF, extract all pages → push into orgPages[]
         → renderThumbnails() builds the draggable grid
       Drag & drop (native HTML5) reorders orgPages[]
       Delete button splices orgPages[], re-renders grid
       "Proceed to Sign" → mergeAndLoad()
         → pdf-lib: create new doc, copyPages() in orgPages order
         → save() → ArrayBuffer → hand to loadPdfBytes() (§4 entry point)
       "Back to Organizer" returns to organizer screen
       "Start Over" resets everything
  ════════════════════════════════════════════════════════════════════ */
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  // Upload area → open multi-file picker
  uploadArea.addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });

  fileInput.addEventListener('change', async () => {
    const files = Array.from(fileInput.files).filter(f => f.type === 'application/pdf');
    if (!files.length) return;
    for (const f of files) {
      if (f.size > 25 * 1024 * 1024) { alert(`"${f.name}" exceeds 25 MB limit — skipped.`); continue; }
    }
    await processUploadedFiles(files.filter(f => f.size <= 25 * 1024 * 1024));
  });

  /* processUploadedFiles — reads each PDF with pdf.js, renders THUMB_SCALE
     thumbnails for every page, and appends them to orgPages[]. */
  async function processUploadedFiles(files) {
    if (!files.length) return;

    showOrganizerScreen();
    showLoader('Rendering thumbnails… Please wait.');

    // Yield one paint cycle so the overlay appears before the heavy work
    await new Promise(r => setTimeout(r, 10));

    orgSpinner.classList.remove('hidden');

    for (const file of files) {
      const bytes  = await file.arrayBuffer();
      const pdfDoc = await pdfjsLib.getDocument({ data: bytes.slice(0) }).promise;

      for (let i = 0; i < pdfDoc.numPages; i++) {
        const page     = await pdfDoc.getPage(i + 1);
        const viewport = page.getViewport({ scale: THUMB_SCALE });

        const tc     = document.createElement('canvas');
        tc.width     = viewport.width;
        tc.height    = viewport.height;
        await page.render({ canvasContext: tc.getContext('2d'), viewport }).promise;

        orgPages.push({
          srcBytes:     bytes,
          srcPageIndex: i,          // 0-based index inside srcBytes
          label:        `${file.name.replace(/\.pdf$/i, '')}  p.${i + 1}`,
          thumbCanvas:  tc,
        });
      }
    }

    orgSpinner.classList.add('hidden');
    hideLoader();
    renderThumbnails();
  }

  /* renderThumbnails — rebuilds the entire #thumb-grid DOM from orgPages[]. */
  function renderThumbnails() {
    thumbGrid.innerHTML = '';
    orgPageCount.textContent = `${orgPages.length} page${orgPages.length !== 1 ? 's' : ''}`;

    if (orgPages.length === 0) {
      orgEmpty.classList.remove('hidden');
      return;
    }
    orgEmpty.classList.add('hidden');

    orgPages.forEach((pg, idx) => {
      const card = document.createElement('div');
      card.className    = 'thumb-card';
      card.draggable    = true;
      card.dataset.idx  = idx;

      // Clone thumb canvas into a display canvas
      const display   = document.createElement('canvas');
      display.width   = pg.thumbCanvas.width;
      display.height  = pg.thumbCanvas.height;
      display.getContext('2d').drawImage(pg.thumbCanvas, 0, 0);

      // Footer: label + delete button
      const footer = document.createElement('div');
      footer.className = 'thumb-footer';

      const labelEl = document.createElement('span');
      labelEl.textContent = pg.label;
      labelEl.style.cssText = 'flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:82px;';

      const delBtn = document.createElement('button');
      delBtn.className = 'thumb-del';
      delBtn.title     = 'Remove this page';
      delBtn.innerHTML = '<span class="material-symbols-outlined" style="font-size:14px">close</span>';
      delBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        orgPages.splice(idx, 1);
        renderThumbnails();
      });

      footer.append(labelEl, delBtn);

      // Drag-handle dots (decorative; whole card is draggable)
      const handle = document.createElement('div');
      handle.className = 'drag-handle';
      for (let d = 0; d < 6; d++) handle.append(document.createElement('span'));

      card.append(display, footer, handle);

      // ── HTML5 Drag & Drop ────────────────────────────────────────
      card.addEventListener('dragstart', (e) => {
        dragSrcIdx = idx;
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', idx); // required for Firefox
      });

      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        thumbGrid.querySelectorAll('.thumb-card').forEach(c => c.classList.remove('drag-over'));
        dragSrcIdx = null;
      });

      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        thumbGrid.querySelectorAll('.thumb-card').forEach(c => c.classList.remove('drag-over'));
        card.classList.add('drag-over');
      });

      card.addEventListener('dragleave', () => {
        card.classList.remove('drag-over');
      });

      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over');
        const targetIdx = parseInt(card.dataset.idx, 10);
        if (dragSrcIdx === null || dragSrcIdx === targetIdx) return;

        // Reorder: remove source, insert before target
        const [moved] = orgPages.splice(dragSrcIdx, 1);
        const insertAt = dragSrcIdx < targetIdx ? targetIdx - 1 : targetIdx;
        orgPages.splice(insertAt, 0, moved);
        renderThumbnails();
      });

      thumbGrid.appendChild(card);
    });
  }

  /* mergeAndLoad — uses pdf-lib to copy pages in current orgPages order
     into a brand-new PDF, then passes the resulting bytes to loadPdfBytes(). */
  async function mergeAndLoad() {
    if (orgPages.length === 0) { alert('Please add at least one page.'); return; }

    const procBtn = document.getElementById('btn-proceed-sign');
    procBtn.disabled = true;

    showLoader('Merging PDFs… Please wait.');
    // Yield one paint cycle so the loader is visible before pdf-lib work begins
    await new Promise(r => setTimeout(r, 10));

    try {
      const mergedDoc = await PDFLib.PDFDocument.create();

      // Group pages by their source bytes to minimise repeated loads
      const srcDocCache = new Map();

      for (const pg of orgPages) {
        let srcDoc = srcDocCache.get(pg.srcBytes);
        if (!srcDoc) {
          // ignoreEncryption: true lets pdf-lib open digitally-signed / encrypted
          // eOffice documents without throwing.  The cryptographic signature is
          // lost (flattened) but all visual content is preserved for merging.
          srcDoc = await PDFLib.PDFDocument.load(pg.srcBytes, { ignoreEncryption: true });
          srcDocCache.set(pg.srcBytes, srcDoc);
        }
        const [copiedPage] = await mergedDoc.copyPages(srcDoc, [pg.srcPageIndex]);
        mergedDoc.addPage(copiedPage);
      }

      const mergedBytes = await mergedDoc.save();
      await loadPdfBytes(mergedBytes.buffer);

    } catch (err) {
      console.error('Merge error:', err);
      alert('Merge failed: ' + err.message);
    } finally {
      hideLoader();
      procBtn.disabled = false;
    }
  }

  /* showOrganizerScreen / showSigningScreen — screen switcher */
  function showOrganizerScreen() {
    pdfPlaceholder.style.display    = 'none';
    organizerScreen.style.display   = 'block';
    pdfWrapper.style.display        = 'none';
    paginationBar.style.display     = 'none';
    pageInfo.style.display          = 'none';
    organizerControls.style.display = 'flex';
    stampControls.style.display     = 'none';
  }

  function showSigningScreen() {
    organizerScreen.style.display   = 'none';
    organizerControls.style.display = 'none';
    stampControls.style.display     = 'flex';
    pdfWrapper.style.display        = 'inline-block';
    paginationBar.style.display     = 'flex';
    pageInfo.style.display          = 'block';
  }

  // "Proceed to Sign" button
  document.getElementById('btn-proceed-sign').addEventListener('click', mergeAndLoad);

  // "Back to Organizer" button (inside stamp-controls)
  document.getElementById('btn-back-organizer').addEventListener('click', () => {
    showOrganizerScreen();
  });

  // "Start Over" — clear everything and return to placeholder
  document.getElementById('btn-organizer-reset').addEventListener('click', () => {
    if (!confirm('Start over? All organizer pages and signatures will be cleared.')) return;
    orgPages = [];
    thumbGrid.innerHTML = '';
    pageObjects.clear();
    pageDimensions.clear();
    pdfDoc   = null;
    pdfBytes = null;
    organizerScreen.style.display   = 'none';
    organizerControls.style.display = 'none';
    stampControls.style.display     = 'none';
    pdfWrapper.style.display        = 'none';
    paginationBar.style.display     = 'none';
    pageInfo.style.display          = 'none';
    pdfPlaceholder.style.display    = 'flex';
    if (fabricCanvas) { fabricCanvas.clear(); fabricCanvas.requestRenderAll(); }
  });


  /* ════════════════════════════════════════════════════════════════════
     §4  PDF LOAD FROM BYTES  →  PDF.JS RENDER
     ──────────────────────────────────────────────────────────────────
     loadPdfBytes() is the single entry point for the signing stage.
     It is called by mergeAndLoad() after the organizer merge, passing
     the ArrayBuffer of the combined PDF.
  ════════════════════════════════════════════════════════════════════ */
  async function loadPdfBytes(buffer) {
    pdfBytes = buffer;

    showLoader('Loading PDF… Please wait.');
    await new Promise(r => setTimeout(r, 10));

    try {
      pdfDoc      = await pdfjsLib.getDocument({ data: buffer.slice(0) }).promise;
      totalPages  = pdfDoc.numPages;
      currentPage = 1;
      pageObjects.clear();
      pageDimensions.clear();

      await renderPage(1);
      initFabricCanvas(pdfCanvas.width, pdfCanvas.height);

      showSigningScreen();
      updatePaginationUI();
    } catch (err) {
      console.error('PDF load error:', err);
      alert('PDF ലോഡ് ചെയ്യാൻ കഴിഞ്ഞില്ല: ' + err.message);
    } finally {
      hideLoader();
    }
  }

  async function renderPage(pageNum) {
    const page     = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: RENDER_SCALE });

    pdfCanvas.width  = viewport.width;
    pdfCanvas.height = viewport.height;

    pageDimensions.set(pageNum, { canvasW: viewport.width, canvasH: viewport.height });

    await page.render({
      canvasContext: pdfCanvas.getContext('2d'),
      viewport,
    }).promise;

    if (fabricCanvas) {
      fabricCanvas.setWidth(viewport.width);
      fabricCanvas.setHeight(viewport.height);
      fabricCanvas.clear();
      Object.assign(fabricCanvas.wrapperEl.style, {
        width:  viewport.width  + 'px',
        height: viewport.height + 'px',
      });
      fabricCanvas.requestRenderAll();
    }
  }


  /* ════════════════════════════════════════════════════════════════════
     §5  FABRIC.JS CANVAS  — created exactly once, resized on page turns
  ════════════════════════════════════════════════════════════════════ */
  function initFabricCanvas(width, height) {
    if (fabricCanvas) return;

    const el  = document.getElementById('fabric-canvas-el');
    el.width  = width;
    el.height = height;

    fabricCanvas = new fabric.Canvas('fabric-canvas-el', {
      width, height,
      selection:              true,
      preserveObjectStacking: true,
      renderOnAddRemove:      true,
      enableRetinaScaling:    false,
    });

    Object.assign(fabricCanvas.wrapperEl.style, {
      position: 'absolute', top: '0', left: '0',
      width: width + 'px', height: height + 'px',
    });

    fabricCanvas.on('selection:created', () => deleteBtnEl.classList.remove('hidden'));
    fabricCanvas.on('selection:updated', () => deleteBtnEl.classList.remove('hidden'));
    fabricCanvas.on('selection:cleared',  () => deleteBtnEl.classList.add('hidden'));

    // ── Ghost follower: track cursor during stamping mode ─────────────
    // mouse:move fires continuously; we simply reposition the ghost and
    // request a render.  The ghost is non-evented so Fabric won't treat
    // it as a hover target and won't reset the cursor.
    fabricCanvas.on('mouse:move', (opt) => {
      if (!stampingMode || !ghostStamp) return;
      const p = opt.pointer;
      ghostStamp.set({
        left: p.x - ghostStamp.getScaledWidth()  / 2,
        top:  p.y - ghostStamp.getScaledHeight() / 2,
      });
      ghostStamp.setCoords();
      fabricCanvas.requestRenderAll();
    });

    // ── Stamping mode: place real stamp at the click point ────────────
    // Capture the pointer position and pending state BEFORE calling
    // exitStampingMode(), which clears them.
    fabricCanvas.on('mouse:down', (opt) => {
      if (!stampingMode || !pendingStampDataUrl) return;

      const pointer = opt.pointer;
      const dataUrl = pendingStampDataUrl;
      const color   = pendingStampColor;

      exitStampingMode();   // removes ghost, resets cursors, clears state

      // Place a fully opaque, interactive copy at the clicked coordinates
      fabric.Image.fromURL(dataUrl, (imgObj) => {
        const MAX_W = fabricCanvas.width * 0.30;
        if (imgObj.width > MAX_W) imgObj.scaleToWidth(MAX_W);

        imgObj.set({
          left:               pointer.x - imgObj.getScaledWidth()  / 2,
          top:                pointer.y - imgObj.getScaledHeight() / 2,
          opacity:            1,
          hasControls:        true,
          hasBorders:         true,
          borderColor:        color,
          cornerColor:        color,
          cornerSize:         10,
          transparentCorners: false,
          cornerStyle:        'circle',
          borderDashArray:    [4, 3],
        });

        fabricCanvas.add(imgObj);
        fabricCanvas.setActiveObject(imgObj);
        fabricCanvas.requestRenderAll();
      }, { crossOrigin: null });
    });

    fabricCanvas.freeDrawingBrush          = new fabric.PencilBrush(fabricCanvas);
    fabricCanvas.freeDrawingBrush.color    = currentInkColor;
    fabricCanvas.freeDrawingBrush.width    = currentBrushSize;
    fabricCanvas.freeDrawingBrush.decimate = 0;
  }

  function saveCurrentPageObjects() {
    if (!fabricCanvas) return;
    pageObjects.set(currentPage, fabricCanvas.toJSON());
  }

  function restorePageObjects(pageNum) {
    if (!fabricCanvas) return;
    fabricCanvas.clear();
    if (pageObjects.has(pageNum)) {
      fabricCanvas.loadFromJSON(pageObjects.get(pageNum), () => fabricCanvas.requestRenderAll());
    }
  }


  /* ════════════════════════════════════════════════════════════════════
     §6  IMAGE-BASED STAMPS  +  IMAGE PROCESSING & STORAGE PIPELINE
  ════════════════════════════════════════════════════════════════════ */
  const BG_THRESHOLD = 240;
  const TARGET_KB    = 15;
  const MIN_SCALE    = 0.40;
  const SCALE_STEP   = 0.10;
  const MAX_STAMP_PX = 200;
  const JPEG_QUALITY = 0.75;

  const STAMP_DEFS = [
    { btnId: 'btn-ink-sig',     label: 'Ink Signature',    color: '#005bbf', storageKey: 'saved_ink'        },
    { btnId: 'btn-desig-seal',  label: 'Designation Seal', color: '#1e7a3c', storageKey: 'saved_desig_seal' },
    { btnId: 'btn-office-seal', label: 'Office Seal',      color: '#b35200', storageKey: 'saved_office_seal'},
  ];

  function removeBackground(imgEl) {
    const c   = document.createElement('canvas');
    c.width   = imgEl.naturalWidth  || imgEl.width;
    c.height  = imgEl.naturalHeight || imgEl.height;
    const ctx = c.getContext('2d');
    ctx.drawImage(imgEl, 0, 0);
    const imageData = ctx.getImageData(0, 0, c.width, c.height);
    const data      = imageData.data;
    for (let i = 0; i < data.length; i += 4) {
      const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      if (lum > BG_THRESHOLD) data[i + 3] = 0;
    }
    ctx.putImageData(imageData, 0, 0);
    return c;
  }

  function autoCrop(srcCanvas) {
    const ctx  = srcCanvas.getContext('2d');
    const { width, height } = srcCanvas;
    const data = ctx.getImageData(0, 0, width, height).data;
    let minX = width, maxX = -1, minY = height, maxY = -1;
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        if (data[(y * width + x) * 4 + 3] > 10) {
          if (x < minX) minX = x; if (x > maxX) maxX = x;
          if (y < minY) minY = y; if (y > maxY) maxY = y;
        }
      }
    }
    if (maxX < 0 || maxY < 0) return srcCanvas;
    const PAD = 2;
    minX = Math.max(0, minX - PAD); minY = Math.max(0, minY - PAD);
    maxX = Math.min(width - 1, maxX + PAD); maxY = Math.min(height - 1, maxY + PAD);
    const cropW = maxX - minX + 1, cropH = maxY - minY + 1;
    const dest = document.createElement('canvas');
    dest.width = cropW; dest.height = cropH;
    dest.getContext('2d').drawImage(srcCanvas, minX, minY, cropW, cropH, 0, 0, cropW, cropH);
    return dest;
  }

  function downscale(srcCanvas) {
    const longest = Math.max(srcCanvas.width, srcCanvas.height);
    if (longest <= MAX_STAMP_PX) return srcCanvas;
    const ratio = MAX_STAMP_PX / longest;
    const destW = Math.max(1, Math.round(srcCanvas.width  * ratio));
    const destH = Math.max(1, Math.round(srcCanvas.height * ratio));
    const dest  = document.createElement('canvas');
    dest.width  = destW; dest.height = destH;
    const ctx   = dest.getContext('2d');
    ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(srcCanvas, 0, 0, destW, destH);
    return dest;
  }

  function compress(scaledCanvas) {
    const origW = scaledCanvas.width, origH = scaledCanvas.height;
    const flat  = document.createElement('canvas');
    flat.width  = origW; flat.height = origH;
    const fc    = flat.getContext('2d');
    fc.fillStyle = '#ffffff'; fc.fillRect(0, 0, origW, origH);
    fc.drawImage(scaledCanvas, 0, 0);
    let scale = 1.0, dataUrl = flat.toDataURL('image/jpeg', JPEG_QUALITY);
    while (dataUrl.length > TARGET_KB * 1024 * 1.37 && scale > MIN_SCALE) {
      scale -= SCALE_STEP;
      const w = Math.max(1, Math.round(origW * scale));
      const h = Math.max(1, Math.round(origH * scale));
      const tmp = document.createElement('canvas');
      tmp.width = w; tmp.height = h;
      const tc  = tmp.getContext('2d');
      tc.fillStyle = '#ffffff'; tc.fillRect(0, 0, w, h);
      tc.imageSmoothingEnabled = true; tc.imageSmoothingQuality = 'high';
      tc.drawImage(flat, 0, 0, w, h);
      dataUrl = tmp.toDataURL('image/jpeg', JPEG_QUALITY);
    }
    console.log(`[Stamp] Final JPEG: ~${Math.round(dataUrl.length * .75 / 1024)} KB (scale ${scale.toFixed(2)})`);
    return dataUrl;
  }

  function processStampImage(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('FileReader failed'));
      reader.onload  = (ev) => {
        const img = new Image();
        img.onerror = () => reject(new Error('Image load failed'));
        img.onload  = () => {
          try {
            resolve(compress(downscale(autoCrop(removeBackground(img)))));
          } catch (e) { reject(e); }
        };
        img.src = ev.target.result;
      };
      reader.readAsDataURL(file);
    });
  }

  /* ─────────────────────────────────────────────────────────────────
     STAMPING MODE
     ──────────────────────────────────────────────────────────────────
     Instead of dropping the stamp at a random centre position, we enter
     a lightweight "stamping mode" where:
       1. The canvas cursor changes to a crosshair.
       2. A dismissable hint banner appears above the canvas.
       3. The next mouse:down on the canvas places the stamp centred
          exactly under the pointer.
       4. Stamping mode exits automatically (or on Escape).

     State: stampingMode flag + pending dataUrl / color stored in closure.
  ───────────────────────────────────────────────────────────────── */
  let stampingMode        = false;
  let pendingStampDataUrl = null;
  let pendingStampColor   = null;
  let ghostStamp          = null;   // semi-transparent Fabric image that follows the cursor

  // Hint banner injected once above #pdf-wrapper
  const stampHint = (() => {
    const el = document.createElement('div');
    el.id = 'stamp-hint';
    el.style.cssText = `
      display: none;
      align-items: center;
      gap: 8px;
      padding: 7px 14px;
      background: #005bbf;
      color: #fff;
      font-family: Inter, sans-serif;
      font-size: 11px;
      font-weight: 700;
      border-radius: 8px;
      box-shadow: 0 4px 14px rgba(0,91,191,.30);
      user-select: none;
      letter-spacing: .02em;
    `;
    el.innerHTML = `
      <span class="material-symbols-outlined" style="font-size:16px;vertical-align:middle;">touch_app</span>
      Click on the document to place the stamp &nbsp;·&nbsp;
      <span id="stamp-hint-cancel"
            style="cursor:pointer;text-decoration:underline;opacity:.85;">Cancel (Esc)</span>
    `;
    // Insert above the paginationBar (which is already inserted above pdf-wrapper sibling)
    paginationBar.parentNode.insertBefore(el, paginationBar);
    el.querySelector('#stamp-hint-cancel').addEventListener('click', exitStampingMode);
    return el;
  })();

  function enterStampingMode(dataUrl, color) {
    if (!fabricCanvas) { alert('ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!'); return; }
    setDrawingMode(false);
    fabricCanvas.discardActiveObject();

    stampingMode        = true;
    pendingStampDataUrl = dataUrl;
    pendingStampColor   = color;

    // ── Fabric cursor fix ────────────────────────────────────────────
    // CSS cursor on the DOM element is overridden by Fabric on every
    // mouse-move.  The correct API is the canvas-level properties:
    fabricCanvas.defaultCursor = 'crosshair';   // cursor over empty canvas
    fabricCanvas.hoverCursor   = 'crosshair';   // cursor when hovering an object

    stampHint.style.display = 'flex';

    // ── Ghost object ─────────────────────────────────────────────────
    // Load the stamp image once; create a semi-transparent, non-
    // interactive Fabric image (the "ghost") that will follow the mouse.
    fabric.Image.fromURL(dataUrl, (imgObj) => {
      if (!stampingMode) return;   // user cancelled while image was loading

      const MAX_W = fabricCanvas.width * 0.30;
      if (imgObj.width > MAX_W) imgObj.scaleToWidth(MAX_W);

      imgObj.set({
        left:             -9999,   // park off-screen until first mouse:move
        top:              -9999,
        opacity:          0.50,    // semi-transparent ghost
        selectable:       false,   // not interactive
        evented:          false,   // doesn't trigger Fabric object events
        hasControls:      false,
        hasBorders:       false,
      });

      ghostStamp = imgObj;
      fabricCanvas.add(ghostStamp);
      fabricCanvas.requestRenderAll();
    }, { crossOrigin: null });
  }

  function exitStampingMode() {
    stampingMode        = false;
    pendingStampDataUrl = null;
    pendingStampColor   = null;

    // Remove ghost from canvas
    if (ghostStamp && fabricCanvas) {
      fabricCanvas.remove(ghostStamp);
      fabricCanvas.requestRenderAll();
    }
    ghostStamp = null;

    // Restore Fabric's default cursor behaviour
    if (fabricCanvas) {
      fabricCanvas.defaultCursor = 'default';
      fabricCanvas.hoverCursor   = 'move';
    }

    stampHint.style.display = 'none';
  }

  // Escape key cancels stamping mode
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && stampingMode) exitStampingMode();
  });

  function addStampToCanvas(dataUrl, color) {
    // Activate stamping mode — the actual placement happens in the
    // mouse:down handler wired up inside initFabricCanvas().
    enterStampingMode(dataUrl, color);
  }

  function buildSavedStampUI(btnId, label, color, storageKey) {
    const sidebarBtn = document.getElementById(btnId);
    if (!sidebarBtn) return;
    let wrapper = document.getElementById(`saved-ui-${storageKey}`);
    if (!wrapper) {
      wrapper    = document.createElement('div');
      wrapper.id = `saved-ui-${storageKey}`;
      sidebarBtn.parentNode.insertBefore(wrapper, sidebarBtn);
    }
    wrapper.innerHTML = '';
    const savedUrl = localStorage.getItem(storageKey);
    if (!savedUrl) { wrapper.style.display = 'none'; return; }

    wrapper.style.cssText = `display:flex;align-items:center;gap:8px;padding:6px 10px;
      margin-bottom:6px;background:#f0f6ff;border:1.5px solid ${color}44;
      border-radius:10px;font-family:Inter,sans-serif;`;

    const thumb = document.createElement('img');
    thumb.src   = savedUrl;
    thumb.style.cssText = `width:44px;height:44px;object-fit:contain;border-radius:6px;
      background:repeating-conic-gradient(#ccc 0% 25%,#fff 0% 50%) 0 0/8px 8px;
      cursor:pointer;border:1px solid ${color}55;flex-shrink:0;`;
    thumb.title = `Click to add ${label} to canvas`;
    thumb.addEventListener('click', () => addStampToCanvas(savedUrl, color));

    const lbl = document.createElement('span');
    lbl.textContent   = label;
    lbl.style.cssText = `flex:1;font-size:11px;font-weight:600;color:${color};line-height:1.3;`;

    const trashBtn = document.createElement('button');
    trashBtn.title       = `Clear saved ${label}`;
    trashBtn.innerHTML   = `<span class="material-symbols-outlined" style="font-size:18px;vertical-align:middle;">delete</span>`;
    trashBtn.style.cssText = `background:none;border:none;cursor:pointer;color:#b91c1c;
      padding:2px 4px;border-radius:6px;transition:background .15s;`;
    trashBtn.addEventListener('mouseover', () => trashBtn.style.background = '#fee2e2');
    trashBtn.addEventListener('mouseout',  () => trashBtn.style.background = 'none');
    trashBtn.addEventListener('click', () => {
      if (!confirm(`"${label}" ഡിലീറ്റ് ചെയ്യണോ?`)) return;
      localStorage.removeItem(storageKey);
      buildSavedStampUI(btnId, label, color, storageKey);
    });

    wrapper.append(thumb, lbl, trashBtn);
    const btnLabelEl = sidebarBtn.querySelector('[data-stamp-label]');
    if (btnLabelEl) btnLabelEl.textContent = `Re-upload ${label}`;
  }

  STAMP_DEFS.forEach(({ btnId, label, color, storageKey }) => {
    const sidebarBtn = document.getElementById(btnId);
    const imgPicker  = document.createElement('input');
    imgPicker.type   = 'file';
    imgPicker.accept = 'image/png,image/jpeg,image/gif,image/webp,image/svg+xml';
    imgPicker.style.display = 'none';
    document.body.appendChild(imgPicker);

    buildSavedStampUI(btnId, label, color, storageKey);

    sidebarBtn.addEventListener('click', () => {
      const savedUrl = localStorage.getItem(storageKey);
      if (savedUrl && fabricCanvas) { addStampToCanvas(savedUrl, color); return; }
      if (!fabricCanvas && !savedUrl) { alert('ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!'); return; }
      imgPicker.value = ''; imgPicker.click();
    });

    imgPicker.addEventListener('change', async () => {
      const file = imgPicker.files[0];
      if (!file) return;
      sidebarBtn.disabled = true; sidebarBtn.style.opacity = '0.6';
      const origHTML = sidebarBtn.innerHTML;
      sidebarBtn.innerHTML = `<span style="font-size:12px;">Processing…</span>`;
      try {
        const processedDataUrl = await processStampImage(file);
        try { localStorage.setItem(storageKey, processedDataUrl); }
        catch (se) { console.warn('localStorage quota exceeded:', se); }
        buildSavedStampUI(btnId, label, color, storageKey);
        if (fabricCanvas) addStampToCanvas(processedDataUrl, color);
        else alert(`${label} saved! Proceed to Sign to stamp it.`);
      } catch (err) {
        console.error('Stamp processing error:', err);
        alert(`"${label}" process ചെയ്യാൻ കഴിഞ്ഞില്ല: ${err.message}`);
      } finally {
        sidebarBtn.innerHTML = origHTML;
        sidebarBtn.disabled  = false; sidebarBtn.style.opacity = '1';
      }
    });
  });


  /* ════════════════════════════════════════════════════════════════════
     §7  FREEHAND DRAW TOOL
  ════════════════════════════════════════════════════════════════════ */
  const toggleDrawBtn  = document.getElementById('btn-toggle-draw');
  const drawBtnLabel   = document.getElementById('draw-btn-label');
  const drawToolbar    = document.getElementById('draw-toolbar');
  const brushSizeInput = document.getElementById('brush-size');
  const brushSizeLbl   = document.getElementById('brush-size-label');
  const clearDrawBtn   = document.getElementById('btn-clear-draw');

  let isDrawingMode    = false;
  let currentInkColor  = '#005bbf';
  let currentBrushSize = 3;

  toggleDrawBtn.addEventListener('click', () => {
    if (!fabricCanvas) { alert('ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!'); return; }
    setDrawingMode(!isDrawingMode);
  });

  function setDrawingMode(active) {
    if (!fabricCanvas) return;
    isDrawingMode = active;
    fabricCanvas.isDrawingMode = active;
    if (active) {
      applyBrushSettings();
      drawBtnLabel.textContent        = 'Stop Drawing';
      toggleDrawBtn.style.background  = '#dc2626';
      toggleDrawBtn.style.color       = '#fff';
      toggleDrawBtn.style.borderColor = '#dc2626';
      drawToolbar.style.display       = 'flex';
      fabricCanvas.wrapperEl.classList.add('drawing-active');
      fabricCanvas.off('path:created');
      fabricCanvas.on('path:created', () => {
        fabricCanvas.discardActiveObject(); fabricCanvas.requestRenderAll();
      });
    } else {
      fabricCanvas.off('path:created');
      drawBtnLabel.textContent        = 'Start Drawing';
      toggleDrawBtn.style.background  = '';
      toggleDrawBtn.style.color       = '';
      toggleDrawBtn.style.borderColor = '';
      drawToolbar.style.display       = 'none';
      fabricCanvas.wrapperEl.classList.remove('drawing-active');
    }
  }

  function applyBrushSettings() {
    if (!fabricCanvas) return;
    const b = fabricCanvas.freeDrawingBrush;
    b.color = currentInkColor; b.width = currentBrushSize; b.decimate = 0;
  }

  document.querySelectorAll('.color-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      currentInkColor = btn.dataset.color;
      document.querySelectorAll('.color-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      if (isDrawingMode) applyBrushSettings();
    });
  });

  brushSizeInput.addEventListener('input', () => {
    currentBrushSize = parseInt(brushSizeInput.value, 10);
    brushSizeLbl.textContent = currentBrushSize;
    if (isDrawingMode) applyBrushSettings();
  });

  clearDrawBtn.addEventListener('click', () => {
    if (!fabricCanvas) return;
    fabricCanvas.getObjects('path').forEach(o => fabricCanvas.remove(o));
    fabricCanvas.requestRenderAll();
  });


  /* ════════════════════════════════════════════════════════════════════
     §8  DELETE  (sidebar button + Delete / Backspace key)
  ════════════════════════════════════════════════════════════════════ */
  deleteBtnEl.addEventListener('click', deleteSelected);

  document.addEventListener('keydown', e => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    if (e.key !== 'Delete' && e.key !== 'Backspace') return;
    deleteSelected();
  });

  function deleteSelected() {
    if (!fabricCanvas) return;
    const sel = fabricCanvas.getActiveObjects();
    if (!sel.length) return;
    sel.forEach(obj => fabricCanvas.remove(obj));
    fabricCanvas.discardActiveObject();
    fabricCanvas.requestRenderAll();
  }


  /* ════════════════════════════════════════════════════════════════════
     §9  DOWNLOAD  —  embed into the real vector PDF via pdf-lib
     ──────────────────────────────────────────────────────────────────
     JPEG overlay drawn ON TOP with BlendMode.Multiply.
     White pixels → invisible (result = background).
     Dark ink → darkened, fully readable.
     Rotation-aware: inverse-rotate snapshot before embedding.
  ════════════════════════════════════════════════════════════════════ */
  document.getElementById('btn-download').addEventListener('click', async () => {
    if (!pdfDoc || !pdfBytes) {
      alert('ഡൗൺലോഡ് ചെയ്യാൻ ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!');
      return;
    }

    saveCurrentPageObjects();

    // const hasObjects = [...pageObjects.values()].some(j => j.objects && j.objects.length > 0);
    // if (!hasObjects) { alert('ദയവായി ആദ്യം ഒരു ഒപ്പ് ചേർക്കുക!'); return; }

    const dlBtn = document.getElementById('btn-download');
    dlBtn.disabled = true;

    showLoader('Generating signed PDF… Please wait.');
    await new Promise(r => setTimeout(r, 10));

    try {
      // ignoreEncryption: true preserves visuals of signed/encrypted PDFs.
      // pdfBytes.slice(0) creates a fresh copy so repeated downloads work.
      const pdfLibDoc = await PDFLib.PDFDocument.load(pdfBytes.slice(0), { ignoreEncryption: true });
      const pdfPages  = pdfLibDoc.getPages();

      for (let p = 1; p <= totalPages; p++) {
        const json = pageObjects.get(p);
        if (!json || !json.objects || json.objects.length === 0) continue;
        const dims = pageDimensions.get(p);
        if (!dims) continue;

        const snapCanvas    = await snapshotFabricPage(json, dims.canvasW, dims.canvasH);
        const pdfPage       = pdfPages[p - 1];
        const mediaBox      = pdfPage.getMediaBox();
        const storedW       = mediaBox.width;
        const storedH       = mediaBox.height;
        const rotation      = pdfPage.getRotation().angle;
        const alignedCanvas = applyInverseRotation(snapCanvas, rotation);

        const jpegDataUrl = alignedCanvas.toDataURL('image/jpeg', 0.82);
        const jpegBase64  = jpegDataUrl.split(',')[1];
        const jpegBytes   = Uint8Array.from(atob(jpegBase64), c => c.charCodeAt(0));

        const embedded = await pdfLibDoc.embedJpg(jpegBytes);
        pdfPage.drawImage(embedded, {
          x: 0, y: 0, width: storedW, height: storedH,
          blendMode: PDFLib.BlendMode.Multiply,
        });
      }

      const finalBytes = await pdfLibDoc.save();
      const blob = new Blob([finalBytes], { type: 'application/pdf' });
      const url  = URL.createObjectURL(blob);
      Object.assign(document.createElement('a'), { href: url, download: 'signed-document.pdf' }).click();
      setTimeout(() => URL.revokeObjectURL(url), 10_000);

    } catch (err) {
      console.error('Download error:', err);
      alert('ഡൗൺലോഡ് പരാജയപ്പെട്ടു: ' + err.message);
    } finally {
      hideLoader();
      dlBtn.innerHTML =
        '<span class="material-symbols-outlined text-base">picture_as_pdf</span> Download Signed PDF';
      dlBtn.disabled = false;
    }
  });


  /* ── HELPER: snapshotFabricPage ──────────────────────────────────── */
  function snapshotFabricPage(fabricJson, canvasW, canvasH) {
    return new Promise((resolve, reject) => {
      try {
        const sc = new fabric.StaticCanvas(null, { width: canvasW, height: canvasH, enableRetinaScaling: false });
        sc.loadFromJSON(fabricJson, () => {
          sc.renderAll();
          const flat    = document.createElement('canvas');
          flat.width    = canvasW; flat.height = canvasH;
          const flatCtx = flat.getContext('2d');
          flatCtx.fillStyle = '#ffffff';
          flatCtx.fillRect(0, 0, canvasW, canvasH);
          flatCtx.drawImage(sc.getElement(), 0, 0);
          resolve(flat);
        });
      } catch (e) { reject(e); }
    });
  }

  /* ── HELPER: applyInverseRotation ────────────────────────────────── */
  function applyInverseRotation(srcCanvas, rotation) {
    if (!rotation || rotation === 0) return srcCanvas;
    const srcW = srcCanvas.width, srcH = srcCanvas.height;
    const dest = document.createElement('canvas');
    const ctx  = dest.getContext('2d');
    if (rotation === 90 || rotation === 270) { dest.width = srcH; dest.height = srcW; }
    else                                      { dest.width = srcW; dest.height = srcH; }
    ctx.save();
    switch (rotation) {
      case 90:  ctx.translate(0, dest.height);        ctx.rotate(-Math.PI / 2); break;
      case 180: ctx.translate(dest.width, dest.height); ctx.rotate(Math.PI);    break;
      case 270: ctx.translate(dest.width, 0);           ctx.rotate(Math.PI / 2); break;
    }
    ctx.drawImage(srcCanvas, 0, 0);
    ctx.restore();
    return dest;
  }

}); // end DOMContentLoaded