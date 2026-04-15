/* ═══════════════════════════════════════════════════════════════════════
   Ink Sign Studio — script.js  (v12)
   ─────────────────────────────────────────────────────────────────────
   §SPA  Tab-switching (Dashboard / Archive)
   §1    PIN screen  +  Forgot-PIN / Reset-App
   §2    State & DOM refs
   §2·5  Global loading overlay
   §3    Pagination bar (injected)
   §0    Document Organizer
         ┌── Organizer Pipeline ───────────────────────────────────────┐
         │  Upload (multiple PDFs) → pdf.js thumbnail render           │
         │  → HTML5 drag-and-drop reorder + ←/→ mobile buttons        │
         │  → page delete                                              │
         │  → page rotate (90° increments, stored in orgPages[].rot)  │
         │  → "Proceed to Sign" → pdf-lib merge (eOffice flatten)     │
         │  → hand off to §4 signing canvas                           │
         └─────────────────────────────────────────────────────────────┘
   §4    PDF load from bytes → pdf.js render
   §5    Fabric.js canvas  (single instance, never disposed)
         allowTouchScrolling = true for mobile page scrolling
   §6    Image-based Stamps  (Ink Signature / Designation Seal / Office Seal)
         ┌── Image Processing & Storage Pipeline ──────────────────────┐
         │  ① Auto Background Removal  — white/light-grey → transparent│
         │  ② Auto-Crop (Bounding Box) — trim invisible padding        │
         │  ③a Downscale              — hard cap at 200 px             │
         │  ③b Compress to JPEG       — quality 0.75, target <15 KB    │
         │  ④  Persistent localStorage                                  │
         │  ⑤  Auto-Load on Start     — sidebar preview + trash button │
         └─────────────────────────────────────────────────────────────┘
   §7    Freehand draw tool
   §7·5  Add Text tool (ghost click-to-place UX, colour-linked to swatches)
   §8    Delete  (button + Delete/Backspace key)
   §9    Download  — pdf-lib, rotation-aware, BlendMode.Multiply on top
                     eOffice form flattening before save
   ═══════════════════════════════════════════════════════════════════════ */

document.addEventListener('DOMContentLoaded', () => {

  /* ════════════════════════════════════════════════════════════════════
     §SPA  TAB SWITCHING
  ════════════════════════════════════════════════════════════════════ */
  function switchTab(targetName) {
    document.querySelectorAll('.nav-tab').forEach(t => {
      t.classList.toggle('active', t.dataset.tab === targetName);
    });
    document.querySelectorAll('.tab-panel').forEach(p => {
      p.classList.toggle('active', p.id === 'tab-' + targetName);
    });
  }

  document.querySelectorAll('.nav-tab').forEach(tab => {
    tab.addEventListener('click', () => switchTab(tab.dataset.tab));
  });

  // Render real archive data whenever the Archive tab is opened
  document.querySelector('[data-tab="archive"]')
    ?.addEventListener('click', renderArchiveTab);

  // Start on dashboard
  switchTab('dashboard');


  /* ════════════════════════════════════════════════════════════════════
     §MOBILE  SIDEBAR TOGGLE
     ──────────────────────────────────────────────────────────────────
     The sidebar is a fixed, slide-in overlay on small screens.
     #btn-mobile-menu (hamburger) toggles it open/closed.
     #mobile-sidebar-backdrop (semi-opaque overlay) closes it on tap.
     We add/remove the `.sidebar-open` class, which the CSS @media
     (max-width: 767px) block converts to translateX(0).
     On md+ the sidebar is in normal flow; .sidebar-open has no effect.
  ════════════════════════════════════════════════════════════════════ */
  const _sidebar   = document.getElementById('app-sidebar');
  const _menuBtn   = document.getElementById('btn-mobile-menu');
  const _backdrop  = document.getElementById('mobile-sidebar-backdrop');

  function _openSidebar() {
    _sidebar?.classList.add('sidebar-open');
    _backdrop?.classList.remove('opacity-0', 'pointer-events-none');
    _backdrop?.classList.add('opacity-100');
    document.body.classList.add('overflow-hidden');
  }
  function _closeSidebar() {
    _sidebar?.classList.remove('sidebar-open');
    _backdrop?.classList.remove('opacity-100');
    _backdrop?.classList.add('opacity-0', 'pointer-events-none');
    document.body.classList.remove('overflow-hidden');
  }

  _menuBtn?.addEventListener('click', () => {
    _sidebar?.classList.contains('sidebar-open') ? _closeSidebar() : _openSidebar();
  });
  _backdrop?.addEventListener('click', _closeSidebar);

  // Auto-close sidebar when a major action is taken on mobile
  ['btn-proceed-sign', 'btn-back-organizer', 'btn-organizer-reset'].forEach(id => {
    document.getElementById(id)?.addEventListener('click', _closeSidebar);
  });


  /* ════════════════════════════════════════════════════════════════════
     §ARCHIVE  HISTORY  —  save & render signed-document records
  ════════════════════════════════════════════════════════════════════ */

  function saveArchiveRecord(name, size) {
    try {
      const history = JSON.parse(localStorage.getItem('iss_archive_history') || '[]');
      history.unshift({ name, date: new Date().toISOString(), size });
      localStorage.setItem('iss_archive_history', JSON.stringify(history));
    } catch (e) {
      console.warn('[Archive] Could not save record:', e);
    }
  }

  function _archiveFormatDate(isoString) {
    try {
      return new Date(isoString).toLocaleDateString('en-IN', {
        day: 'numeric', month: 'short', year: 'numeric',
      });
    } catch (_) { return isoString; }
  }

  function _archiveFormatSize(bytes) {
    if (!bytes || bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024)   return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
  }

  function _buildArchiveRow(record) {
    const row     = document.createElement('div');
    row.className = 'doc-row gap-3 sm:gap-5';
    row.dataset.filename = (record.name || '').toLowerCase();

    const formattedDate = _archiveFormatDate(record.date);
    const formattedSize = _archiveFormatSize(record.size);

    row.innerHTML = `
      <div class="shrink-0 bg-blue-50 rounded-xl flex items-center
                  justify-center text-primary-container"
           style="width:48px;height:48px;min-width:48px">
        <span class="material-symbols-outlined text-xl"
              style="font-variation-settings:'FILL' 1">description</span>
      </div>
      <div class="flex-1 min-w-0">
        <h4 class="font-headline font-bold text-on-surface text-sm leading-tight truncate">
          ${record.name.replace(/</g, '&lt;').replace(/>/g, '&gt;')}
        </h4>
        <p class="text-xs text-on-surface-variant/70 mt-0.5">
          Signed Document · ${formattedSize}
        </p>
      </div>
      <div class="hidden sm:flex px-5 border-x border-outline-variant/15 flex-col gap-0.5 shrink-0">
        <span class="text-[10px] text-on-surface-variant/60 font-bold uppercase tracking-wider">
          Signed Date
        </span>
        <span class="text-sm font-headline font-semibold text-on-surface">
          ${formattedDate}
        </span>
      </div>
      <div class="px-2 sm:px-5 flex items-center shrink-0">
        <span class="stat-chip bg-green-50 text-green-700">
          <span class="w-1.5 h-1.5 bg-green-500 rounded-full"></span>
          <span class="hidden xs:inline">Signed / </span>ഒപ്പിട്ടു
        </span>
      </div>
      <div class="flex gap-1.5 ml-auto shrink-0">
        <button class="w-9 h-9 flex items-center justify-center text-on-surface-variant
                       hover:text-primary hover:bg-surface-container-high
                       rounded-lg transition-colors" title="More options">
          <span class="material-symbols-outlined text-lg">more_vert</span>
        </button>
      </div>
    `;
    return row;
  }

  function renderArchiveTab() {
    const list = document.getElementById('archive-doc-list');
    if (!list) return;

    list.innerHTML = '';

    let history = [];
    try {
      history = JSON.parse(localStorage.getItem('iss_archive_history') || '[]');
    } catch (_) {}

    if (history.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'doc-row flex-col gap-3 items-center justify-center py-10 text-center';
      empty.innerHTML = `
        <span class="material-symbols-outlined text-4xl text-on-surface-variant/25"
              style="font-variation-settings:'FILL' 1">inventory_2</span>
        <p class="text-sm font-headline font-semibold text-on-surface-variant/50">
          No documents in archive
        </p>
        <p class="text-xs text-on-surface-variant/40 max-w-xs leading-relaxed">
          Every PDF you download will appear here automatically.<br>
          ഡൗൺലോഡ് ചെയ്ത ഡോക്യുമെന്റുകൾ ഇവിടെ കാണും.
        </p>
      `;
      list.appendChild(empty);
      return;
    }

    history.forEach(record => list.appendChild(_buildArchiveRow(record)));
  }


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
    if (pinTitle)    pinTitle.innerText    = 'Set New PIN';
    if (pinSubtitle) pinSubtitle.innerText = 'സുരക്ഷയ്ക്കായി ഒരു പുതിയ PIN സൃഷ്ടിക്കുക';
    if (btnText)     btnText.innerText     = 'Save PIN';
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
    /* ── FIX #2: Auto-clear Organizer on unlock ──────────────────────
       Explicitly reset orgPages and the grid DOM so no stale state
       lingers if the PIN screen is dismissed multiple times in a session.
    ─────────────────────────────────────────────────────────────────── */
    orgPages = [];
    if (thumbGrid) thumbGrid.innerHTML = '';
    if (orgPageCount) orgPageCount.textContent = '0 pages';

    pinScreen.style.opacity    = '0';
    pinScreen.style.transition = 'opacity 0.5s';
    setTimeout(() => (pinScreen.style.display = 'none'), 500);
    appDashboard.classList.remove('blur-md', 'pointer-events-none', 'grayscale-[0.2]', 'opacity-40');
  }

  /* ── Reset / Forgot-PIN handler ──────────────────────────────────── */
  function doResetApp() {
    const msg =
      'Reset the app?\n\n' +
      'This will permanently delete your saved PIN, all stored\n' +
      'Ink Signatures / Designation Seals / Office Seals,\n' +
      'and the entire document archive history.\n\n' +
      'The page will reload.';
    if (!confirm(msg)) return;
    [
      'ink_signer_pin',
      'saved_ink',
      'saved_desig_seal',
      'saved_office_seal',
      'iss_archive_history',
    ].forEach(k => localStorage.removeItem(k));
    location.reload();
  }

  document.getElementById('btn-forgot-pin')?.addEventListener('click', doResetApp);
  document.getElementById('btn-reset-app')?.addEventListener('click',  doResetApp);


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

  const RENDER_SCALE = 1.8;
  const THUMB_SCALE  = 0.25;

  let pdfDoc       = null;
  let pdfBytes     = null;
  let currentPage  = 1;
  let totalPages   = 1;
  let fabricCanvas = null;

  const pageObjects    = new Map();
  const pageDimensions = new Map();

  // Organizer state
  // Each entry: { srcBytes, srcPageIndex, label, thumbCanvas, rotation }
  // rotation: 0 | 90 | 180 | 270  (degrees, clockwise)
  let orgPages   = [];
  let dragSrcIdx = null;


  /* ════════════════════════════════════════════════════════════════════
     §2·5  GLOBAL LOADING OVERLAY
  ════════════════════════════════════════════════════════════════════ */
  const loadingOverlay = document.getElementById('loading-overlay');
  const loadingText    = document.getElementById('loading-text');

  function showLoader(msg = 'Please wait…') {
    if (loadingText)    loadingText.textContent = msg;
    if (loadingOverlay) loadingOverlay.classList.add('active');
  }
  function hideLoader() {
    if (loadingOverlay) loadingOverlay.classList.remove('active');
  }


  /* ════════════════════════════════════════════════════════════════════
     §3  PAGINATION BAR  (injected below #pdf-wrapper by JS)
  ════════════════════════════════════════════════════════════════════ */
  const paginationBar = document.createElement('div');
  paginationBar.id    = 'pagination-bar';
  Object.assign(paginationBar.style, {
    display: 'none', alignItems: 'center', justifyContent: 'center',
    gap: '14px', fontFamily: 'Inter, sans-serif', marginTop: '12px',
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
    btnPrev.style.opacity       = currentPage <= 1          ? '0.32' : '1';
    btnPrev.style.pointerEvents = currentPage <= 1          ? 'none'  : 'auto';
    btnNext.style.opacity       = currentPage >= totalPages ? '0.32' : '1';
    btnNext.style.pointerEvents = currentPage >= totalPages ? 'none'  : 'auto';
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
       ← / → mobile buttons also reorder by swapping adjacent items
       Rotate button increments orgPages[n].rotation by 90° (mod 360)
       Delete button splices orgPages[], re-renders grid
       "Proceed to Sign" → mergeAndLoad()
         → pdf-lib: create new doc, flatten eOffice forms, copyPages()
         → apply rotation via copiedPage.setRotation(degrees)
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
      if (f.size > 25 * 1024 * 1024) {
        alert(`"${f.name}" exceeds 25 MB limit — skipped.`);
      }
    }
    const validFiles = files.filter(f => f.size <= 25 * 1024 * 1024);
    if (!validFiles.length) return;

    switchTab('dashboard');
    await processUploadedFiles(validFiles);
  });

  /* processUploadedFiles — reads each PDF with pdf.js, renders THUMB_SCALE
     thumbnails for every page, and appends them to orgPages[].
     Each page entry now includes `rotation: 0` for the rotate tool. */
  async function processUploadedFiles(files) {
    if (!files.length) return;

    showOrganizerScreen();
    showLoader('Rendering thumbnails… Please wait.');
    await new Promise(r => setTimeout(r, 10));

    orgSpinner.classList.remove('hidden');

    for (const file of files) {
      const bytes     = await file.arrayBuffer();
      const pdfDocTmp = await pdfjsLib.getDocument({ data: bytes.slice(0) }).promise;

      for (let i = 0; i < pdfDocTmp.numPages; i++) {
        const page     = await pdfDocTmp.getPage(i + 1);
        const viewport = page.getViewport({ scale: THUMB_SCALE });

        const tc     = document.createElement('canvas');
        tc.width     = viewport.width;
        tc.height    = viewport.height;
        await page.render({ canvasContext: tc.getContext('2d'), viewport }).promise;

        orgPages.push({
          srcBytes:     bytes,
          srcPageIndex: i,
          label:        `${file.name.replace(/\.pdf$/i, '')}  p.${i + 1}`,
          thumbCanvas:  tc,
          rotation:     0,    // ── FIX #3: rotation state (0|90|180|270)
        });
      }
    }

    orgSpinner.classList.add('hidden');
    hideLoader();
    renderThumbnails();
  }

  /* ── FIX #3 HELPER: return a new canvas with src rotated by `degrees` ──
     Handles 90/180/270; returns src unchanged for 0.
     Used to visually display the rotated thumbnail in the organizer grid. */
  function getRotatedCanvas(src, degrees) {
    if (!degrees || degrees === 0) return src;
    const rad  = degrees * Math.PI / 180;
    const swap = degrees === 90 || degrees === 270;
    const destW = swap ? src.height : src.width;
    const destH = swap ? src.width  : src.height;
    const dest  = document.createElement('canvas');
    dest.width  = destW;
    dest.height = destH;
    const ctx   = dest.getContext('2d');
    ctx.save();
    ctx.translate(destW / 2, destH / 2);
    ctx.rotate(rad);
    ctx.drawImage(src, -src.width / 2, -src.height / 2);
    ctx.restore();
    return dest;
  }

  /* renderThumbnails — rebuilds the entire #thumb-grid DOM from orgPages[].
     Each card now has:
       drag handle · display canvas · label
       · ← Move Left · → Move Right  (mobile-friendly reorder buttons)
       · rotate btn · delete btn
     ─────────────────────────────────────────────────────────────────
     MOBILE REORDER FIX: ← and → buttons swap the tapped card with its
     neighbour in orgPages[], then re-render. Disabled (muted) at edges. */
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
      card.className   = 'thumb-card';
      card.draggable   = true;
      card.dataset.idx = idx;

      // ── FIX #3: render the thumbnail with any stored rotation applied ──
      const rotatedSrc = getRotatedCanvas(pg.thumbCanvas, pg.rotation);
      const display    = document.createElement('canvas');
      display.width    = rotatedSrc.width;
      display.height   = rotatedSrc.height;
      display.getContext('2d').drawImage(rotatedSrc, 0, 0);

      // Footer: label + move buttons + rotate button + delete button
      const footer = document.createElement('div');
      footer.className = 'thumb-footer';

      const labelEl = document.createElement('span');
      labelEl.textContent   = pg.label;
      labelEl.style.cssText = 'flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:52px;font-size:10px;';

      // ── MOBILE REORDER: Move Left (←) button ──────────────────────
      const moveLeftBtn = document.createElement('button');
      moveLeftBtn.className = 'thumb-del';
      moveLeftBtn.title     = 'Move Left';
      const isFirst         = idx === 0;
      moveLeftBtn.style.cssText = `color:#005bbf;margin-right:1px;${isFirst ? 'opacity:.22;pointer-events:none;' : ''}`;
      moveLeftBtn.innerHTML = '<span class="material-symbols-outlined" style="font-size:14px">chevron_left</span>';
      moveLeftBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (idx === 0) return;
        [orgPages[idx - 1], orgPages[idx]] = [orgPages[idx], orgPages[idx - 1]];
        renderThumbnails();
      });

      // ── MOBILE REORDER: Move Right (→) button ─────────────────────
      const moveRightBtn = document.createElement('button');
      moveRightBtn.className = 'thumb-del';
      moveRightBtn.title     = 'Move Right';
      const isLast           = idx === orgPages.length - 1;
      moveRightBtn.style.cssText = `color:#005bbf;margin-right:2px;${isLast ? 'opacity:.22;pointer-events:none;' : ''}`;
      moveRightBtn.innerHTML = '<span class="material-symbols-outlined" style="font-size:14px">chevron_right</span>';
      moveRightBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (idx === orgPages.length - 1) return;
        [orgPages[idx], orgPages[idx + 1]] = [orgPages[idx + 1], orgPages[idx]];
        renderThumbnails();
      });

      // ── FIX #3: Rotate button ──────────────────────────────────────
      const rotBtn = document.createElement('button');
      rotBtn.className = 'thumb-del';   // reuse same base style (overridden below)
      rotBtn.title     = 'Rotate 90°';
      rotBtn.style.cssText = 'color:#005bbf;margin-right:2px;';
      rotBtn.innerHTML = '<span class="material-symbols-outlined" style="font-size:14px">rotate_right</span>';
      rotBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        orgPages[idx].rotation = ((orgPages[idx].rotation || 0) + 90) % 360;
        renderThumbnails();
      });

      const delBtn = document.createElement('button');
      delBtn.className = 'thumb-del';
      delBtn.title     = 'Remove this page';
      delBtn.innerHTML = '<span class="material-symbols-outlined" style="font-size:14px">close</span>';
      delBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        orgPages.splice(idx, 1);
        renderThumbnails();
      });

      footer.append(labelEl, moveLeftBtn, moveRightBtn, rotBtn, delBtn);

      // Drag-handle grip dots (decorative)
      const handle = document.createElement('div');
      handle.className = 'drag-handle';
      for (let d = 0; d < 6; d++) handle.append(document.createElement('span'));

      card.append(display, footer, handle);

      // ── HTML5 Drag & Drop ──────────────────────────────────────────
      card.addEventListener('dragstart', (e) => {
        dragSrcIdx = idx;
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', idx);
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

        const [moved] = orgPages.splice(dragSrcIdx, 1);
        const insertAt = dragSrcIdx < targetIdx ? targetIdx - 1 : targetIdx;
        orgPages.splice(insertAt, 0, moved);
        renderThumbnails();
      });

      thumbGrid.appendChild(card);
    });
  }

  /* mergeAndLoad — uses pdf-lib to copy pages in current orgPages order
     into a brand-new PDF, then passes the resulting bytes to loadPdfBytes().

     ── FIX #3: Page rotation export ────────────────────────────────────
     For each page, copiedPage.setRotation(PDFLib.degrees(pg.rotation))
     is called after copying so the organizer rotation is baked into the
     merged PDF before it is handed to the signing canvas.

     eOffice / Digital Signature Flattening:
     srcDoc.getForm().flatten() inside try/catch converts AcroForm fields
     into flat vector graphics, preventing "Invalid Signature" warnings.
  */
  async function mergeAndLoad() {
    if (orgPages.length === 0) { alert('Please add at least one page.'); return; }

    const procBtn = document.getElementById('btn-proceed-sign');
    procBtn.disabled = true;

    showLoader('Merging PDFs… Please wait.');
    await new Promise(r => setTimeout(r, 10));

    try {
      const mergedDoc   = await PDFLib.PDFDocument.create();
      const srcDocCache = new Map();

      for (const pg of orgPages) {
        let srcDoc = srcDocCache.get(pg.srcBytes);
        if (!srcDoc) {
          srcDoc = await PDFLib.PDFDocument.load(pg.srcBytes, { ignoreEncryption: true });
          try { srcDoc.getForm().flatten(); } catch (_) {}
          srcDocCache.set(pg.srcBytes, srcDoc);
        }
        const [copiedPage] = await mergedDoc.copyPages(srcDoc, [pg.srcPageIndex]);

        // ── FIX #3: apply organizer rotation to the exported page ──────
        if (pg.rotation && pg.rotation !== 0) {
          // pdf-lib stores rotation as cumulative; setRotation replaces it.
          // We need to combine any pre-existing page rotation with our added rotation.
          const existingAngle = copiedPage.getRotation().angle || 0;
          const totalAngle    = (existingAngle + pg.rotation) % 360;
          copiedPage.setRotation(PDFLib.degrees(totalAngle));
        }

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

  document.getElementById('btn-proceed-sign').addEventListener('click', mergeAndLoad);

  document.getElementById('btn-back-organizer').addEventListener('click', () => {
    showOrganizerScreen();
  });

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
     ──────────────────────────────────────────────────────────────────
     Calling dispose() + new fabric.Canvas() removes the original
     <canvas> element from the DOM, permanently breaking re-init.
     We avoid this by using setWidth/setHeight + clear() instead.

     allowTouchScrolling = true  lets the browser handle vertical pan
     gestures on mobile when the user is not actively drawing/stamping,
     so the page remains scrollable by touch on the canvas area.
  ════════════════════════════════════════════════════════════════════ */
  function initFabricCanvas(width, height) {
    if (fabricCanvas) return;   // guard: created once only

    const el  = document.getElementById('fabric-canvas-el');
    el.width  = width;
    el.height = height;

    fabricCanvas = new fabric.Canvas('fabric-canvas-el', {
      width,
      height,
      selection:              true,
      preserveObjectStacking: true,
      renderOnAddRemove:      true,
      enableRetinaScaling:    false,
    });

    // ── MOBILE SCROLL FIX: allow touch events to propagate for page scrolling
    //    when the user is not in an active drawing/stamping interaction.
    fabricCanvas.allowTouchScrolling = true;

    Object.assign(fabricCanvas.wrapperEl.style, {
      position: 'absolute', top: '0', left: '0',
      width: width + 'px', height: height + 'px',
    });

    /* ── FIX #4: Selection events — show delete button and expose
       colour toolbar when a text object is selected so the user
       can change text colour without entering draw mode first. ────── */
    fabricCanvas.on('selection:created', () => {
      deleteBtnEl.classList.remove('hidden');
      _syncColorToolbarToSelection();
    });
    fabricCanvas.on('selection:updated', () => {
      deleteBtnEl.classList.remove('hidden');
      _syncColorToolbarToSelection();
    });
    fabricCanvas.on('selection:cleared',  () => {
      deleteBtnEl.classList.add('hidden');
      // Hide colour toolbar unless draw mode is still active
      if (!isDrawingMode) drawToolbar.style.display = 'none';
    });

    // ── Ghost follower: track cursor during stamping OR text-placing mode ──
    fabricCanvas.on('mouse:move', (opt) => {
      const p = opt.pointer;

      if (stampingMode && ghostStamp) {
        ghostStamp.set({
          left: p.x - ghostStamp.getScaledWidth()  / 2,
          top:  p.y - ghostStamp.getScaledHeight() / 2,
        });
        ghostStamp.setCoords();
        fabricCanvas.requestRenderAll();
      }

      if (textPlacingMode && ghostText) {
        ghostText.set({ left: p.x, top: p.y });
        ghostText.setCoords();
        fabricCanvas.requestRenderAll();
      }
    });

    // ── Click handler: place stamp OR place text at exact click point ──
    fabricCanvas.on('mouse:down', (opt) => {

      // Stamping mode: place real stamp at the click point
      if (stampingMode && pendingStampDataUrl) {
        const pointer = opt.pointer;
        const dataUrl = pendingStampDataUrl;
        const color   = pendingStampColor;

        exitStampingMode();

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

        return;
      }

      // Text placing mode: place IText at exact click point and enter editing
      if (textPlacingMode) {
        const pointer  = opt.pointer;
        const fontSize = Math.max(16, Math.round(fabricCanvas.width * 0.025));

        exitTextPlacingMode();

        const textObj = new fabric.IText('Type here', {
          left:               pointer.x,
          top:                pointer.y,
          originX:            'left',
          originY:            'top',
          fontFamily:         'Inter, Arial, sans-serif',
          fontSize,
          fill:               currentInkColor,
          fontWeight:         '600',
          hasControls:        true,
          hasBorders:         true,
          borderColor:        currentInkColor,
          cornerColor:        currentInkColor,
          cornerSize:         10,
          transparentCorners: false,
          cornerStyle:        'circle',
          borderDashArray:    [4, 3],
          editable:           true,
        });

        fabricCanvas.add(textObj);
        fabricCanvas.setActiveObject(textObj);
        textObj.enterEditing();
        textObj.selectAll();
        fabricCanvas.requestRenderAll();

        // Reveal colour toolbar so the user can immediately pick a text colour
        if (!isDrawingMode) drawToolbar.style.display = 'flex';

        return;
      }
    });

    fabricCanvas.freeDrawingBrush          = new fabric.PencilBrush(fabricCanvas);
    fabricCanvas.freeDrawingBrush.color    = currentInkColor;
    fabricCanvas.freeDrawingBrush.width    = currentBrushSize;
    fabricCanvas.freeDrawingBrush.decimate = 0;
  }

  /* ── Show/hide colour toolbar based on what is selected ──────────── */
  function _syncColorToolbarToSelection() {
    if (!fabricCanvas) return;
    const obj    = fabricCanvas.getActiveObject();
    const isText = obj && (obj.type === 'i-text' || obj.type === 'textbox');
    if (isText && !isDrawingMode) {
      // Expose the colour swatches so the user can pick a text colour
      drawToolbar.style.display = 'flex';
      // Sync the active colour button to the text's current fill
      const fill = obj.get('fill') || currentInkColor;
      document.querySelectorAll('.color-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.color === fill);
      });
      // Sync the size slider to reflect the text object's current font size
      const fontSize  = obj.get('fontSize') || 48;
      const sliderVal = Math.round(fontSize / 3);
      const clamped   = Math.max(
        parseInt(brushSizeInput.min, 10) || 1,
        Math.min(parseInt(brushSizeInput.max, 10) || 20, sliderVal)
      );
      brushSizeInput.value     = clamped;
      brushSizeLbl.textContent = clamped;
    } else if (!isText && !isDrawingMode) {
      drawToolbar.style.display = 'none';
    }
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

  /* ── Pipeline step ①: background removal ─────────────────────────── */
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

  /* ── Pipeline step ②: auto-crop to bounding box ─────────────────── */
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
    const PAD  = 2;
    minX = Math.max(0, minX - PAD); minY = Math.max(0, minY - PAD);
    maxX = Math.min(width - 1, maxX + PAD); maxY = Math.min(height - 1, maxY + PAD);
    const cropW = maxX - minX + 1, cropH = maxY - minY + 1;
    const dest  = document.createElement('canvas');
    dest.width  = cropW; dest.height = cropH;
    dest.getContext('2d').drawImage(srcCanvas, minX, minY, cropW, cropH, 0, 0, cropW, cropH);
    return dest;
  }

  /* ── Pipeline step ③a: downscale to MAX_STAMP_PX on longest edge ─── */
  function downscale(srcCanvas) {
    const longest = Math.max(srcCanvas.width, srcCanvas.height);
    if (longest <= MAX_STAMP_PX) return srcCanvas;
    const ratio = MAX_STAMP_PX / longest;
    const destW = Math.max(1, Math.round(srcCanvas.width  * ratio));
    const destH = Math.max(1, Math.round(srcCanvas.height * ratio));
    const dest  = document.createElement('canvas');
    dest.width  = destW; dest.height = destH;
    const ctx   = dest.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(srcCanvas, 0, 0, destW, destH);
    return dest;
  }

  /* ── Pipeline step ③b: compress to JPEG, safety loop ────────────── */
  function compress(scaledCanvas) {
    const origW = scaledCanvas.width, origH = scaledCanvas.height;
    const flat  = document.createElement('canvas');
    flat.width  = origW; flat.height = origH;
    const fc    = flat.getContext('2d');
    fc.fillStyle = '#ffffff'; fc.fillRect(0, 0, origW, origH);
    fc.drawImage(scaledCanvas, 0, 0);

    let scale   = 1.0;
    let dataUrl = flat.toDataURL('image/jpeg', JPEG_QUALITY);

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

  /* ── Full pipeline: File → processed data URL ────────────────────── */
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


  /* ════════════════════════════════════════════════════════════════════
     STAMPING MODE
  ════════════════════════════════════════════════════════════════════ */
  let stampingMode        = false;
  let pendingStampDataUrl = null;
  let pendingStampColor   = null;
  let ghostStamp          = null;

  // ── TEXT PLACING MODE state ────────────────────────────────────────
  let textPlacingMode = false;
  let ghostText       = null;

  // Stamp hint banner — injected once above the paginationBar
  const stampHint = (() => {
    const el = document.createElement('div');
    el.id    = 'stamp-hint';
    el.style.cssText = `
      display: none; align-items: center; gap: 8px;
      padding: 7px 14px; background: #005bbf; color: #fff;
      font-family: Inter, sans-serif; font-size: 11px; font-weight: 700;
      border-radius: 8px; box-shadow: 0 4px 14px rgba(0,91,191,.30);
      user-select: none; letter-spacing: .02em;
    `;
    el.innerHTML = `
      <span class="material-symbols-outlined" style="font-size:16px;vertical-align:middle;">touch_app</span>
      Click on the document to place the stamp &nbsp;·&nbsp;
      <span id="stamp-hint-cancel"
            style="cursor:pointer;text-decoration:underline;opacity:.85;">Cancel (Esc)</span>
    `;
    paginationBar.parentNode.insertBefore(el, paginationBar);
    el.querySelector('#stamp-hint-cancel').addEventListener('click', exitStampingMode);
    return el;
  })();

  // ── Text hint banner — injected above paginationBar (alongside stampHint) ──
  const textHint = (() => {
    const el = document.createElement('div');
    el.id    = 'text-hint';
    el.style.cssText = `
      display: none; align-items: center; gap: 8px;
      padding: 7px 14px; background: #7c3aed; color: #fff;
      font-family: Inter, sans-serif; font-size: 11px; font-weight: 700;
      border-radius: 8px; box-shadow: 0 4px 14px rgba(124,58,237,.30);
      user-select: none; letter-spacing: .02em;
    `;
    el.innerHTML = `
      <span class="material-symbols-outlined" style="font-size:16px;vertical-align:middle;">text_fields</span>
      Click on the document to place text &nbsp;·&nbsp;
      <span id="text-hint-cancel"
            style="cursor:pointer;text-decoration:underline;opacity:.85;">Cancel (Esc)</span>
    `;
    paginationBar.parentNode.insertBefore(el, paginationBar);
    el.querySelector('#text-hint-cancel').addEventListener('click', exitTextPlacingMode);
    return el;
  })();

  function enterStampingMode(dataUrl, color) {
    if (!fabricCanvas) { alert('ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!'); return; }
    setDrawingMode(false);
    exitTextPlacingMode();
    fabricCanvas.discardActiveObject();

    stampingMode        = true;
    pendingStampDataUrl = dataUrl;
    pendingStampColor   = color;

    fabricCanvas.defaultCursor = 'crosshair';
    fabricCanvas.hoverCursor   = 'crosshair';

    stampHint.style.display = 'flex';

    fabric.Image.fromURL(dataUrl, (imgObj) => {
      if (!stampingMode) return;

      const MAX_W = fabricCanvas.width * 0.30;
      if (imgObj.width > MAX_W) imgObj.scaleToWidth(MAX_W);

      imgObj.set({
        left:        -9999,
        top:         -9999,
        opacity:     0.50,
        selectable:  false,
        evented:     false,
        hasControls: false,
        hasBorders:  false,
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

    if (ghostStamp && fabricCanvas) {
      fabricCanvas.remove(ghostStamp);
      fabricCanvas.requestRenderAll();
    }
    ghostStamp = null;

    if (fabricCanvas) {
      fabricCanvas.defaultCursor = 'default';
      fabricCanvas.hoverCursor   = 'move';
    }

    stampHint.style.display = 'none';
  }

  /* ── Text placing mode: enter / exit ─────────────────────────────── */
  function enterTextPlacingMode() {
    if (!fabricCanvas) { alert('ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!'); return; }
    setDrawingMode(false);
    exitStampingMode();
    fabricCanvas.discardActiveObject();

    textPlacingMode = true;

    fabricCanvas.defaultCursor = 'crosshair';
    fabricCanvas.hoverCursor   = 'crosshair';

    textHint.style.display = 'flex';

    // Create semi-transparent ghost text that follows the cursor
    const fontSize = Math.max(16, Math.round(fabricCanvas.width * 0.025));
    const gt = new fabric.IText('Type here', {
      left:        -9999,
      top:         -9999,
      originX:     'left',
      originY:     'top',
      fontFamily:  'Inter, Arial, sans-serif',
      fontSize,
      fill:        currentInkColor,
      fontWeight:  '600',
      opacity:     0.42,
      selectable:  false,
      evented:     false,
      hasControls: false,
      hasBorders:  false,
      editable:    false,
    });
    ghostText = gt;
    fabricCanvas.add(ghostText);
    fabricCanvas.requestRenderAll();
  }

  function exitTextPlacingMode() {
    textPlacingMode = false;

    if (ghostText && fabricCanvas) {
      fabricCanvas.remove(ghostText);
      fabricCanvas.requestRenderAll();
    }
    ghostText = null;

    if (fabricCanvas) {
      fabricCanvas.defaultCursor = 'default';
      fabricCanvas.hoverCursor   = 'move';
    }

    textHint.style.display = 'none';
  }

  // Escape key exits both stamping and text placing modes
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && stampingMode)      exitStampingMode();
    if (e.key === 'Escape' && textPlacingMode)   exitTextPlacingMode();
  });

  function addStampToCanvas(dataUrl, color) {
    enterStampingMode(dataUrl, color);
  }

  /* ── Sidebar saved-stamp preview UI ─────────────────────────────── */
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

    wrapper.style.cssText = `
      display:flex; align-items:center; gap:8px; padding:6px 10px;
      margin-bottom:6px; background:#f0f6ff;
      border:1.5px solid ${color}44; border-radius:10px;
      font-family:Inter,sans-serif;
    `;

    const thumb = document.createElement('img');
    thumb.src   = savedUrl;
    thumb.style.cssText = `
      width:44px; height:44px; object-fit:contain; border-radius:6px;
      background:repeating-conic-gradient(#ccc 0% 25%,#fff 0% 50%) 0 0/8px 8px;
      cursor:pointer; border:1px solid ${color}55; flex-shrink:0;
    `;
    thumb.title = `Click to add ${label} to canvas`;
    thumb.addEventListener('click', () => addStampToCanvas(savedUrl, color));

    const lbl = document.createElement('span');
    lbl.textContent   = label;
    lbl.style.cssText = `flex:1; font-size:11px; font-weight:600; color:${color}; line-height:1.3;`;

    const trashBtn = document.createElement('button');
    trashBtn.title     = `Clear saved ${label}`;
    trashBtn.innerHTML = `<span class="material-symbols-outlined"
      style="font-size:18px;vertical-align:middle;">delete</span>`;
    trashBtn.style.cssText = `
      background:none; border:none; cursor:pointer; color:#b91c1c;
      padding:2px 4px; border-radius:6px; transition:background .15s;
    `;
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

  /* ── Wire each stamp button ──────────────────────────────────────── */
  STAMP_DEFS.forEach(({ btnId, label, color, storageKey }) => {
    const sidebarBtn = document.getElementById(btnId);
    if (!sidebarBtn) return;

    const imgPicker        = document.createElement('input');
    imgPicker.type         = 'file';
    imgPicker.accept       = 'image/png,image/jpeg,image/gif,image/webp,image/svg+xml';
    imgPicker.style.display = 'none';
    document.body.appendChild(imgPicker);

    buildSavedStampUI(btnId, label, color, storageKey);

    sidebarBtn.addEventListener('click', () => {
      const savedUrl = localStorage.getItem(storageKey);
      if (savedUrl && fabricCanvas) { addStampToCanvas(savedUrl, color); return; }
      if (!fabricCanvas && !savedUrl) { alert('ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!'); return; }
      imgPicker.value = '';
      imgPicker.click();
    });

    imgPicker.addEventListener('change', async () => {
      const file = imgPicker.files[0];
      if (!file) return;
      sidebarBtn.disabled      = true;
      sidebarBtn.style.opacity = '0.6';
      const origHTML = sidebarBtn.innerHTML;
      sidebarBtn.innerHTML = `<span style="font-size:12px;">Processing…</span>`;
      try {
        const processedDataUrl = await processStampImage(file);
        try { localStorage.setItem(storageKey, processedDataUrl); }
        catch (se) { console.warn('localStorage quota exceeded:', se); }
        buildSavedStampUI(btnId, label, color, storageKey);
        if (fabricCanvas) addStampToCanvas(processedDataUrl, color);
        else              alert(`${label} saved! Proceed to Sign to stamp it.`);
      } catch (err) {
        console.error('Stamp processing error:', err);
        alert(`"${label}" process ചെയ്യാൻ കഴിഞ്ഞില്ല: ${err.message}`);
      } finally {
        sidebarBtn.innerHTML     = origHTML;
        sidebarBtn.disabled      = false;
        sidebarBtn.style.opacity = '1';
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
    isDrawingMode              = active;
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
        fabricCanvas.discardActiveObject();
        fabricCanvas.requestRenderAll();
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
    const b    = fabricCanvas.freeDrawingBrush;
    b.color    = currentInkColor;
    b.width    = currentBrushSize;
    b.decimate = 0;
  }

  /* ── Colour swatch buttons — linked to both draw brush AND selected text ──
     FIX #4: If an IText / Textbox object is currently selected on the
     canvas, clicking a colour swatch now also updates that object's fill.
  ─────────────────────────────────────────────────────────────────────── */
  document.querySelectorAll('.color-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      currentInkColor = btn.dataset.color;
      document.querySelectorAll('.color-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      // Update brush if in drawing mode
      if (isDrawingMode) applyBrushSettings();

      // ── FIX #4: Update selected text object's fill colour ──────────
      if (fabricCanvas) {
        const activeObj = fabricCanvas.getActiveObject();
        if (activeObj && (activeObj.type === 'i-text' || activeObj.type === 'textbox')) {
          activeObj.set('fill', currentInkColor);
          fabricCanvas.requestRenderAll();
        }
      }
    });
  });

  brushSizeInput.addEventListener('input', () => {
    const sliderVal  = parseInt(brushSizeInput.value, 10);
    currentBrushSize = sliderVal;
    brushSizeLbl.textContent = sliderVal;

    // If a text object is currently selected, drive its fontSize instead
    if (fabricCanvas) {
      const activeObj = fabricCanvas.getActiveObject();
      if (activeObj && (activeObj.type === 'i-text' || activeObj.type === 'textbox')) {
        activeObj.set('fontSize', sliderVal * 3);
        fabricCanvas.requestRenderAll();
        return; // Don't also apply brush settings
      }
    }

    if (isDrawingMode) applyBrushSettings();
  });

  clearDrawBtn.addEventListener('click', () => {
    if (!fabricCanvas) return;
    fabricCanvas.getObjects('path').forEach(o => fabricCanvas.remove(o));
    fabricCanvas.requestRenderAll();
  });


  /* ════════════════════════════════════════════════════════════════════
     §7·5  ADD TEXT TOOL  —  ghost click-to-place UX
     ──────────────────────────────────────────────────────────────────
     Mirrors the Image Stamp flow:
       1. Click "Add Text" → enters textPlacingMode
          · cursor becomes crosshair
          · purple hint banner appears: "Click on the document to place text"
          · a semi-transparent ghost IText follows the cursor
       2. Click anywhere on the canvas → ghost is removed, a real
          editable fabric.IText is placed at the exact pointer coordinate
          and immediately enters editing mode (cursor blinks, text selected)
       3. Colour toolbar is revealed so the user can change text colour
       4. Pressing Esc or clicking "Cancel" exits without placing anything
     ──────────────────────────────────────────────────────────────────
     The colour is seeded from the currently active colour swatch.
  ════════════════════════════════════════════════════════════════════ */
  document.getElementById('btn-add-text')?.addEventListener('click', () => {
    if (!fabricCanvas) { alert('ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!'); return; }
    enterTextPlacingMode();
    // Close mobile sidebar so the canvas is fully visible while placing text
    _closeSidebar();
  });


  /* ════════════════════════════════════════════════════════════════════
     §8  DELETE  (sidebar button + Delete / Backspace key)
  ════════════════════════════════════════════════════════════════════ */
  deleteBtnEl.addEventListener('click', deleteSelected);

  document.addEventListener('keydown', e => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    if (e.key !== 'Delete' && e.key !== 'Backspace') return;
    if (e.target === pinInput) return;
    // Don't intercept Backspace while a fabric IText is being edited
    if (fabricCanvas && fabricCanvas.getActiveObject()?.isEditing) return;
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
       • White pixels (255,255,255) → result = background  (invisible)
       • Dark ink pixels             → result ≈ ink colour (fully visible)
     Mirrors the CSS mix-blend-mode: multiply on the Fabric canvas.
     Rotation-aware: inverse-rotate the snapshot before embedding.

     eOffice / Digital Signature Flattening:
     After loading the merged PDF we call pdfLibDoc.getForm().flatten()
     (wrapped in try/catch) to convert any remaining AcroForm fields
     into flat vector graphics.
  ════════════════════════════════════════════════════════════════════ */
  document.getElementById('btn-download').addEventListener('click', async () => {
    if (!pdfDoc || !pdfBytes) {
      alert('ഡൗൺലോഡ് ചെയ്യാൻ ആദ്യം ഒരു PDF അപ്‌ലോഡ് ചെയ്യുക!');
      return;
    }

    saveCurrentPageObjects();

    const dlBtn = document.getElementById('btn-download');
    dlBtn.disabled = true;

    showLoader('Generating signed PDF… Please wait.');
    await new Promise(r => setTimeout(r, 10));

    try {
      const pdfLibDoc = await PDFLib.PDFDocument.load(
        pdfBytes.slice(0),
        { ignoreEncryption: true }
      );

      // ── eOffice form flattening ──────────────────────────────────
      try { pdfLibDoc.getForm().flatten(); } catch (_) {}

      const pdfPages = pdfLibDoc.getPages();

      for (let p = 1; p <= totalPages; p++) {
        const json = pageObjects.get(p);
        if (!json || !json.objects || json.objects.length === 0) continue;
        const dims = pageDimensions.get(p);
        if (!dims) continue;

        const snapCanvas = await snapshotFabricPage(json, dims.canvasW, dims.canvasH);

        const pdfPage  = pdfPages[p - 1];
        const mediaBox = pdfPage.getMediaBox();
        const storedW  = mediaBox.width;
        const storedH  = mediaBox.height;
        const rotation = pdfPage.getRotation().angle;

        const alignedCanvas = applyInverseRotation(snapCanvas, rotation);

        const jpegDataUrl = alignedCanvas.toDataURL('image/jpeg', 0.82);
        const jpegBase64  = jpegDataUrl.split(',')[1];
        const jpegBytes   = Uint8Array.from(atob(jpegBase64), c => c.charCodeAt(0));

        const embedded = await pdfLibDoc.embedJpg(jpegBytes);
        pdfPage.drawImage(embedded, {
          x:         0,
          y:         0,
          width:     storedW,
          height:    storedH,
          blendMode: PDFLib.BlendMode.Multiply,
        });
      }

      const finalBytes = await pdfLibDoc.save();
      const blob     = new Blob([finalBytes], { type: 'application/pdf' });
      const filename = 'signed-document.pdf';
      const url      = URL.createObjectURL(blob);
      Object.assign(document.createElement('a'), {
        href: url, download: filename,
      }).click();
      setTimeout(() => URL.revokeObjectURL(url), 10_000);

      saveArchiveRecord(filename, blob.size);

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
        const sc = new fabric.StaticCanvas(null, {
          width: canvasW, height: canvasH, enableRetinaScaling: false,
        });
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

    if (rotation === 90 || rotation === 270) {
      dest.width  = srcH;
      dest.height = srcW;
    } else {
      dest.width  = srcW;
      dest.height = srcH;
    }

    ctx.save();
    switch (rotation) {
      case 90:  ctx.translate(0, dest.height);           ctx.rotate(-Math.PI / 2); break;
      case 180: ctx.translate(dest.width, dest.height);  ctx.rotate(Math.PI);      break;
      case 270: ctx.translate(dest.width, 0);            ctx.rotate(Math.PI / 2);  break;
    }
    ctx.drawImage(srcCanvas, 0, 0);
    ctx.restore();
    return dest;
  }

}); // end DOMContentLoaded
