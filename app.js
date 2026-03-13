// ============================================
// 英検スコア管理 — Main Application
// Firebase Auth + Firestore Integration
// ============================================

(function () {
  'use strict';

  // ---- Firebase Init ----
  firebase.initializeApp(firebaseConfig);
  const auth = firebase.auth();
  const db = firebase.firestore();
  const COLLECTION = 'eiken_scores';

  // ---- State ----
  let allData = [];
  let filteredData = [];
  let sortKey = 'name';
  let sortDir = 'asc';

  // ---- DOM refs ----
  const $ = (id) => document.getElementById(id);
  const loginScreen = $('loginScreen');
  const appContainer = $('appContainer');
  const loginBtn = $('loginBtn');
  const loginError = $('loginError');
  const logoutBtn = $('logoutBtn');
  const userInfo = $('userInfo');
  const tableBody = $('tableBody');
  const loadingIndicator = $('loadingIndicator');
  const csvDropToggle = $('csvDropToggle');
  const csvDropOverlay = $('csvDropOverlay');
  const csvDropZone = $('csvDropZone');
  const csvDropClose = $('csvDropClose');
  const csvFileInput = $('csvFileInput');
  const importProgress = $('importProgress');
  const progressFill = $('progressFill');
  const progressText = $('progressText');
  const toastEl = $('toast');
  const filterYear = $('filterYear');
  const filterGrade = $('filterGrade');
  const filterSession = $('filterSession');
  const filterResult = $('filterResult');
  const filterOverall = $('filterOverall');
  const searchInput = $('searchInput');
  const clearFilters = $('clearFilters');
  const filterCount = $('filterCount');
  const exportBtn = $('exportBtn');
  const detailModal = $('detailModal');
  const modalTitle = $('modalTitle');
  const modalBody = $('modalBody');
  const modalClose = $('modalClose');

  // ============================================
  // Authentication
  // ============================================

  const isLocalhost = location.hostname === 'localhost' || location.hostname === '127.0.0.1';

  // Handle redirect result (for production)
  if (!isLocalhost) {
    auth.getRedirectResult().then((result) => {
      if (result.user) {
        console.log('Login success:', result.user.email);
      }
    }).catch((err) => {
      console.error('Redirect error:', err);
      if (loginError) loginError.textContent = 'ログインに失敗しました: ' + err.message;
    });
  }

  auth.onAuthStateChanged((user) => {
    if (user) {
      // Check whitelist
      if (typeof ALLOWED_EMAILS !== 'undefined' && ALLOWED_EMAILS.length > 0) {
        if (!ALLOWED_EMAILS.includes(user.email)) {
          loginError.textContent = 'このアカウントにはアクセス権限がありません。';
          auth.signOut();
          return;
        }
      }
      loginScreen.style.display = 'none';
      appContainer.style.display = 'block';
      userInfo.textContent = user.displayName || user.email;
      loadData();
    } else {
      loginScreen.style.display = 'flex';
      appContainer.style.display = 'none';
    }
  });

  loginBtn.addEventListener('click', async () => {
    loginError.textContent = '';
    const provider = new firebase.auth.GoogleAuthProvider();
    provider.setCustomParameters({ prompt: 'select_account' });

    if (isLocalhost) {
      // Popup for localhost
      try {
        loginBtn.disabled = true;
        loginBtn.textContent = 'ログイン中...';
        await auth.signInWithPopup(provider);
      } catch (err) {
        console.error('Login error:', err);
        loginError.textContent = 'ログインに失敗しました: ' + err.message;
        loginBtn.disabled = false;
        loginBtn.innerHTML = '<svg width="20" height="20" viewBox="0 0 48 48"><path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/><path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/><path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/><path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/></svg> Googleでログイン';
      }
    } else {
      // Redirect for production (Vercel etc.)
      auth.signInWithRedirect(provider);
    }
  });

  logoutBtn.addEventListener('click', () => auth.signOut());

  // ============================================
  // Firestore Data
  // ============================================

  async function loadData() {
    loadingIndicator.classList.remove('hidden');
    try {
      const snapshot = await db.collection(COLLECTION).orderBy('name').get();
      allData = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));

      // If empty, seed from INITIAL_DATA if available
      if (allData.length === 0 && typeof INITIAL_DATA !== 'undefined' && INITIAL_DATA.length > 0) {
        showToast('初期データをインポート中...', 'success');
        await importBatch(INITIAL_DATA);
        return loadData(); // reload after seed
      }

      filteredData = [...allData];
      updateYearFilter();
      updateStats();
      applyFilters();
      loadingIndicator.classList.add('hidden');
    } catch (err) {
      console.error('Data load error:', err);
      loadingIndicator.innerHTML = '<span style="color:var(--fail)">データ読み込みエラー: ' + err.message + '</span>';
    }
  }

  async function importBatch(rows) {
    const BATCH_SIZE = 400; // Firestore batch limit is 500
    let total = rows.length;
    let done = 0;

    for (let i = 0; i < total; i += BATCH_SIZE) {
      const batch = db.batch();
      const chunk = rows.slice(i, i + BATCH_SIZE);
      chunk.forEach((row) => {
        const ref = db.collection(COLLECTION).doc();
        batch.set(ref, sanitizeRow(row));
      });
      await batch.commit();
      done += chunk.length;
      const pct = Math.round((done / total) * 100);
      progressFill.style.width = pct + '%';
      progressText.textContent = done + ' / ' + total + ' 件インポート済み';
    }
  }

  function sanitizeRow(row) {
    const clean = {};
    const fields = [
      'year', 'session', 'grade', 'schoolYear', 'name', 'exemption',
      'primaryResult', 'secondaryResult', 'overallResult',
      'bandPrimary', 'bandSecondary',
      'readingCSE', 'listeningCSE', 'writingCSE', 'primaryCSETotal',
      'speakingCSE', 'totalCSE',
      'readingCEFR', 'listeningCEFR', 'writingCEFR', 'speakingCEFR', 'overallCEFR',
      'readingQ1', 'readingQ2', 'readingQ3', 'readingQ4',
      'readingTotal', 'readingRate',
      'listeningQ1', 'listeningQ2', 'listeningQ3',
      'listeningTotal', 'listeningRate',
      'writingS1Content', 'writingS1Structure', 'writingS1Vocab', 'writingS1Grammar',
      'writingS2Content', 'writingS2Structure', 'writingS2Vocab', 'writingS2Grammar',
      'writingScore', 'writingRate',
      'speakingReading', 'speakingQA', 'speakingAttitude', 'speakingScore'
    ];
    fields.forEach((f) => {
      clean[f] = row[f] !== undefined && row[f] !== null ? row[f] : null;
    });
    return clean;
  }

  // ============================================
  // Stats
  // ============================================

  function updateStats() {
    const data = filteredData;
    $('statTotal').textContent = data.length;

    // 一次合格率 (exclude 一次免除)
    const primaryTargets = data.filter((d) => d.primaryResult && d.primaryResult !== '一次免除');
    const primaryPass = primaryTargets.filter((d) => d.primaryResult === '合格');
    const primaryRate = primaryTargets.length > 0 ? Math.round((primaryPass.length / primaryTargets.length) * 100) : 0;
    $('statPassRate').textContent = primaryRate + '%';

    // 最終合格率 (only those with overallResult value)
    const overallTargets = data.filter((d) => d.overallResult === '合格' || d.overallResult === '不合格');
    const overallPass = overallTargets.filter((d) => d.overallResult === '合格');
    const overallRate = overallTargets.length > 0 ? Math.round((overallPass.length / overallTargets.length) * 100) : 0;
    $('statOverallPass').textContent = overallRate + '%';

    // Unique students
    const names = new Set(data.map((d) => d.name).filter(Boolean));
    $('statStudents').textContent = names.size;
  }

  // ============================================
  // Dynamic Filter Options
  // ============================================

  function updateYearFilter() {
    const years = [...new Set(allData.map((d) => d.year).filter(Boolean))].sort();
    const current = filterYear.value;
    filterYear.innerHTML = '<option value="">全て</option>';
    years.forEach((y) => {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = y + '年度';
      filterYear.appendChild(opt);
    });
    if (current) filterYear.value = current;
  }

  // ============================================
  // Filters & Search
  // ============================================

  function applyFilters() {
    const year = filterYear.value;
    const grade = filterGrade.value;
    const session = filterSession.value;
    const result = filterResult.value;
    const overall = filterOverall.value;
    const query = searchInput.value.trim().toLowerCase();

    filteredData = allData.filter((d) => {
      if (year && String(d.year) !== year) return false;
      if (grade && d.grade !== grade) return false;
      if (session && d.session !== session) return false;
      if (result && d.primaryResult !== result) return false;
      if (overall && d.overallResult !== overall) return false;
      if (query && d.name && !d.name.toLowerCase().includes(query)) return false;
      return true;
    });

    sortData();
    updateStats();
    renderTable();
    filterCount.textContent = filteredData.length + '件表示 / ' + allData.length + '件中';
  }

  [filterYear, filterGrade, filterSession, filterResult, filterOverall].forEach((el) => {
    el.addEventListener('change', applyFilters);
  });

  let searchTimeout;
  searchInput.addEventListener('input', () => {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(applyFilters, 200);
  });

  clearFilters.addEventListener('click', () => {
    filterYear.value = '';
    filterGrade.value = '';
    filterSession.value = '';
    filterResult.value = '';
    filterOverall.value = '';
    searchInput.value = '';
    applyFilters();
  });

  // ============================================
  // Sorting
  // ============================================

  const gradeOrder = { '2級': 1, '準2級+': 2, '準2級': 3, '3級': 4, '4級': 5, '5級': 6 };

  function sortData() {
    filteredData.sort((a, b) => {
      let va = a[sortKey];
      let vb = b[sortKey];

      // Grade ordering
      if (sortKey === 'grade') {
        va = gradeOrder[va] || 99;
        vb = gradeOrder[vb] || 99;
      }

      if (va === null || va === undefined) va = sortDir === 'asc' ? Infinity : -Infinity;
      if (vb === null || vb === undefined) vb = sortDir === 'asc' ? Infinity : -Infinity;

      if (typeof va === 'number' && typeof vb === 'number') {
        return sortDir === 'asc' ? va - vb : vb - va;
      }
      const sa = String(va);
      const sb = String(vb);
      return sortDir === 'asc' ? sa.localeCompare(sb, 'ja') : sb.localeCompare(sa, 'ja');
    });
  }

  document.querySelectorAll('.data-table th[data-sort]').forEach((th) => {
    th.addEventListener('click', () => {
      const key = th.dataset.sort;
      if (sortKey === key) {
        sortDir = sortDir === 'asc' ? 'desc' : 'asc';
      } else {
        sortKey = key;
        sortDir = 'asc';
      }
      // Update header classes
      document.querySelectorAll('.data-table th').forEach((h) => h.classList.remove('sorted-asc', 'sorted-desc'));
      th.classList.add(sortDir === 'asc' ? 'sorted-asc' : 'sorted-desc');
      applyFilters();
    });
  });

  // ============================================
  // Table Rendering
  // ============================================

  function gradeClass(grade) {
    const map = { '2級': 'grade-2', '準2級+': 'grade-p2p', '準2級': 'grade-p2', '3級': 'grade-3', '4級': 'grade-4', '5級': 'grade-5' };
    return map[grade] || '';
  }

  function resultBadge(val) {
    if (!val) return '<span class="badge-none">—</span>';
    if (val === '合格') return '<span class="badge badge-pass">合格</span>';
    if (val === '不合格') return '<span class="badge badge-fail">不合格</span>';
    if (val === '一次免除') return '<span class="badge badge-exempt">免除</span>';
    return '<span class="badge-none">' + val + '</span>';
  }

  function rateCell(rate) {
    if (rate === null || rate === undefined) return '<span class="badge-none">—</span>';
    const pct = Math.round(rate * 100);
    const cls = pct >= 70 ? 'high' : pct >= 50 ? 'mid' : 'low';
    return '<div class="rate-bar">' +
      '<div class="rate-bar-track"><div class="rate-bar-fill ' + cls + '" style="width:' + pct + '%"></div></div>' +
      '<span>' + pct + '%</span></div>';
  }

  function numCell(val) {
    if (val === null || val === undefined) return '<span class="badge-none">—</span>';
    return val;
  }

  function renderTable() {
    if (filteredData.length === 0) {
      tableBody.innerHTML = '<tr><td colspan="21" style="text-align:center;padding:40px;color:var(--text-muted)">該当するデータがありません</td></tr>';
      return;
    }

    const html = filteredData.map((d, i) => {
      return '<tr data-idx="' + i + '">' +
        '<td class="sticky-col">' + (d.name || '—') + '</td>' +
        '<td>' + (d.year || '—') + '</td>' +
        '<td><span class="grade-badge ' + gradeClass(d.grade) + '">' + (d.grade || '—') + '</span></td>' +
        '<td>' + (d.session || '—') + '</td>' +
        '<td>' + (d.schoolYear || '—') + '</td>' +
        '<td>' + resultBadge(d.primaryResult) + '</td>' +
        '<td>' + resultBadge(d.secondaryResult) + '</td>' +
        '<td>' + resultBadge(d.overallResult) + '</td>' +
        '<td>' + (d.bandPrimary || '—') + '</td>' +
        '<td class="num-col">' + numCell(d.readingCSE) + '</td>' +
        '<td class="num-col">' + numCell(d.listeningCSE) + '</td>' +
        '<td class="num-col">' + numCell(d.writingCSE) + '</td>' +
        '<td class="num-col"><strong>' + numCell(d.primaryCSETotal) + '</strong></td>' +
        '<td class="num-col">' + numCell(d.speakingCSE) + '</td>' +
        '<td class="num-col"><strong>' + numCell(d.totalCSE) + '</strong></td>' +
        '<td class="num-col">' + numCell(d.readingTotal) + '</td>' +
        '<td class="num-col">' + rateCell(d.readingRate) + '</td>' +
        '<td class="num-col">' + numCell(d.listeningTotal) + '</td>' +
        '<td class="num-col">' + rateCell(d.listeningRate) + '</td>' +
        '<td class="num-col">' + numCell(d.writingScore) + '</td>' +
        '<td class="num-col">' + numCell(d.speakingScore) + '</td>' +
        '</tr>';
    }).join('');

    tableBody.innerHTML = html;

    // Row click → detail modal
    tableBody.querySelectorAll('tr').forEach((tr) => {
      tr.addEventListener('click', () => {
        const idx = parseInt(tr.dataset.idx);
        showDetail(filteredData[idx]);
      });
    });
  }

  // ============================================
  // Detail Modal
  // ============================================

  function showDetail(d) {
    modalTitle.textContent = d.name + '（' + d.grade + ' ' + d.session + '）';

    const section = (title, rows) => {
      const rowsHtml = rows
        .filter(([, v]) => v !== null && v !== undefined)
        .map(([l, v]) => '<div class="detail-row"><span class="detail-label">' + l + '</span><span class="detail-value">' + v + '</span></div>')
        .join('');
      if (!rowsHtml) return '';
      return '<div class="detail-section"><div class="detail-section-title">' + title + '</div>' + rowsHtml + '</div>';
    };

    const html = '<div class="detail-grid">' +
      '<div>' +
      section('基本情報', [
        ['年度', d.year],
        ['実施回', d.session],
        ['級', d.grade],
        ['学年', d.schoolYear],
        ['一次免除', d.exemption],
      ]) +
      section('合否', [
        ['一次合否', d.primaryResult],
        ['二次合否', d.secondaryResult],
        ['総合合否', d.overallResult],
        ['英検バンド一次', d.bandPrimary],
        ['英検バンド二次', d.bandSecondary],
      ]) +
      section('CEFR', [
        ['Reading', d.readingCEFR],
        ['Listening', d.listeningCEFR],
        ['Writing', d.writingCEFR],
        ['Speaking', d.speakingCEFR],
        ['4技能総合', d.overallCEFR],
      ]) +
      '</div><div>' +
      section('CSEスコア', [
        ['Reading', d.readingCSE],
        ['Listening', d.listeningCSE],
        ['Writing', d.writingCSE],
        ['一次合計', d.primaryCSETotal],
        ['Speaking', d.speakingCSE],
        ['総合', d.totalCSE],
      ]) +
      section('Reading 正答数', [
        ['大問1', d.readingQ1],
        ['大問2', d.readingQ2],
        ['大問3', d.readingQ3],
        ['大問4', d.readingQ4],
        ['合計', d.readingTotal],
        ['正答率', d.readingRate !== null && d.readingRate !== undefined ? Math.round(d.readingRate * 100) + '%' : null],
      ]) +
      section('Listening 正答数', [
        ['大問1', d.listeningQ1],
        ['大問2', d.listeningQ2],
        ['大問3', d.listeningQ3],
        ['合計', d.listeningTotal],
        ['正答率', d.listeningRate !== null && d.listeningRate !== undefined ? Math.round(d.listeningRate * 100) + '%' : null],
      ]) +
      section('Writing スコア', [
        ['設問1 内容', d.writingS1Content],
        ['設問1 構成', d.writingS1Structure],
        ['設問1 語い', d.writingS1Vocab],
        ['設問1 文法', d.writingS1Grammar],
        ['設問2 内容', d.writingS2Content],
        ['設問2 構成', d.writingS2Structure],
        ['設問2 語い', d.writingS2Vocab],
        ['設問2 文法', d.writingS2Grammar],
        ['合計', d.writingScore],
      ]) +
      section('Speaking', [
        ['リーディング', d.speakingReading],
        ['Q&A', d.speakingQA],
        ['アティチュード', d.speakingAttitude],
        ['合計', d.speakingScore],
      ]) +
      '</div></div>';

    modalBody.innerHTML = html;
    detailModal.classList.add('active');
  }

  modalClose.addEventListener('click', () => detailModal.classList.remove('active'));
  detailModal.addEventListener('click', (e) => {
    if (e.target === detailModal) detailModal.classList.remove('active');
  });

  // ============================================
  // CSV Import
  // ============================================

  csvDropToggle.addEventListener('click', () => csvDropOverlay.classList.add('active'));
  csvDropClose.addEventListener('click', () => csvDropOverlay.classList.remove('active'));
  csvDropOverlay.addEventListener('click', (e) => {
    if (e.target === csvDropOverlay) csvDropOverlay.classList.remove('active');
  });

  // Drag & drop
  csvDropZone.addEventListener('dragover', (e) => { e.preventDefault(); csvDropZone.classList.add('drag-over'); });
  csvDropZone.addEventListener('dragleave', () => csvDropZone.classList.remove('drag-over'));
  csvDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    csvDropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleCSVFile(file);
  });

  csvFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleCSVFile(file);
  });

  // CSV column mapping (CSV header → internal field)
  const CSV_HEADER_MAP = {
    '実施年度': 'year',
    '実施回': 'session',
    '級': 'grade',
    '学年': 'schoolYear',
    '氏名': 'name',
    '一次免除の有無': 'exemption',
    '一次合否': 'primaryResult',
    '二次合否': 'secondaryResult',
    '総合合否': 'overallResult',
    '英検バンド一次': 'bandPrimary',
    '英検バンド二次': 'bandSecondary',
    'Reading CSEスコア': 'readingCSE',
    'Listening CSEスコア': 'listeningCSE',
    'Writing CSEスコア': 'writingCSE',
    'CSEスコア 一次合計': 'primaryCSETotal',
    'Speaking CSEスコア': 'speakingCSE',
    '総合CSEスコア': 'totalCSE',
    'Reading CEFR': 'readingCEFR',
    'Listening CEFR': 'listeningCEFR',
    'Writing CEFR': 'writingCEFR',
    'Speaking CEFR': 'speakingCEFR',
    '4技能総合CEFR': 'overallCEFR',
    'Reading正答数（大問1）': 'readingQ1',
    'Reading正答数（大問2）': 'readingQ2',
    'Reading正答数（大問3）': 'readingQ3',
    'Reading正答数（大問4）': 'readingQ4',
    'Listening正答数（大問1）': 'listeningQ1',
    'Listening正答数（大問2）': 'listeningQ2',
    'Listening正答数（大問3）': 'listeningQ3',
    'Writing設問1スコア（内容）': 'writingS1Content',
    'Writing設問1スコア（構成）': 'writingS1Structure',
    'Writing設問1スコア（語い）': 'writingS1Vocab',
    'Writing設問1スコア（文法）': 'writingS1Grammar',
    'Writing設問2スコア（内容）': 'writingS2Content',
    'Writing設問2スコア（構成）': 'writingS2Structure',
    'Writing設問2スコア（語い）': 'writingS2Vocab',
    'Writing設問2スコア（文法）': 'writingS2Grammar',
    'Speaking得点（リーディング）': 'speakingReading',
    'Speaking得点（Q&A）': 'speakingQA',
    'Speaking得点（アティチュード）': 'speakingAttitude',
  };

  // Numeric fields
  const NUMERIC_FIELDS = new Set([
    'year', 'readingCSE', 'listeningCSE', 'writingCSE', 'primaryCSETotal',
    'speakingCSE', 'totalCSE',
    'readingQ1', 'readingQ2', 'readingQ3', 'readingQ4',
    'listeningQ1', 'listeningQ2', 'listeningQ3',
    'writingS1Content', 'writingS1Structure', 'writingS1Vocab', 'writingS1Grammar',
    'writingS2Content', 'writingS2Structure', 'writingS2Vocab', 'writingS2Grammar',
    'speakingReading', 'speakingQA', 'speakingAttitude',
  ]);

  async function handleCSVFile(file) {
    importProgress.style.display = 'block';
    progressFill.style.width = '0%';
    progressText.textContent = 'ファイルを読み込み中...';

    try {
      // Try Shift_JIS (cp932) first, then UTF-8
      let text;
      try {
        text = await readFileAsText(file, 'Shift_JIS');
      } catch {
        text = await readFileAsText(file, 'UTF-8');
      }

      const rows = parseCSV(text);
      if (rows.length < 2) {
        throw new Error('CSVにデータが含まれていません');
      }

      const headers = rows[0];
      const dataRows = rows.slice(1).filter((r) => r.some((c) => c.trim()));

      // Build column mapping
      const colMap = {};
      headers.forEach((h, i) => {
        const field = CSV_HEADER_MAP[h.trim()];
        if (field) colMap[i] = field;
      });

      if (Object.keys(colMap).length < 5) {
        throw new Error('CSVのヘッダーが正しくありません。英検スコアデータのCSVを使用してください。');
      }

      // Convert rows
      const newRecords = [];
      dataRows.forEach((row) => {
        const record = {};
        Object.entries(colMap).forEach(([ci, field]) => {
          let val = row[parseInt(ci)] || '';
          val = val.trim();
          if (val === '' || val === '－' || val === '-' || val === '*') {
            record[field] = null;
          } else if (NUMERIC_FIELDS.has(field)) {
            const num = Number(val);
            record[field] = isNaN(num) ? val : num;
          } else {
            record[field] = val;
          }
        });

        // Compute derived fields
        computeDerived(record);
        newRecords.push(record);
      });

      // Check duplicates
      const existingKeys = new Set(allData.map((d) => makeKey(d)));
      const unique = newRecords.filter((r) => !existingKeys.has(makeKey(r)));
      const dupeCount = newRecords.length - unique.length;

      if (unique.length === 0) {
        showToast('すべてのデータが既に存在しています（' + dupeCount + '件の重複）', 'error');
        importProgress.style.display = 'none';
        return;
      }

      progressText.textContent = unique.length + '件の新規データをインポート中...';

      // Import to Firestore
      await importBatch(unique);

      importProgress.style.display = 'none';
      csvDropOverlay.classList.remove('active');

      let msg = unique.length + '件のデータをインポートしました！';
      if (dupeCount > 0) msg += '（' + dupeCount + '件の重複はスキップ）';
      showToast(msg, 'success');

      // Reload data
      await loadData();

    } catch (err) {
      console.error('CSV import error:', err);
      showToast('インポートエラー: ' + err.message, 'error');
      importProgress.style.display = 'none';
    }
  }

  function makeKey(d) {
    return [d.name, d.grade, d.session, d.year].join('|');
  }

  function computeDerived(r) {
    const g = r.grade || '';

    // readingTotal
    const rqs = [r.readingQ1, r.readingQ2, r.readingQ3, r.readingQ4].filter((v) => typeof v === 'number');
    if (rqs.length > 0 && r.exemption !== '有') {
      r.readingTotal = rqs.reduce((a, b) => a + b, 0);
    }

    // readingRate
    if (typeof r.readingTotal === 'number') {
      const denoms = { '5級': 25, '4級': 35, '3級': 30, '準2級': 29, '準2級+': 29, '2級': 31 };
      const d = denoms[g] || 30;
      r.readingRate = Math.round((r.readingTotal / d) * 10000) / 10000;
    }

    // listeningTotal
    const lqs = [r.listeningQ1, r.listeningQ2, r.listeningQ3].filter((v) => typeof v === 'number');
    if (lqs.length > 0 && r.exemption !== '有') {
      r.listeningTotal = lqs.reduce((a, b) => a + b, 0);
    }

    // listeningRate
    if (typeof r.listeningTotal === 'number') {
      const d = g === '5級' ? 25 : 30;
      r.listeningRate = Math.round((r.listeningTotal / d) * 10000) / 10000;
    }

    // writingScore
    const wfs = [r.writingS1Content, r.writingS1Structure, r.writingS1Vocab, r.writingS1Grammar,
      r.writingS2Content, r.writingS2Structure, r.writingS2Vocab, r.writingS2Grammar].filter((v) => typeof v === 'number');
    if (wfs.length > 0 && g !== '4級' && g !== '5級') {
      r.writingScore = wfs.reduce((a, b) => a + b, 0);
    }

    // speakingScore
    const sfs = [r.speakingReading, r.speakingQA, r.speakingAttitude].filter((v) => typeof v === 'number');
    if (sfs.length > 0) {
      r.speakingScore = sfs.reduce((a, b) => a + b, 0);
    }
  }

  function readFileAsText(file, encoding) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsText(file, encoding);
    });
  }

  function parseCSV(text) {
    const lines = text.split(/\r?\n/);
    return lines.map((line) => {
      const result = [];
      let current = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (inQuotes) {
          if (ch === '"' && line[i + 1] === '"') {
            current += '"';
            i++;
          } else if (ch === '"') {
            inQuotes = false;
          } else {
            current += ch;
          }
        } else {
          if (ch === '"') {
            inQuotes = true;
          } else if (ch === ',') {
            result.push(current);
            current = '';
          } else {
            current += ch;
          }
        }
      }
      result.push(current);
      return result;
    });
  }

  // ============================================
  // CSV Export
  // ============================================

  exportBtn.addEventListener('click', () => {
    const headers = [
      '実施年度', '実施回', '級', '学年', '氏名', '一次免除の有無',
      '一次合否', '二次合否', '総合合否', '英検バンド一次', '英検バンド二次',
      'Reading CSEスコア', 'Listening CSEスコア', 'Writing CSEスコア',
      'CSEスコア一次合計', 'Speaking CSEスコア', '総合CSEスコア',
      'リーディング正答数', 'リーディング正答率', 'リスニング正答数', 'リスニング正答率',
      'ライティングスコア', 'スピーキングスコア'
    ];
    const fields = [
      'year', 'session', 'grade', 'schoolYear', 'name', 'exemption',
      'primaryResult', 'secondaryResult', 'overallResult', 'bandPrimary', 'bandSecondary',
      'readingCSE', 'listeningCSE', 'writingCSE', 'primaryCSETotal', 'speakingCSE', 'totalCSE',
      'readingTotal', 'readingRate', 'listeningTotal', 'listeningRate',
      'writingScore', 'speakingScore'
    ];

    const csvRows = [headers.join(',')];
    filteredData.forEach((d) => {
      const row = fields.map((f) => {
        let v = d[f];
        if (v === null || v === undefined) return '';
        if (f === 'readingRate' || f === 'listeningRate') v = Math.round(v * 100) + '%';
        v = String(v);
        if (v.includes(',') || v.includes('"') || v.includes('\n')) {
          v = '"' + v.replace(/"/g, '""') + '"';
        }
        return v;
      });
      csvRows.push(row.join(','));
    });

    const blob = new Blob(['\uFEFF' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '英検スコアデータ_' + new Date().toISOString().slice(0, 10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
    showToast('CSVをエクスポートしました', 'success');
  });

  // ============================================
  // Toast
  // ============================================

  function showToast(msg, type) {
    toastEl.textContent = msg;
    toastEl.className = 'toast ' + type + ' show';
    setTimeout(() => { toastEl.classList.remove('show'); }, 4000);
  }

  // ============================================
  // Global drag-and-drop (anywhere on page)
  // ============================================

  document.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (appContainer.style.display !== 'none') {
      csvDropOverlay.classList.add('active');
      csvDropZone.classList.add('drag-over');
    }
  });

  document.addEventListener('dragleave', (e) => {
    if (e.relatedTarget === null) {
      csvDropZone.classList.remove('drag-over');
    }
  });

  document.addEventListener('drop', (e) => {
    e.preventDefault();
    csvDropZone.classList.remove('drag-over');
  });

})();
